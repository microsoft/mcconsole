#!/usr/bin/env python
import tkinter as tk
from tkinter import ttk
import csv
import threading
import os
import queue
import sys
import argparse
import msvcrt
from pathlib import Path
import time
import io
import subprocess
import pyautogui  # Used to send keystrokes in Controller mode
import json       # For fallback configuration reading
import re
import random
import socket  # For reverse DNS fallback
import psutil
import win32process 
import win32gui
import win32con
import pygetwindow as gw
import glob
import tempfile
from datetime import datetime


def mask_passwords_in_command(cmd_str):
    """
    Mask passwords in commands for logging and LLM analysis.
    """
    if not cmd_str:
        return cmd_str
    
    # Common password patterns to mask
    patterns = [
        # curl -u username:password
        (r'(-u\s+\w+:)([^\s]+)', r'\1****'),
        # curl --user username:password  
        (r'(--user\s+\w+:)([^\s]+)', r'\1****'),
        # SSH password prompts (if any)
        (r'(password[:\s=]+)([^\s]+)', r'\1****', re.IGNORECASE),
        # Generic PASSWORD= patterns
        (r'(PASSWORD\s*=\s*)([^\s]+)', r'\1****', re.IGNORECASE),
        # Generic PASS= patterns
        (r'(PASS\s*=\s*)([^\s]+)', r'\1****', re.IGNORECASE),
        # Authorization headers
        (r'(Authorization:\s*Basic\s+)([^\s]+)', r'\1****', re.IGNORECASE),
    ]
    
    masked_cmd = cmd_str
    for pattern in patterns:
        if len(pattern) == 2:
            regex, replacement = pattern
            flags = 0
        else:
            regex, replacement, flags = pattern
        
        masked_cmd = re.sub(regex, replacement, masked_cmd, flags=flags)
    
    return masked_cmd

def sanitize_log_content_for_llm(log_content):
    """
    Sanitize log content before sending to LLM service.
    """
    if not log_content:
        return log_content
    
    lines = log_content.split('\n')
    sanitized_lines = []
    
    for line in lines:
        sanitized_line = line
        
        # Patterns for various command formats that might contain passwords
        patterns = [
            # curl commands with -u username:password
            (r'(curl.*?-u\s+\w+:)([^\s]+)', r'\1****'),
            # curl commands with --user username:password
            (r'(curl.*?--user\s+\w+:)([^\s]+)', r'\1****'),
            # RScmCli# prompts with curl commands containing passwords
            (r'(RScmCli#\s+curl.*?-u\s+\w+:)([^\s]+)', r'\1****'),
            # Linux prompts with curl commands
            (r'(root@.*?#\s+curl.*?-u\s+\w+:)([^\s]+)', r'\1****'),
            # Any generic password patterns
            (r'(\s+password[:\s=]+)([^\s]+)', r'\1****', re.IGNORECASE),
        ]
        
        for pattern in patterns:
            if len(pattern) == 2:
                regex, replacement = pattern
                flags = 0
            else:
                regex, replacement, flags = pattern
                
            sanitized_line = re.sub(regex, replacement, sanitized_line, flags=flags)
        
        sanitized_lines.append(sanitized_line)
    
    return '\n'.join(sanitized_lines)

def setup_shell_logging_redirect(console_type, sut_ip):
    """Redirect all print() statements to log file while keeping console open."""
    
    # Create logs directory
    log_dir = Path(r"C:\mcqueen\logs\Mcconsole_log")
    log_dir.mkdir(parents=True, exist_ok=True)
    
    # Generate timestamp for log filename
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    log_file = log_dir / f"shell_mx_{console_type}_{sut_ip}_{timestamp}.log"
    
    # Show user where logs are going before redirecting
    print(f"🚀 Shell MX ({console_type}) Starting for {sut_ip}...")
    print(f"📁 Debug logs will be saved to:")
    print(f"   {log_dir}")
    print(f"📄 Current session log: shell_mx_{console_type}_{sut_ip}_{timestamp}.log")
    print(f"📝 Full log path: {log_file}")
    print(f"\n>>> All further debug output will be redirected to the log file <<<\n")
    
    # Create a custom stdout/stderr that writes to log file only
    class LogRedirect:
        def __init__(self, log_file_path):
            self.log_file = open(log_file_path, 'w', encoding='utf-8', buffering=1)
            
        def write(self, text):
            # Write to log file with timestamp
            if text.strip():  # Only log non-empty lines
                self.log_file.write(f"[{datetime.now().strftime('%H:%M:%S')}] {text}")
            else:
                self.log_file.write(text)
            self.log_file.flush()
                
        def flush(self):
            self.log_file.flush()
    
    # Create redirector
    redirector = LogRedirect(log_file)
    
    # Redirect stdout and stderr to log file
    sys.stdout = redirector
    sys.stderr = redirector
    
    return redirector, log_file

# ─── optional pygame import/init ────────────────────────────────────────
try:
    import pygame
    try:
        pygame.mixer.init(frequency=44100, size=-16, channels=2, buffer=512)
        pygame_available = True
    except Exception as e:
        # mixer failed (no device, driver issue, etc.)
        print(f"[WARN] pygame.mixer init failed, sound disabled: {e}", flush=True)
        pygame_available = False
except ImportError:
    print("[WARN] pygame not installed, sound disabled", flush=True)
    pygame_available = False


#fixing the multi screen black error
#v17: switching from mobaxterm connect by title to also find pid. but some bug, also not seen log in cmd console why moba not opened. 
#v21: worked well with moba
#v22: try to support controller and local console with moba too.
#v24: put moba to 2/3 screen and helper to 1/3. 
#v26: fix pygame play sound, continue if not playble.


# ---------------- Global Variables ----------------
DEBUG = True
highlight_keywords = {}
error_match_text = None      # Widget to display status output
lock = threading.Lock()
helper_window_count = 0
log_file_path = None         # Will store the unique Mobaxterm log file path
last_command = ""            # Last command executed
log_folder = None            # For controller mode log folder
mobaxterm_pid = None
original_commands = []       # Store CSV commands
search_entry = None          # Global reference to search entry widget
# --------------------------------------------------
# --------------------------------------------------


def debug_print(msg):
    if DEBUG:
        msg_str = str(msg)
        
        # Skip excessive window scanning logs for privacy
        if "window: title=" in msg_str and "pid=" in msg_str:
            return  # Don't log window titles that may contain sensitive info
        
        # Skip watch_mobaxterm scanning messages (too verbose)
        if "watch_mobaxterm: scanning" in msg_str:
            return
        
        # Sanitize dispatch messages
        if "Dispatching:" in msg_str:
            parts = msg_str.split("Dispatching: ", 1)
            if len(parts) == 2:
                prefix = parts[0] + "Dispatching: "
                command = parts[1]
                sanitized_command = mask_passwords_in_command(command)
                sanitized_msg = prefix + sanitized_command
            else:
                sanitized_msg = msg_str
        else:
            sanitized_msg = msg_str
        
        print("[DEBUG]", sanitized_msg, flush=True)
        

def clean_output(text):
    """Remove ANSI escape sequences and extra empty lines."""
    text = ansi_escape.sub('', text)
    lines = text.splitlines()
    non_empty_lines = [line for line in lines if line.strip()]
    return "\n".join(non_empty_lines)

def strip_control_chars(s):
    """Remove non-printable control characters except newline and carriage return."""
    return ''.join(ch for ch in s if ch in '\r\n' or 32 <= ord(ch) <= 126)

# Reassign sys.stdout to use UTF-8 with error replacement.
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")

# Add common folder to PYTHONPATH for configuration reading.
sys.path.append(str(Path(__file__).parent.absolute() / "../common"))
from lib_wcs_rm import read_config_m


def parse_args():
    parser = argparse.ArgumentParser(
        description="Interactive SSH or Local Console with dynamic helper window"
    )
    parser.add_argument("--console_type", required=True,
                        help="Console type (e.g., RM, SUT, SOC, Controller)")
    parser.add_argument("--sut_ip", required=True, help="SUT IP for loading configuration")
    parser.add_argument("--local", action="store_true",
                        help="If set for Controller, use a local shell instead of SSH")
    parser.add_argument("--sound", type=str, default="false",
                        help="Enable sound when executing commands (true/false)")
    parser.add_argument("--tab_id", type=int, default=1,
                        help="Optional tab identifier for helper window positioning")
    parser.add_argument("--monitor_id", type=int, default=0,
                        help="Optional monitor identifier (0 for primary monitor)")
    parser.add_argument("--theme", type=str, default=None,
                        help="Theme name for sounds and appearance")
    parser.add_argument("-j", "--jumpbox", type=str, default=None,
                        help="Jumpbox server in format 'user@host:port' or 'user@host' (port defaults to 22)")
    parser.add_argument("--terminal", type=str, default="mobaxterm",
                        choices=["mobaxterm", "securecrt"],
                        help="Terminal emulator to use (default: mobaxterm)")
    return parser.parse_args()



def get_helper_window_counter():
    counter_path = Path(__file__).parent / "helper_window_counter.txt"
    try:
        with open(counter_path, "r+") as f:
            try:
                value = int(f.read().strip())
            except ValueError:
                value = 0
            new_value = value + 1
            f.seek(0)
            f.write(str(new_value))
            f.truncate()
            return new_value
    except FileNotFoundError:
        with open(counter_path, "w") as f:
            f.write("1")
        return 1

def play_sound_pygame(sound_file):
    """
    Load & play any file pygame supports (MP3, WAV, OGG…).
    This uses the single 'music' channel: load() replaces any currently playing track.
    """
    if not pygame_available:
        return

    try:
        pygame.mixer.music.load(sound_file)
        pygame.mixer.music.play()   # non-blocking
    except Exception as e:
        if DEBUG:
            print(f"[ERROR] pygame failed to play {sound_file}: {e}", flush=True)


def play_random_sound(theme_name=None):
    """
    Pick a random .mp3/.wav from the theme directory (if given),
    otherwise from the default sound folder, and play it via pygame.
    """
    # 1) Try theme folder
    candidates = []
    if theme_name:
        theme_dir = os.path.join(os.path.dirname(__file__),
                                 "..", "Theme", theme_name, "Sound")
        if os.path.isdir(theme_dir):
            for fn in os.listdir(theme_dir):
                if fn.lower().endswith((".mp3", ".wav")):
                    candidates.append(os.path.join(theme_dir, fn))

    # 2) Fallback default sound folder
    if not candidates:
        default_dir = os.path.join(os.path.dirname(__file__), "..", "sound")
        if os.path.isdir(default_dir):
            for fn in os.listdir(default_dir):
                if fn.lower().endswith((".mp3", ".wav")):
                    candidates.append(os.path.join(default_dir, fn))

    # 3) No files → bail
    if not candidates:
        if DEBUG:
            print("[WARN] No sound files found; skipping play_random_sound", flush=True)
        return

    # 4) If pygame not ready → bail
    if not pygame_available:
        if DEBUG:
            print("[WARN] play_random_sound skipped: pygame unavailable", flush=True)
        return

    # 5) Pick one and play on a daemon thread
    sound_file = random.choice(candidates)
    if DEBUG:
        print(f"[INFO] Playing via pygame: {sound_file}", flush=True)
    threading.Thread(
        target=lambda: play_sound_pygame(sound_file),
        daemon=True
    ).start()


def refresh_command_tree(commands):
    """
    Refresh the command tree with new commands.
    """
    global tree
    
    if not tree:
        return
        
    # Schedule UI update on main thread
    tree.after(0, lambda: _refresh_tree_ui(commands))

def _refresh_tree_ui(commands):
    """
    Helper function to refresh tree on main thread.
    """
    global tree
    
    if not tree:
        return
    
    # Clear existing items
    tree.delete(*tree.get_children())
    
    # Group commands by category for agent mode
    if agent_mode_active and dynamic_commands:
        # Separate original and dynamic commands
        debug_commands = [cmd for cmd in dynamic_commands if "[DEBUG]" in cmd[0]]
        test_commands = [cmd for cmd in dynamic_commands if "[TEST]" in cmd[0]]
        setup_commands = [cmd for cmd in dynamic_commands if "[SETUP]" in cmd[0]]
        
        # Add category headers and commands
        if debug_commands:
            tree.insert("", tk.END, values=("🔍 DEBUG COMMANDS", "", "", ""), tags=("category",))
            for desc, cmd, kw, hover in debug_commands:
                tags = ("hoverable",) if hover else ()
                tree.insert("", tk.END, values=(desc, cmd, kw, hover), tags=tags)
        
        if test_commands:
            tree.insert("", tk.END, values=("🧪 TEST COMMANDS", "", "", ""), tags=("category",))
            for desc, cmd, kw, hover in test_commands:
                tags = ("hoverable",) if hover else ()
                tree.insert("", tk.END, values=(desc, cmd, kw, hover), tags=tags)
        
        if setup_commands:
            tree.insert("", tk.END, values=("⚙️ SETUP COMMANDS", "", "", ""), tags=("category",))
            for desc, cmd, kw, hover in setup_commands:
                tags = ("hoverable",) if hover else ()
                tree.insert("", tk.END, values=(desc, cmd, kw, hover), tags=tags)
        
        # Add separator and original commands
        if original_commands:
            tree.insert("", tk.END, values=("📋 ORIGINAL COMMANDS", "", "", ""), tags=("category",))
            for desc, cmd, kw, hover in original_commands:
                tags = ("hoverable",) if hover else ()
                tree.insert("", tk.END, values=(desc, cmd, kw, hover), tags=tags)
    else:
        # Standard mode - just load commands as before
        for desc, cmd, kw, hover in commands:
            tags = ("hoverable",) if hover else ()
            tree.insert("", tk.END, values=(desc, cmd, kw, hover), tags=tags)

def simulate_typing_command_terminal(terminal_type, command, host):
    """Simulate typing a command into the terminal window for the given host."""
    try:
        import pygetwindow as gw
    except ImportError:
        print("pygetwindow module not available; cannot simulate typing.")
        return False

    keywords = {"mobaxterm": ["mobaxterm"], "securecrt": ["securecrt", "vshell"]}
    kws = keywords.get(terminal_type.lower(), ["mobaxterm"])

    def find_window_by_substring(sub):
        sub = sub.lower()
        for w in gw.getAllWindows():
            title_lower = w.title.lower()
            if any(kw in title_lower for kw in kws) and sub in title_lower and w.visible:
                return w
        return None

    w = find_window_by_substring(host)
    if w:
        return _activate_and_type(w, command)

    try:
        name, _, _ = socket.gethostbyaddr(host)
        short_name = name.split('.', 1)[0]
        for candidate in (short_name, name):
            w = find_window_by_substring(candidate)
            if w:
                return _activate_and_type(w, command)
    except Exception as e:
        if DEBUG: print("Reverse DNS fallback error:", e)

    global log_file_path
    if log_file_path and os.path.exists(log_file_path):
        try:
            with open(log_file_path, "r", encoding="utf-8", errors="replace") as f:
                content = f.read()
            match = re.search(r"(root@\S+)", content)
            if match:
                alt_host = match.group(1).split('@', 1)[-1]
                w = find_window_by_substring(alt_host)
                if w:
                    return _activate_and_type(w, command)
        except Exception as e:
            if DEBUG: print("Error searching log file for remote hostname:", e)
    print("No Mobaxterm window found for host or extracted hostname:", host)
    return False

def _activate_and_type(window, command):
    """Activate the window and simulate typing the command."""
    try:
        window.activate()
        time.sleep(0.5)  # Allow the window to gain focus
        # Optionally, press CTRL+ALT+F1 to switch to the first tab
        pyautogui.hotkey('ctrl', 'alt', 'f1')
        time.sleep(0.5)  # Allow the tab switch to complete
        pyautogui.typewrite(command + "\n", interval=0)
        info_print("Typed command into window (" + window.title + "): " + command)
        return True
    except Exception as e:
        print("Error typing command in window:", e)
        return False

def info_print(msg):
    if DEBUG:
        # Don't log sensitive window operations
        if "Typed command into window" in str(msg):
            # Just log that a command was sent, not the details
            print("[INFO] Command sent to console", flush=True)
        else:
            print("[INFO]", str(msg), flush=True)

def test_cipher(host, cipher):
    """
    Try connecting with a given cipher. Uses a simple SSH command that exits immediately.
    Adds -o StrictHostKeyChecking=no to avoid interactive prompts.
    
    Returns True if the cipher negotiation was successful.
    (A return containing "Permission denied" is considered success, 
     while a return containing "Unable to negotiate" means the candidate failed.)
    """
    strict_option = "-o StrictHostKeyChecking=no"
    if cipher:
        cmd = f'ssh -o BatchMode=yes -o ConnectTimeout=5 {strict_option} -c {cipher} root@{host} exit'
    else:
        cmd = f'ssh -o BatchMode=yes -o ConnectTimeout=5 {strict_option} root@{host} exit'
    debug_print(f"Running test command: {cmd}")
    result = subprocess.run(cmd, shell=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE, universal_newlines=True)
    stderr = result.stderr.strip()
    debug_print(f"SSH test output stderr: {stderr}")
    if "Unable to negotiate" in stderr:
        return False
    if "Permission denied" in stderr or result.returncode != 0:
        return True
    return result.returncode == 0

def auto_select_cipher(host, candidate_ciphers):
    """
    Given a list of candidate ciphers, test each until one works.
    Returns the first working cipher (or None if none work).
    """
    for cipher in candidate_ciphers:
        debug_print(f"Testing cipher: {cipher}")
        if test_cipher(host, cipher):
            debug_print(f"Using cipher: {cipher}")
            return cipher
    debug_print("No candidate cipher worked.")
    return None


def get_monitor_count():
    """Get the total number of monitors using multiple methods for reliability."""
    try:
        from screeninfo import get_monitors
        monitors = get_monitors()
        if monitors:
            return len(monitors)
    except Exception:
        pass
    
    try:
        import pygetwindow as gw
        screens = gw._pygetwindow._getAllScreens()
        if screens:
            return len(screens)
    except Exception:
        pass
    
    try:
        import win32api
        monitors = win32api.EnumDisplayMonitors()
        if monitors:
            return len(monitors)
    except Exception:
        pass
    
    # If all methods fail, assume single monitor
    return 1

# --- New helper functions for screen/monitor geometry and window positioning --- #

def get_monitor_geometry(monitor_id=0):
    """
    Returns the (x, y, width, height) for the specified monitor.
    monitor_id: 0 for primary display, other values for additional monitors (zero-based).
    Tries multiple methods to ensure reliable multi-monitor detection.
    """
    try:
        # Method 1: Use screeninfo library if available
        from screeninfo import get_monitors
        monitors = get_monitors()
        if monitors and len(monitors) > monitor_id:
            m = monitors[monitor_id]
            debug_print(f"Using screeninfo library: Monitor {monitor_id} found at {m.x},{m.y} with size {m.width}x{m.height}")
            return m.x, m.y, m.width, m.height
        else:
            debug_print(f"screeninfo reports only {len(monitors)} monitors, but requested monitor_id {monitor_id}")
    except Exception as ex:
        debug_print(f"screeninfo method failed: {ex}")
    
    try:
        # Method 2: Use pygetwindow to get screen geometry
        import pygetwindow as gw
        # Get all screens
        screens = gw._pygetwindow._getAllScreens()
        if screens and len(screens) > monitor_id:
            screen = screens[monitor_id]
            x, y, width, height = screen  # Unpack screen geometry
            debug_print(f"Using pygetwindow: Monitor {monitor_id} found at {x},{y} with size {width}x{height}")
            return x, y, width, height
        else:
            debug_print(f"pygetwindow reports only {len(screens)} monitors, but requested monitor_id {monitor_id}")
    except Exception as ex:
        debug_print(f"pygetwindow method failed: {ex}")
    
    try:
        # Method 3: Use win32api as a fallback for Windows
        import win32api
        monitors = win32api.EnumDisplayMonitors()
        if monitors and len(monitors) > monitor_id:
            monitor = monitors[monitor_id]
            info = win32api.GetMonitorInfo(monitor[0])
            # Get the working area (excludes taskbar)
            work_area = info['Work']
            x, y, right, bottom = work_area
            width = right - x
            height = bottom - y
            debug_print(f"Using win32api: Monitor {monitor_id} found at {x},{y} with size {width}x{height}")
            return x, y, width, height
        else:
            debug_print(f"win32api reports only {len(monitors)} monitors, but requested monitor_id {monitor_id}")
    except Exception as ex:
        debug_print(f"win32api method failed: {ex}")
    
    # Fallback to a default value if all methods fail
    default_values = (0, 0, 800, 600)
    debug_print(f"All monitor detection methods failed. Using default values: {default_values}")
    return default_values



def launch_terminal_and_get_pid(cmdline, terminal_type="mobaxterm", timeout=10):
    """Launch terminal via cmdline, then inspect cmd.exe's children for the terminal process."""
    process_names = {"mobaxterm": "mobaxterm.exe", "securecrt": "securecrt.exe"}
    target_process = process_names.get(terminal_type.lower(), "mobaxterm.exe")
    debug_print(f"Looking for process: {target_process}")
    proc = subprocess.Popen(cmdline, shell=True)
    start = time.time()
    while time.time() - start < timeout:
        try:
            parent = psutil.Process(proc.pid)
            for c in parent.children(recursive=True):
                if c.name().lower() == target_process:
                    debug_print(f"Found {terminal_type} PID: {c.pid}")
                    return c.pid
        except psutil.NoSuchProcess:
            break
        time.sleep(0.2)
    debug_print(f"WARNING: {terminal_type} not found in {timeout}s")
    return None

def position_window_by_pid(pid, x, y, w, h, retries=5, delay=0.5):
    """Find the window whose process == pid, move & resize it."""
    for _ in range(retries):
        for win in gw.getAllWindows():
            try:
                _, win_pid = win32process.GetWindowThreadProcessId(win._hWnd)
            except Exception:
                continue
            if win_pid == pid and win.visible:
                # Calculate 2/3 width and height minus 2 lines (assuming ~20px per line)
                width_two_thirds = int(w * 2/3)
                height_minus_lines = h - 40  # Assuming each line is about 20px
                win.moveTo(x, y)
                win.resizeTo(width_two_thirds, height_minus_lines)
                return True
        time.sleep(delay)
    return False

def position_window_by_title(terminal_type, host, x, y, w, h, retries=5, delay=0.5):
    """Fallback: look for terminal window with host substring in title."""
    keywords = {"mobaxterm": ["mobaxterm"], "securecrt": ["securecrt", "vshell"]}
    kws = keywords.get(terminal_type.lower(), ["mobaxterm"])
    target = host.lower()
    for _ in range(retries):
        for win in gw.getAllWindows():
            title_lower = win.title.lower()
            if any(kw in title_lower for kw in kws) and target in title_lower and win.visible:
                width_two_thirds = int(w * 2/3)
                height_minus_lines = h - 40
                win.moveTo(x, y)
                win.resizeTo(width_two_thirds, height_minus_lines)
                return True
        time.sleep(delay)
    return False

def open_helper_window(command_queue, console_type, config, sut_ip, sound_enabled=False, monitor_id=0, theme_name=None):
    """
    Pop up the Tk helper UI, with:
      • "Focus Console" button to raise the paired MobaXterm
      • Search/filter box to quickly find commands
      • Status area for command execution feedback
      • Closing either window tears the other down
    """
    global error_match_text, mobaxterm_pid
    global helper_sound_enabled, helper_theme_name
    global tree, original_commands, search_entry

    # Store sound settings globally
    helper_sound_enabled = sound_enabled
    helper_theme_name = theme_name

    # --- Windows DPI awareness ---
    import ctypes
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass

    # --- Calculate helper-geometry: right 1/3 of chosen monitor ---
    monitor_x, monitor_y, monitor_w, monitor_h = get_monitor_geometry(monitor_id)
    width_one_third = monitor_w // 3
    helper_geometry_str = f"{width_one_third}x{monitor_h}+{monitor_x + monitor_w - width_one_third}+{monitor_y}"

    window = tk.Tk()
    window.title(f"{console_type}-{sut_ip} McQueen Console")
    window.geometry(helper_geometry_str)
    window.tk.call('tk', 'scaling', 1.7)
    window.configure(bg='#1E1E1E')  # Set default background color

    # --- Font zoom support --- (moved up to define current_font_size early)
    import tkinter.font as tkFont
    current_font_size = 10

    def update_fonts():
        f = tkFont.nametofont("TkDefaultFont")
        f.configure(size=current_font_size)
        if error_match_text:
            error_match_text.config(font=("Consolas", current_font_size))

    def zoom_in(event=None):
        nonlocal current_font_size
        current_font_size += 1
        update_fonts()

    def zoom_out(event=None):
        nonlocal current_font_size
        if current_font_size > 1:
            current_font_size -= 1
            update_fonts()

    window.bind("<Control-MouseWheel>", lambda e: zoom_in() if e.delta > 0 else zoom_out())
    window.bind("<Control-plus>", zoom_in)
    window.bind("<Control-minus>", zoom_out)

    # --- Toolbar + Focus button ---
    toolbar = tk.Frame(window, bg='#1E1E1E')
    toolbar.pack(fill=tk.X, padx=10, pady=(5,0))

    # --- Apply theme background to toolbar if available ---
    toolbar_bg_image = None
    if theme_name:
        try:
            from PIL import Image, ImageTk
            theme_dir = os.path.join(os.path.dirname(__file__), "..", "Theme", theme_name, "Picture")
            if os.path.isdir(theme_dir):
                image_files = []
                for ext in ["*.jpg", "*.png", "*.gif"]:
                    image_files.extend(glob.glob(os.path.join(theme_dir, ext)))
                
                if image_files:
                    # Choose a random background image
                    bg_image_path = random.choice(image_files)
                    bg_image = Image.open(bg_image_path)
                    
                    # Get toolbar dimensions
                    toolbar_width = width_one_third - 20  # Adjust for padding
                    toolbar_height = 40  # Estimated toolbar height
                    
                    # Resize the image to fit the toolbar
                    img_width, img_height = bg_image.size
                    
                    # Take a crop of the image (top portion)
                    crop_top = 0
                    crop_height = min(img_height, int(img_height * 0.3))  # Top 30% of image
                    cropped_bg = bg_image.crop((0, crop_top, img_width, crop_height))
                    
                    # Resize to fit toolbar
                    resized_bg = cropped_bg.resize((toolbar_width, toolbar_height), Image.LANCZOS)
                    
                    # Apply a darkening overlay for better text visibility
                    if resized_bg.mode != 'RGBA':
                        resized_bg = resized_bg.convert('RGBA')
                    
                    # Create a new image with RGBA mode and fill with semi-transparent black
                    overlay = Image.new('RGBA', resized_bg.size, (0, 0, 0, 100))  # Semi-transparent black
                    
                    # Alpha composite the images
                    darkened_bg = Image.alpha_composite(resized_bg, overlay)
                    
                    # Convert back to RGB mode for PhotoImage
                    if darkened_bg.mode == 'RGBA':
                        rgb_bg = Image.new('RGB', darkened_bg.size, (0, 0, 0))
                        rgb_bg.paste(darkened_bg, mask=darkened_bg.split()[3])  # Use alpha channel as mask
                        darkened_bg = rgb_bg
                    
                    # Create PhotoImage
                    toolbar_bg_image = ImageTk.PhotoImage(darkened_bg)
                    
                    # Apply as background
                    bg_label = tk.Label(toolbar, image=toolbar_bg_image)
                    bg_label.place(x=0, y=0, relwidth=1, relheight=1)
                    
                    # Keep a reference to prevent garbage collection
                    toolbar.image = toolbar_bg_image
        except Exception as e:
            print(f"[ERROR] Failed to apply theme background: {e}", flush=True)

    def focus_console():
        """Bring the paired MobaXterm window to the front by PID."""
        if not mobaxterm_pid:
            return
        try:
            for w in gw.getAllWindows():
                _, win_pid = win32process.GetWindowThreadProcessId(w._hWnd)
                if win_pid == mobaxterm_pid and w.visible:
                    hwnd = w._hWnd
                    win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
                    win32gui.SetForegroundWindow(hwnd)
                    return
        except Exception as e:
            debug_print(f"focus_console error: {e}")

    
    btn_focus = tk.Button(
        toolbar,
        text="Focus Console",
        command=focus_console,
        bg="#005F9E",           # stronger blue
        fg="white",
        font=("Consolas", 11, "bold"),
        width=15                # make it wide
    )
    btn_focus.pack(side=tk.LEFT, padx=5, pady=5)

    # Agent Mode button removed - functionality is always active now

    # Optional: bind F1 so user can hit the key instead of clicking
    window.bind("<F1>", lambda e: focus_console())

    # --- Command parsing & execution callbacks ---
    def parse_command(cmd_str):
        # first, replace the special SUT_TITLE placeholder with the actual sut_ip
        cmd_str = cmd_str.replace('[[SUT_TITLE]]', sut_ip)
        # then do the rest of your normal config‐based substitutions
        for k, v in config.items():
            cmd_str = cmd_str.replace(f'[[{k}]]', str(v))
        return cmd_str

    def execute_command(cmd_str):
        """
        Execute command but show only descriptions in UI messages.
        """
        global last_command
        
        # Find the description for this command
        description = "Command Execution"  # Default fallback
        for desc, original_cmd, kw, hover in original_commands:
            if original_cmd.strip() == cmd_str.strip():
                description = desc
                break
        
        if error_match_text:
            # Show only the friendly description
            error_match_text.insert(tk.END, f"\n⚡ Executing: {description}\n")
            error_match_text.see(tk.END)
        
        for cmd in [c.strip() for c in cmd_str.split(";;") if c.strip()]:
            # Parse and execute the real command (with passwords)
            real_cmd = parse_command(cmd)
            last_command = real_cmd
            command_queue.put(real_cmd)

            if error_match_text:
                # Show generic message instead of actual command
                error_match_text.insert(tk.END, f"📤 Command sent to console\n")
                error_match_text.see(tk.END)

            if sound_enabled:
                play_random_sound(theme_name)

            time.sleep(0.2)

        if error_match_text:
            error_match_text.insert(tk.END, "✅ Command execution complete.\n")
            error_match_text.insert(tk.END, "-" * 50 + "\n")
            error_match_text.see(tk.END)

    def on_double_click(event):
        sel = tree.selection()
        if sel:
            vals = tree.item(sel[0], "values")
            execute_command(vals[1])

    def on_right_click(event):
        row = tree.identify_row(event.y)
        if not row:
            return
        
        # Select the row
        tree.selection_set(row)
        
        # Get values from the row
        values = tree.item(row, "values")
        if len(values) < 2 or not values[1].strip():
            # Skip if it's a category row without a command
            return
            
        desc, cmd = values[0], values[1]
        
        # Create popup window - much larger and resizable
        popup = tk.Toplevel(window)
        popup.title("Edit Command")
        popup.geometry("1200x500")  # Even larger initial size
        popup.minsize(900, 400)     # Set minimum size
        popup.focus_set()           # Make sure popup gets focus
        
        # Configure popup to be resizable
        popup.columnconfigure(0, weight=1)
        popup.rowconfigure(2, weight=1)  # Make text area row expandable
        
        # Font size control for the popup
        popup_font_size = [12]  # Use list so we can modify it in nested functions
        
        def update_popup_fonts():
            text_widget.config(font=("Consolas", popup_font_size[0]))
            label.config(font=("Arial", popup_font_size[0], "bold"))
            help_label.config(font=("Arial", max(8, popup_font_size[0] - 2)))
        
        def zoom_in_popup(event=None):
            popup_font_size[0] = min(popup_font_size[0] + 1, 24)  # Max size 24
            update_popup_fonts()
        
        def zoom_out_popup(event=None):
            popup_font_size[0] = max(popup_font_size[0] - 1, 8)   # Min size 8
            update_popup_fonts()
        
        # Bind font zoom controls to the popup
        popup.bind("<Control-MouseWheel>", lambda e: zoom_in_popup() if e.delta > 0 else zoom_out_popup())
        popup.bind("<Control-plus>", zoom_in_popup)
        popup.bind("<Control-minus>", zoom_out_popup)
        popup.bind("<Control-Key-0>", lambda e: setattr(popup_font_size, '__setitem__', (0, 12)) or update_popup_fonts())  # Reset to default
        
        # Add label
        label = tk.Label(popup, text=f"Edit Command: {desc}", font=("Arial", 12, "bold"))
        label.grid(row=0, column=0, sticky="ew", padx=10, pady=5)
        
        # Add font control instructions
        font_help = tk.Label(popup, text="Font size: Ctrl+Mouse Wheel, Ctrl +/-, Ctrl+0 to reset", 
                            font=("Arial", 10), fg="blue")
        font_help.grid(row=1, column=0, sticky="ew", padx=10, pady=2)
        
        # Create frame for text widget and scrollbars
        text_frame = tk.Frame(popup)
        text_frame.grid(row=2, column=0, sticky="nsew", padx=10, pady=5)
        text_frame.columnconfigure(0, weight=1)
        text_frame.rowconfigure(0, weight=1)
        
        # Use Text widget with NO word wrapping to keep commands on single line
        # Add horizontal scrollbar for long commands
        text_widget = tk.Text(text_frame, wrap="none", height=10, font=("Consolas", 12))
        text_widget.grid(row=0, column=0, sticky="nsew")
        
        # Add vertical scrollbar
        v_scrollbar = tk.Scrollbar(text_frame, orient="vertical", command=text_widget.yview)
        v_scrollbar.grid(row=0, column=1, sticky="ns")
        text_widget.configure(yscrollcommand=v_scrollbar.set)
        
        # Add horizontal scrollbar
        h_scrollbar = tk.Scrollbar(text_frame, orient="horizontal", command=text_widget.xview)
        h_scrollbar.grid(row=1, column=0, sticky="ew")
        text_widget.configure(xscrollcommand=h_scrollbar.set)
        
        # Insert the command and select all text
        text_widget.insert("1.0", cmd)
        text_widget.focus_set()  # Focus on the text field
        text_widget.tag_add("sel", "1.0", "end-1c")  # Select all text
        
        # Create button frame
        button_frame = tk.Frame(popup)
        button_frame.grid(row=3, column=0, sticky="ew", padx=10, pady=10)
        
        # Function to save and execute the command
        def save_and_exec():
            # Get all text and ensure it's a single line (remove any newlines)
            command_text = text_widget.get("1.0", "end-1c")
            # More aggressive cleaning: remove ALL whitespace duplicates and line breaks
            import re
            # Replace any sequence of whitespace (spaces, tabs, newlines) with single space
            command_text = re.sub(r'\s+', ' ', command_text).strip()
            execute_command(command_text)
            popup.destroy()
        
        def cancel_edit():
            popup.destroy()
        
        # Simple Enter to execute
        def handle_enter(event):
            save_and_exec()
            return "break"  # This prevents the default Enter behavior
        
        # Add buttons with larger fonts
        tk.Button(button_frame, text="Save and Execute", command=save_and_exec, 
                  bg="#28a745", fg="white", font=("Arial", 12, "bold")).pack(side=tk.LEFT, padx=5)
        tk.Button(button_frame, text="Cancel", command=cancel_edit, 
                  bg="#dc3545", fg="white", font=("Arial", 12)).pack(side=tk.LEFT, padx=5)
        
        # Bind keyboard shortcuts
        text_widget.bind("<Return>", handle_enter)     # Enter to execute
        popup.bind("<Escape>", lambda e: cancel_edit())  # Escape to cancel
        
        # Add helpful text at the bottom with larger font
        help_label = tk.Label(popup, text="Tip: Press Enter to execute, Escape to cancel, Ctrl+Mouse Wheel to zoom", 
                             font=("Arial", 10), fg="gray")
        help_label.grid(row=4, column=0, sticky="ew", padx=10, pady=2)

    # --- Create tooltip functionality for hoverable items ---
    class ToolTip:
        def __init__(self, widget):
            self.widget = widget
            self.tipwindow = None
            self.id = None
            self.x = self.y = 0

        def showtip(self, text):
            """Display text in tooltip window"""
            self.text = text
            if self.tipwindow or not self.text:
                return
            
            # Get mouse position for the tooltip
            x = self.widget.winfo_pointerx() + 15
            y = self.widget.winfo_pointery() + 10
            
            # Create tooltip window
            self.tipwindow = tw = tk.Toplevel(self.widget)
            tw.wm_overrideredirect(1)  # Remove window decorations
            tw.wm_geometry(f"+{x}+{y}")
            tw.attributes("-topmost", True)  # Keep tooltip on top
            
            # Create tooltip content
            label = tk.Label(tw, text=self.text, justify=tk.LEFT,
                            background="#ffffe0", relief=tk.SOLID, borderwidth=1,
                            font=("Consolas", 10), wraplength=400)  # Added wraplength for long text
            label.pack(ipadx=5, ipady=5)

        def hidetip(self):
            """Hide the tooltip"""
            tw = self.tipwindow
            self.tipwindow = None
            if tw:
                tw.destroy()

    # --- Commands Treeview ---
    frame = tk.Frame(window, bg='#1E1E1E')
    frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

    style = ttk.Style()
    style.theme_use("clam")
    style.configure("Treeview", background="white", foreground="black",
                    rowheight=25, fieldbackground="white", bordercolor="black",
                    borderwidth=1, relief="solid")
    style.configure("Treeview.Heading", background="lightgray", foreground="black",
                    bordercolor="black", borderwidth=1, relief="raised")
    style.map("Treeview",
              background=[('selected','blue')], foreground=[('selected','white')])
    
    # Add new style for category rows
    style.configure("Category.Treeview.Row", background="#e6f2ff", font=("Arial", 10, "bold"))

    # First, determine the CSV file path to read headers
    helper_folder = os.path.join(os.path.dirname(__file__), "helper")
    prefix = f"{config.get('SUT_SKU')}_" if config.get("SUT_SKU") else ""

    if console_type.lower() == "controller":
        csv_file = os.path.join(helper_folder, f"{prefix}CONTROLLER_commands.csv")
    else:
        csv_file = os.path.join(helper_folder, f"{prefix}{console_type}_commands.csv")

    # Read CSV headers first to set column titles
    column_titles = ["Description", "Command", "Keywords", "HoverText"]  # Default fallback
    if os.path.exists(csv_file):
        try:
            with open(csv_file, newline="", encoding="utf-8", errors="replace") as cf:
                reader = csv.reader(cf)
                header_row = next(reader, None)
                if header_row:
                    # Use CSV headers as column titles, with fallbacks
                    column_titles = []
                    for i, header in enumerate(header_row[:4]):  # Max 4 columns
                        if header.strip():
                            column_titles.append(header.strip())
                        else:
                            # Fallback names if header is empty
                            fallbacks = ["Description", "Command", "Keywords", "HoverText"]
                            column_titles.append(fallbacks[i] if i < len(fallbacks) else f"Column{i+1}")
                    
                    # Ensure we have at least 4 columns
                    while len(column_titles) < 4:
                        fallbacks = ["Description", "Command", "Keywords", "HoverText"]
                        column_titles.append(fallbacks[len(column_titles)])
        except Exception as e:
            print(f"Error reading CSV headers: {e}")

    # Create Treeview with dynamic headers
    tree = ttk.Treeview(frame, columns=("Col1","Col2","Col3","Col4"), show="headings")
    
    # Set dynamic column headers and widths
    tree.heading("Col1", text=column_titles[0]); tree.column("Col1", width=200)
    tree.heading("Col2", text=column_titles[1]); tree.column("Col2", width=350) 
    tree.heading("Col3", text=column_titles[2]); tree.column("Col3", width=200)
    tree.heading("Col4", text=column_titles[3]); tree.column("Col4", width=0, stretch=False)  # Hidden hover column
    
    tree.pack(fill=tk.BOTH, expand=True)
    tree.bind("<Double-1>", on_double_click)
    tree.bind("<Button-3>", on_right_click)

    # Load commands from CSV
    original_commands = []  # Store original commands globally

    # Check if this is an RM session by looking at the CSV filename
    is_rm_session = "RM" in os.path.basename(csv_file)

    print(f"Loading commands from CSV file: {csv_file}")
    if not os.path.exists(csv_file):
        tree.insert("", tk.END, values=("Error", f"CSV not found: {csv_file}", "", ""))
    else:
        try:
            with open(csv_file, newline="", encoding="utf-8", errors="replace") as cf:
                reader = csv.reader(cf)
                row_index = 0
                column_headers = None
                
                for row in reader:
                    # First row: Column headers (skip, already used for treeview headers)
                    if row_index == 0:
                        column_headers = row
                        row_index += 1
                        continue
                    
                    # Skip empty rows
                    if not any(cell.strip() for cell in row):
                        row_index += 1
                        continue
                    
                    # Check if it's a category title (just column A with content)
                    is_category = len(row) >= 1 and row[0].strip() and (len(row) == 1 or not row[1].strip())
                    
                    if is_category:
                        # Create a category header row
                        item_id = tree.insert("", tk.END, values=(row[0], "", "", ""), tags=("category",))
                        
                    elif len(row) >= 2:
                        desc = row[0]
                        cmd  = parse_command(row[1])
                        kw   = row[2].strip() if len(row)>=3 else ""
                        hover_text = row[3].strip() if len(row)>=4 else ""
                        
                        # Store all command details in original_commands
                        original_commands.append((desc, cmd, kw, hover_text))
                        
                        # Insert into tree with appropriate tag if hoverable
                        tags = ("hoverable",) if hover_text else ()
                        item_id = tree.insert("", tk.END, values=(desc, cmd, kw, hover_text), tags=tags)
                    
                    row_index += 1
                
                if not original_commands and not any(tree.exists(item) for item in tree.get_children()):
                    tree.insert("", tk.END, values=("No commands loaded","CSV empty","",""))
        except Exception as e:
            tree.insert("", tk.END, values=("Error loading CSV", str(e), "", ""))

    # Configure styling (add this after the CSV loading)
    tree.tag_configure("header_info", 
                      background="#4CAF50",  # Green background 
                      font=("Arial", 9, "bold"), 
                      foreground="white")    # White text
                      
    # Configure category row styling with padding for visual separation
    tree.tag_configure("category", 
                      background="#004080",  # Darker blue background for better visibility
                      font=("Arial", 10, "bold"), 
                      foreground="white")    # White text for contrast
                      
    # Configure hoverable row styling to indicate it has hover text - make it greener
    tree.tag_configure("hoverable", background="#d0f0d0")  # Soft green color

    # Set up hover tooltips for hoverable rows
    tooltip = ToolTip(tree)
    current_row = None
    
    def on_hover(event):
        nonlocal current_row
        # Get the item under mouse
        row_id = tree.identify_row(event.y)
        
        # Hide tooltip if mouse moved to a different row
        if row_id != current_row:
            tooltip.hidetip()
            current_row = row_id
            
        # If no row or moved to empty area
        if not row_id:
            return
            
        # Get the hover text from the hidden column
        values = tree.item(row_id, "values")
        if len(values) >= 4 and values[3]:
            tooltip.showtip(values[3])
        else:
            tooltip.hidetip()
        
    def on_leave(event):
        # Hide tooltip when mouse leaves the tree
        nonlocal current_row
        tooltip.hidetip()
        current_row = None
    
    # Bind mouse events to the tree
    tree.bind("<Motion>", on_hover)
    tree.bind("<Leave>", on_leave)

    # --- Search/Filter box ---
    search_frame = tk.Frame(frame, bg='#1E1E1E')
    search_frame.pack(fill=tk.X, pady=5)
    tk.Label(search_frame, text="Filter:", bg='#1E1E1E', fg='#E0E0E0',
             font=("Arial", 10, "bold")).pack(side=tk.LEFT, padx=5)
    search_entry = tk.Entry(search_frame, bg='#2A2A2A', fg='#E0E0E0', insertbackground='#E0E0E0')
    search_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)

    def filter_commands(event=None):
        q = search_entry.get().lower()
        if not q:
            # If empty, show all commands
            refresh_command_tree(original_commands)
            return

        tree.delete(*tree.get_children())

        # Keep track of shown categories
        shown_categories = set()
        all_commands = original_commands
        
        # First pass: identify which categories have matches
        for desc, cmd, kw, hover in all_commands:
            if q in desc.lower() or q in cmd.lower() or q in kw.lower():
                # This is a potential match, check if it belongs to a category
                # For simplicity, we'll assume the category is everything before the first ':' in the description
                category = desc.split(':', 1)[0] if ':' in desc else ""
                if category:
                    shown_categories.add(category)
        
        # Second pass: insert category headers and matching items
        current_category = None
        
        for desc, cmd, kw, hover in all_commands:
            # Check if this is a match for the search
            is_match = q in desc.lower() or q in cmd.lower() or q in kw.lower()
            
            # Determine category
            category = desc.split(':', 1)[0] if ':' in desc else ""
            
            # If this item belongs to a category that's not yet inserted and should be shown
            if category and category != current_category and category in shown_categories:
                # Insert the category header
                item_id = tree.insert("", tk.END, values=(category, "", "", ""), tags=("category",))
                current_category = category
            
            # If this item matches the search, insert it
            if is_match:
                tags = ("hoverable",) if hover else ()
                tree.insert("", tk.END, values=(desc, cmd, kw, hover), tags=tags)

    search_entry.bind("<KeyRelease>", filter_commands)

    # --- Auto‐execute first command after 5s ---
    def auto_execute_if_appropriate():
        """
        Modified to look at the SECOND row (index 1) instead of first row for auto-execution.
        First row is now reserved for column headers.
        """
        # Check specifically for B2 (second row, column B) from the CSV
        try:
            with open(csv_file, newline="", encoding="utf-8", errors="replace") as cf:
                reader = csv.reader(cf)
                rows = list(reader)
                
                # Skip first row (headers) and check second row
                if len(rows) >= 2:
                    second_row = rows[1]  # Index 1 = second row
                    # Check if B2 exists and is not empty
                    if len(second_row) > 1 and second_row[1].strip():
                        # Use the second row's B cell (properly parsed)
                        b2_command = parse_command(second_row[1].strip())
                        debug_print(f"Auto-executing B2 command (second row): {b2_command}")
                        execute_command(b2_command)
                    else:
                        debug_print("Skipping auto-execution: B2 (second row) is empty or missing")
                else:
                    debug_print("Skipping auto-execution: CSV has less than 2 rows")
        except Exception as e:
            debug_print(f"Error checking B2 for auto-execution: {e}")
            debug_print("Skipping auto-execution due to error")
    
    window.after(5000, auto_execute_if_appropriate)

    # --- Status/Info area ---
    error_frame = tk.Frame(window, bg='#1E1E1E')
    error_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

    top_line_frame = tk.Frame(error_frame, bg='#1E1E1E')
    top_line_frame.pack(fill=tk.X)

    # Create the "Status:" label
    status_label = tk.Label(top_line_frame, text="Status:", font=("Arial",10,"bold"),
                       bg='#1E1E1E', fg='#E0E0E0')
    status_label.pack(side=tk.LEFT)

    # Create the text area for command output status with scrollbar
    text_frame = tk.Frame(error_frame, bg='#1E1E1E')
    text_frame.pack(fill=tk.BOTH, expand=True)

    error_match_text = tk.Text(text_frame, wrap="word", bg="black", fg="yellow",
                               height=8, font=("Consolas", current_font_size))
    error_match_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

    # Add scrollbar to text area
    scrollbar = tk.Scrollbar(text_frame, orient="vertical", command=error_match_text.yview)
    scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
    error_match_text.configure(yscrollcommand=scrollbar.set)

    # Initial message in the status area
    error_match_text.insert(tk.END, "📋 McConsole Helper Window Ready\n")
    error_match_text.insert(tk.END, "=" * 50 + "\n")
    error_match_text.insert(tk.END, f"Log file: {log_file_path if log_file_path else 'Not set yet'}\n")
    error_match_text.insert(tk.END, "=" * 50 + "\n\n")
    error_match_text.insert(tk.END, "• Double-click any command to execute\n")
    error_match_text.insert(tk.END, "• Right-click for edit/delete options\n")
    error_match_text.insert(tk.END, "• Use search box to filter commands\n")
    
    # --- Theme status indicator ---
    if theme_name:
        theme_label = tk.Label(
            window, 
            text=f"Theme: {theme_name}", 
            bg="#1E1E1E", 
            fg="#00E676",  # Bright green
            font=("Consolas", 8)
        )
        theme_label.pack(side=tk.BOTTOM, anchor='se', padx=5, pady=2)

    # --- Helper→MobaXterm shutdown on helper close ---
    def on_close_helper():
        if mobaxterm_pid:
            try:
                psutil.Process(mobaxterm_pid).terminate()
            except Exception:
                pass
        window.destroy()

    window.protocol("WM_DELETE_WINDOW", on_close_helper)

    def watch_terminal():
        """
        Monitor terminal process (MobaXterm or SecureCRT) with minimal logging for privacy.
        """
        # 1) Wait until we have the PID
        while mobaxterm_pid is None:
            time.sleep(0.1)

        # 2) Give the terminal GUI plenty of time to come up
        time.sleep(20)

        # Determine terminal keywords based on PID process name
        terminal_keywords = ["mobaxterm", "securecrt", "vshell"]

        # 3) Poll every few seconds with minimal logging
        while True:
            found = False
            try:
                for w in gw.getAllWindows():
                    try:
                        _, win_pid = win32process.GetWindowThreadProcessId(w._hWnd)
                        title_lower = w.title.lower()
                        if win_pid == mobaxterm_pid and w.visible and any(kw in title_lower for kw in terminal_keywords):
                            found = True
                            break
                    except Exception:
                        continue  # Skip windows we can't access
            except Exception as e:
                debug_print(f"watch_terminal error: {e}")

            if not found:
                debug_print("Terminal window no longer found - closing helper")
                window.after(0, on_close_helper)
                return

            time.sleep(5)  # Check every 5 seconds instead of 3

    threading.Thread(target=watch_terminal, daemon=True).start()

    # --- Start Tk mainloop ---
    window.mainloop()

def simulate_typing_by_pid(pid, command):
    """
    Bring the window belonging to pid to front via Win32 and type.
    Returns True on success.
    """
    import win32gui, win32con

    for win in gw.getAllWindows():
        try:
            _, win_pid = win32process.GetWindowThreadProcessId(win._hWnd)
        except Exception:
            continue
        if win_pid == pid and win.visible:
            hwnd = win._hWnd
            try:
                win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
                win32gui.SetForegroundWindow(hwnd)
                time.sleep(0.3)
                pyautogui.typewrite(command + "\n", interval=0)
                return True
            except Exception as e:
                debug_print(f"PID‑typing error: {e}")
                return False
    return False

def terminal_connect(config, console_type, command_queue, monitor_id, local, jumpbox=None, terminal_type="mobaxterm"):
    """
    Launch MobaXterm, bind its real PID, wait for its log file to appear,
    position the window, and then dispatch queued commands.
    """
    import os
    import glob
    import time
    from pathlib import Path

    global log_file_path, last_command, mobaxterm_pid

    # ─── pick your startup command ─────────────────────────────────────────
    ctype = console_type.lower()
    if ctype == "controller" or (ctype == "sut" and local):
        exec_cmd = 'cmd.exe /k'
    else:
        host = config.get(f"{console_type.upper()}_IP", "")
        user = config.get(f"{console_type.upper()}_GUESS", "") 
        # exec_cmd = f'ssh {user}@{host}'

        # jumpbox integration
        if jumpbox:
            # Parse jumpbox format: user@host:port or user@host
            if ':' in jumpbox:
                jumpbox_userhost, jumpbox_port = jumpbox.rsplit(':', 1)
                exec_cmd = f'ssh -J {jumpbox_userhost}:{jumpbox_port} {user}@{host}'
                debug_print(f"Using jumpbox connection: {jumpbox_userhost}:{jumpbox_port} -> {user}@{host}")
            else:
                # jumpbox already contains user@host, default port 22
                exec_cmd = f'ssh -J {jumpbox}:22 {user}@{host}'
                debug_print(f"Using jumpbox connection: {jumpbox}:22 -> {user}@{host}")
        else:
            exec_cmd = f'ssh -4 {user}@{host}'
            debug_print(f"Using direct SSH connection: {user}@{host}")

        if console_type.upper() == "RM":
            candidate_ciphers = [
                "aes256-ctr", "aes192-ctr", "aes256-cbc",
                "aes192-cbc", "aes128-cbc", "3des-cbc"
            ]
            cipher = auto_select_cipher(host, candidate_ciphers)
            if cipher:
                exec_cmd += f' -c {cipher}'

    # Build terminal-specific command
    if terminal_type.lower() == "mobaxterm":
        mxt = os.path.abspath(os.path.join("..", "bins", "mx", "MobaXterm.exe"))
        ini = os.path.abspath(os.path.join("..", "bins", "mx", "MobaXterm.ini"))
        cmdline = f'"{mxt}" -i "{ini}" -exec "{exec_cmd}"'
    elif terminal_type.lower() == "securecrt":
        scrt = os.path.abspath(os.path.join("..", "bins", "SCRT", "SecureCRT.exe"))
        if ctype == "controller" or (ctype == "sut" and local):
            #cmdline = f'"{scrt}"'
            cmdline = f'"{scrt}" /S "LocalShell"'
            debug_print(f"Using SecureCRT LocalShell session for Controller mode")
        else:
            # Get password from config (XXX_WHAT)
            password = config.get(f"{console_type.upper()}_WHAT", "")
            
            # Try to use session file first (more reliable than direct connection)
            session_name = f"{console_type.upper()}_{host.replace('.', '_')}"
            session_dir = os.path.abspath(os.path.join("..", "bins", "SCRT", "Sessions"))
            session_file = os.path.join(session_dir, f"{session_name}.ini")
            
            if os.path.exists(session_file):
                # Use saved session (recommended - won't close after 5 seconds)
                cmdline = f'"{scrt}" /S "{session_name}"'
                debug_print(f"Using SecureCRT session: {session_name}")
            else:
                # Direct connection with password from config
                # /ACCEPTHOSTKEYS: automatically accept host keys (no prompt)
                if password:
                    cmdline = f'"{scrt}" /SSH2 /L {user} /PASSWORD {password} /ACCEPTHOSTKEYS {host}'
                    debug_print(f"Using direct connection with password from config")
                else:
                    cmdline = f'"{scrt}" /SSH2 /L {user} /ACCEPTHOSTKEYS {host}'
                    debug_print(f"WARNING: No password found in config ({console_type.upper()}_WHAT)")
                    debug_print(f"SecureCRT will prompt for password manually")
    else:
        raise ValueError(f"Unsupported terminal: {terminal_type}")

    debug_print(f"Launching {terminal_type}: {cmdline}")
    pid = launch_terminal_and_get_pid(cmdline, terminal_type, timeout=10)
    mobaxterm_pid = pid
    debug_print(f"Bound to PID: {pid}")

    # ─── locate its log ────────────────────────────────────────────────────
    # record when we started MobaXterm
    launch_time = time.time()

    # give MobaXterm a moment to begin writing
    #time.sleep(5)

    if not pid:
        debug_print("ERROR: MobaXterm failed to start within 10 s; aborting helper.")

# use a different folder than mcconsole main program folder which is mcconsole
    #log_dir = os.path.join(os.path.dirname(__file__), "..", "logs")
    log_dir = os.path.join(os.path.dirname(__file__), "..", "logs", "Mcconsole_log")

    # retry until we find a log file created after launch_time
    while True:
        candidates = glob.glob(os.path.join(log_dir, "*.log"))
        recent = [
            p for p in candidates
            if os.path.getctime(p) >= launch_time
        ]
        if recent:
            log_file_path = max(recent, key=os.path.getctime)
            info_print(f"Using log file: {log_file_path}")
            break

        debug_print("No log file found after launch; retrying in 3 s…")
        time.sleep(2)

    # ─── position window ───────────────────────────────────────────────────
    mx, my, mw, mh = get_monitor_geometry(monitor_id)
    debug_print(f"monitor geometry: x={mx}, y={my}, w={mw}, h={mh}")

    if pid and position_window_by_pid(pid, mx, my, mw, mh):
        debug_print("Positioned by PID")
    elif position_window_by_title(
            terminal_type,
            "cmd.exe" if (ctype == "controller" or (ctype == "sut" and local))
            else exec_cmd.split()[-1],
            mx, my, mw, mh
         ):
        debug_print("Positioned by title fallback")
    else:
        debug_print("Failed to position window")

    # ─── dispatch queued commands ──────────────────────────────────────────
    while True:
        if not command_queue.empty():
            cmd = command_queue.get().strip()
            last_command = cmd
            debug_print(f"Dispatching: {cmd}")

            # try typing by PID
            if not (pid and simulate_typing_by_pid(pid, cmd)):
                # fallback: find window by title and type
                simulate_typing_command_terminal(terminal_type, cmd,
                    host if ctype != "controller" else ""
                )
        time.sleep(0.1)

def local_shell(command_queue):
    import pyperclip
    cwd = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "controller"))
    try:
        proc = subprocess.Popen(["powershell.exe"], cwd=cwd, 
                                  creationflags=subprocess.CREATE_NEW_CONSOLE)
        print("Controller launched PowerShell in:", cwd)
    except Exception as e:
        print("Controller failed to launch PowerShell:", str(e))
        return

    time.sleep(3)
    while True:
        if not command_queue.empty():
            command = command_queue.get().strip()
            if not simulate_typing_command(command, "powershell"):
                pyperclip.copy(command)
                print("Controller: Command copied to clipboard:", command)
        time.sleep(0.1)

def simulate_typing_command(command, window_title="powershell"):
    try:
        import pygetwindow as gw
    except ImportError:
        print("pygetwindow module not available; cannot simulate typing.")
        return False

    ps_windows = [w for w in gw.getAllWindows() if window_title.lower() in w.title.lower() and w.visible]
    if ps_windows:
        ps_window = ps_windows[0]
        print("Found PowerShell window:", ps_window.title)
        try:
            ps_window.activate()
            time.sleep(0.5)
            pyautogui.typewrite(command + "\n", interval=0)
            print("Typed command into PowerShell:", command)
            return True
        except Exception as e:
            print("Error typing command in PowerShell:", str(e))
            return False
    else:
        print("No PowerShell window found with title containing:", window_title)
    return False

def load_highlight_keywords(csv_file, config):
    global highlight_keywords, reinforcement_text

    def parse_command(command):
        for key, value in config.items():
            command = command.replace(f"[[{key}]]", str(value))
        return command

    try:
        with open(csv_file, newline="", encoding="utf-8", errors="replace") as csvfile:
            reader = csv.reader(csvfile)
            for row in reader:
                if len(row) >= 3:
                    csv_command = row[1].strip()
                    final_command = parse_command(csv_command)
                    keywords = [] if row[2].strip() == "" else [kw.strip() for kw in row[2].split(",") if kw.strip()]
                    highlight_keywords[final_command] = keywords
                    # (any debug prints you need)

        # (optional DEBUG block)
        if DEBUG:
            # print("[DEBUG] Final highlight_keywords dictionary:")
            # ...
            pass

        # **HERE** we must indent the body of the `if reinforcement_text:` block
        if reinforcement_text:
            reinforcement_text.insert(
                tk.END,
                "Loaded Commands and Keywords with parameters replaced.\n"
            )
            reinforcement_text.see(tk.END)

    except Exception as e:
        print(f"Error loading keywords: {e}")

def find_config_file(host):
    sut_dir = Path(os.path.dirname(__file__)) / ".." / "sut"
    pattern = f"settings.{host}.json"
    matches = list(sut_dir.rglob(pattern))
    if len(matches) == 0:
        print(f"No configuration file matching pattern '{pattern}' found in {sut_dir}.")
        return None
    elif len(matches) > 1:
        print(f"Duplicate configuration files found for host '{host}':")
        for match in matches:
            print(match)
        print("Please rename duplicate files to avoid duplication.")
        input("After renaming, press Enter to re-scan...")
        matches = list(sut_dir.rglob(pattern))
        if len(matches) > 1:
            print("Duplicates still exist. Using the first available file.")
        if len(matches) == 0:
            print("No configuration file found after renaming duplicates.")
            return None
    return matches[0]

def setup_securecrt_config_path(config_path=None):
    """
    Configure SecureCRT to use project's configuration directory via registry.
    
    This ensures all team members share the same session configurations,
    making it easier to maintain consistency across different machines.
    
    Args:
        config_path: Path to SecureCRT config directory. 
                    Defaults to C:\mcqueen\bins\SCRT
    
    Returns:
        bool: True if successful, False otherwise
    """
    import winreg
    
    if config_path is None:
        # Default to project's SCRT directory
        config_path = r"C:\mcqueen\bins\SCRT"
    
    # Normalize and verify path
    config_path = os.path.abspath(config_path)
    
    # Ensure the directory exists
    os.makedirs(config_path, exist_ok=True)
    
    try:
        # SecureCRT registry key location
        key_path = r"Software\VanDyke\SecureCRT"
        
        try:
            # Try to open existing key with write permissions
            key = winreg.OpenKey(
                winreg.HKEY_CURRENT_USER,
                key_path,
                0,
                winreg.KEY_SET_VALUE | winreg.KEY_QUERY_VALUE
            )
        except FileNotFoundError:
            # Key doesn't exist, create it
            debug_print("SecureCRT registry key not found, creating...")
            key = winreg.CreateKeyEx(
                winreg.HKEY_CURRENT_USER,
                key_path,
                0,
                winreg.KEY_SET_VALUE
            )
        
        # Read and backup current value
        try:
            current_value, _ = winreg.QueryValueEx(key, "Config Path")
            if current_value != config_path:
                debug_print(f"📝 Current SecureCRT Config Path: {current_value}")
                debug_print(f"📝 Changing to: {config_path}")
                
                # Optional: Save backup of original path
                winreg.SetValueEx(
                    key,
                    "Config Path Backup",
                    0,
                    winreg.REG_SZ,
                    current_value
                )
            else:
                debug_print(f"✅ SecureCRT Config Path already set correctly: {config_path}")
                winreg.CloseKey(key)
                return True
        except FileNotFoundError:
            debug_print("📝 SecureCRT Config Path not set previously, setting now...")
        
        # Set the new Config Path
        winreg.SetValueEx(
            key,
            "Config Path",
            0,
            winreg.REG_SZ,
            config_path
        )
        
        winreg.CloseKey(key)
        
        # Verify the change took effect
        key = winreg.OpenKey(
            winreg.HKEY_CURRENT_USER,
            key_path,
            0,
            winreg.KEY_QUERY_VALUE
        )
        verify_value, _ = winreg.QueryValueEx(key, "Config Path")
        winreg.CloseKey(key)
        
        if verify_value == config_path:
            debug_print(f"✅ SecureCRT Config Path successfully set to: {config_path}")
            debug_print(f"✅ Registry change verified")
            return True
        else:
            debug_print(f"⚠️ Registry verification failed!")
            debug_print(f"   Expected: {config_path}")
            debug_print(f"   Got: {verify_value}")
            return False
            
    except PermissionError:
        debug_print("❌ Permission denied accessing registry")
        debug_print("   Try running the script as administrator")
        return False
    except Exception as e:
        debug_print(f"❌ Error setting SecureCRT Config Path: {type(e).__name__}: {e}")
        import traceback
        debug_print(traceback.format_exc())
        return False

def main():
    args = parse_args()

    # Set up logging redirection early
    log_redirector, log_file_path_shell = setup_shell_logging_redirect(args.console_type, args.sut_ip)

    terminal_type = args.terminal  # Get terminal type
    
    # Configure SecureCRT if it's the selected terminal
    if terminal_type.lower() == "securecrt":
        debug_print("🔧 Configuring SecureCRT registry settings...")
        scrt_config_path = os.path.abspath(os.path.join(
            os.path.dirname(__file__), "..", "bins", "SCRT"
        ))
        
        if not setup_securecrt_config_path(scrt_config_path):
            debug_print("⚠️ Warning: Failed to configure SecureCRT registry")
            debug_print("   SecureCRT may use default user directory for configs")
            debug_print("   Sessions may need to be configured manually")
            # Don't exit, continue anyway
        else:
            debug_print("✅ SecureCRT configured successfully")
    

    if args.console_type.lower() == "controller":
        args.local = True

    sound_enabled = args.sound.lower() == "true"
    theme_name = args.theme
    global log_file_path, log_folder

    # Agent mode is always active now
    agent_mode_active = True

    num_monitors = get_monitor_count()
    debug_print(f"Detected {num_monitors} monitors")

    helper_counter = get_helper_window_counter()
    selected_monitor = (helper_counter - 1) % num_monitors if num_monitors > 0 else args.monitor_id
    debug_print(f"Automatically selected monitor {selected_monitor} (helper window #{helper_counter})")

    if args.console_type.lower() == "controller":
        start_time = time.localtime()
        timestamp = time.strftime("%Y%m%d_%H%M%S", start_time)
        logs_parent = os.path.join(os.path.dirname(os.getcwd()), "logs")
        os.makedirs(logs_parent, exist_ok=True)
        log_folder = os.path.join(logs_parent, f"McConsole_{args.sut_ip}_{timestamp}")
        os.makedirs(log_folder, exist_ok=True)
        file_time = time.strftime("%H%M%S", start_time)
        log_file_path = os.path.join(log_folder, f"McConsole_output_{file_time}.log")
    else:
        log_file_path = None

    # Try to load config for sut_ip
    config = read_config_m(args.sut_ip)
    config_found = isinstance(config, dict)
    if not config_found:
        debug_print(f"Failed to load configuration for {args.sut_ip}; continuing with MobaXterm only.")
        config = {}

    command_queue = queue.Queue()
    
    helper_folder = os.path.join(os.path.dirname(__file__), "helper")
    csv_prefix = f"{config.get('SUT_SKU')}_" if config.get("SUT_SKU") else ""
    csv_file = os.path.join(helper_folder, f"{csv_prefix}{args.console_type.upper()}_commands.csv")
    csv_exists = os.path.exists(csv_file)

    # Always launch MobaXterm (threaded)
    threading.Thread(
        target=terminal_connect,
        args=(config, args.console_type, command_queue, selected_monitor, args.local, args.jumpbox, terminal_type),
        daemon=True
    ).start()

    # Only launch helper window if config AND CSV are present
    if config_found and csv_exists:
        threading.Thread(
            target=lambda: load_highlight_keywords(csv_file, config),
            daemon=True
        ).start()
        try:
            open_helper_window(
                command_queue,
                args.console_type,
                config,
                args.sut_ip,
                sound_enabled,
                monitor_id=selected_monitor,
                theme_name=theme_name
            )
        except KeyboardInterrupt:
            print("Program interrupted. Exiting.", flush=True)
    else:
        debug_print("Helper window not opened: config or CSV not found.")

if __name__ == "__main__":
    main()