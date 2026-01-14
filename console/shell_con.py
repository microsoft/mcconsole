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

def setup_shell_logging_redirect(console_type, sut_ips):
    """Redirect all print() statements to log file while keeping console open."""
    
    # Create logs directory
    log_dir = Path(r"C:\mcqueen\logs\Mcconsole_log")
    log_dir.mkdir(parents=True, exist_ok=True)
    
    # Generate timestamp for log filename
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    sut_list = "_".join(sut_ips)
    log_file = log_dir / f"shell_mx_multi_{console_type}_{sut_list}_{timestamp}.log"
    
    # Show user where logs are going before redirecting
    print(f"🚀 Shell MX Multi ({console_type}) Starting for SUTs: {', '.join(sut_ips)}...")
    print(f"📁 Debug logs will be saved to:")
    print(f"   {log_dir}")
    print(f"📄 Current session log: shell_mx_multi_{console_type}_{sut_list}_{timestamp}.log")
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


# ---------------- Global Variables ----------------
DEBUG = True
highlight_keywords = {}
reinforcement_text = None    # Widget used in the reinforcement window for additional output
error_match_text = None      # Widget to display keyword match output
lock = threading.Lock()
helper_window_count = 0
last_command = ""            # Last command executed (the one we will check reinforcement against)
log_folder = None            # For controller mode log folder
reinforce_btn = None         # The reinforcement button widget (if needed)
sut_configs = {}             # Dictionary to store configurations for each SUT
# --------------------------------------------------


def debug_print(msg):
    if DEBUG:
        msg_str = str(msg)
        
        # Skip excessive window scanning logs for privacy
        if "window: title=" in msg_str and "pid=" in msg_str:
            return  # Don't log window titles that may contain sensitive info
        
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
    import re
    ansi_escape = re.compile(r'\x1B(?:[@-Z\\-_]|\[[0-?]*[ -/]*[@-~])')
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


def parse_args():
    parser = argparse.ArgumentParser(
        description="Interactive Multi-SUT Console with dynamic helper window"
    )
    parser.add_argument("--console_type", required=True,
                        help="Console type (e.g., RM, SUT, SOC, Controller)")
    parser.add_argument("--sut_ip", required=True, nargs='+',
                        help="Multiple SUT IPs for loading configurations")
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
    return parser.parse_args()


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


def launch_cmd_window(sut_ip, user_cmd, config):
    """
    Launch a new cmd window for the specified SUT. 
    Hard-code McController.bat as startup, then execute user command in the same window.
    Capture both command context and output to SUT-specific log files.
    """
    def parse_command_for_sut(cmd_str, sut_config, sut_ip):
        """Parse command with SUT-specific config and replace SUT_TITLE with sut_ip"""
        if not cmd_str or not cmd_str.strip():
            return ""
        
        debug_print(f"Parsing command for {sut_ip}: '{cmd_str}'")
        
        # First, replace SUT_TITLE with the actual sut_ip
        cmd_str = cmd_str.replace('[[SUT_TITLE]]', sut_ip)
        debug_print(f"After SUT_TITLE replacement: '{cmd_str}'")
        
        # Then do the rest of the config-based substitutions
        for k, v in sut_config.items():
            placeholder = f'[[{k}]]'
            if placeholder in cmd_str:
                cmd_str = cmd_str.replace(placeholder, str(v))
                debug_print(f"Replaced {placeholder} with '{v}': '{cmd_str}'")
        
        debug_print(f"Final parsed command: '{cmd_str}'")
        return cmd_str
    
    try:
        # Create SUT-specific log directory and file
        global log_folder
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        if log_folder:
            sut_log_file = os.path.join(log_folder, f"{sut_ip}_{timestamp}.log")
        else:
            # Fallback if log_folder is not set
            base_log_dir = r"C:\mcqueen\logs"
            temp_log_folder = os.path.join(base_log_dir, f"McConsole_Multi_temp_{timestamp}")
            os.makedirs(temp_log_folder, exist_ok=True)
            sut_log_file = os.path.join(temp_log_folder, f"{sut_ip}_{timestamp}.log")

        debug_print(f"Replicating McController.bat logic inline (avoiding cmd.exe /k issue)")
        debug_print(f"Original user command: '{user_cmd}'")
        debug_print(f"Config for {sut_ip}: {list(config.keys())}")
        debug_print(f"SUT log file: {sut_log_file}")
        
        # Parse user command with SUT-specific config
        parsed_user_cmd = parse_command_for_sut(user_cmd, config, sut_ip)
        
        # Create batch file in the controller directory
        controller_dir = r"C:\mcqueen\controller"
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")[:-3]
        batch_filename = f"temp_batch_{sut_ip}_{timestamp}.bat"
        batch_file = os.path.join(controller_dir, batch_filename)
        
        debug_print(f"Creating batch file: {batch_file}")
        debug_print(f"Parsed user cmd: '{parsed_user_cmd}'")
        
        # Ensure controller directory exists
        os.makedirs(controller_dir, exist_ok=True)
        
        with open(batch_file, 'w') as f:
            f.write(f"@echo off\n")
            f.write(f"title {sut_ip} - McQueen Multi-SUT Console\n")
            f.write(f"color 0A\n")  # Green text on black background
            
            # Set up SUT log file variable
            f.write(f'set "SUT_LOG={sut_log_file}"\n')
            
            # Initialize log file
            f.write(f'echo McQueen Multi-SUT Console Session for {sut_ip} > "%SUT_LOG%"\n')
            f.write(f'echo Session started at %DATE% %TIME% >> "%SUT_LOG%"\n')
            f.write(f'echo. >> "%SUT_LOG%"\n')
            
            # Banner with logging
            f.write(f"echo.\n")
            f.write(f"echo ========================================= & echo ========================================= >> \"%SUT_LOG%\"\n")
            f.write(f"echo   McQueen Multi-SUT Console for {sut_ip} & echo   McQueen Multi-SUT Console for {sut_ip} >> \"%SUT_LOG%\"\n")
            f.write(f"echo ========================================= & echo ========================================= >> \"%SUT_LOG%\"\n")
            f.write(f"echo Current directory: %CD% & echo Current directory: %CD% >> \"%SUT_LOG%\"\n")
            f.write(f"echo Current time: %DATE% %TIME% & echo Current time: %DATE% %TIME% >> \"%SUT_LOG%\"\n")
            f.write(f"echo SUT IP: {sut_ip} & echo SUT IP: {sut_ip} >> \"%SUT_LOG%\"\n")
            f.write(f"echo Batch file: {batch_filename} & echo Batch file: {batch_filename} >> \"%SUT_LOG%\"\n")
            f.write(f"echo ========================================= & echo ========================================= >> \"%SUT_LOG%\"\n")
            f.write(f"echo. & echo. >> \"%SUT_LOG%\"\n")
            
            # Replicate McController.bat logic with logging
            f.write(f"echo [STEP 1] Setting up McQueen Python environment... & echo [STEP 1] Setting up McQueen Python environment... >> \"%SUT_LOG%\"\n")
            f.write(f"echo [STEP 1] Replicating McController.bat logic inline & echo [STEP 1] Replicating McController.bat logic inline >> \"%SUT_LOG%\"\n")
            f.write(f"echo. & echo. >> \"%SUT_LOG%\"\n")
            
            # Set MCQ_BASE to parent of controller directory
            f.write(f"for %%I in (\"%CD%\\..\") do set \"MCQ_BASE=%%~fI\"\n")
            f.write(f"echo [STEP 1] MCQ_BASE set to: %MCQ_BASE% & echo [STEP 1] MCQ_BASE set to: %MCQ_BASE% >> \"%SUT_LOG%\"\n")
            
            # Add Python and Scripts to PATH
            f.write(f"set \"PATH=%MCQ_BASE%\\python;%MCQ_BASE%\\python\\Scripts;%PATH%\"\n")
            f.write(f"echo [STEP 1] Updated PATH with McQueen Python & echo [STEP 1] Updated PATH with McQueen Python >> \"%SUT_LOG%\"\n")
            f.write(f"echo. & echo. >> \"%SUT_LOG%\"\n")
            
            # Verify Python environment with output capture
            f.write(f"echo [STEP 1] McQueen Python env activated: & echo [STEP 1] McQueen Python env activated: >> \"%SUT_LOG%\"\n")
            f.write(f"echo [STEP 1]   python is \"%MCQ_BASE%\\python\\python.exe\" & echo [STEP 1]   python is \"%MCQ_BASE%\\python\\python.exe\" >> \"%SUT_LOG%\"\n")
            f.write(f"echo [STEP 1]   pip    is \"%MCQ_BASE%\\python\\Scripts\\pip.exe\" & echo [STEP 1]   pip    is \"%MCQ_BASE%\\python\\Scripts\\pip.exe\" >> \"%SUT_LOG%\"\n")
            f.write(f"echo. & echo. >> \"%SUT_LOG%\"\n")
            f.write(f"echo [STEP 1] Verifying Python and pip locations: & echo [STEP 1] Verifying Python and pip locations: >> \"%SUT_LOG%\"\n")
            
            # Capture where python output
            f.write(f"where python 2>nul && (\n")
            f.write(f"  for /f \"tokens=*\" %%a in ('where python 2^>nul') do (\n")
            f.write(f"    echo [STEP 1] Python found: %%a & echo [STEP 1] Python found: %%a >> \"%SUT_LOG%\"\n")
            f.write(f"  )\n")
            f.write(f") || (\n")
            f.write(f"  echo [STEP 1] WARNING: python not found in PATH & echo [STEP 1] WARNING: python not found in PATH >> \"%SUT_LOG%\"\n")
            f.write(f")\n")
            
            # Capture where pip output
            f.write(f"where pip 2>nul && (\n")
            f.write(f"  for /f \"tokens=*\" %%a in ('where pip 2^>nul') do (\n")
            f.write(f"    echo [STEP 1] Pip found: %%a & echo [STEP 1] Pip found: %%a >> \"%SUT_LOG%\"\n")
            f.write(f"  )\n")
            f.write(f") || (\n")
            f.write(f"  echo [STEP 1] WARNING: pip not found in PATH & echo [STEP 1] WARNING: pip not found in PATH >> \"%SUT_LOG%\"\n")
            f.write(f")\n")
            
            f.write(f"echo. & echo. >> \"%SUT_LOG%\"\n")
            f.write(f"echo [STEP 1] McQueen environment setup completed & echo [STEP 1] McQueen environment setup completed >> \"%SUT_LOG%\"\n")
            f.write(f"echo. & echo. >> \"%SUT_LOG%\"\n")
            
            # Now execute user command in the McQueen environment
            if parsed_user_cmd and parsed_user_cmd.strip():
                f.write(f"echo [STEP 2] Executing user command in McQueen environment... & echo [STEP 2] Executing user command in McQueen environment... >> \"%SUT_LOG%\"\n")
                f.write(f"echo [STEP 2] Target SUT: {sut_ip} & echo [STEP 2] Target SUT: {sut_ip} >> \"%SUT_LOG%\"\n")
                
                # Handle multiple commands separated by ;;
                user_commands = [cmd.strip() for cmd in parsed_user_cmd.split(";;") if cmd.strip()]
                f.write(f"echo [STEP 2] Found {len(user_commands)} command(s) to execute & echo [STEP 2] Found {len(user_commands)} command(s) to execute >> \"%SUT_LOG%\"\n")
                f.write(f"echo. & echo. >> \"%SUT_LOG%\"\n")
                
                for i, cmd in enumerate(user_commands, 1):
                    f.write(f"echo [STEP 2.{i}] Executing: {cmd} & echo [STEP 2.{i}] Executing: {cmd} >> \"%SUT_LOG%\"\n")
                    f.write(f"echo [STEP 2.{i}] Starting at %TIME%... & echo [STEP 2.{i}] Starting at %TIME%... >> \"%SUT_LOG%\"\n")
                    
                    # Execute command and capture output to both console and log
                    # Windows does not have 'tee' by default; use a FOR loop to echo output to both console and log
                    f.write(f'echo [STEP 2.{i}] Command Output: >> "%SUT_LOG%"\n')
                    # f.write(f'for /f "delims=" %%A in (\'{cmd} 2^>^&1\') do (echo %%A & echo %%A >> "%SUT_LOG%")\n')

                    # f.write(f'for /f "delims=" %%A in (\'cmd /c "{cmd}" 2^>^&1\') do (echo %%A & echo %%A >> "%SUT_LOG%")\n')

                    if cmd.lower().startswith("curl "):
                        # # Use for /f only for curl to tee live output to console + log
                        # f.write(f'for /f "delims=" %%A in (\'{cmd} 2^>^&1\') do (echo %%A & echo %%A >> "%SUT_LOG%")\n')

                        # Wrap curl in cmd /c to protect special characters
                        safe_cmd = cmd.replace('"', '""')  # escape quotes for batch
                        f.write(f'for /f "delims=" %%A in (\'cmd /c "{safe_cmd} 2^>^&1"\') do (echo %%A & echo %%A >> "%SUT_LOG%")\n')
                    else:
                        # For everything else (python, exe, bat, etc.) just redirect safely
                        # f.write(f'{cmd} >> "%SUT_LOG%" 2>&1\n')
                        f.write(f'{cmd}\n')

                    f.write(f"echo. & echo. >> \"%SUT_LOG%\"\n")
                    f.write(f"echo [STEP 2.{i}] Completed at %TIME% with errorlevel: %ERRORLEVEL% & echo [STEP 2.{i}] Completed at %TIME% with errorlevel: %ERRORLEVEL% >> \"%SUT_LOG%\"\n")
                    
                    # Add small delay between commands to avoid overwhelming servers
                    if i < len(user_commands):
                        f.write(f"echo [STEP 2.{i}] Waiting 2 seconds before next command... & echo [STEP 2.{i}] Waiting 2 seconds before next command... >> \"%SUT_LOG%\"\n")
                        f.write(f"timeout /t 2 /nobreak >nul\n")
                    f.write(f"echo. & echo. >> \"%SUT_LOG%\"\n")
            else:
                f.write(f"echo [STEP 2] No user command to execute. & echo [STEP 2] No user command to execute. >> \"%SUT_LOG%\"\n")
                f.write(f"echo. & echo. >> \"%SUT_LOG%\"\n")
            
            # Footer
            f.write(f"echo ========================================= & echo ========================================= >> \"%SUT_LOG%\"\n")
            f.write(f"echo All commands completed for {sut_ip} & echo All commands completed for {sut_ip} >> \"%SUT_LOG%\"\n")
            f.write(f"echo Final completion time: %DATE% %TIME% & echo Final completion time: %DATE% %TIME% >> \"%SUT_LOG%\"\n")
            f.write(f"echo Log saved to: %SUT_LOG% & echo Log saved to: %SUT_LOG% >> \"%SUT_LOG%\"\n")
            f.write(f"echo ========================================= & echo ========================================= >> \"%SUT_LOG%\"\n")
            f.write(f"echo.\n")
            f.write(f"echo Press any key to close this window...\n")
            f.write(f"pause\n")
            
            # Create a separate cleanup batch file to delete this one after a delay
            cleanup_filename = f"cleanup_{sut_ip}_{timestamp}.bat"
            cleanup_file = os.path.join(controller_dir, cleanup_filename)
            f.write(f"echo Scheduling cleanup of temporary files...\n")
            f.write(f"start \"\" /min cmd /c \"timeout /t 3 /nobreak >nul && del \"{batch_filename}\" 2>nul && del \"{cleanup_filename}\" 2>nul\"\n")
        
        # Launch the cmd window with the batch file - use full path and run from controller directory
        subprocess.Popen(['cmd', '/k', batch_filename], 
                        creationflags=subprocess.CREATE_NEW_CONSOLE,
                        cwd=controller_dir)  # Set working directory to controller
        
        # Schedule cleanup of any leftover temp batch files from previous runs
        cleanup_thread = threading.Thread(
            target=cleanup_old_batch_files,
            args=(controller_dir,),
            daemon=True
        )
        cleanup_thread.start()
        
        debug_print(f"Successfully launched cmd window for {sut_ip}")
        debug_print(f"  - Batch file: {batch_file}")
        debug_print(f"  - Working directory: {controller_dir}")
        debug_print(f"  - SUT log file: {sut_log_file}")
        debug_print(f"  - McController logic: Replicated inline (no direct .bat call)")
        debug_print(f"  - Parsed user cmd: '{parsed_user_cmd}'")
        
    except Exception as e:
        debug_print(f"Error launching cmd window for {sut_ip}: {e}")
        import traceback
        debug_print(f"Full traceback: {traceback.format_exc()}")



def cleanup_old_batch_files(controller_dir):
    """
    Clean up old temporary batch files that might be left behind.
    Run this on a background thread with a delay.
    """
    try:
        time.sleep(5)  # Wait 5 seconds to ensure current batch files are running
        
        import glob
        pattern = os.path.join(controller_dir, "temp_batch_*.bat")
        old_files = glob.glob(pattern)
        
        for file_path in old_files:
            try:
                # Check if file is older than 1 minute
                file_age = time.time() - os.path.getctime(file_path)
                if file_age > 60:  # 1 minute
                    os.remove(file_path)
                    debug_print(f"Cleaned up old batch file: {os.path.basename(file_path)}")
            except Exception as e:
                debug_print(f"Could not remove {file_path}: {e}")
        
        # Also clean up old cleanup files
        cleanup_pattern = os.path.join(controller_dir, "cleanup_*.bat")
        old_cleanup_files = glob.glob(cleanup_pattern)
        
        for file_path in old_cleanup_files:
            try:
                file_age = time.time() - os.path.getctime(file_path)
                if file_age > 60:  # 1 minute
                    os.remove(file_path)
                    debug_print(f"Cleaned up old cleanup file: {os.path.basename(file_path)}")
            except Exception as e:
                debug_print(f"Could not remove {file_path}: {e}")
                
    except Exception as e:
        debug_print(f"Error in cleanup_old_batch_files: {e}")


def open_helper_window_multi(console_type, sut_configs, sut_ips, sound_enabled=False, monitor_id=0, theme_name=None):
    """
    Multi-SUT helper window that launches multiple cmd windows when commands are executed.
    This version has no Focus Console button and is not bound to any MobaXterm window.
    """
    global reinforcement_text, error_match_text, reinforce_btn
    global helper_sound_enabled, helper_theme_name

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
    sut_list = ", ".join(sut_ips)
    window.title(f"Multi-SUT Controller - {sut_list}")
    window.geometry(helper_geometry_str)
    window.tk.call('tk', 'scaling', 1.7)
    window.configure(bg='#1E1E1E')

    # --- Font zoom support ---
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

    # --- Toolbar (without Focus Console button) ---
    toolbar = tk.Frame(window, bg='#1E1E1E')
    toolbar.pack(fill=tk.X, padx=10, pady=(5,0))

    # Add multi-SUT status label
    status_label = tk.Label(
        toolbar,
        text=f"Multi-SUT Controller ({len(sut_ips)} SUTs)",
        bg="#1E1E1E",
        fg="#00E676",
        font=("Consolas", 12, "bold")
    )
    status_label.pack(side=tk.LEFT, padx=5, pady=5)

    # --- Command execution callbacks ---
    def execute_command_multi(cmd_str):
        """
        Execute command for all SUTs by launching multiple cmd windows.
        McController.bat is hard-coded as the startup command.
        """
        global last_command
        
        # Find the description for this command
        description = "Command Execution"
        for desc, original_cmd, kw, hover in all_commands:
            if original_cmd.strip() == cmd_str.strip():
                description = desc
                break
        
        if error_match_text:
            error_match_text.insert(tk.END, f"\n⚡ Executing: {description}\n")
            error_match_text.insert(tk.END, f"🎯 Targeting {len(sut_ips)} SUTs: {', '.join(sut_ips)}\n")
            error_match_text.insert(tk.END, f"🔧 Using inline McController logic (no .bat call)\n")
            error_match_text.see(tk.END)
        
        # Launch cmd windows for each SUT
        for sut_ip in sut_ips:
            if sut_ip in sut_configs:
                debug_print(f"Launching command for {sut_ip}: user='{cmd_str}'")
                launch_cmd_window(sut_ip, cmd_str, sut_configs[sut_ip])
                if error_match_text:
                    error_match_text.insert(tk.END, f"📤 Launched cmd window for {sut_ip}\n")
                    
                    # Show what the parsed commands will look like
                    def quick_parse(cmd_str, config, sut_ip):
                        if not cmd_str:
                            return "No command"
                        parsed = cmd_str.replace('[[SUT_TITLE]]', sut_ip)
                        for k, v in config.items():
                            parsed = parsed.replace(f'[[{k}]]', str(v))
                        return parsed[:100] + "..." if len(parsed) > 100 else parsed
                    
                    sample_user = quick_parse(cmd_str, sut_configs[sut_ip], sut_ip)
                    error_match_text.insert(tk.END, f"   📋 McController: Inline logic (Python env setup)\n")
                    error_match_text.insert(tk.END, f"   📋 User Cmd: {sample_user}\n")
                    error_match_text.insert(tk.END, f"   📁 Config loaded: {len(sut_configs[sut_ip])} parameters\n")
                    error_match_text.see(tk.END)
                
                if sound_enabled:
                    play_random_sound(theme_name)
                
                time.sleep(0.5)  # Small delay between launches
            else:
                if error_match_text:
                    error_match_text.insert(tk.END, f"❌ No config found for {sut_ip}\n")
                    error_match_text.see(tk.END)
        
        last_command = cmd_str
        
        if error_match_text:
            error_match_text.insert(tk.END, "✅ Multi-SUT command execution complete.\n")
            error_match_text.insert(tk.END, "💡 Each window sets up McQueen Python env, then runs your command\n")
            error_match_text.insert(tk.END, "-" * 50 + "\n")
            error_match_text.see(tk.END)

    def on_double_click(event):
        sel = tree.selection()
        if sel:
            vals = tree.item(sel[0], "values")
            if len(vals) >= 2 and vals[1].strip():  # Ensure there's a command
                execute_command_multi(vals[1])

    def on_right_click(event):
        row = tree.identify_row(event.y)
        if not row:
            return
        
        tree.selection_set(row)
        values = tree.item(row, "values")
        if len(values) < 2 or not values[1].strip():
            return
            
        desc, cmd = values[0], values[1]
        
        # Create popup window for command editing
        popup = tk.Toplevel(window)
        popup.title("Edit Multi-SUT Command")
        popup.geometry("1200x600")
        popup.minsize(900, 400)
        popup.focus_set()
        
        popup.columnconfigure(0, weight=1)
        popup.rowconfigure(2, weight=1)
        
        # Font size control
        popup_font_size = [12]
        
        def update_popup_fonts():
            text_widget.config(font=("Consolas", popup_font_size[0]))
            label.config(font=("Arial", popup_font_size[0], "bold"))
            help_label.config(font=("Arial", max(8, popup_font_size[0] - 2)))
        
        def zoom_in_popup(event=None):
            popup_font_size[0] = min(popup_font_size[0] + 1, 24)
            update_popup_fonts()
        
        def zoom_out_popup(event=None):
            popup_font_size[0] = max(popup_font_size[0] - 1, 8)
            update_popup_fonts()
        
        popup.bind("<Control-MouseWheel>", lambda e: zoom_in_popup() if e.delta > 0 else zoom_out_popup())
        popup.bind("<Control-plus>", zoom_in_popup)
        popup.bind("<Control-minus>", zoom_out_popup)
        
        # Add labels and info
        label = tk.Label(popup, text=f"Edit Multi-SUT Command: {desc}", font=("Arial", 12, "bold"))
        label.grid(row=0, column=0, sticky="ew", padx=10, pady=5)
        
        info_label = tk.Label(popup, text=f"Will execute on {len(sut_ips)} SUTs: {', '.join(sut_ips)}", 
                             font=("Arial", 10), fg="blue")
        info_label.grid(row=1, column=0, sticky="ew", padx=10, pady=2)
        
        # Text frame with scrollbars
        text_frame = tk.Frame(popup)
        text_frame.grid(row=2, column=0, sticky="nsew", padx=10, pady=5)
        text_frame.columnconfigure(0, weight=1)
        text_frame.rowconfigure(0, weight=1)
        
        text_widget = tk.Text(text_frame, wrap="none", height=15, font=("Consolas", 12))
        text_widget.grid(row=0, column=0, sticky="nsew")
        
        v_scrollbar = tk.Scrollbar(text_frame, orient="vertical", command=text_widget.yview)
        v_scrollbar.grid(row=0, column=1, sticky="ns")
        text_widget.configure(yscrollcommand=v_scrollbar.set)
        
        h_scrollbar = tk.Scrollbar(text_frame, orient="horizontal", command=text_widget.xview)
        h_scrollbar.grid(row=1, column=0, sticky="ew")
        text_widget.configure(xscrollcommand=h_scrollbar.set)
        
        text_widget.insert("1.0", cmd)
        text_widget.focus_set()
        text_widget.tag_add("sel", "1.0", "end-1c")
        
        # Button frame
        button_frame = tk.Frame(popup)
        button_frame.grid(row=3, column=0, sticky="ew", padx=10, pady=10)
        
        def save_and_exec():
            command_text = text_widget.get("1.0", "end-1c")
            command_text = re.sub(r'\s+', ' ', command_text).strip()
            execute_command_multi(command_text)
            popup.destroy()
        
        def cancel_edit():
            popup.destroy()
        
        def handle_enter(event):
            save_and_exec()
            return "break"
        
        tk.Button(button_frame, text="Save and Execute on All SUTs", command=save_and_exec, 
                  bg="#28a745", fg="white", font=("Arial", 12, "bold")).pack(side=tk.LEFT, padx=5)
        tk.Button(button_frame, text="Cancel", command=cancel_edit, 
                  bg="#dc3545", fg="white", font=("Arial", 12)).pack(side=tk.LEFT, padx=5)
        
        text_widget.bind("<Return>", handle_enter)
        popup.bind("<Escape>", lambda e: cancel_edit())
        
        help_label = tk.Label(popup, text="Tip: Press Enter to execute, Escape to cancel, Ctrl+Mouse Wheel to zoom", 
                             font=("Arial", 10), fg="gray")
        help_label.grid(row=4, column=0, sticky="ew", padx=10, pady=2)

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

    # Load CSV for the first SUT (assume all SUTs use the same command structure)
    helper_folder = os.path.join(os.path.dirname(__file__), "helper")
    first_sut = sut_ips[0]
    first_config = sut_configs.get(first_sut, {})
    prefix = f"{first_config.get('SUT_SKU')}_" if first_config.get("SUT_SKU") else ""
    
    csv_file = os.path.join(helper_folder, f"{prefix}CONTROLLER_commands.csv")

    # Read CSV headers
    column_titles = ["Description", "Command", "Keywords", "HoverText"]
    if os.path.exists(csv_file):
        try:
            with open(csv_file, newline="", encoding="utf-8", errors="replace") as cf:
                reader = csv.reader(cf)
                header_row = next(reader, None)
                if header_row:
                    column_titles = []
                    for i, header in enumerate(header_row[:4]):
                        if header.strip():
                            column_titles.append(header.strip())
                        else:
                            fallbacks = ["Description", "Command", "Keywords", "HoverText"]
                            column_titles.append(fallbacks[i] if i < len(fallbacks) else f"Column{i+1}")
                    
                    while len(column_titles) < 4:
                        fallbacks = ["Description", "Command", "Keywords", "HoverText"]
                        column_titles.append(fallbacks[len(column_titles)])
        except Exception as e:
            print(f"Error reading CSV headers: {e}")

    # Create Treeview
    tree = ttk.Treeview(frame, columns=("Col1","Col2","Col3","Col4"), show="headings")
    
    tree.heading("Col1", text=column_titles[0]); tree.column("Col1", width=200)
    tree.heading("Col2", text=column_titles[1]); tree.column("Col2", width=350) 
    tree.heading("Col3", text=column_titles[2]); tree.column("Col3", width=200)
    tree.heading("Col4", text=column_titles[3]); tree.column("Col4", width=0, stretch=False)
    
    tree.pack(fill=tk.BOTH, expand=True)
    tree.bind("<Double-1>", on_double_click)
    tree.bind("<Button-3>", on_right_click)

    # Load commands from CSV
    all_commands = []

    print(f"Loading commands from CSV file: {csv_file}")
    if not os.path.exists(csv_file):
        tree.insert("", tk.END, values=("Error", f"CSV not found: {csv_file}", "", ""))
    else:
        try:
            with open(csv_file, newline="", encoding="utf-8", errors="replace") as cf:
                reader = csv.reader(cf)
                row_index = 0
                
                for row in reader:
                    if row_index == 0:  # Skip header
                        row_index += 1
                        continue
                    
                    if not any(cell.strip() for cell in row):  # Skip empty rows
                        row_index += 1
                        continue
                    
                    # Check if it's a category title
                    is_category = len(row) >= 1 and row[0].strip() and (len(row) == 1 or not row[1].strip())
                    
                    if is_category:
                        tree.insert("", tk.END, values=(row[0], "", "", ""), tags=("category",))
                    elif len(row) >= 2:
                        desc = row[0]
                        cmd = row[1]  # Don't parse here, will parse per SUT during execution
                        kw = row[2].strip() if len(row) >= 3 else ""
                        hover_text = row[3].strip() if len(row) >= 4 else ""
                        
                        all_commands.append((desc, cmd, kw, hover_text))
                        
                        tags = ("hoverable",) if hover_text else ()
                        tree.insert("", tk.END, values=(desc, cmd, kw, hover_text), tags=tags)
                    
                    row_index += 1
                
                if not all_commands and not any(tree.exists(item) for item in tree.get_children()):
                    tree.insert("", tk.END, values=("No commands loaded","CSV empty","",""))
        except Exception as e:
            tree.insert("", tk.END, values=("Error loading CSV", str(e), "", ""))

    # Configure styling
    tree.tag_configure("category", 
                      background="#004080", 
                      font=("Arial", 10, "bold"), 
                      foreground="white")
    tree.tag_configure("hoverable", background="#d0f0d0")

    # --- Search functionality ---
    search_frame = tk.Frame(frame, bg='#1E1E1E')
    search_frame.pack(fill=tk.X, pady=5)
    tk.Label(search_frame, text="Search:", bg='#1E1E1E', fg='#E0E0E0').pack(side=tk.LEFT, padx=5)
    search_entry = tk.Entry(search_frame, bg='#2A2A2A', fg='#E0E0E0', insertbackground='#E0E0E0')
    search_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)

    def filter_commands(event=None):
        q = search_entry.get().lower()
        tree.delete(*tree.get_children())
        
        shown_categories = set()
        
        # First pass: identify categories with matches
        for desc, cmd, kw, hover in all_commands:
            if q in desc.lower() or q in cmd.lower() or q in kw.lower():
                category = desc.split(':', 1)[0] if ':' in desc else ""
                if category:
                    shown_categories.add(category)
        
        # Second pass: insert categories and matching items
        current_category = None
        
        for desc, cmd, kw, hover in all_commands:
            is_match = q in desc.lower() or q in cmd.lower() or q in kw.lower()
            category = desc.split(':', 1)[0] if ':' in desc else ""
            
            if category and category != current_category and category in shown_categories:
                tree.insert("", tk.END, values=(category, "", "", ""), tags=("category",))
                current_category = category
            
            if is_match:
                tags = ("hoverable",) if hover else ()
                tree.insert("", tk.END, values=(desc, cmd, kw, hover), tags=tags)

    search_entry.bind("<KeyRelease>", filter_commands)
    search_entry.bind("<Return>", filter_commands)

    # --- Error match area ---
    error_frame = tk.Frame(window, bg='#1E1E1E')
    error_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
    
    top_line_frame = tk.Frame(error_frame, bg='#1E1E1E')
    top_line_frame.pack(fill=tk.X)
    
    ai_label = tk.Label(top_line_frame, text="Multi-SUT Console:", font=("Arial",10,"bold"), 
                       bg='#1E1E1E', fg='#E0E0E0')
    ai_label.pack(side=tk.LEFT)
    
    text_frame = tk.Frame(error_frame, bg='#1E1E1E')
    text_frame.pack(fill=tk.BOTH, expand=True)
    
    error_match_text = tk.Text(text_frame, wrap="word", bg="black", fg="yellow",
                               height=8, font=("Consolas", current_font_size))
    error_match_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

    scrollbar = tk.Scrollbar(text_frame, orient="vertical", command=error_match_text.yview)
    scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
    error_match_text.configure(yscrollcommand=scrollbar.set)

    # Initial message
    error_match_text.insert(tk.END, "🚀 Multi-SUT Controller Ready\n")
    error_match_text.insert(tk.END, "=" * 50 + "\n")
    error_match_text.insert(tk.END, f"Configured SUTs: {', '.join(sut_ips)}\n")
    error_match_text.insert(tk.END, f"Total SUTs: {len(sut_ips)}\n")
    error_match_text.insert(tk.END, "=" * 50 + "\n\n")
    error_match_text.insert(tk.END, "1. Double-click any command to execute on all SUTs\n")
    error_match_text.insert(tk.END, "2. Each SUT will get its own cmd window\n")
    error_match_text.insert(tk.END, "3. Startup command runs first, then selected command\n\n")

    # --- Theme status indicator ---
    if theme_name:
        theme_label = tk.Label(
            window, 
            text=f"Theme: {theme_name} | Multi-SUT Mode", 
            bg="#1E1E1E", 
            fg="#00E676",
            font=("Consolas", 8)
        )
        theme_label.pack(side=tk.BOTTOM, anchor='se', padx=5, pady=2)

    def on_close_helper():
        window.destroy()

    window.protocol("WM_DELETE_WINDOW", on_close_helper)
    window.mainloop()


def load_all_sut_configs(sut_ips):
    """
    Load configurations for all specified SUT IPs.
    Returns a dictionary with sut_ip as key and config as value.
    """
    configs = {}
    
    for sut_ip in sut_ips:
        debug_print(f"Loading configuration for {sut_ip}")
        config = read_config_m(sut_ip)
        
        if isinstance(config, dict):
            configs[sut_ip] = config
            debug_print(f"Successfully loaded config for {sut_ip}")
            debug_print(f"  Config keys: {list(config.keys())}")
            debug_print(f"  Sample values: SUT_IP={config.get('SUT_IP', 'N/A')}, SLOT_ID={config.get('SLOT_ID', 'N/A')}")
        else:
            debug_print(f"Failed to load configuration for {sut_ip}: {config}")
            # Create a minimal config with just the SUT_IP
            configs[sut_ip] = {"SUT_IP": sut_ip}
            debug_print(f"Created minimal config for {sut_ip}")
    
    return configs


def main():
    args = parse_args()

    # Set up logging redirection early
    log_redirector, log_file_path_shell = setup_shell_logging_redirect(args.console_type, args.sut_ip)

    # Force controller type behavior for multi-SUT
    if args.console_type.lower() == "controller":
        args.local = True

    sound_enabled = args.sound.lower() == "true"
    theme_name = args.theme
    global log_folder, sut_configs

    num_monitors = get_monitor_count()
    debug_print(f"Detected {num_monitors} monitors")

    helper_counter = get_helper_window_counter()
    selected_monitor = (helper_counter - 1) % num_monitors if num_monitors > 0 else args.monitor_id
    debug_print(f"Automatically selected monitor {selected_monitor} (helper window #{helper_counter})")

    # Create log folder for controller mode
    if args.console_type.lower() == "controller":
        start_time = time.localtime()
        timestamp = time.strftime("%Y%m%d_%H%M%S", start_time)
        logs_parent = os.path.join(os.path.dirname(os.getcwd()), "logs")
        os.makedirs(logs_parent, exist_ok=True)
        sut_list = "_".join(args.sut_ip)
        log_folder = os.path.join(logs_parent, f"McConsole_Multi_{sut_list}_{timestamp}")
        os.makedirs(log_folder, exist_ok=True)

    # Load configurations for all SUTs
    debug_print(f"Loading configurations for SUTs: {args.sut_ip}")
    sut_configs = load_all_sut_configs(args.sut_ip)
    
    # Check if we have at least one valid config
    valid_configs = [sut for sut, config in sut_configs.items() if len(config) > 1]  # More than just SUT_IP
    
    if not valid_configs:
        debug_print("No valid configurations found for any SUT. Please check your SUT configuration files.")
        return
    
    debug_print(f"Successfully loaded configurations for: {valid_configs}")
    
    # Check for CSV file existence (use first SUT's config)
    helper_folder = os.path.join(os.path.dirname(__file__), "helper")
    first_sut = args.sut_ip[0]
    first_config = sut_configs.get(first_sut, {})
    csv_prefix = f"{first_config.get('SUT_SKU')}_" if first_config.get("SUT_SKU") else ""
    csv_file = os.path.join(helper_folder, f"{csv_prefix}CONTROLLER_commands.csv")
    csv_exists = os.path.exists(csv_file)
    
    if not csv_exists:
        debug_print(f"CSV file not found: {csv_file}")
        debug_print("Helper window will not be opened.")
        return

    debug_print(f"Found CSV file: {csv_file}")
    
    try:
        # Launch the multi-SUT helper window
        open_helper_window_multi(
            args.console_type,
            sut_configs,
            args.sut_ip,
            sound_enabled,
            monitor_id=selected_monitor,
            theme_name=theme_name
        )
    except KeyboardInterrupt:
        print("Program interrupted. Exiting.", flush=True)


if __name__ == "__main__":
    main()