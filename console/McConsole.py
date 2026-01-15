

from collections import defaultdict
import logging
import subprocess
import sys
import socket
import shutil
import time
import tkinter as tk
# from tkinter import ttk
from pathlib import Path
import os
from tkinter import simpledialog
import psutil
import json
import glob
import threading
import queue
from concurrent.futures import ThreadPoolExecutor, as_completed
from tkinter import messagebox
import paramiko
import random
import urllib.request
import re
import pygame
import webbrowser
import ssl
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from datetime import datetime
from ttkbootstrap.scrolled import ScrolledFrame
# On Windows, optionally auto-switch code page:
if os.name == 'nt':
    os.system('chcp 65001 > nul')

# Now reconfigure Python's streams to UTF-8
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding='utf-8')
    sys.stderr.reconfigure(encoding='utf-8')

def check_and_print_env_var_instructions():
    """Check if credential environment variables are set and display status for each."""
    
    # List of possible environment variables to check
    env_vars_to_check = [
        'RM_GUESS', 'RM_WHAT',
        'SUT_GUESS', 'SUT_WHAT', 
        'SOC_GUESS', 'SOC_WHAT',
        'BMC_GUESS', 'BMC_WHAT'
    ]
    
    print("🔍 Credential Environment Variables Status:")
    for var in env_vars_to_check:
        if os.environ.get(var):
            print(f"   ✓ {var}=SET")
        else:
            print(f"   ✗ {var}=NOT SET")
    
    print("\n📋 Priority Logic (McConsole credential resolution):")
    print("   1️⃣  settings.{SUT_name}.json file values (HIGHEST PRIORITY)")
    print("   2️⃣  Environment variables (if config file field is empty)")
    print("   3️⃣  Empty/Missing (if neither is set)")
    print("\n💡 Recommendation:")
    print("   • Put credentials in config files for SUT-specific values and convenience, also integration with McController")
    print("   • Use environment variables for security")
    print("   • Environment variables are IGNORED if config file has values")
    
    # Show helpful commands
    print("\n💡 Commands:")
    print("   • Check variables: set | findstr \"GUESS WHAT\"")
    print("   • Set variables: set SUT_GUESS=username & set SUT_WHAT=password")
    
    # ADD THIS IMPORTANT SECTION:
    print("\n⚠️  IMPORTANT: After setting environment variables, you MUST run McConsole")
    print("   from the SAME CMD window where you set the variables:")
    print("   • .\\mcconsole.bat  OR  python .\\mcconsole.py")
    print("   • Variables set with 'set' only exist in the current CMD session!")
    print()


def get_credential_from_env(service_type, credential_type):
    """
    Get credential from environment variables.
    
    Args:
        service_type: Service type like 'RM', 'SUT', 'SOC', 'BMC', etc.
        credential_type: Either 'GUESS' (username) or 'WHAT' (password)
    
    Returns:
        The credential value from environment variables or None if not found
    """
    # Use the same naming convention as config files: SERVICE_GUESS, SERVICE_WHAT
    env_var_name = f"{service_type}_{credential_type}"
    return os.environ.get(env_var_name)

def setup_logging_redirect():
    """Redirect all print() statements to log file while keeping console open."""
    
    # Create logs directory
    log_dir = Path(r"C:\mcqueen\logs\Mcconsole_log")
    log_dir.mkdir(parents=True, exist_ok=True)
    
    # Generate timestamp for log filename
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    log_file = log_dir / f"McConsole_debug_{timestamp}.log"
    
    # Show user where logs are going before redirecting
    print(f"🚀 McConsole Starting...")
    print(f"📁 Debug logs will be saved to:")
    print(f"   {log_dir}")
    print(f"📄 Current session log: McConsole_debug_{timestamp}.log")
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

# Call logging setup before any other code that might print
# But check environment variables first so they print to console
check_and_print_env_var_instructions()
log_redirector, log_file_path = setup_logging_redirect()


logging.getLogger("paramiko.transport").setLevel(logging.WARNING)

NO_WINDOW = subprocess.CREATE_NO_WINDOW

try:
    from PIL import Image, ImageTk
except ImportError:
    print("Pillow is required to display JPEG images. Install with 'pip install pillow'.")
    sys.exit(1)

# importing new node window manager
from new_node_window import NodeWindowManager

# importing led status manager
from led_status_manager import LEDStatusManager

#from led_status_checker import check_led_status

#python_exe = str(Path(__file__).parent.parent / "python" / "python.exe")

# Pick up the distributed venv under C:\mcqueen\python\.venv
base = Path(r"C:\mcqueen\python")  # adjust if your layout differs
#venv_python = base / ".venv" / "Scripts" / "python.exe"
venv_python = base / "python.exe"

if not venv_python.exists():
    print(f"Error: venv python not found at {venv_python}")
    print('Please ensure that you have extracted the McConsole ZIP package under "C:\\mcqueen" and use 7-Zip to extract it.')
    input("Press Enter to exit...")
    sys.exit(1)

python_exe = str(venv_python)



if not Path(python_exe).exists():
    print(f"Error: Python executable not found at {python_exe}")
    sys.exit(1)


def initialize_mobaxterm():
    mx_dir    = Path(__file__).parent.parent / "bins" / "mx"
    slash_dir = mx_dir / "slash"
    mobax_exe = mx_dir / "Mobaxterm.exe"
    mobax_ini = mx_dir / "mobaxterm.ini"

    need_init = (not slash_dir.exists()) or (slash_dir.is_dir() and not any(slash_dir.iterdir()))
    if need_init and mobax_exe.exists() and mobax_ini.exists():
        print(">> Initializing portable MobaXterm (slash/ missing or empty)…")
        try:
            subprocess.Popen(
                [str(mobax_exe), "-i", str(mobax_ini)],
                creationflags=NO_WINDOW       # ← use your NO_WINDOW here
            )
        except Exception as e:
            print(f"Failed to initialize MobaXterm: {e}", file=sys.stderr)


def sync_mobaxterm_bookmarks():
    """
    Merge only the [Bookmarks] entries from the current user's roaming
    MobaXterm.ini into bins\\mx\\MobaXterm.ini, skipping duplicates and
    leaving every other section untouched. Falls back to 'mbcs' encoding.
    """
    user_ini = Path(os.environ["APPDATA"]) / "MobaXterm" / "MobaXterm.ini"
    bin_ini  = Path(__file__).parent.parent / "bins" / "mx" / "MobaXterm.ini"
    if not (user_ini.is_file() and bin_ini.is_file()):
        return

    def _read_lines(p: Path, keepends=False):
        # try UTF-8 first, then Windows ANSI ('mbcs')
        text = None
        for enc in ("utf-8", "mbcs"):
            try:
                return p.read_text(encoding=enc).splitlines(keepends=keepends)
            except UnicodeDecodeError:
                continue
        # as last resort, raw bytes→latin1
        return p.read_bytes().decode("latin1").splitlines(keepends=keepends)

    # --- read user [Bookmarks]
    bookmarks = {}
    for line in _read_lines(user_ini, keepends=False):
        s = line.strip()
        if s.upper() == "[BOOKMARKS]":
            in_bm = True
            continue
        if 'in_bm' in locals() and in_bm:
            if s.startswith("["):
                break
            if "=" in line:
                k, v = line.split("=", 1)
                bookmarks[k.strip()] = v.rstrip()

    # --- merge into bin ini ---
    bin_lines = _read_lines(bin_ini, keepends=True)
    out, in_bm, seen = [], False, set()

    for ln in bin_lines:
        s = ln.strip()
        if s.upper() == "[BOOKMARKS]":
            in_bm = True
            out.append(ln)
            continue

        if in_bm:
            if s.startswith("["):
                # inject any new entries before leaving section
                for k, v in bookmarks.items():
                    if k not in seen:
                        out.append(f"{k}={v}\n")
                in_bm = False
                out.append(ln)
            else:
                out.append(ln)
                if "=" in ln:
                    seen.add(ln.split("=", 1)[0].strip())
            continue

        out.append(ln)

    # if file ended inside [Bookmarks], flush remaining
    if in_bm:
        for k, v in bookmarks.items():
            if k not in seen:
                out.append(f"{k}={v}\n")

    # write back using ANSI to match original
    bin_ini.write_text("".join(out), encoding="mbcs")


def _relay_output(proc, prefix=""):
    """Read the child‐proc's stdout and print each line with a prefix."""
    for line in proc.stdout:
        print(f"{prefix}{line}", end="", flush=True)
    proc.stdout.close()


class SessionManager:
    def __init__(self):
        print("\n=== Initializing Console Session Manager ===")
        
        try:
            pygame.mixer.init()
            self.audio_enabled = True
        except pygame.error as e:
            print("Warning: Audio device initialization failed. Running without audio. Error:", e)
            self.audio_enabled = False
            
        sync_mobaxterm_bookmarks()    

        # self.root = tk.Tk()
        # self.root = ttk.Window(themename="darkly")


        saved_ui_theme = self.load_ui_theme_setting()
        self.root = ttk.Window(themename=saved_ui_theme)

        # --- CHANGE 1: Define the icon path as an instance variable ---
        self.icon_path = Path(__file__).parent / "mc.ico"
        
        self.setup_window_icon()

        self.child_processes = []
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)
        
        # ---------------------------------------------------------------------
        # Read local version from version.txt
        # ---------------------------------------------------------------------
        version_file = Path(__file__).parent / "version.txt"
        self.local_version = "0.0"
        if version_file.exists():
            try:
                with open(version_file, 'r') as vf:
                    self.local_version = vf.read().strip()
            except Exception as e:
                print("Error reading version file:", e)

        # ---------------------------------------------------------------------
        # Set window title and background
        # ---------------------------------------------------------------------
        self.root.title(f"McConsole - {self.local_version}")
        # self.root.configure(bg='#1E1E1E')
        

        self.style = self.root.style
        

        self.winscp_path = str(Path(__file__).parent.parent / "bins" / "WinSCP" / "WinSCP.exe")
        self.winscp_available = Path(self.winscp_path).exists()

        self.status_queue = queue.Queue()
        self.led_status_queue = queue.Queue()
        self.led_status_manager = LEDStatusManager(self.led_status_queue)

        # Added jumpbox configuration loading
        self.jumpbox_config_file = Path(__file__).parent / "jumpbox_config.json"
        self.jumpbox_config = self.load_jumpbox_config()


        print("Loading SUT configurations data from the settings files...")
        self.sut_configs = self.load_sut_configs()

        self.checkbuttons = {}
        self.manual_checkbox_state = defaultdict(dict)      # new dict to manage checkbox state
        self.status_labels = {}
        self.led_status_labels = {}
        self.folder_frames = {}


        self.folder_label_vars = {}
        # --- Multi-select nodes state ---
        self.selected_nodes = set()  # Set of (folder, identifier)
        self.selected_node_frames = {}  # (folder, identifier): row_frame
        self.selection_bar = None
        self.selection_bar_nodes = []
        self.ip_label_widgets = {}
        

        self.folder_expanded = {}
        self.expand_state_file = Path(__file__).parent / "expand_state.json"
        self.folder_expanded_state = {}
        self.load_expand_state()

        # Initialize new node window with a refresh callback
        self.node_window_manager = NodeWindowManager(self.root, self.refresh_folder, icon_path=self.icon_path)

        print("Creating UI for console window...")
        self.create_widgets()

        self.notepadpp_path = self.find_notepadpp()

        print("Starting background checker for connection...")
        self.start_background_checker()

        print("Starting background checker for LED status...")
        self.start_led_status_checker()  # New function to start LED status checks

        self.bgm_playlist = []
        self.bgm_index = 0
        self.bgm_active = False

        print("Initialization is complete now!\n")

        self.root.protocol("WM_DELETE_WINDOW", self.on_close)
 

    def _validate_widget_references(self):
        """
        Clean up stale widget references that may have been destroyed by refresh_folder().
        This is a defensive method that ensures we don't operate on destroyed widgets.
        """
        # Clean up selected_node_frames
        stale_keys = []
        for key, frame in self.selected_node_frames.items():
            try:
                frame.winfo_exists()
            except tk.TclError:
                # Widget is destroyed, mark for removal
                stale_keys.append(key)
        
        for key in stale_keys:
            del self.selected_node_frames[key]
            # Also remove from selected_nodes if it exists there
            self.selected_nodes.discard(key)
        
        # Clean up ip_label_widgets
        stale_keys = []
        for key, label in self.ip_label_widgets.items():
            try:
                label.winfo_exists()
            except tk.TclError:
                stale_keys.append(key)
        
        for key in stale_keys:
            del self.ip_label_widgets[key]

    def _safe_configure_widget(self, widget, **kwargs):
        """
        Safely configure a widget, handling the case where it might be destroyed.
        Returns True if successful, False if widget is destroyed.
        """
        try:
            if widget.winfo_exists():
                widget.configure(**kwargs)
                return True
        except (tk.TclError, AttributeError):
            pass
        return False

    # Replace the existing _on_node_row_click method
    def _on_node_row_click(self, event, folder, identifier, row_frame):
        """Enhanced node row click handler with stale reference protection"""
        # First, validate and clean up any stale references
        self._validate_widget_references()
        
        # Check if the clicked frame is still valid
        try:
            if not row_frame.winfo_exists():
                print(f"Warning: Clicked on stale widget reference for {identifier}")
                return
        except tk.TclError:
            print(f"Warning: Clicked on destroyed widget for {identifier}")
            return
        
        # Support Ctrl+Click for multi-select
        ctrl = (event.state & 0x0004) != 0  # Windows Ctrl mask
        key = (folder, identifier)
        
        if ctrl:
            if key in self.selected_nodes:
                self.selected_nodes.remove(key)
            else:
                self.selected_nodes.add(key)
        else:
            # Single click: clear all, select only this
            self.selected_nodes = {key}
        
        self._update_node_selection_visuals()
        self._show_selection_bar_if_needed()



    def _show_selection_bar_if_needed(self):
        """Show modal window for multi-node selection with reference validation"""
        # Clean up any stale references first
        self._validate_widget_references()
        
        if not self.selected_nodes:
            if hasattr(self, 'selection_modal') and self.selection_modal:
                self._close_modal()
            return

        # Check if modal exists AND is still valid
        modal_exists = (hasattr(self, 'selection_modal') and 
                    self.selection_modal and 
                    self._is_modal_valid())
        
        if modal_exists:
            self._update_modal_content()
        else:
            self._create_selection_modal()

    def _is_modal_valid(self):
        """Check if the modal window is still valid."""
        try:
            # Return True if the widget exists, False otherwise.
            return self.selection_modal and self.selection_modal.winfo_exists()
        except tk.TclError:
            # This can happen if the widget is destroyed but reference still exists
            return False

    def _create_selection_modal(self):
        """Create modal window for multi-node selection"""
        # Create modal window
        self.selection_modal = ttk.Toplevel(self.root)
        self.selection_modal.title("Multi-SUT Console")
        self.selection_modal.geometry("440x380")
        self.selection_modal.minsize(440, 380)

        # Set modal properties - NO grab_set() to allow main window interaction
        # self.selection_modal.transient(self.root)
        self.selection_modal.resizable(True, True)
    
        
        # Set window icon if available
        if hasattr(self, 'icon_path') and self.icon_path.exists():
            try:
                self.selection_modal.iconbitmap(str(self.icon_path))
            except:
                pass
        
        # Position at bottom center of parent window
        # self._position_modal_bottom_center()
        self.selection_modal.after(50, self._position_modal_bottom_center)

        # Create content
        self._create_modal_content()
        
        # Bind window close events
        self.selection_modal.bind('<Escape>', lambda e: self._close_modal())
        self.selection_modal.protocol("WM_DELETE_WINDOW", self._close_modal)
        
        # Keep modal on top but allow interaction with main window
        self.selection_modal.attributes('-topmost', True)
        self.selection_modal.focus_set()


    def _position_modal_bottom_center(self):
        self.selection_modal.update_idletasks()
        self.root.update_idletasks()  # ensure root is placed

        modal_width = self.selection_modal.winfo_width()
        modal_height = self.selection_modal.winfo_height()

        parent_x = self.root.winfo_x()
        parent_y = self.root.winfo_y()
        parent_width = self.root.winfo_width()
        parent_height = self.root.winfo_height()

        # fallback if root hasn’t reported coords yet
        if parent_x <= 0 and parent_y <= 0:
            parent_x = (self.root.winfo_screenwidth() // 2) - (parent_width // 2)
            parent_y = (self.root.winfo_screenheight() // 2) - (parent_height // 2)

        x = parent_x + (parent_width - modal_width) // 2
        y = parent_y + parent_height - modal_height - 50
        self.selection_modal.geometry(f"{modal_width}x{modal_height}+{x}+{y}")


       
    def _create_modal_content(self):
        """Create the content inside the modal - styled like Auto Refresh modal"""
        # Main container with padding
        main_frame = ttk.Frame(self.selection_modal, padding=20)
        main_frame.pack(fill='both', expand=True)
        
        # Title header - similar to "Auto Refresh Configuration"
        title_label = ttk.Label(main_frame, text="Multi-SUT Console Selection", 
                            font=("Segoe UI", 14, "bold"),
                            foreground=self.root.style.colors.get('info') or '#8E44AD')
        title_label.pack(anchor='w', pady=(0, 20))
        
        # Selection summary
        num_selected = len(self.selected_nodes)
        summary_text = f"{num_selected} node{'s' if num_selected != 1 else ''} selected"
        summary_label = ttk.Label(main_frame, text=summary_text, 
                                font=("Segoe UI", 10))
        summary_label.pack(anchor='w', pady=(0, 15))
        
        # Bottom buttons frame - MOVE THIS BEFORE nodes section and pack at bottom
        self.button_frame = ttk.Frame(main_frame)  # Make it instance variable
        self.button_frame.pack(side='bottom', fill='x', pady=(10, 0))
        
        self.button_frame.columnconfigure(0, weight=1)
        self.button_frame.columnconfigure(1, weight=1)
        
        # Styled buttons
        clear_btn = ttk.Button(self.button_frame, text="Clear All", 
                            bootstyle="secondary",
                            command=self._clear_all_selections)
        clear_btn.grid(row=0, column=0, sticky="ew", padx=(0, 10))
        
        open_btn = ttk.Button(self.button_frame, text="Open Multi-SUT Console", 
                            bootstyle="success",
                            command=self._launch_multi_sut_console)
        open_btn.grid(row=0, column=1, sticky="ew")
        
        # Nodes container with border - pack AFTER buttons with remaining space
        nodes_section = ttk.LabelFrame(main_frame, text="Selected Nodes", padding=15)
        nodes_section.pack(fill='both', expand=True, pady=(0, 10))  # Reduced bottom padding
        
        # Scrollable nodes container with fixed height
        self.nodes_container = ScrolledFrame(nodes_section, autohide=True, height=120)  # Reduced height
        self.nodes_container.pack(fill='both', expand=True)
        
        # Enhanced mouse wheel binding
        self._bind_modal_mousewheel()
        
        # Update content
        self._update_modal_content()


    def _update_modal_content(self):
        """Update the modal content when selections change"""
        # Case 1: No modal exists → create a fresh one
        if not hasattr(self, 'selection_modal') or not self.selection_modal:
            self._create_selection_modal()
            return

        # Case 2: Modal exists but is destroyed → recreate it
        try:
            if not self.selection_modal.winfo_exists():
                self._create_selection_modal()
                return
        except:
            self._create_selection_modal()
            return

        # Case 3: Nodes container missing or destroyed → recreate content
        if not hasattr(self, 'nodes_container') or not self.nodes_container:
            self._create_modal_content()
            return

        try:
            self.nodes_container.winfo_exists()
        except:
            self._create_modal_content()
            return

        # --- Now safe to clear and repopulate ---
        for child in self.nodes_container.winfo_children():
            child.destroy()

        # Repopulate nodes
        inner_frame = ttk.Frame(self.nodes_container)
        inner_frame.pack(fill='x', padx=5)

        for key in sorted(list(self.selected_nodes), key=lambda k: k[1]):
            folder, identifier = key
            node_row = ttk.Frame(inner_frame, padding=(1, 2))
            node_row.pack(fill='x', pady=1)

            remove_btn = ttk.Button(node_row, text="✕",
                                    bootstyle="danger-outline",
                                    padding=(2, 0),
                                    command=lambda k=key: self._remove_selected_node(k))
            remove_btn.pack(side='left', padx=(0, 8))

            # remove_btn = ttk.Label(node_row, text="✕", foreground="red", cursor="hand2")
            # remove_btn.bind("<Button-1>", lambda e, k=key: self._remove_selected_node(k))
            # remove_btn.pack(side='left', padx=(0, 8))


            node_label = ttk.Label(node_row, text=identifier, anchor='w',
                                font=("Segoe UI", 9))
            node_label.pack(side='left', fill='x', expand=True)

        # Update summary
        try:
            for child in self.selection_modal.winfo_children():
                for subchild in child.winfo_children():
                    if isinstance(subchild, ttk.Label) and "selected" in subchild.cget("text"):
                        num_selected = len(self.selected_nodes)
                        subchild.configure(text=f"{num_selected} node{'s' if num_selected != 1 else ''} selected")
                        break
        except:
            pass


    def _bind_modal_mousewheel(self):
        """Enhanced mouse wheel binding for modal"""
        def _on_mousewheel(event):
            try:
                # Get the canvas from ScrolledFrame
                for child in self.nodes_container.winfo_children():
                    if hasattr(child, 'yview'):
                        child.yview_scroll(int(-1 * (event.delta / 120)), "units")
                        break
            except:
                pass
        
        # Bind to modal and container
        self.selection_modal.bind("<MouseWheel>", _on_mousewheel)
        self.nodes_container.bind("<MouseWheel>", _on_mousewheel)


    def _close_modal(self):
        """Close the selection modal, clear selections, and perform cleanup."""
        # First, clear the selection data and update the main UI visuals.
        self.selected_nodes.clear()
        
        # Safely update visuals only if we have valid references
        try:
            self._update_node_selection_visuals()
        except tk.TclError:
            # If visuals can't be updated due to destroyed widgets, just clear the references
            self.selected_node_frames.clear()
            self.ip_label_widgets.clear()

        # Now, handle the destruction of the modal window if it exists.
        if hasattr(self, 'selection_modal') and self.selection_modal:
            try:
                if self.selection_modal.winfo_exists():
                    self.selection_modal.destroy()
            except tk.TclError:  
                pass  
            
            # Clean up all references
            self.selection_modal = None
            self.nodes_container = None
            self.button_frame = None

    def _remove_selected_node(self, key):
        """Remove a node from selection and update modal"""
        self.selected_nodes.discard(key)
        self._update_node_selection_visuals()
        
        # Update modal or close if empty
        if not self.selected_nodes:
            self._close_modal()
        else:
            self._update_modal_content()

    def _launch_multi_sut_console(self):
        """Launch multi-SUT console and close modal"""
        sut_ids = [identifier for (folder, identifier) in self.selected_nodes]
        if not sut_ids:
            return
        
        cmd = [sys.executable, "./shell_con.py", "--console_type", "controller", "--sut_ip"] + sut_ids
        try:
            subprocess.Popen(cmd, cwd=os.path.dirname(__file__))
            # Close modal after successful launch
            self._close_modal()
        except Exception as e:
            print(f"Error launching multi-SUT console: {e}")

    def _clear_all_selections(self):
        """Clear all selected nodes and close modal"""
        self.selected_nodes.clear()
        self._update_node_selection_visuals()
        self._close_modal()


    def _update_node_selection_visuals(self):
        """State-driven visual update - render current selection onto existing widgets"""
        # Iterate through all current widget references
        for key, ip_label in self.ip_label_widgets.items():
            try:
                # Check state: is this node selected?
                if key in self.selected_nodes:
                    # Apply selected style
                    self._safe_configure_widget(ip_label, foreground=self.root.style.colors.primary)
                else:
                    # Apply default style  
                    self._safe_configure_widget(ip_label, foreground="")
            except:
                # Widget invalid - skip it, will be cleaned up naturally
                continue

    # Enhanced cleanup in on_close method
    def on_close(self):
        """Enhanced cleanup when closing main window"""
        # Close selection modal if open
        if hasattr(self, 'selection_modal') and self.selection_modal:
            try:
                self.selection_modal.destroy()
            except:
                pass

    def _on_folder_refreshed(self):
        """Simplified cleanup - no state management needed"""
        # Just close modal, state is preserved elsewhere
        if hasattr(self, 'selection_modal') and self.selection_modal:
            try:
                if self.selection_modal.winfo_exists():
                    self._close_modal()
            except tk.TclError:
                pass

    def setup_window_icon(self):
        """Set up window icon using the stored instance path."""
        
        # Use the instance variable defined in __init__
        ico_path = self.icon_path
        
        print(f"Looking for icon file: {ico_path} - {'EXISTS' if ico_path.exists() else 'NOT FOUND'}")
        
        if ico_path.exists():
            try:
                # Apply to the main root window
                self.root.iconbitmap(str(ico_path))
                print(f"Successfully set icon on root window: {ico_path}")
                return True
            except Exception as e:
                print(f"Failed to set icon: {e}")
        else:
            print(f"Icon file not found: {ico_path}")
        
        return False

    def track_process(self, proc: subprocess.Popen):
        """Keep proc in our list while it's alive, then remove it."""
        self.child_processes.append(proc)
        def reaper():
            proc.wait()
            if proc in self.child_processes:
                self.child_processes.remove(proc)
        threading.Thread(target=reaper, daemon=True).start()

    # --- Added for remote version check ---
    def parse_version(self, ver_str):
        """
        Parse a version string (ignoring any leading non-digit characters)
        and return a tuple of integers.
        Example: "v3.2" -> (3, 2), "1.9.0" -> (1, 9, 0)
        """
        match = re.search(r'(\d+(?:\.\d+)+)', ver_str)
        if match:
            return tuple(int(part) for part in match.group(1).split('.'))
        return (0,)

    def on_close(self):
        # 1. Save UI state
        try:
            self.save_expand_state()
        except Exception as e:
            print("Warning saving state on close:", e)
            
            
                # 3. Kill only the MobaXterm processes this script spawned
        parent = psutil.Process(os.getpid())
        for child in parent.children(recursive=True):
            try:
                if child.name().lower().startswith('mobaxterm'):
                    child.terminate()
                    child.wait(timeout=2)
            except (psutil.NoSuchProcess, psutil.TimeoutExpired, Exception):
                try:
                    child.kill()
                except (psutil.NoSuchProcess, Exception):
                    pass    

        # 2. Terminate all explicitly tracked subprocesses
        for proc in list(self.child_processes):
            try:
                if proc.poll() is None:
                    proc.terminate()
                    # give it a moment to die gracefully
                    proc.wait(timeout=2)
            except (psutil.NoSuchProcess, subprocess.TimeoutExpired, Exception):
                # either it's already gone, or didn't terminate in time
                try:
                    proc.kill()
                except (psutil.NoSuchProcess, Exception):
                    pass

        ##3. Kill MobaXterm.exe and every child it spawned
        # for p in psutil.process_iter(['name']):
            # name = (p.info.get('name') or "").lower()
            # if name.startswith('mobaxterm'):
              ##  include the parent and _all_ descendants
                # to_kill = [p] + p.children(recursive=True)
                # for proc in to_kill:
                    # try:
                        # proc.terminate()
                        # proc.wait(timeout=2)
                    # except (psutil.NoSuchProcess, psutil.TimeoutExpired, Exception):
                        # try:
                            # proc.kill()
                        # except (psutil.NoSuchProcess, Exception):
                            # pass




        # 4. Also kill any other descendants of this process
        parent = psutil.Process(os.getpid())
        for child in parent.children(recursive=True):
            try:
                child.terminate()
                child.wait(timeout=2)
            except (psutil.NoSuchProcess, psutil.TimeoutExpired, Exception):
                try:
                    child.kill()
                except (psutil.NoSuchProcess, Exception):
                    pass

        # 5. Wait a moment, force-kill if needed
        gone, alive = psutil.wait_procs(parent.children(recursive=True), timeout=3)
        for p in alive:
            try:
                p.kill()
            except (psutil.NoSuchProcess, Exception):
                pass

        # 6. Finally destroy the GUI and exit immediately
        try:
            self.root.destroy()
        except:
            pass
        os._exit(0)





    def add_to_favorites(self, file_path):
        """
        Move a settings.json file to the Favorites folder and update metadata.

        Args:
            file_path (_type_): Path to the settings.json file to move
        """
        favorites_dir = Path(__file__).parent.parent / "sut" / "Favorites"
        favorites_dir.mkdir(parents=True, exist_ok=True)

        metadata_file = favorites_dir / "favorites_metadata.json"

        if metadata_file.exists():
            with open(metadata_file, "r") as f:
                try:
                    metadata = json.load(f)
                except json.JSONDecodeError:
                    metadata = {}
        else:
            metadata = {}

        file_path_obj = Path(file_path)

        existing_favorites = list(favorites_dir.glob('*'))
        for fav_file in existing_favorites:
            try:
                with open(file_path, 'r') as original, open(fav_file, 'r') as favorite:
                    if original.read() == favorite.read():
                        messagebox.showinfo("Duplicate Favorite",
                                            f"The configuration for {file_path_obj.name} is already in Favorites.")
                        return
            except Exception as e:
                print(f"Error comparing files: {e}")

        try:
            metadata[file_path_obj.name] = str(file_path_obj.parent)

            destination = favorites_dir / file_path_obj.name
            shutil.move(str(file_path), str(destination))

            with open(metadata_file, "w") as f:
                json.dump(metadata, f, indent=4)

            self.sut_configs = self.load_sut_configs()
            self.update_folder_content('Favorites')

            try:
                relative_folder = Path(file_path_obj.parent).relative_to(Path(__file__).parent.parent / "sut")
                folder_key = str(relative_folder) if str(relative_folder) != "." else "Root"
                self.update_folder_content(folder_key)
            except Exception as e:
                print(f"Error updating original folder UI: {e}")

            print(f"Moved {file_path_obj.name} to Favorites and updated metadata.")
        except Exception as e:
            print(f"Error moving file to Favorites: {e}")

    def remove_from_favorites(self, file_path):
        """
        Remove a file from the Favorites folder and restore it to its original location.

        Args:
            file_path (_type_): Path to the settings.json file to remove from Favorites
        """
        try:
            file_path_obj = Path(file_path)
            favorites_dir = file_path_obj.parent

            if 'Favorites' not in str(favorites_dir):
                print("File is not in Favorites directory")
                return

            metadata_file = favorites_dir / "favorites_metadata.json"
            metadata = {}

            if metadata_file.exists():
                try:
                    with open(metadata_file, "r") as f:
                        metadata = json.load(f)
                except json.JSONDecodeError:
                    pass  # Leave metadata as empty {}

            original_folder = metadata.get(file_path_obj.name)
            moved_to_root = False

            if not original_folder:
                messagebox.showwarning(
                    "Missing Metadata",
                    f"Original location for {file_path_obj.name} not found. Moving it to the Root folder."
                )
                original_folder = Path(__file__).parent.parent / "sut"
                moved_to_root = True

            Path(original_folder).mkdir(parents=True, exist_ok=True)
            destination = Path(original_folder) / file_path_obj.name
            shutil.move(str(file_path), str(destination))

            # Clean up metadata if applicable
            if file_path_obj.name in metadata:
                metadata.pop(file_path_obj.name, None)
                with open(metadata_file, "w") as f:
                    json.dump(metadata, f, indent=4)

            # Reload configs after the file has been moved
            self.sut_configs = self.load_sut_configs()

            # Refresh Favorites folder UI
            self.update_folder_content("Favorites")

            # Determine the destination folder key for UI (Root or relative path)
            try:
                relative_folder = Path(original_folder).relative_to(Path(__file__).parent.parent / "sut")
                folder_key = str(relative_folder) if str(relative_folder) != "." else "Root"
            except Exception as e:
                print(f"Error resolving folder key: {e}")
                folder_key = "Root"

            for key in {"Favorites", folder_key}:
                if key in self.folder_frames:
                    try:
                        self.update_folder_content(key)
                        self.refresh_folder(key)
                        self._on_folder_refreshed()
                    except Exception as e:
                        print(f"Error updating UI for folder {key}: {e}")

            print(f"Removed {file_path_obj.name} from Favorites and restored to {folder_key}")

        except Exception as e:
            print(f"Error removing from Favorites: {e}")


    def duplicate_node(self, file_path, folder=None):
        """
        Duplicates a node within the same folder and updates the UI.
        
        Args:
            file_path: Path to the configuration file of the node to duplicate
            folder: Folder containing the node (optional, will be determined from file_path if not provided)
        """
        try:
            sut_dir = Path(__file__).parent.parent / "sut"
            file_path_obj = Path(file_path)
            
            try:
                relative_folder = file_path_obj.parent.relative_to(sut_dir)
                actual_folder = str(relative_folder) if str(relative_folder) != "." else "Root"
            except ValueError:
                print(f"Error: File path {file_path} not within sut directory")
                raise ValueError("File path not within sut directory")
            
            print(f"File path: {file_path}")
            print(f"Original folder parameter: {folder}")
            print(f"Determined folder: '{actual_folder}'")
            
            folder = actual_folder
            
            with open(file_path, 'r') as f:
                original_config = json.load(f)
            
            # Get the original identifier
            original_identifier = None
            normalized_file_path = os.path.normpath(file_path)
            
            # First check if the folder exists in sut_configs
            if folder not in self.sut_configs:
                print(f"Warning: Folder '{folder}' not found in sut_configs")
                print(f"Available folders: {list(self.sut_configs.keys())}")
                # Try to find the correct folder
                for potential_folder in self.sut_configs:
                    for identifier, config_data in self.sut_configs[potential_folder].items():
                        normalized_config_path = os.path.normpath(config_data['file_path'])
                        if normalized_config_path == normalized_file_path:
                            folder = potential_folder
                            original_identifier = identifier
                            print(f"Found matching file in folder: {folder}, identifier: {original_identifier}")
                            break
                    if original_identifier:
                        break
            else:
                # If folder exists, look for matching file path
                for identifier, config_data in self.sut_configs[folder].items():
                    normalized_config_path = os.path.normpath(config_data['file_path'])
                    if normalized_config_path == normalized_file_path:
                        original_identifier = identifier
                        break
                        
            if not original_identifier:
                print(f"Debug: Could not find original identifier for file path: {file_path}")
                print(f"Available paths in folder '{folder}':")
                if folder in self.sut_configs:
                    for identifier, config_data in self.sut_configs[folder].items():
                        print(f"  - {config_data['file_path']}")
                raise ValueError("Could not find original identifier for file path")
        
            # Create a new unique identifier with numbering
            new_identifier = self.generate_unique_identifier(original_identifier, folder)
            
            file_path_obj = Path(file_path)
            dir_path = file_path_obj.parent
            new_file_name = f"settings.{new_identifier}.json"
            new_file_path = str(dir_path / new_file_name)
        
            
            with open(new_file_path, 'w') as f:
                json.dump(original_config, f, indent=2)
            
            if folder not in self.sut_configs:
                self.sut_configs[folder] = {}
            
            config_data = {
                'file_path': new_file_path,
                'identifier': new_identifier
            }
            
            for service_type in [k for k in self.sut_configs[folder][original_identifier] 
                                if k not in ['file_path', 'identifier']]:
                config_data[service_type] = self.sut_configs[folder][original_identifier][service_type]
                
                # Also copy any associated GUESS and WHAT fields
                guess_key = f"{service_type}_GUESS"
                what_key = f"{service_type}_WHAT"
                
                if guess_key in self.sut_configs[folder][original_identifier]:
                    config_data[guess_key] = self.sut_configs[folder][original_identifier][guess_key]
                    
                if what_key in self.sut_configs[folder][original_identifier]:
                    config_data[what_key] = self.sut_configs[folder][original_identifier][what_key]
            
            # Add the new configuration to sut_configs
            self.sut_configs[folder][new_identifier] = config_data
            
            # 8. Update the UI to show the new node
            self.update_folder_content(folder)
            
            # print(f"Successfully duplicated node: {original_identifier} → {new_identifier}")
            print(f"Successfully duplicated node: {original_identifier} -> {new_identifier}")
            
        except Exception as e:
            print(f"Error duplicating node: {e}")
            messagebox.showerror("Duplication Error", f"Failed to duplicate node: {str(e)}")

    def generate_unique_identifier(self, original_identifier, folder):
        """
        Generates a unique identifier for a duplicated node by adding numbering.
        
        Args:
            original_identifier: The identifier of the original node
            folder: The folder where the node will be placed
        
        Returns:
            A unique identifier string with numbering
        """
        # Check if the identifier already contains numbering like "(1)"
        base_name = original_identifier
        current_num = 0
        
        match = re.search(r'\((\d+)\)$', original_identifier)
        if match:
            # Extract the base name and current number
            current_num = int(match.group(1))
            base_name = original_identifier[:match.start()].strip()
        
        # Generate candidate names with incremented numbers
        while True:
            current_num += 1
            candidate_name = f"{base_name} ({current_num})"
            
            # Check if this name is already used in the folder
            if candidate_name not in self.sut_configs.get(folder, {}):
                return candidate_name

    def save_duplicate_config(self, file_path, config_data):
        """Save a duplicated configuration to a new file."""
        try:
            save_data = config_data.copy()
            if 'file_path' in save_data:
                del save_data['file_path']
            
            os.makedirs(os.path.dirname(file_path), exist_ok=True)
            
            with open(file_path, 'w') as f:
                json.dump(save_data, f, indent=4)
            
            return True
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save duplicate configuration: {str(e)}")
            return False


    def delete_node(self, file_path):
        """
        Deletes a node from the UI and removes its file from the filesystem.
        
        Args:
            file_path: Path to the configuration file of the node to delete
        """
        try:
            # 1. Determine which folder the node belongs to
            sut_dir = Path(__file__).parent.parent / "sut"
            file_path_obj = Path(file_path)
            
            try:
                relative_folder = file_path_obj.parent.relative_to(sut_dir)
                folder = str(relative_folder) if str(relative_folder) != "." else "Root"
            except ValueError:
                print(f"Error: File path {file_path} not within sut directory")
                raise ValueError("File path not within sut directory")
            
            # 2. Find the identifier of the node
            identifier = None
            normalized_file_path = os.path.normpath(file_path)
            
            # Look through all folders to find the matching file path
            for potential_folder in self.sut_configs:
                for potential_identifier, config_data in self.sut_configs[potential_folder].items():
                    normalized_config_path = os.path.normpath(config_data['file_path'])
                    if normalized_config_path == normalized_file_path:
                        folder = potential_folder
                        identifier = potential_identifier
                        break
                if identifier:
                    break
                    
            if not identifier:
                print(f"Debug: Could not find identifier for file path: {file_path}")
                raise ValueError("Could not find identifier for file path")
                
            # 3. Ask for confirmation before deleting
            result = messagebox.askyesno(
                "Confirm Deletion", 
                f"Are you sure you want to delete node '{identifier}'?\nThis action cannot be undone.",
                icon='warning'
            )
            
            if not result:
                return  # User canceled the operation
                
            # 4. Remove the node from the in-memory configuration
            if folder in self.sut_configs and identifier in self.sut_configs[folder]:
                del self.sut_configs[folder][identifier]
                
                # Also clean up related data structures
                if file_path in self.checkbuttons:
                    del self.checkbuttons[file_path]
                if file_path in self.status_labels:
                    del self.status_labels[file_path]
                if file_path in self.led_status_labels:
                    del self.led_status_labels[file_path]
                
                # 5. Delete the file from disk
                os.remove(file_path)
                print(f"Successfully deleted node file: {file_path}")
                
                # 6. Update the UI to reflect the deletion
                self.update_folder_content(folder)
                
                # Show success message
                messagebox.showinfo("Node Deleted", f"Successfully deleted node '{identifier}'")
            else:
                raise ValueError(f"Node '{identifier}' not found in folder '{folder}'")
                
        except Exception as e:
            print(f"Error deleting node: {e}")
            messagebox.showerror("Deletion Error", f"Failed to delete node: {str(e)}")

    # new function to rename a node from UI
    def rename_node(self, file_path):
        """
        Rename a node (settings file) and update the UI
        
        Args:
            file_path: The path to the settings file to rename
        """
        # Get current file info
        file_path_obj = Path(file_path)
        old_filename = file_path_obj.name
        folder_path = file_path_obj.parent
        
        # Extract current identifier from filename
        current_identifier = file_path_obj.stem.split("settings.")[1]
        
        # Find which folder contains this node
        containing_folder = None
        for folder, nodes in self.sut_configs.items():
            if any(data.get('file_path') == file_path for _, data in nodes.items()):
                containing_folder = folder
                break
        
        if not containing_folder:
            messagebox.showerror("Error", "Could not find the containing folder for this node.")
            return
        
        # Show dialog to get new name
        new_identifier = simpledialog.askstring(
            "Rename Node", 
            "Enter new name for the node:", 
            initialvalue=current_identifier
        )

        # Check if user cancelled
        if new_identifier is None:
            return
        
        new_identifier = new_identifier.strip()
        
        # Validate new name
        if not new_identifier.strip():
            messagebox.showerror("Error", "Name cannot be empty.")
            return
        
        # Check for invalid characters in the new name
        invalid_chars = ['<', '>', ':', '"', '/', '\\', '|', '?', '*']
        if any(char in new_identifier for char in invalid_chars):
            messagebox.showerror(
                "Error", 
                f"Name cannot contain any of these characters: {''.join(invalid_chars)}"
            )
            return
        
        # Check for duplicates in the same folder
        new_filename = f"settings.{new_identifier}.json"
        new_path = folder_path / new_filename
        
        if new_path.exists():
            messagebox.showerror(
                "Error", 
                f"A node with the name '{new_identifier}' already exists in this folder."
            )
            return
        
        try:
            # Rename the file on disk
            file_path_obj.rename(new_path)
            
            # Update internal data structures
            
            # First, remove the old entry from sut_configs
            old_node_data = self.sut_configs[containing_folder].pop(current_identifier, None)
            
            # If the node was found, update its path and add it back under the new identifier
            if old_node_data:
                old_node_data['file_path'] = str(new_path)
                old_node_data['identifier'] = new_identifier
                self.sut_configs[containing_folder][new_identifier] = old_node_data
                
                # Update Favorites if needed
                if 'Favorites' in self.sut_configs and current_identifier in self.sut_configs['Favorites']:
                    fav_data = self.sut_configs['Favorites'].pop(current_identifier, None)
                    if fav_data:
                        fav_data['file_path'] = str(new_path)
                        fav_data['identifier'] = new_identifier
                        self.sut_configs['Favorites'][new_identifier] = fav_data
                
                # Update UI dictionaries
                if file_path in self.checkbuttons:
                    self.checkbuttons[str(new_path)] = self.checkbuttons.pop(file_path)
                
                if file_path in self.status_labels:
                    self.status_labels[str(new_path)] = self.status_labels.pop(file_path)
                    # Update identifier in status_labels
                    for service_type, info in self.status_labels[str(new_path)].items():
                        if isinstance(info, dict) and 'identifier' in info:
                            info['identifier'] = new_identifier
                
                if file_path in self.led_status_labels:
                    self.led_status_labels[str(new_path)] = self.led_status_labels.pop(file_path)
                
                # Refresh UI for both containing folder and Favorites if needed
                # self.refresh_folder(containing_folder)
                self.update_folder_content(containing_folder)
                if 'Favorites' in self.sut_configs and new_identifier in self.sut_configs['Favorites']:
                    # self.refresh_folder('Favorites')
                    self.update_folder_content('Favorites')
                    
                messagebox.showinfo("Success", f"Node renamed to '{new_identifier}' successfully.")
            else:
                # This should not happen unless there's a synchronization issue
                messagebox.showerror(
                    "Error", 
                    "Node was renamed on disk, but could not be updated in UI. Please refresh."
                )
                # self.refresh_all_folders()
        
        except PermissionError:
            messagebox.showerror(
                "Permission Error", 
                "Could not rename the file. Make sure it's not in use by another process."
            )
        except OSError as e:
            messagebox.showerror("Error", f"Could not rename the file: {str(e)}")

    def find_notepadpp(self):
        possible_notepad_paths = [
            r"C:\Program Files\Notepad++\notepad++.exe",
            r"C:\Program Files (x86)\Notepad++\notepad++.exe"
        ]
        for path in possible_notepad_paths:
            if os.path.exists(path):
                return path
        return None

    
    def load_sut_configs(self):
        sut_dir = Path(__file__).parent.parent / "sut"

        favorites_dir = sut_dir / "Favorites"
        favorites_dir.mkdir(parents=True, exist_ok=True)

        # Initializing including empty folders
        sut_configs = {}
        for folder_path in sut_dir.glob('**/'):
            if folder_path.is_dir():
                relative_folder = folder_path.relative_to(sut_dir)
                folder_key = str(relative_folder) if str(relative_folder) != "." else "Root"
                sut_configs.setdefault(folder_key, {}) 

        sut_files = glob.glob(str(sut_dir / "**" / "settings.*.json"), recursive=True)
        # sut_configs = {}
        for file_path in sut_files:
            try:
                file_path_obj = Path(file_path)
                relative_folder = file_path_obj.parent.relative_to(sut_dir)
                folder_key = str(relative_folder) if str(relative_folder) != "." else "Root"
                identifier = file_path_obj.stem.split("settings.")[1]

                with open(file_path, 'r') as f:
                    config = json.load(f)

                config_data = {
                    'file_path': file_path,
                    'identifier': identifier
                }

                service_types = set()
                for key in config:
                    if key.endswith('_IP') or key.endswith('_GUESS') or key.endswith('_WHAT'):
                        service_type = key.split('_')[0]
                        service_types.add(service_type)

                for service_type in service_types:
                    ip_key = f"{service_type}_IP"
                    service_ip = config.get(ip_key)

                    if service_ip is None and service_type in config:
                        service_ip = config.get(service_type)
                        config_data[service_type] = service_ip
                    else:
                        config_data[service_type] = service_ip

                    # Handle _GUESS (username) with environment variable fallback
                    guess_value = config.get(f"{service_type}_GUESS")
                    if not guess_value:  # If empty or None in config
                        guess_value = get_credential_from_env(service_type, 'GUESS')
                        if guess_value:
                            print(f"   Using environment variable for {service_type}_GUESS in {identifier}")
                    config_data[f"{service_type}_GUESS"] = guess_value

                    # Handle _WHAT (password) with environment variable fallback  
                    what_value = config.get(f"{service_type}_WHAT")
                    if not what_value:  # If empty or None in config
                        what_value = get_credential_from_env(service_type, 'WHAT')
                        if what_value:
                            print(f"   Using environment variable for {service_type}_WHAT in {identifier}")
                    config_data[f"{service_type}_WHAT"] = what_value

                # sut_configs.setdefault(folder_key, {})[identifier] = config_data
                sut_configs[folder_key][identifier] = config_data
            except Exception as e:
                print(f"Error loading {file_path}: {e}")
                continue

        return dict(sorted(sut_configs.items()))
        
            
    def update_scrollbar(self):
        """Show scrollbar only when content exceeds visible area and manage mouse wheel scrolling."""
        try:
            self.canvas.update_idletasks()  # Force geometry update
            canvas_height = self.canvas.winfo_height()
            content_height = self.canvas.bbox("all")[3] if self.canvas.bbox("all") else 0
            
            if content_height > canvas_height:
                # Content needs scrolling, display scrollbar and enable mouse wheel
                self.canvas.config(yscrollcommand=self.scrollbar.set)
                self.scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
                self.enable_mouse_wheel()
            else:
                # No scrolling needed, hide scrollbar and disable mouse wheel
                self.canvas.config(yscrollcommand=None)
                if self.scrollbar.winfo_ismapped():
                    self.scrollbar.pack_forget()
                self.disable_mouse_wheel()
        except (AttributeError, TypeError):
            # Handle case when canvas or content isn't fully initialized
            pass

    def enable_mouse_wheel(self):
        """Enable mouse wheel scrolling for the canvas."""
        self.canvas.bind_all("<MouseWheel>", self.on_mouse_wheel)
        
    def disable_mouse_wheel(self):
        """Disable mouse wheel scrolling for the canvas."""
        self.canvas.unbind_all("<MouseWheel>")
        
    def on_mouse_wheel(self, event):
        """Handle mouse wheel scrolling."""
        self.canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        
    
    def show_disclaimer(self):
        """Read and display Readme.txt in a modal window."""
        path = Path(__file__).parent / "Readme.txt"
        text = "Readme file not found."
        if path.exists():
            try:
                text = path.read_text(encoding="utf-8")
            except Exception as e:
                text = f"Failed to load disclaimer: {e}"

        dlg = tk.Toplevel(self.root)
        dlg.title("Disclaimer")
        if hasattr(self, 'icon_path') and self.icon_path.exists():
            try:
                dlg.iconbitmap(str(self.icon_path))
            except Exception as e:
                print(f"Failed to set icon on disclaimer dialog: {e}")
        dlg.transient(self.root)
        dlg.grab_set()
        dlg.geometry("800x600")

        txt = tk.Text(dlg, wrap="word", borderwidth=0, highlightthickness=0)
        txt.insert("1.0", text)
        txt.configure(state="disabled")
        txt.pack(fill=BOTH, expand=True, padx=10, pady=10)

        close = ttk.Button(dlg, text="Close", bootstyle="secondary", command=dlg.destroy)
        close.pack(pady=(0,10))   
        
     

    def create_widgets(self):
        # self.root.geometry("1024x768")
        self.root.geometry("1200x768")      # To accommodate jumpbox button
        # self.root.resizable(False, False)
        main_frame = ttk.Frame(self.root)
        main_frame.pack(fill=BOTH, expand=True, padx=10, pady=10)

        self.themes = self.load_themes()
        
        self.theme_name = self.load_theme_setting()

        # Load and apply saved UI theme
        # saved_ui_theme = self.load_ui_theme_setting()
        # self.style.theme_use(saved_ui_theme)

        self.configure_fonts()

        self.settings_frame = ttk.Frame(main_frame)
        self.settings_frame.pack(fill=X, pady=(0, 5))
        
        self.sound_effect_var = tk.BooleanVar(value=False)
        self.bgm_var = tk.BooleanVar(value=False)

        # —————— SETTINGS DROPDOWN BUTTON ——————
        settings_menu = tk.Menu(self.root, tearoff=0)
        settings_menu.add_command(label="🎵 BGM", command=self.show_theme_settings)
        settings_menu.add_command(label="🎨 Theme", command=self.change_theme)
        settings_menu.add_command(label="📖 Readme", command=self.show_disclaimer)

        settings_btn = ttk.Menubutton(
            self.settings_frame,
            # text="⚙️",
            text="🛠",
            bootstyle="info",
            direction="below"
        )
        settings_btn.pack(side=LEFT, padx=(0, 10))
        settings_btn["menu"] = settings_menu
        self.create_tooltip(settings_btn, "Application preferences and settings")

        # — BGM button: icon only
        # bgm_btn = ttk.Button(
        #     self.settings_frame,
        #     text=" 🎵 ",
        #     bootstyle="info",
        #     command=self.show_theme_settings,
        #     cursor="hand2"
        # )


        # bgm_btn.pack(side=LEFT, padx=(0, 10))
        # self.create_tooltip(bgm_btn, "Apply new BGM music to McConsole")

        
        # theme_btn = ttk.Button(
        #     self.settings_frame, 
        #     text=" 🎨 ", 
        #     bootstyle="info",
        #     cursor="hand2",
        #     command=self.change_theme
        # )
        # theme_btn.pack(side=LEFT, padx=(0, 10))
        # self.create_tooltip(theme_btn, "Apply new theme to McConsole")


        
        # disclaimer_btn = ttk.Button(
        #     self.settings_frame,
        #     text=" 📖 ",
        #     bootstyle="info",
        #     command=self.show_disclaimer,
        #     cursor="hand2"
        # )
        # disclaimer_btn.pack(side=LEFT, padx=(0, 10))
        # self.create_tooltip(disclaimer_btn, "Show Readme")

        new_folder_btn = ttk.Button(
            self.settings_frame, 
            text="📁 New", 
            bootstyle="info",
            command=self.show_new_folder_dialog,
            cursor="hand2"
        )
        new_folder_btn.pack(side=LEFT, padx=(0, 10))
        self.create_tooltip(new_folder_btn, "Create new SUT folder in the UI")

        self.auto_check_btn = ttk.Button(
            self.settings_frame, 
            text="⏳ Refresh", 
            bootstyle="info", 
            command=self.show_auto_check_settings,
            cursor="hand2"
        )
        self.auto_check_btn.pack(side=LEFT, padx=(0, 10))
        self.create_tooltip(self.auto_check_btn, "Periodically auto-refresh ping/SSH status for nodes in selected folders")   
        self.update_auto_check_button_text()   


        # —————— Jumpbox button ——————
        self.jumpbox_btn = ttk.Button(
            self.settings_frame,
            text="🌐 Jumpbox",
            bootstyle="info",
            command=self.open_jumpbox_settings,
            cursor="hand2"
        )
        self.jumpbox_btn.pack(side=tk.LEFT, padx=(0, 10))

        # create a tooltip for jumpbox button which will show if the jumpbox is enabled and what's the jumpbox IP
        self.jumpbox_config_file = Path(__file__).parent / "jumpbox_config.json"
        self.jumpbox_config = self.load_jumpbox_config()
        # jumpbox_tooltip_text = (
        #     f"Jumpbox is {'enabled' if self.jumpbox_config['enabled'] else 'disabled'}.\n"
        #     f"IP: {self.jumpbox_config['jumpbox_ip']}\n"
        #     # f"Port: {self.jumpbox_config['jumpbox_port']}\n"
        #     # f"Username: {self.jumpbox_config['jumpbox_username']}\n"
        #     f"Last updated: {self.jumpbox_config['last_updated']}"
        # )
        # self.create_tooltip(self.jumpbox_btn, jumpbox_tooltip_text)

        self.jumpbox_tooltip = self.create_tooltip(self.jumpbox_btn, "Loading...")
        self.update_jumpbox_button_text()

        
        # new button for mobaxterm controller console
        mobaxterm_btn = ttk.Button(
            self.settings_frame,
            text="📟 MobaX",
            bootstyle="info",
            command=self.launch_mobaxterm,
            cursor="hand2"
        )
        mobaxterm_btn.pack(side=LEFT, padx=(0, 10))
        self.create_tooltip(mobaxterm_btn, "Open MobaXterm controller window")


        sep = ttk.Separator(main_frame, orient=HORIZONTAL, bootstyle="secondary")
        sep.pack(fill=X, pady=1)  

        info_frame = ttk.Frame(main_frame)
        info_frame.pack(fill=tk.X, padx=10, pady=(5, 10))

        info_label = ttk.Label(
            info_frame, 
            text="💡 Tip: Hold Ctrl + Click to select multiple nodes for Multiple SUT command execution ",
            font=("Consolas", 10, "bold"),
            bootstyle="warning"  
        )
        info_label.pack(side=tk.LEFT)

        scroll_container = ttk.Frame(main_frame)
        scroll_container.pack(fill=tk.BOTH, expand=True)
        

        self.scrollbar = ttk.Scrollbar(scroll_container, orient=VERTICAL, bootstyle="round")
        
        
        self.canvas = tk.Canvas(scroll_container, highlightthickness=0)
        self.canvas.pack(side=LEFT, fill=BOTH, expand=True)
        
        self.scrollbar.pack(side=RIGHT, fill=Y)
        self.scrollbar.config(command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=self.scrollbar.set)
        
        # Content frame with modern styling
        self.content_frame = ttk.Frame(self.canvas)
        canvas_window = self.canvas.create_window((0, 0), window=self.content_frame, anchor="nw")
        
        self.canvas.bind('<Configure>', lambda e: (
            self.canvas.itemconfig(canvas_window, width=self.canvas.winfo_width()),
            self.update_scrollbar()
        ))
        
        self.content_frame.bind('<Configure>', lambda e: (
            self.canvas.configure(scrollregion=self.canvas.bbox("all")),
            self.update_scrollbar()
        ))
        
        if 'Favorites' not in self.sut_configs:
            self.sut_configs['Favorites'] = {}

        self.build_folder_ui()
        self.root.update_idletasks()
        self.root.minsize(self.root.winfo_width(), min(800, self.root.winfo_height()))
        self.update_scrollbar()
        
        if self.theme_name and self.theme_name != "None":
            self.apply_theme(self.theme_name)
        
        self.setup_root_context_menu()
        
        # Initialize auto-check after UI is built
        self.init_auto_check()
    
        self._apply_canvas_theme()

    def update_jumpbox_button_text(self):
        """Update jumpbox button text and style based on enabled state"""
        if self.jumpbox_config['enabled']:
            # text = "🌐 Jumpbox ✅"
            text = "🌐 Jumpbox ✓"
            # text = "🌐 Jumpbox ●"
            # text = "🌐 Jumpbox ◉"
            # style = "success"
            tooltip_text = (
                f"Jumpbox is ENABLED\n"
                f"IP: {self.jumpbox_config['jumpbox_ip']}\n"
                # f"Port: {self.jumpbox_config['jumpbox_port']}\n"
                # f"Username: {self.jumpbox_config['jumpbox_username']}\n"
                f"Last updated: {self.jumpbox_config['last_updated']}"
            )
        else:
            # text = "🌐 Jumpbox ❌"
            text = "🌐 Jumpbox ☓"
            # text = "🌐 Jumpbox ○"
            # text = "🌐 Jumpbox ◯"
            # style = "secondary"
            tooltip_text = (
                f"Jumpbox is DISABLED\n"
                f"Click to configure jumpbox settings\n"
                f"Last updated: {self.jumpbox_config['last_updated']}"
            )
        
        # self.jumpbox_btn.configure(text=text, bootstyle=style)
        self.jumpbox_btn.configure(text=text)
        # Update tooltip text
        self.create_tooltip(self.jumpbox_btn, tooltip_text)

    def load_jumpbox_config(self):
        """Load jumpbox configuration from JSON file"""
        default_config = {
            "enabled": False,
            "jumpbox_ip": "",
            "jumpbox_port": "22",
            "jumpbox_username": "root",
            "last_updated": ""
        }
        
        if self.jumpbox_config_file.exists():
            try:
                with open(self.jumpbox_config_file, 'r') as f:
                    config = json.load(f)
                    # Ensure all required keys exist
                    for key, default_value in default_config.items():
                        if key not in config:
                            config[key] = default_value
                    return config
            except Exception as e:
                print(f"Error loading jumpbox config: {e}")
        
        return default_config

    def save_jumpbox_config(self):
        """Save jumpbox configuration to JSON file"""
        try:
            # self.jumpbox_config["last_updated"] = datetime.now().isoformat()
            self.jumpbox_config["last_updated"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            with open(self.jumpbox_config_file, 'w') as f:
                json.dump(self.jumpbox_config, f, indent=4)
            # print("Jumpbox configuration saved successfully")
            
            self.update_jumpbox_button_text()
            
            # Show success message
            self.test_status_var.set("✅ Configuration saved successfully!")
            self.jumpbox_window.after(2000, lambda: self.test_status_var.set(""))

        except Exception as e:
            print(f"Error saving jumpbox config: {e}")

    def adjust_jumpbox_window_height(self):
        """Adjust jumpbox window height based on test message + form size"""
        self.jumpbox_window.update_idletasks()
        
        content_height = self.jumpbox_window.winfo_reqheight()
        screen_height = self.jumpbox_window.winfo_screenheight()
        
        min_height = 750  # minimum usable height
        max_height = int(screen_height * 0.85)
        
        new_height = min(max(content_height, min_height), max_height)
        current_width = max(650, self.jumpbox_window.winfo_width())

        self.jumpbox_window.geometry(f"{current_width}x{new_height}")

    def open_jumpbox_settings(self):
        """Open jumpbox server settings window"""
        # Check if window already exists
        if hasattr(self, 'jumpbox_window') and self.jumpbox_window.winfo_exists():
            self.jumpbox_window.lift()
            self.jumpbox_window.focus()
            return
        
        self.jumpbox_window = ttk.Toplevel(self.root)
        self.jumpbox_window.title("Jumpbox Server Settings")
        self.jumpbox_window.geometry("650x800")
        self.jumpbox_window.minsize(612, 782)
        self.jumpbox_window.resizable(True, True)
        
        # Set window icon if available
        if hasattr(self, 'icon_path') and self.icon_path.exists():
            try:
                self.jumpbox_window.iconbitmap(str(self.icon_path))
            except:
                pass
        
        # Make window modal
        self.jumpbox_window.transient(self.root)
        # self.jumpbox_window.grab_set()
        
        # Main frame with padding
        main_frame = ttk.Frame(self.jumpbox_window, padding=20)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Title
        title_label = ttk.Label(
            main_frame, 
            text="Jumpbox Server Configuration", 
            font=("Consolas", 14, "bold"),
            bootstyle="info"
        )
        title_label.pack(pady=(0, 20))
        
        # Enable/Disable jumpbox
        enable_frame = ttk.Frame(main_frame)
        enable_frame.pack(fill=tk.X, pady=(0, 15))
        
        self.jumpbox_enabled_var = tk.BooleanVar(value=self.jumpbox_config["enabled"])
        enable_cb = ttk.Checkbutton(
            enable_frame,
            text="Enable Jumpbox Server for all SSH connections",
            variable=self.jumpbox_enabled_var,
            bootstyle="success-round-toggle",
            command=self.on_jumpbox_toggle,
            cursor="hand2"
        )
        enable_cb.pack(anchor=tk.W)
        
        # Settings frame
        self.settings_frame = ttk.LabelFrame(main_frame, text="Jumpbox Server Details", padding=15)
        self.settings_frame.pack(fill=tk.X, pady=(0, 20))
        
        # IP Address
        ip_frame = ttk.Frame(self.settings_frame)
        ip_frame.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Label(ip_frame, text="IP Address:", font=("Consolas", 10)).pack(anchor=tk.W)
        self.jumpbox_ip_var = tk.StringVar(value=self.jumpbox_config["jumpbox_ip"])
        ip_entry = ttk.Entry(
            ip_frame, 
            textvariable=self.jumpbox_ip_var, 
            font=("Consolas", 10),
            bootstyle="info"
        )
        ip_entry.pack(fill=tk.X, pady=(5, 0))
        
        # Port
        port_frame = ttk.Frame(self.settings_frame)
        port_frame.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Label(port_frame, text="Port:", font=("Consolas", 10)).pack(anchor=tk.W)
        self.jumpbox_port_var = tk.StringVar(value=self.jumpbox_config["jumpbox_port"])
        port_entry = ttk.Entry(
            port_frame, 
            textvariable=self.jumpbox_port_var, 
            font=("Consolas", 10),
            bootstyle="info",
            width=10
        )
        port_entry.pack(anchor=tk.W, pady=(5, 0))
        
        # Username
        username_frame = ttk.Frame(self.settings_frame)
        username_frame.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Label(username_frame, text="Username:", font=("Consolas", 10)).pack(anchor=tk.W)
        self.jumpbox_username_var = tk.StringVar(value=self.jumpbox_config["jumpbox_username"])
        username_entry = ttk.Entry(
            username_frame, 
            textvariable=self.jumpbox_username_var, 
            font=("Consolas", 10),
            bootstyle="info"
        )
        username_entry.pack(fill=tk.X, pady=(5, 0))
        
        # Password field (for testing only)
        password_frame = ttk.Frame(self.settings_frame)
        password_frame.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Label(password_frame, text="Password (for connectivity check only):", font=("Consolas", 10)).pack(anchor=tk.W)
        
        # frame for password entry and show/hide button
        password_input_frame = ttk.Frame(password_frame)
        password_input_frame.pack(fill=tk.X, pady=(5, 0))
        
        self.jumpbox_password_var = tk.StringVar()
        self.jumpbox_password_entry = ttk.Entry(
            password_input_frame, 
            textvariable=self.jumpbox_password_var, 
            font=("Consolas", 10),
            bootstyle="info",
            show="*"  # Masking password
        )
        self.jumpbox_password_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        # Show/Hide password button
        self.password_visible = False
        self.toggle_password_btn = ttk.Button(
            password_input_frame,
            text="👁",
            width=3,
            command=self.toggle_password_visibility,
            cursor="hand2",
            bootstyle="secondary"
        )
        self.toggle_password_btn.pack(side=tk.RIGHT, padx=(5, 0))
        
        # Password note
        password_note_frame = ttk.Frame(password_frame)
        password_note_frame.pack(fill=tk.X, pady=(5, 0))
        
        password_note_text = "🔒 This password is only used for connection testing and will NOT be saved."
        password_note_label = ttk.Label(
            password_note_frame, 
            text=password_note_text, 
            font=("Consolas", 8),
            foreground="orange",
            wraplength=self.jumpbox_window.winfo_width() - 80,
            justify=tk.LEFT
        )
        password_note_label.pack(anchor=tk.W)

        # # Note about password
        # note_frame = ttk.Frame(self.settings_frame)
        # note_frame.pack(fill=tk.X, pady=(10, 0))
        
        # note_text = "📝 Note: Password authentication will be handled by MobaXterm.\nEnsure your jumpbox server is configured for SSH key authentication\nor MobaXterm has the credentials saved."
        # note_label = ttk.Label(
        #     note_frame, 
        #     text=note_text, 
        #     font=("Consolas", 9),
        #     foreground="gray",
        #     # wraplength=450,
        #     wraplength=self.jumpbox_window.winfo_width() - 80,
        #     justify=tk.LEFT
        # )
        # note_label.pack(anchor=tk.W)
        
        # Test connection frame
        test_frame = ttk.Frame(main_frame)
        test_frame.pack(fill=tk.X, pady=(0, 15))
        
        test_btn = ttk.Button(
            test_frame,
            text="Test Jumpbox Connection",
            bootstyle="warning",
            command=self.test_jumpbox_connection,
            cursor="hand2"
        )
        test_btn.pack(anchor=tk.W)
        
        # Status label for test results
        self.test_status_var = tk.StringVar()
        self.test_status_label = ttk.Label(
            test_frame, 
            textvariable=self.test_status_var,
            font=("Consolas", 9),
            wraplength=self.jumpbox_window.winfo_width() - 80,
            justify=tk.LEFT
        )
        self.test_status_label.pack(fill=tk.X, pady=(5, 0))
        
        # Buttons frame
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=(20, 0))
        
        # Save button
        save_btn = ttk.Button(
            button_frame,
            text="Save Configuration",
            bootstyle="success",
            command=self.save_jumpbox_settings,
            cursor="hand2"
        )
        save_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        # Cancel button
        cancel_btn = ttk.Button(
            button_frame,
            text="Cancel",
            bootstyle="secondary",
            command=self.jumpbox_window.destroy,
            cursor="hand2"
        )
        cancel_btn.pack(side=tk.LEFT)
        
        # Reset button
        reset_btn = ttk.Button(
            button_frame,
            text="Reset to Defaults",
            bootstyle="danger",
            command=self.reset_jumpbox_settings,
            cursor="hand2"
        )
        reset_btn.pack(side=tk.RIGHT)
        
        # Initialize UI state
        self.on_jumpbox_toggle()
        
        # Center the window
        self.jumpbox_window.update_idletasks()
        x = (self.jumpbox_window.winfo_screenwidth() // 2) - (self.jumpbox_window.winfo_width() // 2)
        y = (self.jumpbox_window.winfo_screenheight() // 2) - (self.jumpbox_window.winfo_height() // 2)
        self.jumpbox_window.geometry(f"+{x}+{y}")

    def toggle_password_visibility(self):
        """Toggle password visibility in the entry field"""
        if self.password_visible:
            self.jumpbox_password_entry.configure(show="*")
            self.toggle_password_btn.configure(text="👁")
            self.password_visible = False
        else:
            self.jumpbox_password_entry.configure(show="")
            self.toggle_password_btn.configure(text="🙈")
            self.password_visible = True

    def on_jumpbox_toggle(self):
        """Handle jumpbox enable/disable toggle"""
        enabled = self.jumpbox_enabled_var.get()
        
        # Enable/disable all widgets in settings frame
        for child in self.settings_frame.winfo_children():
            self.toggle_widget_state(child, "normal" if enabled else "disabled")

        if hasattr(self, 'jumpbox_password_entry'):
            if self.jumpbox_enabled_var.get():
                self.jumpbox_password_entry.configure(state="normal")
                self.toggle_password_btn.configure(state="normal")
            else:
                self.jumpbox_password_entry.configure(state="disabled")
                self.toggle_password_btn.configure(state="disabled")
                # Clear password when disabled
                self.jumpbox_password_var.set("")

        self.jumpbox_config['enabled'] = enabled
        if hasattr(self, 'jumpbox_btn'):
            self.update_jumpbox_button_text()
    
    def toggle_widget_state(self, widget, state):
        """Recursively toggle widget state"""
        try:
            if hasattr(widget, 'configure'):
                if isinstance(widget, (ttk.Entry, ttk.Checkbutton, ttk.Button)):
                    widget.configure(state=state)
            # Handle child widgets
            for child in widget.winfo_children():
                self.toggle_widget_state(child, state)
        except:
            pass

    def test_jumpbox_connection(self):
        """Test TCP and SSH connection to jumpbox server"""
        if not self.jumpbox_enabled_var.get():
            self.test_status_var.set("⚠️ Jumpbox is disabled")
            return

        jumpbox_ip = self.jumpbox_ip_var.get().strip()
        jumpbox_port = self.jumpbox_port_var.get().strip()
        jumpbox_username = self.jumpbox_username_var.get().strip()
        jumpbox_password = self.jumpbox_password_var.get()

        if not jumpbox_ip:
            self.test_status_var.set("❌ Please enter jumpbox IP address")
            return

        if not jumpbox_port.isdigit():
            self.test_status_var.set("❌ Invalid port number")
            return

        if not jumpbox_username:
            self.test_status_var.set("❌ Please enter jumpbox username")
            return
        
        if not jumpbox_password:
            self.test_status_var.set("❌ Please enter password for connection testing")
            return

        self.test_status_var.set("🔄 Testing connection...")
        self.jumpbox_window.update()

        def test_connection():
            tcp_success = False
            ssh_success = False
            ssh_error = None

            # Step 1: TCP Port Test
            try:
                sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
                sock.settimeout(5)
                result = sock.connect_ex((jumpbox_ip, int(jumpbox_port)))
                sock.close()
                tcp_success = (result == 0)
            except Exception as e:
                tcp_success = False
                ssh_error = str(e)

            if tcp_success:
                # ssh_cmd = [
                #     "ssh",
                #     "-o", "BatchMode=yes",
                #     "-p", jumpbox_port,
                #     f"{jumpbox_username}@{jumpbox_ip}",
                #     "exit"
                # ]
                # try:
                #     result = subprocess.run(ssh_cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=7)
                #     ssh_success = (result.returncode == 0)
                #     if not ssh_success:
                #         ssh_error = result.stderr.decode().strip()
                # except Exception as e:
                #     ssh_error = str(e)

                try:
                    ssh_client = paramiko.SSHClient()
                    ssh_client.set_missing_host_key_policy(paramiko.AutoAddPolicy())
                    
                    # Test SSH connection with password
                    ssh_client.connect(
                        hostname=jumpbox_ip,
                        port=int(jumpbox_port),
                        username=jumpbox_username,
                        password=jumpbox_password,
                        timeout=7,
                        allow_agent=False,
                        look_for_keys=False
                    )
                    
                    # Testing a simple command
                    stdin, stdout, stderr = ssh_client.exec_command('echo "test"')
                    stdout.read()
                    ssh_client.close()
                    ssh_success = True
                    
                except paramiko.AuthenticationException:
                    ssh_error = "Authentication failed: Invalid username or password"
                except paramiko.SSHException as e:
                    ssh_error = f"SSH connection error: {str(e)}"
                except Exception as e:
                    ssh_error = f"Connection error: {str(e)}"


            def update_ui():
                """Update UI with test results"""
                if not tcp_success:
                    self.test_status_var.set("❌ TCP port test failed! Jumpbox server not reachable.")
                elif ssh_success:
                    self.test_status_var.set("✅ TCP + SSH authentication successful! Jumpbox is ready.")
                else:
                    msg = "⚠️ TCP port is reachable, but SSH authentication failed.\n"
                    msg += f"Details: {ssh_error or 'Unknown error'}\n"
                    msg += "Please check your username and password."
                    self.test_status_var.set(msg)

                self.test_status_label.configure(wraplength=self.jumpbox_window.winfo_width() - 80)
                self.adjust_jumpbox_window_height()
                
                # Clearing password after test for security
                self.jumpbox_password_var.set("")
            
            self.root.after(0, update_ui)

        threading.Thread(target=test_connection, daemon=True).start()

    def save_jumpbox_settings(self):
        """Save jumpbox settings and close window"""
        # Validate inputs
        if self.jumpbox_enabled_var.get():
            jumpbox_ip = self.jumpbox_ip_var.get().strip()
            jumpbox_port = self.jumpbox_port_var.get().strip()
            jumpbox_username = self.jumpbox_username_var.get().strip()
            
            if not jumpbox_ip:
                messagebox.showerror("Validation Error", "Please enter jumpbox IP address")
                return
            
            if not jumpbox_port.isdigit():
                messagebox.showerror("Validation Error", "Port must be a number")
                return
            
            if not jumpbox_username:
                messagebox.showerror("Validation Error", "Please enter jumpbox username")
                return
        
        # Update configuration
        self.jumpbox_config.update({
            "enabled": self.jumpbox_enabled_var.get(),
            "jumpbox_ip": self.jumpbox_ip_var.get().strip(),
            "jumpbox_port": self.jumpbox_port_var.get().strip(),
            "jumpbox_username": self.jumpbox_username_var.get().strip()
        })
        
        self.save_jumpbox_config()
        
        # Show success message
        status = "enabled" if self.jumpbox_config["enabled"] else "disabled"
        messagebox.showinfo("Success", f"Jumpbox configuration saved successfully!\nJumpbox is now {status}.")
        
        self.jumpbox_password_var.set("") # Clearing password field for security
        self.jumpbox_window.destroy()

    def reset_jumpbox_settings(self):
        """Reset jumpbox settings to defaults"""
        if messagebox.askyesno("Confirm Reset", "Are you sure you want to reset jumpbox settings to defaults?"):
            self.jumpbox_enabled_var.set(False)
            self.jumpbox_ip_var.set("")
            self.jumpbox_port_var.set("22")
            self.jumpbox_username_var.set("root")
            self.test_status_var.set("")
            self.on_jumpbox_toggle()

            if hasattr(self, 'jumpbox_password_var'):
                self.jumpbox_password_var.set("")

    # def get_jumpbox_command_prefix(self, target_ip, target_username="root"):
    #     """Generate jumpbox command prefix if enabled"""
    #     if not self.jumpbox_config["enabled"]:
    #         return None
        
    #     jumpbox_ip = self.jumpbox_config["jumpbox_ip"]
    #     jumpbox_port = self.jumpbox_config["jumpbox_port"]
    #     jumpbox_username = self.jumpbox_config["jumpbox_username"]
        
    #     if not jumpbox_ip:
    #         return None
        
    #     return f"ssh -J {jumpbox_username}@{jumpbox_ip}:{jumpbox_port} {target_username}@{target_ip}"

    # new function to launch mobaxterm with controller window
    def launch_mobaxterm(self):
        try:
            subprocess.Popen(["python", "./shell_mx.py", "--console_type", "controller", "--sut_ip", "unknown"], shell=True)
        except Exception as e:
            print(f"Failed to launch MobaXterm: {e}")


    def configure_fonts(self):
        """Configure consistent fonts across all themes"""
        
        button_font = ("Consolas", 10, "bold")
        label_font = ("Consolas", 10, "bold")
        checkbutton_font = ("Consolas", 10, "bold")
        
        self.style.configure("TButton", font=button_font)
        self.style.configure("info.TButton", font=button_font)
        
        self.style.configure("TCheckbutton", font=checkbutton_font)
        self.style.configure("info.Round.Toggle", font=checkbutton_font)
        
        self.style.configure("TLabel", font=label_font)

    def change_theme(self):
        """Change the application theme and maintain consistent font styling"""
        themes = ["cyborg", "darkly", "solar", "cosmo", "pulse"]
        current_theme = self.style.theme_use()
        current_index = themes.index(current_theme) if current_theme in themes else 0
        next_theme = themes[(current_index + 1) % len(themes)]
        
        self.style.theme_use(next_theme)
        
        self.configure_fonts()
        
        # Re-apply canvas theme
        self._apply_canvas_theme()
        
        # self.save_ui_theme_setting(next_theme)

        current_theme_setting = self.load_theme_setting()  # Get current theme setting
        self.save_theme_setting(current_theme_setting, next_theme)  # Save both
        
        self._update_node_selection_visuals()
        
        print(f"Theme changed to: {next_theme}")

    def _apply_canvas_theme(self):
        """Apply theme colors to canvas based on current ttkbootstrap theme"""
        colors = self.style.colors
        self.canvas.configure(
            bg=colors.bg,
            highlightcolor=colors.border,
            insertbackground=colors.fg
        )

    def load_expand_state(self):
        """Load folder expansion states and auto-check settings from JSON."""
        self.auto_check_interval_ms = 60000  # Default 1 minute
        self.auto_check_settings = {
            "enabled": False,
            "interval_ms": 60000,
            "folders": []
        }
        
        if self.expand_state_file.exists():
            try:
                with open(self.expand_state_file, 'r') as f:
                    data = json.load(f)
                    self.folder_expanded_state = data.get("folder_expanded_state", {})
                    
                    if "auto_check_settings" in data:
                        self.auto_check_settings = data["auto_check_settings"]
                        self.auto_check_interval_ms = self.auto_check_settings["interval_ms"]
                    else:
                        self.auto_check_interval_ms = int(data.get("auto_check_interval_ms", 60000))
                        self.auto_check_settings["interval_ms"] = self.auto_check_interval_ms
            except Exception as e:
                print("Warning: could not load expand_state.json:", e)
    
    def save_expand_state(self):
        """Save the folder expansion states and auto-check settings to JSON."""
        data = {
            "folder_expanded_state": self.folder_expanded_state,
            "auto_check_settings": self.auto_check_settings
        }
        try:
            with open(self.expand_state_file, 'w') as f:
                json.dump(data, f, indent=2)
        except Exception as e:
            print("Warning: could not save expand_state.json:", e)

    def update_auto_check_button_text(self):
        """Update the auto-check button text based on enabled state."""
        if self.auto_check_settings["enabled"]:
            text = "⏳ Refresh ✓"
            # style = "success"
            tooltip_text = (
                f"Auto Refresh is ENABLED\n"
                f"Interval: {self.auto_check_settings['interval_ms'] // 1000} seconds\n"
                f"Folders: {', '.join(self.auto_check_settings['folders']) if self.auto_check_settings['folders'] else 'All expanded folders'}"
            )
        else:
            text = "⏳ Refresh ☓"
            # style = "secondary"
            tooltip_text = (
                f"Auto Refresh is DISABLED\n"
                f"Click to configure auto refresh settings"
            )
        
        self.auto_check_btn.configure(text=text)
        self.create_tooltip(self.auto_check_btn, tooltip_text)

    def show_auto_check_settings(self):
        """Show the auto-check settings dialog."""
        
        dialog = ttk.Toplevel(self.root)
        dialog.title("Auto Refresh Settings")
        if hasattr(self, 'icon_path') and self.icon_path.exists():
            try:
                dialog.iconbitmap(str(self.icon_path))
            except Exception as e:
                print(f"Failed to set icon on auto refresh dialog: {e}")
        dialog.geometry("450x660")
        dialog.minsize(450, 660)
        dialog.resizable(True, True)
        dialog.transient(self.root)
        # dialog.grab_set()
        
        # Center the dialog
        dialog.update_idletasks()
        x = (dialog.winfo_screenwidth() // 2) - (dialog.winfo_width() // 2)
        y = (dialog.winfo_screenheight() // 2) - (dialog.winfo_height() // 2)
        dialog.geometry(f"+{x}+{y}")
        
        # Main frame with padding
        main_frame = ttk.Frame(dialog, padding=20)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Title
        title_label = ttk.Label(
            main_frame, 
            text="Auto Refresh Configuration", 
            font=("Consolas", 14, "bold"),
            bootstyle="info"
        )
        title_label.pack(pady=(0, 25))
        
        # Enable/Disable frame
        enable_frame = ttk.Frame(main_frame)
        enable_frame.pack(fill=tk.X, pady=(0, 20))
        
        self.dialog_enabled_var = tk.BooleanVar(value=self.auto_check_settings["enabled"])
        enable_cb = ttk.Checkbutton(
            enable_frame, 
            text="Enable Auto Refresh", 
            variable=self.dialog_enabled_var,
            bootstyle="success-round-toggle",
            command=self.toggle_dialog_controls
        )
        enable_cb.pack(anchor=tk.W)
        
        # Settings frame
        self.settings_frame = ttk.LabelFrame(main_frame, text="Refresh Settings", padding=15)
        self.settings_frame.pack(fill=tk.X, pady=(0, 20))
        
        # Interval setting
        interval_frame = ttk.Frame(self.settings_frame)
        interval_frame.pack(fill=tk.X, pady=(0, 15))
        
        ttk.Label(interval_frame, text="Refresh Interval:", font=("Consolas", 10)).pack(anchor=tk.W)
        
        interval_input_frame = ttk.Frame(interval_frame)
        interval_input_frame.pack(fill=tk.X, pady=(5, 0))
        
        self.interval_var = tk.StringVar(value=str(self.auto_check_settings["interval_ms"] // 1000))
        self.interval_spinbox = ttk.Spinbox(
            interval_input_frame, 
            from_=10, 
            to=3600, 
            textvariable=self.interval_var, 
            width=12,
            font=("Consolas", 10),
            bootstyle="info"
        )
        self.interval_spinbox.pack(side=tk.LEFT)
        
        ttk.Label(interval_input_frame, text="seconds", font=("Consolas", 10)).pack(side=tk.LEFT, padx=(10, 0))

        # Folder selection frame
        folders_frame = ttk.LabelFrame(main_frame, text="Select Folders to Monitor", padding=15)
        # folders_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 20))
        folders_frame.pack(fill=tk.X, pady=(0, 20))

        # Select All checkbox
        select_all_frame = ttk.Frame(folders_frame)
        select_all_frame.pack(fill=tk.X, pady=(0, 10))

        self.select_all_var = tk.BooleanVar()
        select_all_cb = ttk.Checkbutton(
            select_all_frame,
            text="Select All Folders",
            variable=self.select_all_var,
            bootstyle="warning-round-toggle",
            command=self.toggle_select_all_folders
        )
        select_all_cb.pack(anchor=tk.W)

        # Add separator
        separator = ttk.Separator(folders_frame, orient='horizontal')
        separator.pack(fill=tk.X, pady=(5, 10))

        # Get only expanded folders from the current state
        expanded_folders = []
        for folder_name, is_expanded in self.folder_expanded_state.items():
            if is_expanded and folder_name in self.sut_configs and self.sut_configs[folder_name]:
                expanded_folders.append(folder_name)

        expanded_folders.sort()

        if not expanded_folders:
            no_folders_label = ttk.Label(
                folders_frame, 
                text="* No expanded folders available.\nExpand folder(s) that need refresh first.",
                font=("Consolas", 9),
                foreground="gray",
                wraplength=400,
                justify=tk.LEFT
            )
            no_folders_label.pack(pady=20)
            self.folder_vars = {}
            self.folder_checkboxes = []
        else:
            # Creating scrollable frame for folder checkboxes
            max_visible_items = 6
            needs_scrollbar = len(expanded_folders) > max_visible_items
            
            if needs_scrollbar:
                scroll_container = ttk.Frame(folders_frame)
                # scroll_container.pack(fill=tk.BOTH, expand=True)
                scroll_container.pack(fill=tk.X)
                
                # listbox
                folder_listframe = ttk.Frame(scroll_container)
                folder_listframe.pack(fill=tk.BOTH, expand=True)
                
                # scrollable text widget
                scroll_frame = ttk.Frame(folder_listframe)
                scroll_frame.pack(fill=tk.BOTH, expand=True)
                
                # Scrollbar
                scrollbar = ttk.Scrollbar(scroll_frame, bootstyle="round")
                scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
                
                checkbox_canvas = tk.Canvas(
                    scroll_frame, 
                    yscrollcommand=scrollbar.set,
                    # height=200,  
                    height=min(150, len(expanded_folders) * 25 + 10), 
                    highlightthickness=0
                )
                checkbox_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
                scrollbar.config(command=checkbox_canvas.yview)
                
                checkbox_frame = ttk.Frame(checkbox_canvas)
                canvas_window = checkbox_canvas.create_window((0, 0), window=checkbox_frame, anchor="nw")
            else:
                # Simple frame without scrolling
                checkbox_frame = ttk.Frame(folders_frame)
                checkbox_frame.pack(fill=tk.BOTH, expand=True)
            
            # Create checkboxes
            self.folder_vars = {}
            self.folder_checkboxes = []
            enabled_folders = self.auto_check_settings.get("folders", [])
            
            for folder_name in expanded_folders:
                var = tk.BooleanVar(value=folder_name in enabled_folders)
                self.folder_vars[folder_name] = var
                
                cb = ttk.Checkbutton(
                    checkbox_frame, 
                    text=folder_name, 
                    variable=var, 
                    bootstyle="info-round-toggle",
                    command=self.update_select_all_state
                )
                cb.pack(anchor=tk.W, padx=5, pady=2, fill=tk.X)
                self.folder_checkboxes.append(cb)
            
            self.update_select_all_state()

            # Configure scrolling if needed
            if needs_scrollbar:
                def configure_scroll_region(event=None):
                    checkbox_canvas.configure(scrollregion=checkbox_canvas.bbox("all"))
                
                def configure_canvas_width(event):
                    canvas_width = event.width
                    checkbox_canvas.itemconfig(canvas_window, width=canvas_width)
                
                checkbox_frame.bind("<Configure>", configure_scroll_region)
                checkbox_canvas.bind("<Configure>", configure_canvas_width)
                
                # Mouse wheel scrolling
                def on_mousewheel(event):
                    checkbox_canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
                
                def bind_mousewheel(widget):
                    widget.bind("<MouseWheel>", on_mousewheel)
                    widget.bind("<Button-4>", lambda e: checkbox_canvas.yview_scroll(-1, "units"))
                    widget.bind("<Button-5>", lambda e: checkbox_canvas.yview_scroll(1, "units"))
                    
                    for child in widget.winfo_children():
                        bind_mousewheel(child)
                
                bind_mousewheel(checkbox_canvas)
                bind_mousewheel(checkbox_frame)
        
        # Buttons frame
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=(20, 0))
        
        # Save button
        save_btn = ttk.Button(
            button_frame, 
            text="Save Config", 
            bootstyle="success",
            command=lambda: self.save_auto_check_settings(dialog)
        )
        save_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        # Cancel button
        cancel_btn = ttk.Button(
            button_frame, 
            text="Cancel", 
            bootstyle="secondary",
            command=dialog.destroy
        )
        cancel_btn.pack(side=tk.LEFT)
        
        # Store controls for enable/disable functionality
        # self.dialog_controls = [self.settings_frame]
        self.dialog_controls = [self.settings_frame, select_all_cb]
        if hasattr(self, 'folder_checkboxes'):
            self.dialog_controls.extend(self.folder_checkboxes)
        
        # Initial state
        self.toggle_dialog_controls()

    def toggle_dialog_controls(self):
        """Enable/disable dialog controls based on the enabled checkbox."""
        state = tk.NORMAL if self.dialog_enabled_var.get() else tk.DISABLED
        
        # Enable/disable the entire settings frame
        if hasattr(self, 'settings_frame'):
            try:
                # For LabelFrame, we need to handle children
                for child in self.settings_frame.winfo_children():
                    self._set_widget_state_recursive(child, state)
            except tk.TclError:
                pass
        
        # Special handling for folder checkboxes
        if hasattr(self, 'folder_checkboxes'):
            for cb in self.folder_checkboxes:
                try:
                    cb.configure(state=state)
                except tk.TclError:
                    pass

    def _set_widget_state_recursive(self, widget, state):
        """Recursively set state for all child widgets."""
        try:
            if hasattr(widget, 'configure'):
                widget.configure(state=state)
        except tk.TclError:
            pass
        
        # Recursively handle children
        for child in widget.winfo_children():
            self._set_widget_state_recursive(child, state)

    def toggle_select_all_folders(self):
        """Toggle all folder checkboxes based on select all state."""
        if not hasattr(self, 'folder_vars'):
            return
        
        select_all_state = self.select_all_var.get()
        for var in self.folder_vars.values():
            var.set(select_all_state)

    def update_select_all_state(self):
        """Update select all checkbox state based on individual folder selections."""
        if not hasattr(self, 'folder_vars') or not hasattr(self, 'select_all_var'):
            return
        
        all_selected = all(var.get() for var in self.folder_vars.values())
        any_selected = any(var.get() for var in self.folder_vars.values())
        
        if all_selected:
            self.select_all_var.set(True)
        elif not any_selected:
            self.select_all_var.set(False)

    def save_auto_check_settings(self, dialog):
        """Save the auto-refresh settings and apply theme."""
        try:
            # Validate interval
            interval_seconds = int(self.interval_var.get())
            if interval_seconds < 10:
                interval_seconds = 10
            elif interval_seconds > 3600:
                interval_seconds = 3600
                
            # Get selected folders
            selected_folders = [folder for folder, var in self.folder_vars.items() if var.get()]
            
            # Validation: If auto-check is enabled, ensure at least one folder is selected
            if self.dialog_enabled_var.get() and not selected_folders:
                tk.messagebox.showwarning(
                    "No Folders Selected", 
                    "Please select at least one folder to enable auto refresh functionality.\n\n"
                    "Either:\n"
                    "• Select one or more folders from the list, or\n"
                    "• Disable auto refresh if you don't want this feature"
                )
                return  # Avoid closing the dialog box
            
            # Update settings
            self.auto_check_settings = {
                "enabled": self.dialog_enabled_var.get(),
                "interval_ms": interval_seconds * 1000,
                "folders": selected_folders
            }
            
            self.save_expand_state()
            
            self.apply_auto_check_settings()

            self.update_auto_check_button_text()
            
            if self.auto_check_settings["enabled"]:
                folder_count = len(selected_folders)
                folder_list = ", ".join(selected_folders)
                print(f"Auto Refresh enabled for {folder_count} folder(s) with {interval_seconds}s interval")
                print(f"Monitoring folders: {folder_list}")
            else:
                print("Auto Refresh disabled")
            
            dialog.destroy()
            
        except ValueError:
            tk.messagebox.showerror(
                "Invalid Input", 
                "Please enter a valid interval between 10 and 3600 seconds."
            )

    def init_auto_check(self):
        """Initialize auto-check based on saved settings."""
        self.auto_check_running = False
        if self.auto_check_settings["enabled"] and self.auto_check_settings["folders"]:
            self.apply_auto_check_settings()

    def apply_auto_check_settings(self):
        """Apply the current auto-check settings."""
        self.auto_check_running = False  # Stop any existing auto-check
        
        if self.auto_check_settings["enabled"] and self.auto_check_settings["folders"]:
            self.auto_check_running = True
            self.schedule_auto_check()
            print(f"Auto Refresh started for folders: {', '.join(self.auto_check_settings['folders'])}")
        else:
            print("Auto Refresh stopped")

    def schedule_auto_check(self):
        """Schedule the auto-check with the new logic."""
        if not self.auto_check_running:
            return
            
        def auto_check_loop():
            if not self.auto_check_running:
                return
                
            interval = self.auto_check_settings["interval_ms"]
            enabled_folders = self.auto_check_settings["folders"]
            
            print(f"DEBUG: Auto Refresh triggered for folders: {enabled_folders}, next check in {interval} ms")
            
            # Only check enabled folders that are expanded
            for folder in enabled_folders:
                if folder in self.folder_expanded and self.folder_expanded[folder]:
                    self.refresh_folder(folder, is_auto_check=True)
                    self._on_folder_refreshed()
            
            self.root.after(interval, auto_check_loop)
        
        self.root.after(self.auto_check_settings["interval_ms"], auto_check_loop)

    def sanitize_folder_name(self, name):
        """Sanitize the folder name to remove unsafe characters and normalize spacing."""
        name = name.strip()  # Trim leading/trailing spaces
        name = re.sub(r"[^a-zA-Z0-9 _-]", "", name)  # Remove disallowed characters
        name = re.sub(r"\s+", " ", name)  # Collapse multiple spaces into one
        return name.strip()    


    def create_new_folder(self, folder_name):
        """Create a new SUT folder on disk and add it to the UI without rebuilding everything."""
        if not folder_name or folder_name.strip() == "":
            messagebox.showerror("Error", "Folder name cannot be empty")
            return
        sanitized_name = self.sanitize_folder_name(folder_name)
        
        # Check for existing folder with the same name
        sut_dir = Path(__file__).parent.parent / "sut"
        new_folder_path = sut_dir / sanitized_name
        
        if sanitized_name in self.sut_configs or new_folder_path.exists():
            messagebox.showerror("Error", f"A folder named '{sanitized_name}' already exists")
            return
            
        try:
            # Create the directory
            new_folder_path.mkdir(exist_ok=True)
            
            # Update our internal data structure
            self.sut_configs[sanitized_name] = {}
            
            # Add this folder to the UI (targeted update)
            self.add_new_folder_to_ui(sanitized_name)
            
            # Show success message
            messagebox.showinfo("Success", f"Folder '{sanitized_name}' created successfully")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to create folder: {str(e)}")


    def add_new_folder_to_ui(self, folder_name):
        """
        Add a new folder to the UI by rebuilding the folder structure safely
        and then refreshing the content of all expanded folders.
        """

        if folder_name not in self.sut_configs:
            self.sut_configs[folder_name] = {}

        self.folder_expanded_state[folder_name] = True
        self.save_expand_state()

        self.build_folder_ui()
        
        print("Rebuilding UI complete. Refreshing content of expanded folders...")
        for folder, is_expanded in self.folder_expanded.items():
            if is_expanded:
                print(f"-> Refreshing folder: {folder}")
                self.refresh_folder(folder) 
                self._on_folder_refreshed()

        self.root.update_idletasks()
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        self.update_scrollbar()

    def show_new_folder_dialog(self):
        """Show dialog to get new folder name from user."""
        dialog = tk.Toplevel(self.root)
        dialog.title("New Folder")
        if hasattr(self, 'icon_path') and self.icon_path.exists():
            try:
                dialog.iconbitmap(str(self.icon_path))
            except Exception as e:
                print(f"Failed to set icon on new folder dialog: {e}")
        dialog.geometry("300x160")
        dialog.minsize(300,160)
        # dialog.resizable(False, False)
        dialog.transient(self.root)
        dialog.grab_set()
        
        # Apply the current theme to the dialog
        # dialog.configure(bg="#1E1E1E")  # Default dark theme
        
        # Center the dialog
        dialog.update_idletasks()
        x = self.root.winfo_x() + (self.root.winfo_width() - dialog.winfo_width()) // 2
        y = self.root.winfo_y() + (self.root.winfo_height() - dialog.winfo_height()) // 2
        dialog.geometry(f"+{x}+{y}")
        
        # Create frame for content
        frame = ttk.Frame(dialog)
        frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Label
        ttk.Label(frame, text="Enter folder name:").pack(anchor="w", pady=(0, 5))
        
        # Entry field
        entry = ttk.Entry(frame, width=30)
        entry.pack(fill="x", pady=(0, 10))
        entry.focus_set()
        
        # Buttons frame
        btn_frame = ttk.Frame(frame)
        btn_frame.pack(fill="x")
        
        # Create and Cancel buttons
        create_btn = ttk.Button(
            btn_frame, text="Create", bootstyle="info",
            command=lambda: [self.create_new_folder(entry.get()), dialog.destroy()]
        )
        create_btn.pack(side=tk.RIGHT, padx=(5, 0))
        
        cancel_btn = ttk.Button(
            btn_frame, text="Cancel", bootstyle="secondary",
            command=dialog.destroy
        )
        cancel_btn.pack(side=tk.RIGHT)
        
        # Bind Enter key to create button
        entry.bind("<Return>", lambda e: [self.create_new_folder(entry.get()), dialog.destroy()])
        
        dialog.protocol("WM_DELETE_WINDOW", dialog.destroy)  # Handle window close

    def setup_root_context_menu(self):
        """Set up a context menu for the main content area."""
        # Create a context menu for empty space
        self.root_context_menu = tk.Menu(self.root, tearoff=0)
        self.root_context_menu.add_command(
            label="New Folder",
            command=self.show_new_folder_dialog
        )
        # self.root_context_menu.add_command(
        #     label="Refresh All",
        #     command=self.refresh_all_folders
        # )
        
        # Bind the context menu to right-click on main content areas
        self.content_frame.bind("<Button-3>", self.show_root_context_menu)
        self.canvas.bind("<Button-3>", self.show_root_context_menu)
        
    def show_root_context_menu(self, event):
        """Show the root context menu on right-click."""
        self.root_context_menu.tk_popup(event.x_root, event.y_root)

    def save_theme_setting(self, theme_name, ui_theme=None):
        """Save the selected theme to a settings file."""
        settings_file = Path(__file__).parent / "theme_setting.json"
        
        # Load existing settings first
        settings = {}
        if settings_file.exists():
            try:
                with open(settings_file, 'r') as f:
                    settings = json.load(f)
            except Exception as e:
                print(f"Error loading existing settings: {e}")
        
        # Update the settings
        settings["theme"] = theme_name
        if ui_theme is not None:
            settings["ui_theme"] = ui_theme
        
        try:
            with open(settings_file, 'w') as f:
                json.dump(settings, f, indent=2)
            print(f"Theme setting saved: {theme_name}" + (f", UI theme: {ui_theme}" if ui_theme else ""))
        except Exception as e:
            print(f"Error saving theme setting: {e}")

    def load_theme_setting(self):
        """Load the saved theme setting if available."""
        settings_file = Path(__file__).parent / "theme_setting.json"
        if settings_file.exists():
            try:
                with open(settings_file, 'r') as f:
                    settings = json.load(f)
                    theme_name = settings.get("theme", "None")
                    print(f"Loaded saved theme setting: {theme_name}")
                    return theme_name
            except Exception as e:
                print(f"Error loading theme setting: {e}")
        return "None"  # Default to no theme if no setting is found

    def load_ui_theme_setting(self):
        """Load the saved UI theme setting if available."""
        settings_file = Path(__file__).parent / "theme_setting.json"
        if settings_file.exists():
            try:
                with open(settings_file, 'r') as f:
                    settings = json.load(f)
                    ui_theme = settings.get("ui_theme", "darkly")
                    print(f"Loaded saved UI theme setting: {ui_theme}")
                    return ui_theme
            except Exception as e:
                print(f"Error loading UI theme setting: {e}")
        return "darkly"  # Default to darkly theme

    def build_folder_ui(self):
        # Reset or rebuild folders
        for widget in self.content_frame.winfo_children():
            widget.destroy()

        folders_in_order = ['Favorites'] + sorted(
            [f for f in self.sut_configs if f != 'Favorites'],
            key=lambda x: (x != 'Root', x)
        )

        grid_row = 0
        for folder in folders_in_order:
            self.add_folder_section(folder, grid_row)
            grid_row += 2

    def add_folder_section(self, folder, grid_row):
        """
        Add a section for a specific folder in the UI.

        Args:
            folder (_type_): _folder name to add
            grid_row (_type_): _row index for grid placement
        """
        self.folder_expanded[folder] = self.folder_expanded_state.get(folder, True)
        arrow_symbol = "▼ " if self.folder_expanded[folder] else "► "
        self.folder_label_vars[folder] = tk.StringVar(value=arrow_symbol + folder)

        header_frame = ttk.Frame(self.content_frame)
        # header_frame.grid(row=grid_row, column=0, sticky="ew", padx=3, pady=(2, 0))  
        header_frame.grid(row=grid_row, column=0, sticky="ew", padx=6, pady=(4, 2))  

        folder_label = ttk.Label(
            header_frame, textvariable=self.folder_label_vars[folder],
            font=("Consolas", 10, "bold"),
            cursor="hand2"
        )
        folder_label.grid(row=0, column=0, sticky="w", padx=(0, 10))
        folder_label.bind("<Button-1>", lambda e, f=folder: self.toggle_folder(f))

        folder_context_menu = tk.Menu(self.root, tearoff=0)
        folder_context_menu.add_command(
            label="Add new Node",
            command=lambda f=folder: self.node_window_manager.open_new_node_window(f)
        )
        folder_label.bind('<Button-3>', lambda e, menu=folder_context_menu: self.show_context_menu(e, menu))

        # Smaller add button
        add_node_btn = ttk.Button(
            header_frame, text="➕", bootstyle="info",
            width=3, padding=3,
            command=lambda f=folder: self.node_window_manager.open_new_node_window(f)
        )
        add_node_btn.grid(row=0, column=1, padx=(0, 6))
        self.create_tooltip(add_node_btn, f"Add new node to {folder}")

        # Smaller refresh button
        refresh_btn = ttk.Button(
            header_frame, text="↻", bootstyle="success",
            width=3, padding=3,
            command=lambda f=folder: self.refresh_folder(f)
        )
        refresh_btn.grid(row=0, column=2)
        self.create_tooltip(refresh_btn, f"Refresh {folder} folder")

        folder_content = ttk.Frame(self.content_frame)
        folder_content.grid(row=grid_row + 1, column=0, sticky="ew", padx=20, pady=(0, 10))
        self.folder_frames[folder] = folder_content

        if not self.folder_expanded[folder]:
            folder_content.grid_remove()

        self.populate_folder_content(folder)


    def update_folder_content(self, folder):
        """
        Update the content of a specific folder without affecting other folders.
        """
        if folder not in self.folder_frames:
            return

        # Remember which nodes had active connections
        active_connections = {}
        for identifier, config_data in self.sut_configs.get(folder, {}).items():
            file_path = config_data['file_path']
            if file_path in self.status_labels:
                active_services = {}
                for service_type, info in self.status_labels[file_path].items():
                    if hasattr(info, 'get') and info.get('is_accessible'):
                        active_services[service_type] = True
                if active_services:
                    active_connections[file_path] = active_services

        # Clear existing content
        for widget in self.folder_frames[folder].winfo_children():
            widget.destroy()

        # Repopulate the folder content
        self.populate_folder_content(folder)

        # Only refresh connections if the folder is expanded
        if folder in self.folder_expanded and self.folder_expanded[folder]:
            self.refresh_folder(folder)
            self._on_folder_refreshed()

    def add_console_to_node(self, file_path):
        """Open the node window with existing node data pre-populated"""
        # Extract folder name and identifier from file_path
        sut_dir = Path(__file__).parent.parent / "sut"
        relative_path = Path(file_path).parent.relative_to(sut_dir)
        folder_name = str(relative_path) if str(relative_path) != "." else "Root"
        
        # Load existing config
        try:
            with open(file_path, 'r') as f:
                existing_config = json.load(f)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load node configuration: {str(e)}")
            return
        
        # Open the node window with existing data
        self.node_window_manager.open_new_node_window(folder_name, existing_config=existing_config, file_path=file_path)

    def populate_folder_content(self, folder):
        """
        Helper method to populate a folder with configuration items
        """
        if folder not in self.folder_frames or folder not in self.sut_configs:
            return
        
        # Calculate the maximum number of service types to determine button column positioning
        max_service_types = 0
        for identifier, config_data in self.sut_configs.get(folder, {}).items():
            service_count = 0
            for key in config_data:
                if key not in ['file_path', 'identifier'] and not key.endswith('_GUESS') and not key.endswith('_WHAT'):
                    if config_data.get(key):
                        service_count += 1
            max_service_types = max(max_service_types, service_count)

        button_start_column = max_service_types + 3

        # Create a row for each configuration item

        for identifier, config_data in self.sut_configs.get(folder, {}).items():
            file_path = config_data['file_path']
            display_text = identifier
            is_truncated = len(display_text) > 18
            display_text = identifier if not is_truncated else display_text[:15] + "..."

            key = (folder, identifier)
            initial_bootstyle = "primary" if key in self.selected_nodes else "default"
            row_frame = ttk.Frame(self.folder_frames[folder], bootstyle=initial_bootstyle)
            row_frame.pack(fill=tk.X, pady=0)
            
            # Register for selection logic
            self.selected_node_frames[key] = row_frame

            led_status_label = ttk.Label(
                row_frame, text="◉", font=("Consolas", 14), foreground="#FFA726"
            )
            led_status_label.grid(row=0, column=0, padx=(0, 3))

            self.led_status_labels[file_path] = led_status_label
            self.create_tooltip(led_status_label, f"LED Status: Checking...")

            ip_label = ttk.Label(
                row_frame, text=display_text, font=("Consolas", 10, "bold"),
                cursor="hand2", anchor='w', width=20
            )
            ip_label.grid(row=0, column=1, padx=(5,10), sticky='w')
            ip_label.bind('<Double-Button-1>', lambda e, fp=file_path: self.edit_settings_file(fp))

            # Store the ip_label widget for later access in _update_node_selection_visuals
            self.ip_label_widgets[key] = ip_label # Add this line

            if is_truncated:
                self.create_tooltip(ip_label, identifier)

            # Context menu based on folder type
            context_menu = tk.Menu(self.root, tearoff=0)
            context_menu.add_command(label="Edit Settings", command=lambda fp=file_path: self.edit_settings_file(fp))
            context_menu.add_command(label="Add other Consoles", command=lambda fp=file_path: self.add_console_to_node(fp))
            context_menu.add_command(label="Rename Node", command=lambda fp=file_path: self.rename_node(fp))
            if folder == 'Favorites':
                context_menu.add_command(label="Remove from Favorites", command=lambda fp=file_path: self.remove_from_favorites(fp))
            else:
                context_menu.add_command(label="Add to Favorites", command=lambda fp=file_path: self.add_to_favorites(fp))
            context_menu.add_command(label="Duplicate Node", command=lambda fp=file_path: self.duplicate_node(fp))
            context_menu.add_command(label="Delete Node", command=lambda fp=file_path: self.delete_node(fp))
            ip_label.bind('<Button-3>', lambda e, menu=context_menu: self.show_context_menu(e, menu))

            self.checkbuttons[file_path] = {}
            self.status_labels[file_path] = {}

            # Determine service types to display
            default_service_types = ['RM', 'SOC', 'SUT']
            detected_service_types = []
            
            for key_config in config_data:
                if key_config in ['file_path', 'identifier']:
                    continue
                if not key_config.endswith('_GUESS') and not key_config.endswith('_WHAT'):
                    service_type = key_config
                    if service_type not in detected_service_types and config_data.get(service_type):
                        detected_service_types.append(service_type)

            # Prioritize default service types
            service_types = []
            for service in default_service_types:
                if service in detected_service_types:
                    service_types.append(service)
                    detected_service_types.remove(service)

            service_types.extend(sorted(detected_service_types))

            # Create UI elements for each service type
            for col_idx, service_type in enumerate(service_types, start=2):
                service_ip = config_data.get(service_type);
                if not service_ip:
                    continue

                username = config_data.get(f'{service_type}_GUESS')
                password = config_data.get(f'{service_type}_WHAT')

                service_frame = ttk.Frame(row_frame)
                service_frame.grid(row=0, column=col_idx, padx=(3, 8), pady=1, sticky='w')  # Reduced padding

                
                status_label = ttk.Label(
                    service_frame, text="●", font=("Consolas", 14), bootstyle="warning"
                )
                status_label.pack(side=tk.LEFT, padx=(0,2))  # Reduced padding

                self.status_labels[file_path][service_type] = {
                    'label': status_label,
                    'ip': service_ip,
                    'username': username,
                    'password': password,
                    'tooltip': f"Status: Checking...\nIP: {service_ip}",
                    'identifier': identifier
                }
                self.create_tooltip(status_label, f"Status: Checking...\nIP: {service_ip}")

                cb_var = tk.BooleanVar(value=False)
                self.checkbuttons[file_path][service_type] = cb_var
                if file_path not in self.manual_checkbox_state:
                    self.manual_checkbox_state[file_path] = {}

                self.manual_checkbox_state[file_path][service_type] = None

                def make_toggle_handler(fp, st, var):
                    def handler(*args):
                        self.manual_checkbox_state[fp][st] = var.get()
                    return handler

                cb_var.trace_add("write", make_toggle_handler(file_path, service_type, cb_var))


                cb = ttk.Checkbutton(
                    service_frame, text=service_type, variable=cb_var, bootstyle="info-round-toggle", state="disabled"
                ) 
                cb.pack(side=tk.LEFT)

            # Spacer to push buttons to the right
            spacer_frame = ttk.Frame(row_frame)
            spacer_frame.grid(row=0, column=button_start_column-1, padx=(20, 25), pady=1, sticky='ew')
            row_frame.grid_columnconfigure(button_start_column-1, weight=1, minsize=50)

            # Create button frame
            button_frame = ttk.Frame(row_frame)
            button_frame.grid(row=0, column=button_start_column, padx=(5, 10), pady=1, sticky='e')

           
            ssh_btn = ttk.Button(
                button_frame, text="SSH", bootstyle="info", padding=(8, 4),
                command=lambda fp=file_path: self.connect_selected(fp)
            )
            ssh_btn.pack(side=tk.LEFT, padx=(0, 3))

            controller_btn = ttk.Button(
                button_frame, text="Control", bootstyle="info", padding=(8, 4),
                command=lambda fp=file_path: self.connect_controller(fp)
            )
            controller_btn.pack(side=tk.LEFT, padx=(3, 0))  # Reduced padding

            # --- Click binding ---
            def bind_recursive(widget):
                if "TFrame" in widget.winfo_class() or "TLabel" in widget.winfo_class():
                    widget.bind('<Button-1>', lambda e, f=folder, i=identifier, rf=row_frame: self._on_node_row_click(e, f, i, rf))
                
                for child in widget.winfo_children():
                    bind_recursive(child)
            
            bind_recursive(row_frame)

            # Add a subtle separator after each row (except the last one)
            if list(self.sut_configs.get(folder, {}).keys()).index(identifier) < len(self.sut_configs.get(folder, {})) - 1:
                separator = ttk.Separator(self.folder_frames[folder], orient='horizontal')
                separator.pack(fill=tk.X, padx=10, pady=1)

    def toggle_folder(self, folder):
        frame = self.folder_frames[folder]
        currently_expanded = self.folder_expanded[folder]
        if currently_expanded:
            frame.grid_remove()
            self.folder_label_vars[folder].set("► " + folder)
            self.folder_expanded[folder] = False
        else:
            frame.grid()
            self.folder_label_vars[folder].set("▼ " + folder)
            self.folder_expanded[folder] = True

        self.folder_expanded_state[folder] = self.folder_expanded[folder]
        
        # Update scrollbar visibility after toggling folder
        self.root.after(100, self.update_scrollbar)

    def edit_settings_file(self, file_path):
        settings_file = Path(file_path)
        try:
            if self.notepadpp_path and os.path.exists(self.notepadpp_path):
                subprocess.Popen([self.notepadpp_path, str(settings_file)])
            else:
                subprocess.Popen(['notepad', str(settings_file)])
        except Exception as e:
            print(f"Error opening settings file {settings_file}: {e}")

    def show_context_menu(self, event, menu):
        try:
            menu.tk_popup(event.x_root, event.y_root)
        finally:
            menu.grab_release()

    def create_tooltip(self, widget, text):
        def show_tooltip(event):
            tooltip = tk.Toplevel(self.root)
            tooltip.wm_overrideredirect(True)
            tooltip.wm_geometry(f"+{event.x_root + 10}+{event.y_root + 10}")
            tooltip.configure(bg="#000000")  # fallback bg

            frame = ttk.Frame(tooltip, bootstyle="secondary")  
            frame.pack()

            label = ttk.Label(
                frame,
                text=text,
                bootstyle="inverse", 
                justify="left",
                padding=6,
                wraplength=250
            )
            label.pack()

            def hide_tooltip():
                tooltip.destroy()

            widget.tooltip = tooltip
            widget.bind('<Leave>', lambda e: hide_tooltip())
            tooltip.bind('<Leave>', lambda e: hide_tooltip())

        widget.bind('<Enter>', show_tooltip)



    def start_background_checker(self):
        def check_all_connections():
            with ThreadPoolExecutor() as executor:
                futures = []
                for folder, expanded in self.folder_expanded.items():
                    if expanded:
                        for identifier, config_data in self.sut_configs.get(folder, {}).items():
                            file_path = config_data['file_path']
                            node_key = file_path
                            if node_key in self.status_labels:
                                for service_type, info in self.status_labels[node_key].items():
                                    futures.append(
                                        executor.submit(
                                            self.check_single_connection,
                                            identifier,
                                            file_path,
                                            service_type,
                                            info['ip'],
                                            info['username'],
                                            info['password']
                                        )
                                    )
                for future in as_completed(futures):
                    try:
                        result = future.result()
                        self.status_queue.put(result)
                    except Exception as e:
                        print(f"Error in connection check: {e}")

        def process_batch_ui_updates():
            updates = []
            try:
                while not self.status_queue.empty():
                    updates.append(self.status_queue.get_nowait())
                for update in updates:
                    node_key = update['node_id']
                    service_type = update['service_type']
                    is_accessible = update['is_accessible']
                    status_msg = update['status_msg']
                    if node_key in self.status_labels:
                        status_info = self.status_labels[node_key][service_type]
                        status_label = status_info['label']
                        if not status_label.winfo_exists():
                            continue
                        # status_label.configure(bootstyle="success" if is_accessible else "danger")
                        status_label.configure(foreground="#00E676" if is_accessible else "#FF1744")
                        status_info['tooltip'] = f"Status: {status_msg}\nIP: {status_info['ip']}"
                        self.create_tooltip(status_label, status_info['tooltip'])

                        # Handling checkboxes with manual toggles by user
                        manual_state = self.manual_checkbox_state.get(node_key, {}).get(service_type)
                        if manual_state is None:
                            try:
                                self.checkbuttons[node_key][service_type].set(is_accessible)
                            except KeyError:
                                print(f"Warning: Missing checkbox for {node_key} / {service_type}")


                        for w in status_label.master.winfo_children():
                            if isinstance(w, ttk.Checkbutton) and w.cget('text') == service_type:
                                # w.configure(state="normal" if is_accessible else "disabled")
                                w.configure(state="normal") # --> Keeping the checkboxes always clickable
                                break
            except queue.Empty:
                pass
            self.root.after(200, process_batch_ui_updates)

        threading.Thread(target=check_all_connections, daemon=True).start()
        self.root.after(200, process_batch_ui_updates)

    def check_single_connection(self, identifier, file_path, service_type, ip, username, password):
        if not all([ip, username, password]):
            return {
                'node_id': file_path,
                'file_path': file_path,
                'service_type': service_type,
                'is_accessible': False,
                'status_msg': "Missing credentials"
            }
        # is_accessible, status_msg = self.check_ssh_connection(ip, username, password)
        is_accessible, status_msg = self.check_ssh_connection(
            ip, username, password, identifier=identifier, service_type=service_type
        )

        return {
            'node_id': file_path,
            'file_path': file_path,
            'service_type': service_type,
            'is_accessible': is_accessible,
            'status_msg': status_msg
        }

    def check_ssh_connection(self, ip, username, password,
                            identifier="", service_type="", max_retries=2):
        context = f"[Node: {identifier} | Console: {service_type} | IP: {ip}]"
        print(f"\n{context} Starting connection check…")

        if not ip:
            print("\nNo IP provided")
            return False, "No IP provided"

        # ping_cmd = (["ping", "-n", "1", "-w", "1000", ip]
        #             if os.name == 'nt' else ["ping", "-c", "1", ip])

        # prioritizing ipv4 for SSH connections
        ping_cmd = (["ping", "-4", "-n", "1", "-w", "1000", ip]
            if os.name == 'nt' else ["ping", "-4", "-c", "1", ip])


        ping_success = False
        for attempt in range(max_retries + 1):
            try:
                print(f"\n{context} Pinging (Attempt {attempt+1})…")
                resp = subprocess.run(ping_cmd,
                                    stdout=subprocess.PIPE,
                                    stderr=subprocess.PIPE,
                                    timeout=5)
                if resp.returncode == 0:
                    ping_success = True
                    break
            except subprocess.TimeoutExpired:
                print(f"\n{context} Ping timed out (Attempt {attempt+1})")

            if attempt < max_retries:
                time.sleep(0.5 * (2 ** attempt))

        if not ping_success:
            print(f"\n{context} Ping failed after {max_retries+1} attempts — proceeding with SSH anyway")

        
        # Resolving the IP address to an IPv4 address
        try:
            ipv4_address = socket.gethostbyname(ip)
        except socket.gaierror as e:
            print(f"\n{context} Failed to resolve IPv4 address: {e}")
            return False, f"Failed to resolve IPv4 address: {e}"


        print(f"\n{context} Attempting SSH connection…")
        ssh = paramiko.SSHClient()
        ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())
        try:
            ssh.connect(
                # hostname=ip,
                hostname=ipv4_address,
                port=22,
                username=username,
                password=password,
                timeout=5,
                banner_timeout=5
            )
            ssh.get_transport().set_keepalive(30)
            print(f"\n{context} SSH connection successful")
            return True, "SSH Connected"
        except (paramiko.SSHException,
                paramiko.ssh_exception.NoValidConnectionsError,
                TimeoutError,
                OSError) as e:
            print(f"\n{context} SSH connection failed: {e}")
            return False, f"SSH connection failed: {e}"
        finally:
            ssh.close()



    def refresh_folder(self, folder, is_auto_check=False):
        """
        Refresh a folder by reloading configurations and rebuilding UI if needed.
        Targeted refresh for other remaining folders.
        """
        # print(f"DEBUG: Starting refresh for folder: {folder}")
        if not is_auto_check:
            print(f"DEBUG: Starting refresh for folder: {folder}")

        sut_dir = Path(__file__).parent.parent / "sut"

        # Handle folder that might not exist in configs yet
        if folder not in self.sut_configs:
            if folder == "Favorites":
                folder_path = sut_dir / "Favorites"
            elif folder == "Root":
                folder_path = sut_dir
            else:
                folder_path = sut_dir / folder

            # Create folder if it doesn't exist
            folder_path.mkdir(exist_ok=True)
            
            # Initialize empty config for this folder
            self.sut_configs[folder] = {}
            
            # Make sure the folder is present in UI
            if folder not in self.folder_frames:
                # Find the position to insert the new folder
                folders_in_order = ['Favorites'] + sorted(
                    [f for f in self.sut_configs if f != 'Favorites'],
                    key=lambda x: (x != 'Root', x)
                )
                insert_pos = folders_in_order.index(folder)
                
                # Add the folder section to UI
                grid_row = insert_pos * 2
                self.add_folder_section(folder, grid_row)
                
                # Rearrange subsequent folders if needed
                for i, f in enumerate(folders_in_order[insert_pos+1:], start=insert_pos+1):
                    if f in self.folder_frames:
                        frame = self.folder_frames[f].master
                        frame.grid(row=i*2, column=0)
                        self.folder_frames[f].grid(row=i*2+1, column=0)
                
            if not is_auto_check:
                print(f"DEBUG: Created new folder {folder}")
            # return

        # Store references to current configurations
        old_configs = self.sut_configs.get(folder, {}).copy()
        print(f"DEBUG: Found {len(old_configs)} existing configs in folder {folder}")
        
        # Determine folder path based on folder name
        if folder == "Favorites":
            folder_path = sut_dir / "Favorites"
        elif folder == "Root":
            folder_path = sut_dir
        else:
            folder_path = sut_dir / folder
        
        # Find all config files in this folder
        sut_files = glob.glob(str(folder_path / "settings.*.json"))
        print(f"DEBUG: Found {len(sut_files)} config files in {folder_path}")
        
        # Load each config file in this folder
        new_folder_configs = {}
        for file_path in sut_files:
            try:
                file_path_obj = Path(file_path)
                if "settings." in file_path_obj.stem:
                    identifier = file_path_obj.stem.split("settings.")[1]
                else:
                    identifier = file_path_obj.stem

                with open(file_path, 'r') as f:
                    config = json.load(f)

                config_data = {
                    'file_path': file_path,
                    'identifier': identifier
                }

                # Extract service types
                service_types = set()
                for key in config:
                    if key.endswith('_IP') or key.endswith('_GUESS') or key.endswith('_WHAT'):
                        service_type = key.split('_')[0]
                        service_types.add(service_type)

                print(f"DEBUG: Detected service types for {identifier}: {service_types}")

                for service_type in service_types:
                    ip_key = f"{service_type}_IP"
                    service_ip = config.get(ip_key)

                    if service_ip is None and service_type in config:
                        service_ip = config.get(service_type)
                        config_data[service_type] = service_ip
                    else:
                        config_data[service_type] = service_ip

                    # Handle _GUESS (username) with environment variable fallback
                    guess_value = config.get(f"{service_type}_GUESS")
                    if not guess_value:  # If empty or None in config
                        guess_value = get_credential_from_env(service_type, 'GUESS')
                        if guess_value:
                            print(f"   Using environment variable for {service_type}_GUESS in {identifier}")
                    config_data[f"{service_type}_GUESS"] = guess_value

                    # Handle _WHAT (password) with environment variable fallback  
                    what_value = config.get(f"{service_type}_WHAT")
                    if not what_value:  # If empty or None in config
                        what_value = get_credential_from_env(service_type, 'WHAT')
                        if what_value:
                            print(f"   Using environment variable for {service_type}_WHAT in {identifier}")
                    config_data[f"{service_type}_WHAT"] = what_value

                new_folder_configs[identifier] = config_data
            except Exception as e:
                print(f"DEBUG: Error loading {file_path}: {e}")
                continue
        
        # ----- TARGETED UPDATES: ANALYZE WHAT CHANGED -----
        
        # 1. Get list of added nodes (in new configs but not old)
        added_nodes = {id: config for id, config in new_folder_configs.items() 
                    if id not in old_configs}
        
        # 2. Get list of removed nodes (in old configs but not new)
        removed_nodes = {id: config for id, config in old_configs.items() 
                        if id not in new_folder_configs}
        
        # 3. Get list of modified nodes (services changed)
        modified_nodes = {}
        for id, new_config in new_folder_configs.items():
            if id in old_configs:
                old_config = old_configs[id]
                
                # Compare service sets
                old_services = {k for k in old_config 
                            if k not in ['file_path', 'identifier'] 
                            and not k.endswith('_GUESS') 
                            and not k.endswith('_WHAT')
                            and old_config.get(k)}
                            
                new_services = {k for k in new_config 
                            if k not in ['file_path', 'identifier'] 
                            and not k.endswith('_GUESS') 
                            and not k.endswith('_WHAT')
                            and new_config.get(k)}
                
                if old_services != new_services:
                    modified_nodes[id] = new_config
        
        # Update the master configuration dictionary with new data
        self.sut_configs[folder] = new_folder_configs
        
        # ----- TARGETED UI UPDATES -----
        
        # Handle case of an empty folder
        if not new_folder_configs and folder in self.folder_frames:
            # Clear any existing content but keep the folder itself
            for widget in self.folder_frames[folder].winfo_children():
                widget.destroy()
            print(f"DEBUG: Folder {folder} is now empty")
            return
            
        # Check if we need to create the folder frames
        if folder not in self.folder_frames:
            print(f"DEBUG: Creating UI for new folder {folder}")
            folders_in_order = ['Favorites'] + sorted(
                [f for f in self.sut_configs if f != 'Favorites'],
                key=lambda x: (x != 'Root', x)
            )
            insert_pos = folders_in_order.index(folder)
            grid_row = insert_pos * 2
            self.add_folder_section(folder, grid_row)
            
        # ----- PERFORM TARGETED UPDATES -----
        
        # need_rebuild = len(added_nodes) > 0 or len(removed_nodes) > 0 or len(modified_nodes) > 0

        # need_rebuild = (not is_auto_check) and (
        #     len(added_nodes) > 0 or len(removed_nodes) > 0 or len(modified_nodes) > 0
        # )

        need_rebuild = (
            (not is_auto_check and (len(added_nodes) > 0 or len(removed_nodes) > 0 or len(modified_nodes) > 0))
            or (is_auto_check and len(added_nodes) > 0)  # <-- allow UI population for new SUTs during auto-check
        )
        
        if need_rebuild:
            print(f"DEBUG: Performing targeted updates for folder {folder}")
            print(f"DEBUG: Added: {len(added_nodes)}, Removed: {len(removed_nodes)}, Modified: {len(modified_nodes)}")
            # Remember current selections to restore after rebuild
            preserved_selections = self.selected_nodes.copy()
            
            # Only rebuild this specific folder's content
            if folder in self.folder_frames:
                # Clear this folder's frame
                for widget in self.folder_frames[folder].winfo_children():
                    widget.destroy()
                
                # Clean up tracking data for removed nodes
                for id, config in removed_nodes.items():
                    file_path = config['file_path']
                    if file_path in self.status_labels:
                        del self.status_labels[file_path]
                    if file_path in self.checkbuttons:
                        del self.checkbuttons[file_path]
                    if file_path in self.led_status_labels:
                        del self.led_status_labels[file_path]
                    if file_path in self.manual_checkbox_state:
                        del self.manual_checkbox_state[file_path]
                
                # Rebuild only this folder's content
                self.populate_folder_content(folder)
                self.selected_nodes = preserved_selections
                self._update_node_selection_visuals()
        else:
            if not is_auto_check:
                print(f"DEBUG: No changes detected in folder {folder}, just refreshing statuses")
            # Reset status indicators for existing services to indicate checking
            for identifier, config_data in self.sut_configs[folder].items():
                file_path = config_data['file_path']
                if file_path in self.status_labels:
                    self.reset_status_indicators(file_path)
                # Reset LED status indicators if they exist
                if file_path in self.led_status_labels:
                    led_label = self.led_status_labels[file_path]
                    led_label.configure(foreground="#FFA726")  # Orange/amber for "checking"
                    self.create_tooltip(led_label, "LED Status: Checking...")
        
        # ----- START CONNECTION CHECKS -----
        
        # Only check connections if the folder is expanded
        if folder in self.folder_expanded and self.folder_expanded[folder]:
            def check_folder_connections():
                with ThreadPoolExecutor() as executor:
                    futures = []
                    for identifier, config_data in self.sut_configs[folder].items():
                        file_path = config_data['file_path']
                        
                        # Handle connection status checks
                        if file_path in self.status_labels:
                            for service_type, info in self.status_labels[file_path].items():
                                service_ip = config_data.get(service_type)
                                username = config_data.get(f'{service_type}_GUESS')
                                password = config_data.get(f'{service_type}_WHAT')
                                
                                if service_ip:
                                    self.status_labels[file_path][service_type].update({
                                        'ip': service_ip,
                                        'username': username,
                                        'password': password
                                    })
                                    
                                    futures.append(
                                        executor.submit(
                                            self.check_single_connection,
                                            identifier,
                                            file_path,
                                            service_type,
                                            service_ip,
                                            username,
                                            password
                                        )
                                    )
                                    if not is_auto_check:
                                        print(f"DEBUG: Checking connection for {identifier}, service {service_type}, IP {service_ip}")
                        
                        # Handle LED status checks
                        if file_path in self.led_status_labels:
                            futures.append(
                                executor.submit(
                                    self.led_status_manager.check_led_status_background,
                                    file_path
                                )
                            )
                            if not is_auto_check:
                                print(f"DEBUG: Checking LED status for {identifier}")
                    
                    # Process all futures
                    for future in as_completed(futures):
                        try:
                            result = future.result()
                            if result:
                                if 'service_type' in result:
                                    self.status_queue.put(result)
                                elif 'status' in result:
                                    self.led_status_queue.put(result)
                        except Exception as e:
                            print(f"DEBUG: Error in connection check: {e}")

            threading.Thread(target=check_folder_connections, daemon=True).start()
        
        self.root.after(300, self.update_scrollbar)
        self._on_folder_refreshed()


    def reset_status_indicators(self, file_path):
        """Reset status indicators to 'checking' state."""
        for service_type, info in self.status_labels[file_path].items():
            info['label'].configure(foreground="#FFA726")
            info['tooltip'] = f"Status: Checking...\nIP: {info['ip']}"
            self.create_tooltip(info['label'], info['tooltip'])


    def start_led_status_checker(self):
        """Start the LED status checker and process updates."""
        # Start the periodic check in the LED manager
        self.led_status_manager.start_periodic_check(
            self.sut_configs, 
            self.folder_expanded, 
            self.led_status_labels
        )
        
        # Process LED status updates in the UI
        def process_led_status_updates():
            updates = []
            try:
                while not self.led_status_queue.empty():
                    updates.append(self.led_status_queue.get_nowait())
                for update in updates:
                    file_path = update['file_path']
                    status = update['status']
                    details = update['details']
                    if file_path in self.led_status_labels:
                        led_label = self.led_status_labels[file_path]
                        if not led_label.winfo_exists():
                            continue
                        
                        # Set LED color based on status
                        if status == "GREEN":
                            led_label.configure(foreground="#00E676")  # Green
                        elif status == "RED":
                            led_label.configure(foreground="#FF1744")  # Red
                        else:  # GREY for unknown/missing SKU/connection failures
                            led_label.configure(foreground="#BDBDBD")  # Grey
                    
                        tooltip_text = f"LED Status: {status}\n{details}"
                        self.create_tooltip(led_label, tooltip_text)
            except queue.Empty:
                pass
            self.root.after(200, process_led_status_updates)
            
        # Start processing UI updates
        self.root.after(200, process_led_status_updates)



    def load_themes(self):
        """
        Scan the Theme directory for available themes.
        Each theme should have BGM, Picture, and Sound subfolders.
        """
        theme_dir = Path(__file__).parent.parent / "Theme"
        themes = {"None": None}  # Default option for no theme
        
        if theme_dir.exists():
            for item in theme_dir.iterdir():
                if item.is_dir():
                    # Check if the folder has the required subfolders
                    has_bgm = (item / "BGM").exists()
                    has_picture = (item / "Picture").exists()
                    has_sound = (item / "Sound").exists()
                    
                    if has_bgm or has_picture or has_sound:
                        themes[item.name] = str(item)
        
        return themes
        

    def apply_theme(self, theme_name):
        """
        Apply the selected theme (only to the top bar):
        1. Set the background image for the top bar only
        2. Configure BGM folder
        3. Set sound effect settings
        """
        if theme_name == "None" or not theme_name:
            # Reset to default
            self.settings_frame.configure(style="Custom.TFrame")
            for widget in self.settings_frame.winfo_children():
                if isinstance(widget, tk.Label):
                    widget.configure(bg="#1E1E1E")
                if isinstance(widget, tk.Button):
                    pass  # Keep original button styling
                    
            if hasattr(self, 'top_bg_label'):
                self.top_bg_label.destroy()
                delattr(self, 'top_bg_label')
                
            self.theme_name = None
            self.bgm_var.set(False)
            self.stop_bgm()
            return
        
        theme_path = Path(self.themes[theme_name])
        self.theme_name = theme_name
        
        # 1. Set background image if available (only for the top bar)
        picture_dir = theme_path / "Picture"
        if picture_dir.exists():
            background_images = []
            for ext in ["*.jpg", "*.png", "*.gif"]:
                background_images.extend(list(picture_dir.glob(ext)))
            
            if background_images:
                # Choose a random background image
                bg_image_path = random.choice(background_images)
                try:
                    # Load the image for the top bar
                    bg_image = Image.open(bg_image_path)
                    
                    # Get top bar dimensions
                    self.settings_frame.update_idletasks()
                    bar_width = self.settings_frame.winfo_width() 
                    bar_height = self.settings_frame.winfo_height()
                    
                    # Use reasonable defaults if dimensions are zero
                    if bar_width <= 1:
                        bar_width = self.root.winfo_width()
                    if bar_height <= 1:
                        bar_height = 40  # Default height for top bar
                    
                    # Resize the image to fit the top bar dimensions
                    img_width, img_height = bg_image.size
                    
                    # Take a crop of the image (top portion for visual interest)
                    crop_top = 0
                    crop_height = min(img_height, int(img_height * 0.3))  # Top 30% of image
                    cropped_bg = bg_image.crop((0, crop_top, img_width, crop_height))
                    
                    # Resize the cropped image to fit the top bar
                    resized_bg = cropped_bg.resize((bar_width, bar_height), Image.LANCZOS)
                    
                    # Apply a slight darkening overlay for better text visibility
                    if resized_bg.mode != 'RGBA':
                        resized_bg = resized_bg.convert('RGBA')
                    
                    overlay = Image.new('RGBA', resized_bg.size, (0, 0, 0, 100))  # 40% black
                    darkened_bg = Image.alpha_composite(resized_bg, overlay)
                    
                    # Create the PhotoImage for the top bar
                    self.top_bg_photo = ImageTk.PhotoImage(darkened_bg)
                    
                    # Create or update the background label for the top bar
                    if hasattr(self, 'top_bg_label'):
                        self.top_bg_label.destroy()
                    
                    # Create a background label for the settings frame
                    self.top_bg_label = tk.Label(self.settings_frame, image=self.top_bg_photo)
                    self.top_bg_label.place(x=0, y=0, relwidth=1, relheight=1)
                    self.top_bg_label.lower()  # Send to back
                    
                    # Make settings frame widgets have transparent backgrounds
                    for widget in self.settings_frame.winfo_children():
                        if isinstance(widget, tk.Label):
                            widget.configure(bg="")  # Empty bg for transparency
                        # Keep button styling for better visibility
                    
                    print(f"Applied theme to top bar: {theme_name}")
                except Exception as e:
                    print(f"Error setting top bar background: {e}")
        
        # 2. Configure BGM folder
        bgm_dir = theme_path / "BGM"
        if bgm_dir.exists():
            self.bgm_playlist = []
            files = list(bgm_dir.glob("*.flac")) + list(bgm_dir.glob("*.mp3")) + list(bgm_dir.glob("*.wav"))
            if files:
                self.bgm_playlist = [str(f) for f in files]
                random.shuffle(self.bgm_playlist)
                self.bgm_index = 0
                self.bgm_var.set(True)
                self.bgm_active = True
                self.play_current_bgm()
                print(f"Set BGM folder to theme: {theme_name}")
        
        # 3. Sound effect is enabled when a theme is selected
        self.sound_effect_var.set(True)


    def show_theme_settings(self):
        """
        Open a dialog to select the theme
        """
        settings_dialog = tk.Toplevel(self.root)
        settings_dialog.title("Theme Settings")
        if hasattr(self, 'icon_path') and self.icon_path.exists():
            try:
                settings_dialog.iconbitmap(str(self.icon_path))
            except Exception as e:
                print(f"Failed to set icon on settings dialog: {e}")
        settings_dialog.geometry("400x400")
        settings_dialog.minsize(400,400)
        # settings_dialog.configure(bg='#1E1E1E')
        # settings_dialog.resizable(False, False)
        
        # Set the dialog to be modal
        settings_dialog.transient(self.root)
        settings_dialog.grab_set()
        
        # Center the dialog on the parent window
        parent_x = self.root.winfo_rootx()
        parent_y = self.root.winfo_rooty()
        parent_width = self.root.winfo_width()
        parent_height = self.root.winfo_height()
        
        dialog_width = 400
        dialog_height = 300
        
        x = parent_x + (parent_width - dialog_width) // 2
        y = parent_y + (parent_height - dialog_height) // 2
        
        settings_dialog.geometry(f"+{x}+{y}")
        
        # Theme selection
        theme_frame = ttk.Frame(settings_dialog)
        theme_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        theme_label = ttk.Label(theme_frame, text="Select Theme:")
        theme_label.pack(anchor="w", pady=(0, 10))
        
        # Get current theme
        current_theme = self.theme_name if hasattr(self, 'theme_name') else "None"
        
        # Create theme selection variable
        self.theme_var = tk.StringVar(value=current_theme)
        
        # Create listbox for themes
        theme_listbox = tk.Listbox(
            theme_frame, 
            bg='#2A2A2A', 
            fg='#E0E0E0', 
            selectbackground='#3F51B5',
            selectforeground='#FFFFFF',
            font=("Consolas", 10),
            height=10,
            relief=tk.FLAT
        )
        theme_listbox.pack(fill=tk.BOTH, expand=True)
        
        # Add themes to listbox
        for idx, theme_name in enumerate(self.themes.keys()):
            theme_listbox.insert(tk.END, theme_name)
            if theme_name == current_theme:
                theme_listbox.selection_set(idx)
        
        # Button frame
        button_frame = ttk.Frame(settings_dialog)
        button_frame.pack(fill=tk.X, pady=10, padx=20)
        
        # Apply button
        apply_btn = ttk.Button(
            button_frame, 
            text="Apply", 
            bootstyle="info",
            command=lambda: self.apply_theme_from_dialog(theme_listbox.get(theme_listbox.curselection()), settings_dialog)
        )
        apply_btn.pack(side=RIGHT, padx=5)
        
        # Cancel button
        cancel_btn = ttk.Button(
            button_frame, 
            text="Cancel", 
            bootstyle="secondary",
            command=settings_dialog.destroy
        )
        cancel_btn.pack(side=RIGHT, padx=5)

    def apply_theme_from_dialog(self, theme_name, dialog):
        """Apply the selected theme and close the dialog"""
        if theme_name:
            self.apply_theme(theme_name)
            self.save_theme_setting(theme_name)  # Add this line to save the setting
        dialog.destroy()


    def connect_selected(self, file_path):
        selected = [st for st, var in self.checkbuttons[file_path].items() if var.get()]
        if not selected:
            return

        try:
            with open(file_path, 'r') as f:
                settings = json.load(f)
        except Exception as e:
            print(f"Error reading settings file {file_path}: {e}")
            return

        identifier = Path(file_path).stem.split("settings.")[1]
        sound_value = "true" if self.sound_effect_var.get() else "false"
        
        # Pass theme name if one is selected
        theme_name = getattr(self, 'theme_name', None)

        def launch_one(idx):
            if idx >= len(selected):
                return
            st = selected[idx]
            
            # Check if jumpbox is enabled and configured
            if self.jumpbox_config["enabled"] and self.jumpbox_config["jumpbox_ip"]:
                # For jumpbox connections, we need the actual target IP
                target_ip = self.get_target_ip_from_settings(settings, st)
                if not target_ip:
                    print(f"No IP found for service type {st} in jumpbox mode")
                    print(f"Available keys in settings: {list(settings.keys())}")
                    self.root.after(10000, lambda: launch_one(idx+1))
                    return
                
                # Use MobaXterm with jumpbox
                # self.launch_mobaxterm_with_jumpbox(target_ip, identifier, st)

                cmd = [
                    python_exe, "shell_mx.py",
                    "--console_type", st,
                    "--sut_ip", identifier,
                    "--sound", sound_value,
                    "-j",  # Added the jumpbox parameter
                    f"{self.jumpbox_config['jumpbox_username']}@{self.jumpbox_config['jumpbox_ip']}:{self.jumpbox_config['jumpbox_port']}"
                ]
            else:
                # Use existing method (original behavior)
                cmd = [
                    python_exe, "shell_mx.py",
                    "--console_type", st,
                    "--sut_ip", identifier,
                    "--sound", sound_value
                ]
                
            # Add theme parameter if available
            if theme_name:
                cmd.extend(["--theme", theme_name])
                
            proc = subprocess.Popen(
                cmd,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                creationflags=NO_WINDOW,
                universal_newlines=True
            )
            self.track_process(proc)
            threading.Thread(
                target=_relay_output,
                args=(proc, f"[{identifier}/{st}] "),
                daemon=True
            ).start()
            
            # Schedule next launch
            self.root.after(10000, lambda: launch_one(idx+1))

        launch_one(0)

    def get_target_ip_from_settings(self, settings, service_type):
        """Extract target IP from settings based on service type"""
        # Try different possible key formats based on your JSON structure
        possible_keys = [
            f"{service_type}_IP",  # RM_IP, SUT_IP, SOC_IP
            service_type,          # RM, SUT, SOC (fallback)
            f"{service_type.lower()}_ip",  # rm_ip, sut_ip, soc_ip
            f"{service_type.upper()}_IP"   # Ensure uppercase
        ]
        
        for key in possible_keys:
            if key in settings and settings[key]:
                return settings[key]
        
        return None


    def connect_controller(self, file_path):
        try:
            with open(file_path, 'r') as f:
                settings = json.load(f)
        except Exception as e:
            print(f"Error reading settings file {file_path}: {e}")
            return

        identifier = Path(file_path).stem.split("settings.")[1]
        sound_value = "true" if self.sound_effect_var.get() else "false"
        starting_dir = str(Path(__file__).parent.absolute())
        
        # Pass theme name if one is selected
        theme_name = getattr(self, 'theme_name', None)
        
        cmd = [
            python_exe, "shell_mx.py",
            "--console_type", "controller",
            "--sut_ip", identifier,
            "--sound", sound_value
        ]
        
        # Add theme parameter if available
        if theme_name:
            cmd.extend(["--theme", theme_name])
            
        proc = subprocess.Popen(cmd, cwd=starting_dir)
        self.track_process(proc)

    def connect_winscp(self, file_path):
        selected = [st for st, var in self.checkbuttons[file_path].items() if var.get()]
        if not selected:
            return

        procs = []
        try:
            with open(file_path, 'r') as f:
                settings = json.load(f)
            for service_type in selected:
                if service_type == 'RM':
                    target_ip = settings.get('RM_IP')
                    username = settings.get('RM_GUESS')
                    password = settings.get('RM_WHAT')
                elif service_type == 'SOC':
                    target_ip = settings.get('SOC_IP')
                    username = settings.get('SOC_GUESS')
                    password = settings.get('SOC_WHAT')
                elif service_type == 'SUT':
                    target_ip = settings.get('SUT_IP')
                    username = settings.get('SUT_GUESS')
                    password = settings.get('SUT_WHAT')
                else:
                    continue

                if not all([target_ip, username, password]):
                    print(f"Missing credentials for {service_type} in {file_path}")
                    continue

                encoded_password = urllib.parse.quote(password, safe="")
                session_url = f"sftp://{username}:{encoded_password}@{target_ip}/"
                proc = subprocess.Popen([self.winscp_path, session_url])
                procs.append(proc)
                print(f"Launched WinSCP for {service_type} in {file_path}")
        except Exception as e:
            print(f"Error launching WinSCP for {file_path}: {e}")
        finally:
            # track them all
            self.child_processes.extend(procs)


    def toggle_bgm(self):
        if self.bgm_var.get():
            self.start_bgm()
        else:
            self.stop_bgm()

    def start_bgm(self):
        bgm_folder = Path(__file__).parent.parent / "sound" / "BGM"
        files = list(bgm_folder.glob("*.flac")) + list(bgm_folder.glob("*.mp3")) + list(bgm_folder.glob("*.wav"))
        if not files:
            print("No background music files found in:", bgm_folder)
            return
        self.bgm_playlist = [str(f) for f in files]
        random.shuffle(self.bgm_playlist)
        self.bgm_index = 0
        self.bgm_active = True
        self.play_current_bgm()

    def play_current_bgm(self):
        if not self.audio_enabled or not self.bgm_active or not self.bgm_playlist:
            return
        current_track = self.bgm_playlist[self.bgm_index]
        try:
            pygame.mixer.music.load(current_track)
            pygame.mixer.music.set_volume(0.3)  # Set volume to 30%
            pygame.mixer.music.play()
        except Exception as e:
            print(f"Error playing {current_track}: {e}")
        self.root.after(1000, self.check_bgm_end)
        
    def check_bgm_end(self):
        if not self.audio_enabled or not self.bgm_active:
            return
        if not pygame.mixer.music.get_busy():
            self.bgm_index = (self.bgm_index + 1) % len(self.bgm_playlist)
            self.play_current_bgm()
        else:
            self.root.after(1000, self.check_bgm_end)

    def stop_bgm(self):
        self.bgm_active = False
        if self.audio_enabled:
            pygame.mixer.music.stop()

    def run(self):
        self.root.update_idletasks()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry(f'+{x}+{y}')
        self.root.mainloop()

def reset_helper_window_counter():
    counter_path = Path(__file__).parent / "helper_window_counter.txt"
    if counter_path.exists():
        counter_path.unlink()
        print("Helper window counter reset.")

def main():
    
    # --- enforce running from C:\mcqueen\console only ---
    script_path = Path(__file__).resolve()
    expected_dir = Path(r"C:\mcqueen\console").resolve()
    if script_path.parent.resolve() != expected_dir:
        print(
            "ERROR: McConsole must be located in C:\\mcqueen\\console.\n"
            "Please place the mcconsole*.zip under C:\\mcqueen and extract it so\n"
            "that you end up with this script at C:\\mcqueen\\console."
        )
        # give the user a chance to read the message
        if os.name == 'nt':
            os.system('pause')
        else:
            input("Press Enter to exit...")
        sys.exit(1)
    # --- end location check ---

    
    initialize_mobaxterm()
    reset_helper_window_counter()
    app = SessionManager()
    app.run()

if __name__ == "__main__":
    reset_helper_window_counter()

    
    main()