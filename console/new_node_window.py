import tkinter as tk
# from tkinter import ttk, messagebox
import re
import ttkbootstrap as ttk
from tkinter import messagebox
from pathlib import Path
import json
from ttkbootstrap.style import Style

class NodeWindowManager:
    def __init__(self, root, refresh_callback=None, icon_path=None):
        
        self.root = root
        self.refresh_callback = refresh_callback
        self.icon_path = icon_path  
        self.current_windows = []

        self.MAX_CONSOLES = 5
        

    def style_entry_by_theme(self, entry_widget, placeholder=False):
        style = Style()
        theme = style.theme.name

        # Dark themes from ttkbootstrap
        dark_themes = {
            "darkly", "superhero", "cyborg", "solar", "vapor"
        }

        is_dark = theme in dark_themes

        if placeholder:
            fg_color = "#999999" if is_dark else "#666666"
        else:
            fg_color = "#FFFFFF" if is_dark else "#000000"

        entry_widget.config(foreground=fg_color)


    
    def open_new_node_window(self, folder_name, existing_config=None, file_path=None):
       
        top = ttk.Toplevel(self.root)
        if existing_config:
            top.title(f"Edit Node - {folder_name}")
            title_text = f"Edit Node in {folder_name}"
        else:
            top.title(f"Add New Node - {folder_name}")
            title_text = f"Add Node to {folder_name}"
        # setting icon for this window
        if hasattr(self, 'icon_path') and self.icon_path.exists():
            try:
                top.iconbitmap(str(self.icon_path))
            except Exception as e:
                print(f"Failed to set icon on new node dialog: {e}")
        top.geometry("800x650") # Changed default window size
        top.minsize(800, 650)
        
        top.resizable(True, True)

        self.current_windows.append(top)

        # Add some padding
        main_frame = ttk.Frame(top)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        # Title label
        title_label = ttk.Label(
            main_frame, 
            text=title_text, 
            font=("Arial", 14, "bold")
        )
        title_label.pack(pady=(0, 15))
        
        entries = {}
        
        # Identifier frame with validation
        id_frame = ttk.Frame(main_frame)
        id_frame.pack(fill=tk.X, pady=(5, 10))
        
        ttk.Label(
            id_frame, 
            text="Node Identifier or IP: *",
            font=("Arial", 10, "bold")
        ).pack(anchor=tk.W)
        
        id_entry = ttk.Entry(id_frame, width=40)
        id_entry.pack(fill=tk.X, pady=(5, 0))

        # Handle existing config or default placeholder
        if existing_config:
            # Extract identifier from file_path and pre-populate
            identifier = Path(file_path).stem.split("settings.")[1]
            id_entry.insert(0, identifier)
            # id_entry.config(foreground="#FFFFFF")
            self.style_entry_by_theme(id_entry, placeholder=False)
            id_entry.config(state="disabled")
        else:
            id_entry.insert(0, "e.g. 192.168.1.100 or server-name")
            # id_entry.config(foreground="#999999")
            self.style_entry_by_theme(id_entry, placeholder=True)

        entries['identifier'] = id_entry
        
        # Placeholder handling for identifier
        def on_id_entry_focus_in(event):
            if id_entry.get() == "e.g. 192.168.1.100 or server-name":
                id_entry.delete(0, tk.END)
                # id_entry.config(foreground="#FFFFFF")
                self.style_entry_by_theme(id_entry, placeholder=False)

        
        def on_id_entry_focus_out(event):
            if not id_entry.get():
                id_entry.insert(0, "e.g. 192.168.1.100 or server-name")
                # id_entry.config(foreground="#999999")
                self.style_entry_by_theme(id_entry, placeholder=True)

        if not existing_config:
            id_entry.bind("<FocusIn>", on_id_entry_focus_in)
            id_entry.bind("<FocusOut>", on_id_entry_focus_out)
        

        # SUT_SKU field (required for LED status)
        sku_frame = ttk.Frame(main_frame)
        sku_frame.pack(fill=tk.X, pady=(0, 10))
        ttk.Label(
            sku_frame,
            text="SUT SKU: *",
            font=("Arial", 10, "bold")
        ).pack(anchor=tk.W)
        sku_entry = ttk.Entry(sku_frame, width=40)
        sku_entry.pack(fill=tk.X, pady=(5, 0))
        sku_entry.insert(0, "e.g. CAAA, CAAAA, DAAAA")
        self.style_entry_by_theme(sku_entry, placeholder=True)
        entries['sut_sku'] = sku_entry
        # Placeholder handling for SUT_SKU
        def on_sku_entry_focus_in(event):
            if sku_entry.get() == "e.g. CAAA, CAAAA, DAAAA":
                sku_entry.delete(0, tk.END)
                self.style_entry_by_theme(sku_entry, placeholder=False)
        def on_sku_entry_focus_out(event):
            if not sku_entry.get():
                sku_entry.insert(0, "e.g. CAAA, CAAAA, DAAAA")
                self.style_entry_by_theme(sku_entry, placeholder=True)
        sku_entry.bind("<FocusIn>", on_sku_entry_focus_in)
        sku_entry.bind("<FocusOut>", on_sku_entry_focus_out)

        # SUT_OS field (new, after SKU)
        os_frame = ttk.Frame(main_frame)
        os_frame.pack(fill=tk.X, pady=(0, 10))
        ttk.Label(
            os_frame,
            text="SUT OS:",
            font=("Arial", 10, "bold")
        ).pack(anchor=tk.W)
        os_entry = ttk.Entry(os_frame, width=40)
        os_entry.pack(fill=tk.X, pady=(5, 0))
        os_entry.insert(0, "e.g. linux, windows")
        self.style_entry_by_theme(os_entry, placeholder=True)
        entries['sut_os'] = os_entry
        # Placeholder handling for SUT_OS
        def on_os_entry_focus_in(event):
            if os_entry.get() == "e.g. linux, windows":
                os_entry.delete(0, tk.END)
                self.style_entry_by_theme(os_entry, placeholder=False)
        def on_os_entry_focus_out(event):
            if not os_entry.get():
                os_entry.insert(0, "e.g. linux, windows")
                self.style_entry_by_theme(os_entry, placeholder=True)
        os_entry.bind("<FocusIn>", on_os_entry_focus_in)
        os_entry.bind("<FocusOut>", on_os_entry_focus_out)

        # Pre-populate fields if editing existing node
        if existing_config:
            # Extract identifier from file_path
            identifier = Path(file_path).stem.split("settings.")[1]
            id_entry.delete(0, tk.END)
            id_entry.insert(0, identifier)
            # id_entry.config(foreground="#FFFFFF")
            self.style_entry_by_theme(id_entry, placeholder=False)
            
            # Pre-populate SKU
            if 'SUT_SKU' in existing_config:
                sku_entry.delete(0, tk.END)
                sku_entry.insert(0, existing_config['SUT_SKU'])
                self.style_entry_by_theme(sku_entry, placeholder=False)
            # Pre-populate SUT_OS
            if 'SUT_OS' in existing_config:
                os_entry.delete(0, tk.END)
                os_entry.insert(0, existing_config['SUT_OS'])
                self.style_entry_by_theme(os_entry, placeholder=False)
        
        # Services section
        services_label = ttk.Label(
            main_frame, 
            text="Consoles", 
            # background="#2D2D30", 
            # foreground="#FFFFFF",
            font=("Arial", 12, "bold")
        )
        services_label.pack(anchor=tk.W, pady=(10, 5))
        
        ttk.Label(
            main_frame, 
            text="Add at least one console to connect to this node", 
            # background="#2D2D30", 
            # foreground="#AAAAAA",
            font=("Arial", 9)
        ).pack(anchor=tk.W, pady=(0, 10))
        
        # Frame to contain service rows
        services_outer_frame = ttk.Frame(main_frame)
        services_outer_frame.pack(fill=tk.BOTH, expand=True, pady=5)

        services_container = ttk.Frame(services_outer_frame)
        services_container.pack(fill=tk.BOTH, expand=True)
        
        # Add header for services
        header_frame = ttk.Frame(services_container)
        header_frame.pack(fill=tk.X, pady=(0, 5))
        
        service_headers = ["Console Type", "IP Address", "Username", "Password"]
        widths = [15, 20, 15, 15]
        
        for i, header in enumerate(service_headers):
            ttk.Label(
                header_frame, 
                text=header, 
                # background="#2D2D30", 
                # foreground="#DDDDDD",
                font=("Arial", 9, "bold"),
                width=widths[i]
            ).pack(side=tk.LEFT, padx=3)
              
        row_counter = [0] 
        service_entries = {}  
        
        def add_service_row():
            # Generate unique ID for this row
            row_id = f"row_{row_counter[0]}"
            row_counter[0] += 1
            
            # Main row for service details
            row_frame = ttk.Frame(services_container)
            # row_frame.pack(fill=tk.X, pady=2)
            row_frame.pack(fill=tk.X, pady=(5, 10))
            
            # Service Type Entry
            service_type_entry = ttk.Entry(row_frame, width=10)
            service_type_entry.pack(side=tk.LEFT, padx=3)
            service_type_entry.insert(0, "RM/SOC/SUT")
            # service_type_entry.config(foreground="#999999")
            self.style_entry_by_theme(service_type_entry, placeholder=True)
            
            # Service IP Entry
            service_ip_entry = ttk.Entry(row_frame, width=20)
            service_ip_entry.pack(side=tk.LEFT, padx=3)
            service_ip_entry.insert(0, "e.g. 192.168.1.100")
            # service_ip_entry.config(foreground="#999999")
            self.style_entry_by_theme(service_ip_entry, placeholder=True)
            
            # Username Entry
            username_entry = ttk.Entry(row_frame, width=15)
            username_entry.pack(side=tk.LEFT, padx=3)
            username_entry.insert(0, "username")
            # username_entry.config(foreground="#999999")
            self.style_entry_by_theme(username_entry, placeholder=True)
            
            # Password Entry
            password_entry = ttk.Entry(row_frame, width=15)
            password_entry.pack(side=tk.LEFT, padx=3)
            password_entry.insert(0, "password")
            # password_entry.config(foreground="#999999")
            self.style_entry_by_theme(password_entry, placeholder=True)
            
            # Store entries in a dictionary with row_id as key
            service_entries[row_id] = {
                'frame': row_frame,
                'type': service_type_entry,
                'ip': service_ip_entry,
                'username': username_entry,
                'password': password_entry,
                'slot_id': None, 
            }
            
            # Delete button
            delete_btn = ttk.Button(
                row_frame, 
                text="✕", 
                width=3,
                command=lambda rid=row_id: remove_service_row(rid),
                bootstyle="danger-outline"
            )
            delete_btn.pack(side=tk.LEFT, padx=(5, 0))
            
            def on_service_type_focus_in(event):
                if service_type_entry.get() == "RM/SOC/SUT":
                    service_type_entry.delete(0, tk.END)
                    # service_type_entry.config(foreground="#FFFFFF")
                    self.style_entry_by_theme(service_type_entry, placeholder=False)
            
            def on_service_type_focus_out(event):
                if not service_type_entry.get():
                    service_type_entry.insert(0, "RM/SOC/SUT")
                    # service_type_entry.config(foreground="#999999")
                    self.style_entry_by_theme(service_type_entry, placeholder=True)
                else:
                    # Check if service type is "RM" and create SLOT_ID field if needed
                    service_type = service_type_entry.get().strip().upper()
                    if service_type == "RM":
                        create_slot_id_row(row_id)
                    else:
                        # If service type changed from RM to something else, remove the SLOT_ID field
                        remove_slot_id_row(row_id)
            
            service_type_entry.bind("<FocusIn>", on_service_type_focus_in)
            service_type_entry.bind("<FocusOut>", on_service_type_focus_out)
            
            # Placeholder behavior for service IP
            def on_service_ip_focus_in(event):
                if service_ip_entry.get() == "e.g. 192.168.1.100":
                    service_ip_entry.delete(0, tk.END)
                    # service_ip_entry.config(foreground="#FFFFFF")
                    self.style_entry_by_theme(service_ip_entry, placeholder=False)
            
            def on_service_ip_focus_out(event):
                if not service_ip_entry.get():
                    service_ip_entry.insert(0, "e.g. 192.168.1.100")
                    # service_ip_entry.config(foreground="#999999")
                    self.style_entry_by_theme(service_ip_entry, placeholder=True)
            
            service_ip_entry.bind("<FocusIn>", on_service_ip_focus_in)
            service_ip_entry.bind("<FocusOut>", on_service_ip_focus_out)
            
            # Placeholder behavior for username
            def on_username_focus_in(event):
                if username_entry.get() == "username":
                    username_entry.delete(0, tk.END)
                    # username_entry.config(foreground="#FFFFFF")
                    self.style_entry_by_theme(username_entry, placeholder=False)
            
            def on_username_focus_out(event):
                if not username_entry.get():
                    username_entry.insert(0, "username")
                    # username_entry.config(foreground="#999999")
                    self.style_entry_by_theme(username_entry, placeholder=True)
            
            username_entry.bind("<FocusIn>", on_username_focus_in)
            username_entry.bind("<FocusOut>", on_username_focus_out)
            
            # Placeholder behavior for password
            def on_password_focus_in(event):
                if password_entry.get() == "password":
                    password_entry.delete(0, tk.END)
                    # password_entry.config(foreground="#FFFFFF")
                    self.style_entry_by_theme(password_entry, placeholder=False)
            
            def on_password_focus_out(event):
                if not password_entry.get():
                    password_entry.insert(0, "password")
                    # password_entry.config(foreground="#999999")
                    self.style_entry_by_theme(password_entry, placeholder=True)
            
            password_entry.bind("<FocusIn>", on_password_focus_in)
            password_entry.bind("<FocusOut>", on_password_focus_out)

            top.after(100, adjust_window_size)
            update_ui_state()
            return row_id
        
        
        def create_slot_id_row(row_id):
            # Check if this service row exists
            if row_id not in service_entries:
                return
                
            # Check if SLOT_ID row already exists for this service
            service_row = service_entries[row_id]
            if service_row['slot_id'] is not None:
                return
                
            # Create a new frame for SLOT_ID that will appear below the service row
            slot_frame = ttk.Frame(services_container)
            
            # slot_frame.pack(fill=tk.X, padx=(20, 0), pady=(0, 5))
            slot_frame.pack(fill=tk.X, padx=(20, 0), pady=(0, 10))
            
            ttk.Label(
                slot_frame, 
                text="SLOT ID: *", 
                # background="#2D2D30", 
                # foreground="#FFFFFF",
                font=("Arial", 9)
            ).pack(side=tk.LEFT, padx=(0, 5))
            
            slot_id_entry = ttk.Entry(slot_frame, width=10)
            slot_id_entry.pack(side=tk.LEFT)
            slot_id_entry.insert(0, "e.g. 1")
            # slot_id_entry.config(foreground="#999999")
            self.style_entry_by_theme(slot_id_entry, placeholder=True)
            
            # Save references to slot_frame and slot_id_entry
            service_entries[row_id]['slot_id'] = {
                'frame': slot_frame,
                'entry': slot_id_entry
            }
            
            # Placeholder behavior for SLOT_ID
            def on_slot_id_focus_in(event):
                if slot_id_entry.get() == "e.g. 1":
                    slot_id_entry.delete(0, tk.END)
                    # slot_id_entry.config(foreground="#FFFFFF")
                    self.style_entry_by_theme(slot_id_entry, placeholder=False)
            
            def on_slot_id_focus_out(event):
                if not slot_id_entry.get():
                    slot_id_entry.insert(0, "e.g. 1")
                    # slot_id_entry.config(foreground="#999999")
                    self.style_entry_by_theme(slot_id_entry, placeholder=True)
            
            slot_id_entry.bind("<FocusIn>", on_slot_id_focus_in)
            slot_id_entry.bind("<FocusOut>", on_slot_id_focus_out)

            top.after(100, adjust_window_size)

        def remove_slot_id_row(row_id):
            # Check if this service row exists and has a slot_id
            if row_id in service_entries and service_entries[row_id]['slot_id'] is not None:
                slot_data = service_entries[row_id]['slot_id']
                slot_data['frame'].destroy()
                service_entries[row_id]['slot_id'] = None

            top.after(100, adjust_window_size)
        
        def remove_service_row(row_id):
            if row_id in service_entries:
                # Remove the slot_id frame if it exists
                remove_slot_id_row(row_id)
                
                # Remove the service frame
                service_entries[row_id]['frame'].destroy()
                
                # Remove from the dictionary
                del service_entries[row_id]

                # Adjust window size after removing a row
                top.after(100, adjust_window_size)
                update_ui_state()
        

        def adjust_window_size():
            top.update_idletasks()
            
            num_services = len(service_entries)
            
            slot_id_count = sum(1 for service in service_entries.values() if service['slot_id'] is not None)
            
            base_height = 600   # increased base height to accommodate SUT OS
            
            service_height = 50
            
            slot_id_height = 40
            
            calculated_height = base_height + (num_services * service_height) + (slot_id_count * slot_id_height)
            
            min_height = 650
            max_height = int(top.winfo_screenheight() * 0.85)
            
            new_height = max(min_height, min(calculated_height, max_height))
            
            current_width = max(800, top.winfo_width())
            
            top.geometry(f"{current_width}x{new_height}")

        
        # Buttons frame
        buttons_frame = ttk.Frame(main_frame)
        buttons_frame.pack(fill=tk.X, pady=15)

        
        add_service_btn = ttk.Button(
            buttons_frame, 
            text="+ Add Another Console", 
            # command=add_service_row,
            # command=add_service_with_limit,
            bootstyle="info"
        )
        add_service_btn.pack(side=tk.LEFT)

        # NOW define the functions that reference add_service_btn
        def update_ui_state():
            if len(service_entries) >= self.MAX_CONSOLES:
                show_inline_warning()
            else:
                hide_inline_warning()

        def show_inline_warning():
            if not hasattr(top, 'warning_frame'):
                top.warning_frame = ttk.Frame(main_frame)
                
                # Warning icon + message
                warning_icon = ttk.Label(top.warning_frame, text="⚠️", font=("Arial", 12))
                warning_icon.pack(side=tk.LEFT, padx=(0, 5))
                
                self.warning_label = ttk.Label(
                    top.warning_frame,
                    text=f"Console limit reached. Maximum {self.MAX_CONSOLES} consoles allowed per node.",
                    bootstyle="warning",
                    font=("Arial", 9)
                )
                self.warning_label.pack(side=tk.LEFT)
            
            # Place AFTER buttons_frame instead of before
            top.warning_frame.pack(fill=tk.X, after=buttons_frame)
            add_service_btn.configure(state="disabled", text="+ Add Another Console") # Keep original text


        def hide_inline_warning():
            if hasattr(top, 'warning_frame'):
               top.warning_frame.pack_forget()
            add_service_btn.configure(state="normal", text="+ Add Another Console")

        def add_service_with_limit():
            if len(service_entries) >= self.MAX_CONSOLES:
                show_inline_warning()
                return
            add_service_row()
            update_ui_state()
        
        add_service_btn.configure(command=add_service_with_limit)

        # add_service_row()
        # update_ui_state()

        if existing_config:
            # Pre-populate existing services
            for key, value in existing_config.items():
                if key.endswith('_IP') and value:
                    service_type = key.split('_')[0]
                    row_id = add_service_row()
                    
                    # Populate the row
                    service_entries[row_id]['type'].delete(0, tk.END)
                    service_entries[row_id]['type'].insert(0, service_type)
                    # service_entries[row_id]['type'].config(foreground="#FFFFFF")
                    self.style_entry_by_theme(service_entries[row_id]['type'], placeholder=False)
                    
                    service_entries[row_id]['ip'].delete(0, tk.END)
                    service_entries[row_id]['ip'].insert(0, value)
                    # service_entries[row_id]['ip'].config(foreground="#FFFFFF")
                    self.style_entry_by_theme(service_entries[row_id]['ip'], placeholder=False)
                    
                    # Populate username and password if they exist
                    username_key = f"{service_type}_GUESS"
                    if username_key in existing_config and existing_config[username_key]:
                        service_entries[row_id]['username'].delete(0, tk.END)
                        service_entries[row_id]['username'].insert(0, existing_config[username_key])
                        # service_entries[row_id]['username'].config(foreground="#FFFFFF")
                        self.style_entry_by_theme(service_entries[row_id]['username'], placeholder=False)
                    
                    password_key = f"{service_type}_WHAT"
                    if password_key in existing_config and existing_config[password_key]:
                        service_entries[row_id]['password'].delete(0, tk.END)
                        service_entries[row_id]['password'].insert(0, existing_config[password_key])
                        # service_entries[row_id]['password'].config(foreground="#FFFFFF")
                        self.style_entry_by_theme(service_entries[row_id]['password'], placeholder=False)
                    
                    # Handle SLOT_ID for RM
                    if service_type == "RM" and 'SLOT_ID' in existing_config:
                        create_slot_id_row(row_id)
                        service_entries[row_id]['slot_id']['entry'].delete(0, tk.END)
                        service_entries[row_id]['slot_id']['entry'].insert(0, existing_config['SLOT_ID'])
                        # service_entries[row_id]['slot_id']['entry'].config(foreground="#FFFFFF")
                        self.style_entry_by_theme(service_entries[row_id]['slot_id']['entry'], placeholder=False)
        else:
            add_service_row()

        update_ui_state()
        
        cancel_btn = ttk.Button(
            buttons_frame, 
            text="Cancel", 
            command=top.destroy,
            bootstyle="secondary"
        )
        cancel_btn.pack(side=tk.RIGHT, padx=5)
        
        def validate_ip(ip):
            """Validate if string is a valid IP address"""
            try:
                parts = ip.split('.')
                if len(parts) != 4:
                    return False
                for part in parts:
                    if not 0 <= int(part) <= 255:
                        return False
                return True
            except (ValueError, AttributeError):
                return False
            
        def validate_hostname(hostname):
            if len(hostname) > 255:
                return False
            if hostname[-1] == ".":
                hostname = hostname[:-1]
            allowed = re.compile(r"^(?!-)[A-Z\d-]{1,63}(?<!-)$", re.IGNORECASE)
            return all(allowed.match(part) for part in hostname.split("."))
        
        def sanitize_filename(name):
            """Remove characters that are invalid in filenames"""
            invalid_chars = '<>:"/\\|?*'
            for char in invalid_chars:
                name = name.replace(char, '_')
            return name
        
        def check_duplicate_file(identifier, folder):
            """Check if a file with this identifier already exists"""
            config_path = Path(__file__).parent.parent / "sut" / folder
            filename = f"settings.{identifier}.json"
            return (config_path / filename).exists()
        
        def validate_slot_id(value):
            """Validate if the value is a valid integer"""
            try:
                if value and value != "e.g. 1":
                    int(value)
                    return True
                return False
            except ValueError:
                return False
        
        def save_node():
            # Get identifier and remove placeholder if present
            identifier = id_entry.get().strip()
            if identifier == "e.g. 192.168.1.100 or server-name":
                identifier = ""
                
            # Get SUT_SKU and remove placeholder if present
            sut_sku = sku_entry.get().strip()
            if sut_sku == "e.g. CAAA, CAAAA, DAAAA":
                sut_sku = ""
            # Get SUT_OS and remove placeholder if present
            sut_os = os_entry.get().strip()
            if sut_os == "e.g. linux, windows":
                sut_os = ""
                
            # Validate required fields
            if not identifier:
                messagebox.showerror("Error", "Node Identifier/IP is required.")
                id_entry.focus_set()
                return
                
            if not sut_sku:
                messagebox.showerror("Error", "SUT SKU is required.")
                sku_entry.focus_set()
                return
         
            if not sut_os:
                messagebox.showerror("Error", "SUT OS is not specified.")
                os_entry.focus_set()
                return
            # Sanitize filename
            safe_identifier = sanitize_filename(identifier)
            if safe_identifier != identifier:
                if not messagebox.askyesno(
                    "Warning", 
                    f"Identifier contains invalid characters for a filename.\n\n"
                    f"Do you want to continue using '{safe_identifier}' instead?"
                ):
                    return
                identifier = safe_identifier
                
            # Check for duplicates
            # Only check for duplicates if it's a new node or identifier changed
            if not existing_config or (existing_config and Path(file_path).stem.split("settings.")[1] != identifier):
                if check_duplicate_file(identifier, folder_name):
                    if not messagebox.askyesno(
                        "Warning", 
                        f"A node with identifier '{identifier}' already exists in {folder_name}.\n\n"
                        f"Do you want to overwrite it?"
                    ):
                        return
            
            # Validate services and collect config
            valid_services = False
            config = {"SUT_SKU": sut_sku}
            if sut_os:
                config["SUT_OS"] = sut_os.lower()
            rm_service_exists = False
            
            for row_id, service_data in service_entries.items():
                stype = service_data['type'].get().strip()
                ip = service_data['ip'].get().strip()
                user = service_data['username'].get().strip()
                pw = service_data['password'].get().strip()
                
                # Skip placeholder values
                if stype in ["RM/SOC/SUT", ""] or ip in ["e.g. 192.168.1.100", ""]:
                    continue
                    
                    
                if not (validate_ip(ip) or validate_hostname(ip)):
                    messagebox.showerror(
                        "Invalid Address", 
                        f"'{ip}' is neither a valid IPv4 address nor a valid hostname.\n\n"
                        f"Please enter a valid IP (xxx.xxx.xxx.xxx) or hostname (e.g., myhost.corp.com)."
                    )
                    return

                # Check if service type is RM and validate SLOT_ID
                stype_upper = stype.upper()
                if stype_upper == "RM":
                    rm_service_exists = True
                    
                    if service_data['slot_id'] is None:
                        messagebox.showerror("Error", "SLOT ID is required for RM service.")
                        return
                        
                    slot_id_entry = service_data['slot_id']['entry']
                    slot_id = slot_id_entry.get().strip()
                    
                    if not validate_slot_id(slot_id):
                        messagebox.showerror("Error", "SLOT ID must be a valid integer.")
                        slot_id_entry.focus_set()
                        return
                        
                    # Add SLOT_ID to the configuration
                    config["SLOT_ID"] = slot_id
                    
                config[f"{stype}_IP"] = ip
                if user and user != "username":
                    config[f"{stype}_GUESS"] = user
                if pw and pw != "password":
                    config[f"{stype}_WHAT"] = pw
                    
                valid_services = True
                
            if not valid_services:
                messagebox.showerror(
                    "No Consoles", 
                    "Please add at least one console with a console type and IP address."
                )
                return
                
            # Save configuration
            try:
                # config_path = Path(__file__).parent.parent / "sut" / folder_name
                if folder_name.lower() == "root":
                    config_path = Path(__file__).parent.parent / "sut"
                else:
                    config_path = Path(__file__).parent.parent / "sut" / folder_name
                config_path.mkdir(parents=True, exist_ok=True)
                
                filename = f"settings.{identifier}.json"
                filepath = config_path / filename
                
                with open(filepath, 'w') as f:
                    json.dump(config, f, indent=2)
                    
                # Success message with details
                services_text = ", ".join(key.split('_')[0] for key in config.keys() if key.endswith('_IP'))
                
                success_msg = f"Node '{identifier}' added to {folder_name} with consoles: {services_text}."
                if rm_service_exists:
                    success_msg += f"\nSLOT ID: {config.get('SLOT_ID')}"
                    
                messagebox.showinfo("Success", success_msg)
                top.destroy()
                
                if self.refresh_callback and callable(self.refresh_callback):
                    self.refresh_callback(folder_name)
                    
            except Exception as e:
                messagebox.showerror("Error", f"Failed to save node: {str(e)}")
        
        save_btn_text = "Update Node" if existing_config else "Save Node"
        save_btn = ttk.Button(
            buttons_frame, 
            text=save_btn_text, 
            command=save_node,
            bootstyle="info"
        )
        save_btn.pack(side=tk.RIGHT)
        

        # Center the window on screen
        top.update_idletasks()
        width = top.winfo_width()
        height = top.winfo_height()
        x = (top.winfo_screenwidth() // 2) - (width // 2)
        y = (top.winfo_screenheight() // 2) - (height // 2)
        top.geometry(f'+{x}+{y}')
       
        def on_window_close():
            if top in self.current_windows:
                self.current_windows.remove(top)
            top.destroy()

        top.protocol("WM_DELETE_WINDOW", on_window_close)
        cancel_btn.config(command=on_window_close)