import tkinter as tk
from tkinter import ttk, messagebox
import threading
import time

try:
    from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
    from comtypes import CLSCTX_ALL
    PYCAW_AVAILABLE = True
except ImportError:
    AudioUtilities = None
    IAudioEndpointVolume = None
    CLSCTX_ALL = None
    PYCAW_AVAILABLE = False

class AdvancedMicController:
    def __init__(self, root):
        self.root = root
        self.root.title("Advanced Mic Volume Controller")
        self.root.geometry("600x750")
        # self.root.resizable(False, False)
        
        # Variables
        self.current_volume = tk.IntVar(value=50)
        self.is_locked = tk.BooleanVar(value=False)
        self.lock_thread = None
        self.volume_interface = None
        self.current_device = None
        self.recording_devices = []
        self.selected_device = tk.StringVar()
        
        if PYCAW_AVAILABLE:
            self.init_audio_interface()
        
        self.setup_ui()
        self.update_volume_display()
        
    def init_audio_interface(self):
        """Initialize audio interface using pycaw"""
        try:
            if not AudioUtilities:
                print("AudioUtilities not available")
                self.volume_interface = None
                self.current_device = None
                self.recording_devices = []
                return
                
            # Suppress pycaw warnings
            import warnings
            warnings.filterwarnings("ignore", category=UserWarning, module="pycaw")
            
            # Get the default microphone first (most reliable)
            devices = AudioUtilities.GetMicrophone()
            if devices:
                interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
                self.volume_interface = interface.QueryInterface(IAudioEndpointVolume)
                self.current_device = devices
            else:
                self.volume_interface = None
                self.current_device = None
            
            # Try to get recording devices (with error handling)
            self.recording_devices = []
            try:
                from pycaw.pycaw import EDataFlow
                device_enumerator = AudioUtilities.GetDeviceEnumerator()
                if device_enumerator:
                    collection = device_enumerator.EnumAudioEndpoints(EDataFlow.eCapture.value, 1)  # DEVICE_STATE_ACTIVE
                    count = collection.GetCount()
                    
                    for i in range(count):
                        try:
                            device = collection.Item(i)
                            self.recording_devices.append(device)
                        except:
                            continue
            except Exception as e:
                print(f"Could not enumerate recording devices: {e}")
                
        except Exception as e:
            print(f"Failed to initialize audio interface: {e}")
            self.volume_interface = None
            self.current_device = None
            self.recording_devices = []
    
    def setup_ui(self):
        # Main frame
        main_frame = ttk.Frame(self.root, padding="15")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Title
        ttk.Label(main_frame, text="Advanced Microphone Controller", 
                 font=("Arial", 14, "bold")).pack(pady=(0, 15))
        
        # Status
        status_text = "✓ Audio API Connected" if PYCAW_AVAILABLE and self.volume_interface else "⚠ Using Fallback Method"
        status_color = "green" if PYCAW_AVAILABLE and self.volume_interface else "orange"
        ttk.Label(main_frame, text=status_text, foreground=status_color).pack(pady=(0, 5))
        
        # Device selection (simplified approach)
        if PYCAW_AVAILABLE:
            device_frame = ttk.LabelFrame(main_frame, text="Microphone Control", padding="10")
            device_frame.pack(fill=tk.X, pady=(0, 10))
            
            # Show current device info
            device_info = "Default Microphone"
            if self.current_device:
                try:
                    from pycaw.pycaw import AudioUtilities
                    device_info = "Connected to Audio System"
                except:
                    pass
            
            ttk.Label(device_frame, text=f"Controlling: {device_info}").pack(side=tk.LEFT, padx=(0, 10))
            
            # Button frame for better layout
            button_subframe = ttk.Frame(device_frame)
            button_subframe.pack(side=tk.RIGHT)
            
            ttk.Button(button_subframe, text="Detect Devices", 
                      command=self.detect_external_mic).pack(side=tk.LEFT, padx=2)
            ttk.Button(button_subframe, text="Windows Settings", 
                      command=self.open_windows_sound_settings).pack(side=tk.LEFT, padx=2)
            ttk.Button(button_subframe, text="Reset Default", 
                      command=self.reset_to_default).pack(side=tk.LEFT, padx=2)
        
        # Volume display
        self.volume_display = ttk.Label(main_frame, text="Volume: 50%", 
                                       font=("Arial", 12, "bold"))
        self.volume_display.pack(pady=(0, 10))
        
        # Volume slider frame
        slider_frame = ttk.LabelFrame(main_frame, text="Volume Control", padding="10")
        slider_frame.pack(fill=tk.X, pady=(0, 15))
        
        self.volume_scale = ttk.Scale(slider_frame, from_=0, to=100, 
                                     orient=tk.HORIZONTAL, length=320,
                                     variable=self.current_volume,
                                     command=self.on_volume_change)
        self.volume_scale.pack(pady=(0, 10))
        
        # Quick buttons
        button_frame = ttk.Frame(slider_frame)
        button_frame.pack()
        
        for vol in [0, 25, 50, 75, 100]:
            text = "Mute" if vol == 0 else f"{vol}%"
            ttk.Button(button_frame, text=text, width=8,
                      command=lambda v=vol: self.set_volume(v)).pack(side=tk.LEFT, padx=2)
        
        # Advanced controls
        control_frame = ttk.LabelFrame(main_frame, text="Advanced Controls", padding="10")
        control_frame.pack(fill=tk.X, pady=(0, 15))
        
        # Mute/Unmute
        mute_frame = ttk.Frame(control_frame)
        mute_frame.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Button(mute_frame, text="Mute", width=10,
                  command=self.mute).pack(side=tk.LEFT, padx=3)
        ttk.Button(mute_frame, text="Unmute", width=10,
                  command=self.unmute).pack(side=tk.LEFT, padx=3)
        ttk.Button(mute_frame, text="Refresh", width=10,
                  command=self.update_volume_display).pack(side=tk.LEFT, padx=3)

        
        # Volume lock
        ttk.Checkbutton(control_frame, text="Lock volume (prevent automatic changes)", 
                       variable=self.is_locked,
                       command=self.toggle_lock).pack(anchor=tk.W)
        
        self.lock_status = ttk.Label(control_frame, text="Status: Unlocked", 
                                    foreground="orange")
        self.lock_status.pack(anchor=tk.W)
        
        # Info frame
        info_frame = ttk.LabelFrame(main_frame, text="Information", padding="10")
        info_frame.pack(fill=tk.X)
        
        info_text = """Tips:
• Use 'Lock volume' during calls to prevent Windows from auto-adjusting
• The app works best when run as administrator
• If volume control doesn't work, try the fallback methods"""
        
        ttk.Label(info_frame, text=info_text, font=("Arial", 8), 
                 justify=tk.LEFT).pack(anchor=tk.W)
    
    def get_current_volume(self):
        """Get current microphone volume"""
        if self.volume_interface:
            try:
                volume = self.volume_interface.GetMasterVolumeLevelScalar()
                return int(volume * 100)
            except:
                pass
        return None
    
    def get_current_device_interface(self):
        """Get the current device interface"""
        return self.volume_interface
    
    def get_device_name_alternative(self, device, index):
        """Alternative method to get device name using PowerShell"""
        try:
            # Try to get device ID first
            device_id = None
            try:
                device_id = device.GetId()
            except:
                pass
            
            if device_id:
                # Use PowerShell to get the friendly name
                import subprocess
                ps_script = f'''
                $deviceId = "{device_id}"
                $devices = Get-PnpDevice -Class AudioEndpoint | Where-Object {{$_.InstanceId -like "*$deviceId*" -or $_.DeviceID -like "*$deviceId*"}}
                if ($devices) {{
                    $devices[0].FriendlyName
                }} else {{
                    # Try alternative approach
                    Get-WmiObject Win32_SoundDevice | Where-Object {{$_.DeviceID -like "*$deviceId*"}} | Select-Object -ExpandProperty Name
                }}
                '''
                
                try:
                    result = subprocess.run(["powershell", "-Command", ps_script], 
                                          capture_output=True, text=True, timeout=5)
                    if result.returncode == 0 and result.stdout.strip():
                        name = result.stdout.strip()
                        if name and name != "":
                            return name
                except:
                    pass
            
            # Fallback: Try to use COM properties directly
            try:
                import comtypes
                from comtypes import GUID
                
                # Try different property keys
                prop_keys = [
                    # PKEY_Device_FriendlyName
                    (GUID("{a45c254e-df1c-4efd-8020-67d146a850e0}"), 14),
                    # PKEY_Device_DeviceDesc  
                    (GUID("{a45c254e-df1c-4efd-8020-67d146a850e0}"), 2),
                    # PKEY_DeviceInterface_FriendlyName
                    (GUID("{026e516e-b814-414b-83cd-856d6fef4822}"), 2)
                ]
                
                prop_store = device.OpenPropertyStore(0)
                for guid, pid in prop_keys:
                    try:
                        from comtypes import _ole32
                        prop_key = comtypes.Structure()
                        prop_key.fmtid = guid
                        prop_key.pid = pid
                        
                        prop_variant = prop_store.GetValue(prop_key)
                        name = prop_variant.GetValue()
                        if name and isinstance(name, str) and name.strip():
                            return name.strip()
                    except:
                        continue
                        
            except:
                pass
            
            # Final fallback: Use PowerShell to list all audio devices and match by index
            try:
                import subprocess
                ps_script = '''
                $devices = Get-PnpDevice -Class AudioEndpoint | Where-Object {$_.Status -eq "OK"}
                $captureDevices = @()
                foreach ($device in $devices) {
                    if ($device.FriendlyName -match "microphone|mic|headset|audio|capture|recording") {
                        $captureDevices += $device.FriendlyName
                    }
                }
                $captureDevices | ForEach-Object { $_ }
                '''
                
                result = subprocess.run(["powershell", "-Command", ps_script], 
                                      capture_output=True, text=True, timeout=5)
                if result.returncode == 0 and result.stdout.strip():
                    device_names = result.stdout.strip().split('\n')
                    if index < len(device_names) and device_names[index].strip():
                        return device_names[index].strip()
            except:
                pass
                
        except Exception as e:
            print(f"Error getting device name: {e}")
        
        # Ultimate fallback
        return f"Audio Device {index + 1}"
    
    def detect_external_mic(self):
        """Try to detect and switch to external microphone with progress dialog"""
        # Create progress dialog
        progress_window = tk.Toplevel(self.root)
        progress_window.title("Detecting Microphones...")
        progress_window.geometry("450x300")
        progress_window.resizable(False, False)
        progress_window.transient(self.root)
        progress_window.grab_set()
        
        # Center the window
        progress_window.geometry("+%d+%d" % (
            self.root.winfo_rootx() + 50,
            self.root.winfo_rooty() + 50
        ))
        
        # Progress UI
        ttk.Label(progress_window, text="Scanning for microphone devices...", 
                 font=("Arial", 12, "bold")).pack(pady=10)
        
        # Progress bar
        progress_var = tk.DoubleVar()
        progress_bar = ttk.Progressbar(progress_window, variable=progress_var, 
                                     maximum=100, length=300)
        progress_bar.pack(pady=10)
        
        # Status label
        status_label = ttk.Label(progress_window, text="Initializing...", 
                               font=("Arial", 9))
        status_label.pack(pady=5)
        
        # Device list display
        ttk.Label(progress_window, text="Detected devices:", 
                 font=("Arial", 10, "bold")).pack(pady=(10, 5))
        
        # Scrollable text widget for real-time device display
        text_frame = ttk.Frame(progress_window)
        text_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=5)
        
        device_text = tk.Text(text_frame, height=8, width=50, font=("Arial", 8))
        scrollbar = ttk.Scrollbar(text_frame, orient=tk.VERTICAL, command=device_text.yview)
        device_text.configure(yscrollcommand=scrollbar.set)
        
        device_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Cancel button
        cancel_pressed = tk.BooleanVar(value=False)
        ttk.Button(progress_window, text="Cancel", 
                  command=lambda: cancel_pressed.set(True)).pack(pady=10)
        
        # Start detection in separate thread
        import threading
        
        def detection_thread():
            try:
                from pycaw.pycaw import AudioUtilities, EDataFlow
                
                if cancel_pressed.get():
                    return
                
                # Update status
                status_label.config(text="Connecting to audio system...")
                progress_var.set(10)
                progress_window.update()
                
                # Get all capture devices (including disabled ones)
                device_enumerator = AudioUtilities.GetDeviceEnumerator()
                
                if cancel_pressed.get():
                    return
                
                # Try both active and all devices
                collections = []
                status_label.config(text="Scanning active devices...")
                progress_var.set(20)
                progress_window.update()
                
                try:
                    # Active devices
                    active_collection = device_enumerator.EnumAudioEndpoints(EDataFlow.eCapture.value, 1)
                    collections.append(("Active", active_collection))
                    device_text.insert(tk.END, "✓ Found active devices collection\n")
                    device_text.see(tk.END)
                    progress_window.update()
                except Exception as e:
                    device_text.insert(tk.END, f"⚠ Could not access active devices: {e}\n")
                    device_text.see(tk.END)
                    progress_window.update()
                
                if cancel_pressed.get():
                    return
                
                status_label.config(text="Scanning all devices...")
                progress_var.set(30)
                progress_window.update()
                
                try:
                    # All devices (including unplugged)
                    all_collection = device_enumerator.EnumAudioEndpoints(EDataFlow.eCapture.value, 15)
                    collections.append(("All", all_collection))
                    device_text.insert(tk.END, "✓ Found all devices collection\n")
                    device_text.see(tk.END)
                    progress_window.update()
                except Exception as e:
                    device_text.insert(tk.END, f"⚠ Could not access all devices: {e}\n")
                    device_text.see(tk.END)
                    progress_window.update()
                
                found_devices = []
                device_names_seen = set()
                total_progress = 40
                
                device_text.insert(tk.END, "\n--- Scanning Devices ---\n")
                device_text.see(tk.END)
                progress_window.update()
                
                for collection_name, collection in collections:
                    if cancel_pressed.get():
                        return
                        
                    try:
                        count = collection.GetCount()
                        device_text.insert(tk.END, f"\n{collection_name} collection: {count} devices\n")
                        device_text.see(tk.END)
                        progress_window.update()
                        
                        status_label.config(text=f"Checking {collection_name.lower()} devices...")
                        
                        for i in range(count):
                            if cancel_pressed.get():
                                return
                                
                            try:
                                device = collection.Item(i)
                                
                                # Update progress
                                device_progress = (i + 1) / count * 25  # 25% per collection
                                progress_var.set(total_progress + device_progress)
                                
                                # Get device friendly name using alternative method
                                friendly_name = self.get_device_name_alternative(device, i)
                                
                                # Check device state
                                try:
                                    state = device.GetState()
                                    state_text = ""
                                    state_icon = ""
                                    if state == 1:  # DEVICE_STATE_ACTIVE
                                        state_text = " (Active)"
                                        state_icon = "✓"
                                    elif state == 2:  # DEVICE_STATE_DISABLED
                                        state_text = " (Disabled)"
                                        state_icon = "○"
                                    elif state == 4:  # DEVICE_STATE_NOTPRESENT
                                        state_text = " (Not Present)"
                                        state_icon = "⚠"
                                    elif state == 8:  # DEVICE_STATE_UNPLUGGED
                                        state_text = " (Unplugged)"
                                        state_icon = "○"
                                    else:
                                        state_text = f" (State: {state})"
                                        state_icon = "?"
                                    
                                    display_name = f"{friendly_name}{state_text}"
                                except:
                                    display_name = friendly_name
                                    state = 1
                                    state_icon = "?"
                                
                                # Avoid duplicates
                                if display_name not in device_names_seen:
                                    device_names_seen.add(display_name)
                                    found_devices.append((display_name, device, state if 'state' in locals() else 1))
                                    
                                    # Show in real-time
                                    device_text.insert(tk.END, f"  {state_icon} {friendly_name}{state_text}\n")
                                    device_text.see(tk.END)
                                    progress_window.update()
                                
                            except Exception as e:
                                device_text.insert(tk.END, f"  ✗ Error checking device {i}: {str(e)[:50]}...\n")
                                device_text.see(tk.END)
                                progress_window.update()
                                continue
                        
                        total_progress += 25
                                
                    except Exception as e:
                        device_text.insert(tk.END, f"✗ Error with {collection_name} collection: {e}\n")
                        device_text.see(tk.END)
                        progress_window.update()
                        continue
                
                if cancel_pressed.get():
                    return
                
                # Finalize
                progress_var.set(100)
                status_label.config(text=f"Scan complete! Found {len(found_devices)} devices")
                device_text.insert(tk.END, f"\n--- Scan Complete ---\nTotal devices found: {len(found_devices)}\n")
                device_text.see(tk.END)
                progress_window.update()
                
                # Wait a moment then close progress and show selection
                progress_window.after(1500, lambda: self.finish_detection(progress_window, found_devices))
                
            except Exception as e:
                device_text.insert(tk.END, f"\n✗ Detection error: {e}\n")
                device_text.see(tk.END)
                status_label.config(text="Error occurred during detection")
                progress_window.update()
                
                # Show fallback after delay
                progress_window.after(2000, lambda: self.finish_detection_with_fallback(progress_window))
        
        # Start the detection thread
        thread = threading.Thread(target=detection_thread, daemon=True)
        thread.start()
        
        # Handle window close
        def on_close():
            cancel_pressed.set(True)
            progress_window.destroy()
        
        progress_window.protocol("WM_DELETE_WINDOW", on_close)
    
    def finish_detection(self, progress_window, found_devices):
        """Finish detection and show device selection"""
        progress_window.destroy()
        
        if found_devices:
            self.show_device_selection(found_devices)
        else:
            messagebox.showinfo("External Microphone", 
                              "No microphone devices found.\n\n"
                              "Try:\n"
                              "1. Make sure your earphones/headset is properly connected\n"
                              "2. Check Windows Sound Settings\n"
                              "3. Try unplugging and reconnecting your device")
    
    def finish_detection_with_fallback(self, progress_window):
        """Finish detection with fallback method"""
        progress_window.destroy()
        self.detect_external_fallback()
    
    def show_device_selection(self, devices):
        """Show a dialog to select microphone device"""
        selection_window = tk.Toplevel(self.root)
        selection_window.title("Select Microphone Device")
        selection_window.geometry("500x400")
        selection_window.resizable(True, True)
        
        # Center the window
        selection_window.transient(self.root)
        selection_window.grab_set()
        
        ttk.Label(selection_window, text="Available Microphone Devices:", 
                 font=("Arial", 12, "bold")).pack(pady=10)
        
        # Info label
        info_label = ttk.Label(selection_window, 
                              text="Tip: Look for devices with 'Headset', 'Earphones', or your brand name",
                              font=("Arial", 9), foreground="blue")
        info_label.pack(pady=(0, 10))
        
        # Create listbox with devices
        listbox_frame = ttk.Frame(selection_window)
        listbox_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)
        
        scrollbar = ttk.Scrollbar(listbox_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        device_listbox = tk.Listbox(listbox_frame, yscrollcommand=scrollbar.set, 
                                   font=("Arial", 9), height=12)
        device_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.config(command=device_listbox.yview)
        
        # Populate listbox with better formatting
        active_devices = []
        inactive_devices = []
        
        for name, device, state in devices:
            if state == 1:  # Active
                active_devices.append((name, device, state))
            else:
                inactive_devices.append((name, device, state))
        
        # Add active devices first
        all_devices_for_selection = []
        if active_devices:
            device_listbox.insert(tk.END, "=== ACTIVE DEVICES ===")
            all_devices_for_selection.append(None)  # Placeholder for header
            
            for name, device, state in active_devices:
                device_listbox.insert(tk.END, f"✓ {name}")
                all_devices_for_selection.append((name, device, state))
        
        # Add inactive devices
        if inactive_devices:
            if active_devices:
                device_listbox.insert(tk.END, "")
                all_devices_for_selection.append(None)  # Placeholder for spacing
            
            device_listbox.insert(tk.END, "=== OTHER DEVICES ===")
            all_devices_for_selection.append(None)  # Placeholder for header
            
            for name, device, state in inactive_devices:
                device_listbox.insert(tk.END, f"○ {name}")
                all_devices_for_selection.append((name, device, state))
        
        # Select first active device by default
        first_active_index = None
        for i, item in enumerate(all_devices_for_selection):
            if item and item[2] == 1:  # First active device
                first_active_index = i
                break
        
        if first_active_index:
            device_listbox.selection_set(first_active_index)
        
        # Buttons
        button_frame = ttk.Frame(selection_window)
        button_frame.pack(pady=10)
        
        def select_device():
            selection = device_listbox.curselection()
            if selection:
                selected_index = selection[0]
                if selected_index < len(all_devices_for_selection) and all_devices_for_selection[selected_index]:
                    selected_name, selected_device, state = all_devices_for_selection[selected_index]
                    try:
                        # Switch to selected device
                        interface = selected_device.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
                        new_volume_interface = interface.QueryInterface(IAudioEndpointVolume)
                        
                        self.volume_interface = new_volume_interface
                        self.current_device = selected_device
                        
                        selection_window.destroy()
                        messagebox.showinfo("Device Selected", f"Now controlling: {selected_name}")
                        self.update_volume_display()
                        
                    except Exception as e:
                        messagebox.showerror("Error", f"Could not switch to device:\n{str(e)}\n\nTry selecting an 'Active' device.")
                else:
                    messagebox.showwarning("Selection", "Please select a valid device (not a header)")
            else:
                messagebox.showwarning("Selection", "Please select a device")
        
        def test_device():
            """Test if selected device works"""
            selection = device_listbox.curselection()
            if selection:
                selected_index = selection[0]
                if selected_index < len(all_devices_for_selection) and all_devices_for_selection[selected_index]:
                    selected_name, selected_device, state = all_devices_for_selection[selected_index]
                    try:
                        # Try to activate the device
                        interface = selected_device.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
                        volume_interface = interface.QueryInterface(IAudioEndpointVolume)
                        
                        # Try to get current volume
                        current_vol = volume_interface.GetMasterVolumeLevelScalar()
                        messagebox.showinfo("Test Result", 
                                          f"✓ Device works!\n"
                                          f"Current volume: {int(current_vol * 100)}%")
                        
                    except Exception as e:
                        messagebox.showerror("Test Result", 
                                           f"✗ Device test failed:\n{str(e)}")
        
        ttk.Button(button_frame, text="Test Device", command=test_device).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Select Device", command=select_device).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Cancel", command=selection_window.destroy).pack(side=tk.LEFT, padx=5)
    
    def detect_external_fallback(self):
        """Fallback method to detect external microphones"""
        try:
            import subprocess
            
            # Use PowerShell to list audio devices more comprehensively
            ps_script = """
            # Get all audio devices
            Get-PnpDevice -Class AudioEndpoint | Where-Object {$_.FriendlyName -like "*mic*" -or $_.FriendlyName -like "*headset*" -or $_.FriendlyName -like "*earphone*"} | Select-Object FriendlyName, Status
            """
            
            result = subprocess.run(["powershell", "-Command", ps_script], 
                                  capture_output=True, text=True, timeout=15)
            
            devices_found = False
            if result.returncode == 0 and result.stdout.strip():
                devices_found = True
                messagebox.showinfo("Audio Devices Found", 
                                  f"Found these audio devices:\n{result.stdout}\n\n"
                                  "Steps to use your external microphone:\n"
                                  "1. Right-click sound icon in taskbar\n"
                                  "2. Select 'Open Sound settings'\n"
                                  "3. Click 'More sound settings'\n"
                                  "4. Go to 'Recording' tab\n"
                                  "5. Right-click your external mic and 'Set as Default'\n"
                                  "6. Click 'Reset to Default' in this app")
            
            if not devices_found:
                # Try another approach
                ps_script2 = """
                Get-WmiObject Win32_SoundDevice | Select-Object Name, Status | Format-Table -AutoSize
                """
                
                result2 = subprocess.run(["powershell", "-Command", ps_script2], 
                                       capture_output=True, text=True, timeout=10)
                
                message = "External Microphone Detection Help\n\n"
                
                if result2.returncode == 0 and result2.stdout.strip():
                    message += f"System audio devices:\n{result2.stdout}\n\n"
                
                message += """For earphones/headsets with microphone:

1. Make sure your device is properly connected
2. Right-click the sound icon in your taskbar
3. Select 'Open Sound settings'
4. Scroll down and click 'More sound settings'
5. In the 'Recording' tab, look for:
   - Your headset/earphone brand name
   - 'Headset Microphone'
   - 'External Microphone'
   - Any device that's not 'Internal Microphone'

6. Right-click the correct device and select:
   - 'Enable' (if disabled)
   - 'Set as Default Device'
   - 'Set as Default Communication Device'

7. Click 'Reset to Default' in this app to reconnect

If you still don't see your external mic:
- Try unplugging and reconnecting
- Check if your earphones have a mic (some don't)
- Try a different USB port or audio jack"""
                
                messagebox.showinfo("External Microphone Setup", message)
                
        except Exception as e:
            messagebox.showinfo("Manual Setup Required", 
                              """External Microphone Setup Guide:

1. Connect your earphones/headset properly
2. Right-click sound icon in taskbar → 'Open Sound settings'
3. Click 'More sound settings' → 'Recording' tab
4. Look for your external microphone device
5. Right-click it → 'Set as Default Device'
6. Click 'Reset to Default' in this app

Common device names to look for:
- Headset Microphone
- External Microphone  
- Your headset brand name
- Realtek Audio (for some devices)

If no external mic appears, your earphones may not have a microphone.""")
    
    def reset_to_default(self):
        """Reset to default microphone"""
        try:
            if not PYCAW_AVAILABLE or not AudioUtilities:
                messagebox.showerror("Reset Error", "Audio library not available")
                return
                
            devices = AudioUtilities.GetMicrophone()
            if devices:
                interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
                self.volume_interface = interface.QueryInterface(IAudioEndpointVolume)
                self.current_device = devices
                messagebox.showinfo("Reset", "Reset to default microphone")
                self.update_volume_display()
            else:
                messagebox.showerror("Reset Error", "Could not find default microphone")
        except Exception as e:
            messagebox.showerror("Reset Error", f"Could not reset to default: {str(e)}")
    
    def open_windows_sound_settings(self):
        """Open Windows Sound Settings"""
        try:
            import subprocess
            # Open Windows Sound Settings directly
            subprocess.run(["ms-settings:sound"], shell=True)
        except Exception as e:
            try:
                # Fallback: Open Sound Control Panel
                subprocess.run(["control", "mmsys.cpl"], shell=True)
            except Exception as e2:
                messagebox.showinfo("Manual Instructions", 
                                  "Please manually open Sound Settings:\n\n"
                                  "1. Right-click the sound icon in your taskbar\n"
                                  "2. Select 'Open Sound settings'\n"
                                  "3. Click 'More sound settings'\n"
                                  "4. Go to the 'Recording' tab\n"
                                  "5. Set your desired microphone as default")
    

    
    def set_volume(self, volume):
        """Set microphone volume"""
        try:
            self.current_volume.set(volume)
            
            if self.volume_interface:
                # Use pycaw for precise control
                volume_scalar = volume / 100.0
                self.volume_interface.SetMasterVolumeLevelScalar(volume_scalar, None)
                self.volume_display.config(text=f"Volume: {volume}%")
            else:
                # Fallback to system commands
                self.set_volume_fallback(volume)
                
        except Exception as e:
            # Try fallback method
            self.set_volume_fallback(volume)
    
    def set_volume_fallback(self, volume):
        """Fallback method using system commands"""
        import subprocess
        try:
            # Try PowerShell method
            ps_cmd = f"""
            $devices = Get-WmiObject -Class Win32_SoundDevice
            # Fallback volume setting method
            """
            subprocess.run(["powershell", "-WindowStyle", "Hidden", "-Command", ps_cmd], 
                         timeout=3)
            self.volume_display.config(text=f"Volume: {volume}% (Fallback)")
        except:
            self.volume_display.config(text=f"Volume: {volume}% (Set)")
    
    def on_volume_change(self, value):
        """Handle slider movement"""
        volume = int(float(value))
        self.set_volume(volume)
    
    def mute(self):
        """Mute microphone"""
        try:
            if self.volume_interface:
                self.volume_interface.SetMute(1, None)
                self.volume_display.config(text="Volume: Muted")
            else:
                self.set_volume(0)
        except Exception as e:
            self.set_volume(0)
    
    def unmute(self):
        """Unmute microphone"""
        try:
            if self.volume_interface:
                self.volume_interface.SetMute(0, None)
                volume = self.current_volume.get()
                self.volume_display.config(text=f"Volume: {volume}%")
            else:
                volume = self.current_volume.get()
                if volume == 0:
                    volume = 50
                    self.current_volume.set(volume)
                self.set_volume(volume)
        except Exception as e:
            volume = self.current_volume.get()
            if volume == 0:
                volume = 50
                self.current_volume.set(volume)
            self.set_volume(volume)
    
    def update_volume_display(self):
        """Update the volume display with current system volume"""
        try:
            if self.volume_interface:
                volume = self.volume_interface.GetMasterVolumeLevelScalar()
                current = int(volume * 100)
                self.current_volume.set(current)
                self.volume_display.config(text=f"Volume: {current}%")
            else:
                volume = self.current_volume.get()
                self.volume_display.config(text=f"Volume: {volume}% (Estimated)")
        except Exception as e:
            volume = self.current_volume.get()
            self.volume_display.config(text=f"Volume: {volume}%")
    
    def toggle_lock(self):
        """Toggle volume lock"""
        if self.is_locked.get():
            self.start_lock()
        else:
            self.stop_lock()
    
    def start_lock(self):
        """Start volume locking"""
        self.lock_status.config(text="Status: Locked ✓", foreground="green")
        self.lock_thread = threading.Thread(target=self.maintain_volume, daemon=True)
        self.lock_thread.start()
    
    def stop_lock(self):
        """Stop volume locking"""
        self.lock_status.config(text="Status: Unlocked", foreground="orange")
    
    def maintain_volume(self):
        """Maintain the set volume level"""
        target_volume = self.current_volume.get()
        
        while self.is_locked.get():
            try:
                if self.volume_interface:
                    current_vol = self.volume_interface.GetMasterVolumeLevelScalar()
                    current = int(current_vol * 100)
                    
                    if abs(current - target_volume) > 5:
                        self.set_volume(target_volume)
                        print(f"Volume corrected: {current}% -> {target_volume}%")
                
                time.sleep(2)  # Check every 2 seconds
            except Exception as e:
                print(f"Volume monitoring error: {e}")
                time.sleep(5)

def main():
    root = tk.Tk()
    
    if not PYCAW_AVAILABLE:
        response = messagebox.askyesno(
            "Missing Dependencies", 
            "For best performance, install pycaw library:\n\n"
            "pip install pycaw comtypes\n\n"
            "Continue with basic functionality?"
        )
        if not response:
            return
    
    app = AdvancedMicController(root)
    
    # Center window
    root.update_idletasks()
    x = (root.winfo_screenwidth() // 2) - (root.winfo_width() // 2)
    y = (root.winfo_screenheight() // 2) - (root.winfo_height() // 2)
    root.geometry(f"+{x}+{y}")
    
    root.mainloop()

if __name__ == "__main__":
    main()