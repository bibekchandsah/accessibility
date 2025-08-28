import tkinter as tk
from tkinter import ttk, messagebox
import subprocess
import threading
import time
import re
import json
import os
import ctypes
import sys

# Try to import notification libraries
try:
    from plyer import notification
    PLYER_AVAILABLE = True
except ImportError:
    PLYER_AVAILABLE = False

try:
    import win10toast
    WIN10TOAST_AVAILABLE = True
except ImportError:
    WIN10TOAST_AVAILABLE = False

class AdvancedCameraController:
    def __init__(self, root):
        self.root = root
        self.root.title("Advanced Camera Controller")
        self.root.geometry("705x690")
        
        # Variables
        self.cameras = []
        self.selected_camera = tk.StringVar()
        self.refresh_in_progress = False
        self.auto_refresh = tk.BooleanVar(value=False)
        self.show_all_devices = tk.BooleanVar(value=False)
        self.is_admin = self.check_admin_privileges()
        
        # Settings file
        self.settings_file = "camera_settings.json"
        self.load_settings()
        
        self.setup_ui()
        self.refresh_cameras()
        
        # Auto-refresh timer
        self.schedule_auto_refresh()
        
        # Setup keyboard shortcuts
        self.setup_keyboard_shortcuts()
        
        # Notification system
        self.setup_notifications()
        
        # Show startup notification
        self.root.after(1000, lambda: self.show_startup_notification())
    
    def check_admin_privileges(self):
        """Check if running with administrator privileges"""
        try:
            return ctypes.windll.shell32.IsUserAnAdmin()
        except:
            return False
    
    def setup_notifications(self):
        """Setup notification system"""
        # Test notification system
        if PLYER_AVAILABLE:
            self.notification_method = "plyer"
        elif WIN10TOAST_AVAILABLE:
            self.notification_method = "win10toast"
            self.toaster = win10toast.ToastNotifier()
        else:
            self.notification_method = "fallback"
    
    def toggle_notifications(self):
        """Toggle notification system"""
        self.save_settings()
        if self.notifications_enabled.get():
            self.show_notification("Camera Controller", "Notifications enabled", "info")
        else:
            # Don't show notification when disabling notifications
            pass
    
    def show_startup_notification(self):
        """Show startup notification with shortcuts info"""
        if self.is_admin:
            message = "Ready! Use Shift+Alt+E to enable, Shift+Alt+D to disable cameras"
        else:
            message = "Ready in limited mode! Use Shift+Alt+E to enable cameras"
        
        self.show_notification("Camera Controller Started", message, "info")
    
    def show_notification(self, title, message, icon_type="info"):
        """Show system notification"""
        if not self.notifications_enabled.get():
            return
        
        try:
            if self.notification_method == "plyer":
                notification.notify(
                    title=title,
                    message=message,
                    app_name="Camera Controller",
                    timeout=3
                )
            elif self.notification_method == "win10toast":
                self.toaster.show_toast(
                    title,
                    message,
                    duration=3,
                    threaded=True
                )
            else:
                # Fallback: Use Windows balloon tip
                self.show_balloon_notification(title, message)
        except Exception as e:
            print(f"Notification error: {e}")
    
    def show_balloon_notification(self, title, message):
        """Fallback notification using Windows balloon tip"""
        try:
            # Create a simple balloon notification using PowerShell
            ps_script = f'''
            Add-Type -AssemblyName System.Windows.Forms
            $balloon = New-Object System.Windows.Forms.NotifyIcon
            $balloon.Icon = [System.Drawing.SystemIcons]::Information
            $balloon.BalloonTipTitle = "{title}"
            $balloon.BalloonTipText = "{message}"
            $balloon.Visible = $true
            $balloon.ShowBalloonTip(3000)
            Start-Sleep -Seconds 4
            $balloon.Dispose()
            '''
            
            subprocess.run(["powershell", "-WindowStyle", "Hidden", "-Command", ps_script], 
                          creationflags=subprocess.CREATE_NO_WINDOW)
        except:
            pass
    
    def setup_keyboard_shortcuts(self):
        """Setup keyboard shortcuts"""
        # Bind keyboard shortcuts to the main window
        self.root.bind('<Shift-Alt-KeyPress-e>', lambda e: self.shortcut_enable_camera())
        self.root.bind('<Shift-Alt-KeyPress-E>', lambda e: self.shortcut_enable_camera())
        self.root.bind('<Shift-Alt-KeyPress-d>', lambda e: self.shortcut_disable_camera())
        self.root.bind('<Shift-Alt-KeyPress-D>', lambda e: self.shortcut_disable_camera())
        
        # Make sure the window can receive focus for shortcuts
        self.root.focus_set()
        
        # Add shortcuts info to title
        original_title = self.root.title()
        self.root.title(f"{original_title} - Shortcuts: Shift+Alt+E (Enable), Shift+Alt+D (Disable)")
    
    def shortcut_enable_camera(self):
        """Enable camera via keyboard shortcut"""
        camera = self.get_selected_camera()
        if camera:
            self.enable_camera()
            self.show_notification("Camera Controller", f"Enabling {camera['name']}...", "info")
        else:
            # Enable all cameras if none selected
            if self.cameras:
                self.enable_all_cameras()
                self.show_notification("Camera Controller", "Enabling all cameras...", "info")
            else:
                self.show_notification("Camera Controller", "No cameras found to enable", "warning")
    
    def shortcut_disable_camera(self):
        """Disable camera via keyboard shortcut"""
        if not self.is_admin:
            self.show_notification("Camera Controller", "Administrator privileges required to disable cameras", "warning")
            return
            
        camera = self.get_selected_camera()
        if camera:
            self.disable_camera()
            self.show_notification("Camera Controller", f"Disabling {camera['name']}...", "info")
        else:
            # Disable all cameras if none selected
            if self.cameras:
                enabled_cameras = [c for c in self.cameras if c['status'] == 'OK']
                if enabled_cameras:
                    self.disable_all_cameras()
                    self.show_notification("Camera Controller", "Disabling all cameras...", "info")
                else:
                    self.show_notification("Camera Controller", "All cameras are already disabled", "info")
            else:
                self.show_notification("Camera Controller", "No cameras found to disable", "warning")
        
    def setup_ui(self):
        # Main frame
        main_frame = ttk.Frame(self.root, padding="15")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Title and status
        title_frame = ttk.Frame(main_frame)
        title_frame.pack(fill=tk.X, pady=(0, 15))
        
        ttk.Label(title_frame, text="Advanced Camera Controller", 
                 font=("Arial", 14, "bold")).pack(side=tk.LEFT)
        
        # Admin status and main status
        status_frame = ttk.Frame(title_frame)
        status_frame.pack(side=tk.RIGHT)
        
        admin_text = "🔒 Administrator" if self.is_admin else "⚠️ Limited Mode"
        admin_color = "green" if self.is_admin else "orange"
        ttk.Label(status_frame, text=admin_text, foreground=admin_color, 
                 font=("Arial", 9)).pack(anchor=tk.E)
        
        self.status_label = ttk.Label(status_frame, text="Ready", foreground="green")
        self.status_label.pack(anchor=tk.E)
        
        # Options frame
        options_frame = ttk.LabelFrame(main_frame, text="Options", padding="10")
        options_frame.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Checkbutton(options_frame, text="Auto-refresh every 5 seconds", 
                       variable=self.auto_refresh,
                       command=self.toggle_auto_refresh).pack(side=tk.LEFT, padx=(0, 20))
        
        ttk.Checkbutton(options_frame, text="Show all imaging devices", 
                       variable=self.show_all_devices,
                       command=self.refresh_cameras).pack(side=tk.LEFT, padx=(0, 20))
        
        # Notification toggle
        self.notifications_enabled = tk.BooleanVar(value=True)
        ttk.Checkbutton(options_frame, text="Show notifications", 
                       variable=self.notifications_enabled,
                       command=self.toggle_notifications).pack(side=tk.LEFT)
        
        # Camera list frame
        camera_frame = ttk.LabelFrame(main_frame, text="Connected Cameras", padding="10")
        camera_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 15))
        
        # Listbox with scrollbar and details
        list_frame = ttk.Frame(camera_frame)
        list_frame.pack(fill=tk.BOTH, expand=True)
        
        # Camera listbox
        listbox_frame = ttk.Frame(list_frame)
        listbox_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        scrollbar = ttk.Scrollbar(listbox_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.camera_listbox = tk.Listbox(listbox_frame, yscrollcommand=scrollbar.set, 
                                        font=("Arial", 9), height=10)
        self.camera_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.config(command=self.camera_listbox.yview)
        
        # Bind selection event
        self.camera_listbox.bind('<<ListboxSelect>>', self.on_camera_select)
        
        # Details panel
        details_frame = ttk.LabelFrame(list_frame, text="Device Details", padding="10")
        details_frame.pack(side=tk.RIGHT, fill=tk.Y, padx=(10, 0))
        
        self.details_text = tk.Text(details_frame, width=30, height=10, 
                                   font=("Arial", 8), wrap=tk.WORD)
        details_scrollbar = ttk.Scrollbar(details_frame, orient=tk.VERTICAL, 
                                         command=self.details_text.yview)
        self.details_text.configure(yscrollcommand=details_scrollbar.set)
        
        self.details_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        details_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Refresh button
        ttk.Button(camera_frame, text="🔄 Refresh Camera List", 
                  command=self.refresh_cameras).pack(pady=(10, 0))
        
        # Control buttons frame
        control_frame = ttk.LabelFrame(main_frame, text="Camera Controls", padding="10")
        control_frame.pack(fill=tk.X, pady=(0, 15))
        
        # Primary controls
        primary_frame = ttk.Frame(control_frame)
        primary_frame.pack(pady=(0, 10))
        
        enable_btn = ttk.Button(primary_frame, text="✓ Enable Selected", width=18,
                               command=self.enable_camera)
        enable_btn.pack(side=tk.LEFT, padx=3)
        
        disable_btn = ttk.Button(primary_frame, text="✗ Disable Selected", width=18,
                                command=self.disable_camera)
        disable_btn.pack(side=tk.LEFT, padx=3)
        
        # Add tooltip for non-admin users
        if not self.is_admin:
            self.add_tooltip(disable_btn, "Requires Administrator privileges to disable devices")
        ttk.Button(primary_frame, text="📷 Test Camera", width=15,
                  command=self.test_camera).pack(side=tk.LEFT, padx=3)
        ttk.Button(primary_frame, text="🔍 Diagnose", width=12,
                  command=self.diagnose_camera).pack(side=tk.LEFT, padx=3)
        
        # Secondary controls
        secondary_frame = ttk.Frame(control_frame)
        secondary_frame.pack()
        
        ttk.Button(secondary_frame, text="Enable All", width=12,
                  command=self.enable_all_cameras).pack(side=tk.LEFT, padx=3)
        ttk.Button(secondary_frame, text="Disable All", width=12,
                  command=self.disable_all_cameras).pack(side=tk.LEFT, padx=3)
        ttk.Button(secondary_frame, text="Device Manager", width=15,
                  command=self.open_device_manager).pack(side=tk.LEFT, padx=3)
        ttk.Button(secondary_frame, text="Privacy Settings", width=15,
                  command=self.open_camera_settings).pack(side=tk.LEFT, padx=3)
        
        # Info frame
        info_frame = ttk.LabelFrame(main_frame, text="Information", padding="10")
        info_frame.pack(fill=tk.X)
        
        if self.is_admin:
            info_text = """💡 Tips & Shortcuts:
• Full device control available (Administrator mode)
• Keyboard shortcuts: Shift+Alt+E (Enable), Shift+Alt+D (Disable)
• Notifications show camera status changes
• Use 'Test Camera' to verify camera functionality
• Check Device Manager if cameras don't appear"""
        else:
            info_text = """⚠️ Limited Mode (Not Administrator):
• Can view and enable cameras
• Keyboard shortcuts: Shift+Alt+E (Enable only)
• Cannot disable cameras (requires Administrator privileges)
• Right-click this app → "Run as administrator" for full control
• Use 'Test Camera' to verify camera functionality"""
        
        ttk.Label(info_frame, text=info_text, font=("Arial", 8), 
                 justify=tk.LEFT).pack(anchor=tk.W)
        
        # Add "Run as Admin" button for non-admin users
        if not self.is_admin:
            ttk.Button(info_frame, text="🔒 Restart as Administrator", 
                      command=self.restart_as_admin).pack(pady=(10, 0))
    
    def load_settings(self):
        """Load settings from file"""
        try:
            if os.path.exists(self.settings_file):
                with open(self.settings_file, 'r') as f:
                    settings = json.load(f)
                    self.auto_refresh.set(settings.get('auto_refresh', False))
                    self.show_all_devices.set(settings.get('show_all_devices', False))
                    if hasattr(self, 'notifications_enabled'):
                        self.notifications_enabled.set(settings.get('notifications_enabled', True))
        except:
            pass
    
    def save_settings(self):
        """Save settings to file"""
        try:
            settings = {
                'auto_refresh': self.auto_refresh.get(),
                'show_all_devices': self.show_all_devices.get(),
                'notifications_enabled': self.notifications_enabled.get()
            }
            with open(self.settings_file, 'w') as f:
                json.dump(settings, f)
        except:
            pass
    
    def update_status(self, message, color="black"):
        """Update status label"""
        self.status_label.config(text=message, foreground=color)
        self.root.update()
    
    def toggle_auto_refresh(self):
        """Toggle auto-refresh functionality"""
        self.save_settings()
        if self.auto_refresh.get():
            self.update_status("Auto-refresh enabled", "blue")
        else:
            self.update_status("Auto-refresh disabled", "orange")
    
    def schedule_auto_refresh(self):
        """Schedule auto-refresh if enabled"""
        if self.auto_refresh.get() and not self.refresh_in_progress:
            self.refresh_cameras()
        
        # Schedule next refresh
        self.root.after(5000, self.schedule_auto_refresh)
    
    def refresh_cameras(self):
        """Refresh the list of cameras"""
        if self.refresh_in_progress:
            return
            
        self.refresh_in_progress = True
        self.update_status("Scanning for cameras...", "blue")
        
        def scan_thread():
            try:
                cameras = self.get_camera_devices()
                
                # Update UI in main thread
                self.root.after(0, lambda: self.update_camera_list(cameras))
                
            except Exception as e:
                self.root.after(0, lambda: self.update_status(f"Error: {str(e)}", "red"))
            finally:
                self.refresh_in_progress = False
        
        threading.Thread(target=scan_thread, daemon=True).start()
    
    def get_camera_devices(self):
        """Get list of camera devices using PowerShell"""
        cameras = []
        
        try:
            if self.show_all_devices.get():
                # Show all imaging devices
                ps_script = '''
                # Get all imaging and camera devices
                $devices = @()
                
                # Camera class devices
                $cameras = Get-PnpDevice -Class Camera
                foreach ($camera in $cameras) {
                    $devices += [PSCustomObject]@{
                        Name = $camera.FriendlyName
                        InstanceId = $camera.InstanceId
                        Status = $camera.Status
                        Class = "Camera"
                        Present = $camera.Present
                    }
                }
                
                # Image class devices
                $imageDevices = Get-PnpDevice -Class Image
                foreach ($device in $imageDevices) {
                    $devices += [PSCustomObject]@{
                        Name = $device.FriendlyName
                        InstanceId = $device.InstanceId
                        Status = $device.Status
                        Class = "Image"
                        Present = $device.Present
                    }
                }
                
                # USB Video devices
                $usbDevices = Get-PnpDevice | Where-Object {
                    $_.FriendlyName -like "*camera*" -or 
                    $_.FriendlyName -like "*webcam*" -or
                    $_.FriendlyName -like "*video*" -and
                    $_.FriendlyName -notlike "*audio*"
                }
                foreach ($device in $usbDevices) {
                    $devices += [PSCustomObject]@{
                        Name = $device.FriendlyName
                        InstanceId = $device.InstanceId
                        Status = $device.Status
                        Class = "USB"
                        Present = $device.Present
                    }
                }
                
                # Output unique devices
                $devices | Sort-Object Name -Unique | ForEach-Object {
                    "$($_.Name)|$($_.InstanceId)|$($_.Status)|$($_.Class)|$($_.Present)"
                }
                '''
            else:
                # Show only camera devices
                ps_script = '''
                $cameras = Get-PnpDevice -Class Camera
                foreach ($camera in $cameras) {
                    "$($camera.FriendlyName)|$($camera.InstanceId)|$($camera.Status)|Camera|$($camera.Present)"
                }
                '''
            
            result = subprocess.run(["powershell", "-Command", ps_script], 
                                  capture_output=True, text=True, timeout=15)
            
            if result.returncode == 0:
                lines = result.stdout.strip().split('\n')
                seen_devices = set()
                
                for line in lines:
                    if line.strip() and '|' in line:
                        parts = line.strip().split('|')
                        if len(parts) >= 4:
                            name = parts[0].strip()
                            instance_id = parts[1].strip()
                            status = parts[2].strip()
                            device_class = parts[3].strip()
                            present = parts[4].strip() if len(parts) > 4 else "True"
                            
                            # Avoid duplicates
                            if instance_id not in seen_devices and name:
                                seen_devices.add(instance_id)
                                cameras.append({
                                    'name': name,
                                    'instance_id': instance_id,
                                    'status': status,
                                    'class': device_class,
                                    'present': present
                                })
            
        except Exception as e:
            print(f"Error getting camera devices: {e}")
        
        return cameras
    
    def update_camera_list(self, cameras):
        """Update the camera listbox"""
        self.cameras = cameras
        self.camera_listbox.delete(0, tk.END)
        
        if not cameras:
            self.camera_listbox.insert(tk.END, "No cameras found")
            self.update_status("No cameras detected", "orange")
        else:
            for camera in cameras:
                # Status icons
                if camera['status'] == "OK":
                    status_icon = "✓"
                    status_text = "Enabled"
                elif camera['status'] == "Error":
                    status_icon = "✗"
                    status_text = "Disabled"
                else:
                    status_icon = "?"
                    status_text = camera['status']
                
                # Present indicator
                present_icon = "" if camera['present'] == "True" else " (Not Present)"
                
                display_text = f"{status_icon} {camera['name']} ({status_text}){present_icon}"
                self.camera_listbox.insert(tk.END, display_text)
            
            self.update_status(f"Found {len(cameras)} camera device(s)", "green")
        
        # Clear details
        self.details_text.delete(1.0, tk.END)
    
    def on_camera_select(self, event):
        """Handle camera selection"""
        camera = self.get_selected_camera()
        if camera:
            self.show_camera_details(camera)
    
    def show_camera_details(self, camera):
        """Show detailed information about selected camera"""
        self.details_text.delete(1.0, tk.END)
        
        details = f"""Device: {camera['name']}

Status: {camera['status']}
Class: {camera['class']}
Present: {camera['present']}

Instance ID:
{camera['instance_id']}

Actions Available:
• Enable/Disable device
• Test functionality
• Open in Device Manager
"""
        
        self.details_text.insert(1.0, details)
    
    def get_selected_camera(self):
        """Get the currently selected camera"""
        selection = self.camera_listbox.curselection()
        if selection and self.cameras:
            index = selection[0]
            if index < len(self.cameras):
                return self.cameras[index]
        return None
    
    def enable_camera(self):
        """Enable the selected camera"""
        camera = self.get_selected_camera()
        if not camera:
            messagebox.showwarning("Selection", "Please select a camera to enable")
            return
        
        self.update_status(f"Enabling {camera['name']}...", "blue")
        
        def enable_thread():
            try:
                success = self.change_device_state(camera['instance_id'], enable=True)
                
                if success:
                    self.root.after(0, lambda: self.update_status(f"Enabled {camera['name']}", "green"))
                    self.root.after(0, lambda: self.show_notification("Camera Enabled", f"✓ {camera['name']} is now enabled", "info"))
                    self.root.after(2000, self.refresh_cameras)  # Refresh after 2 seconds
                else:
                    self.root.after(0, lambda: self.update_status("Failed to enable camera - try running as Administrator", "red"))
                    self.root.after(0, lambda: self.show_notification("Enable Failed", f"✗ Could not enable {camera['name']}", "warning"))
                    
            except Exception as e:
                self.root.after(0, lambda: self.update_status(f"Error: {str(e)}", "red"))
        
        threading.Thread(target=enable_thread, daemon=True).start()
    
    def disable_camera(self):
        """Disable the selected camera"""
        camera = self.get_selected_camera()
        if not camera:
            messagebox.showwarning("Selection", "Please select a camera to disable")
            return
        
        # Check admin privileges
        if not self.is_admin:
            if messagebox.askyesno("Administrator Required", 
                                 "Disabling cameras requires Administrator privileges.\n\n"
                                 "Would you like to restart the application as Administrator?"):
                self.restart_as_admin()
            return
        
        # Confirm disable action
        if not messagebox.askyesno("Confirm", 
                                 f"Disable camera '{camera['name']}'?\n\n"
                                 "This will make the camera unavailable to all applications."):
            return
        
        self.update_status(f"Disabling {camera['name']}...", "blue")
        
        def disable_thread():
            try:
                success = self.change_device_state(camera['instance_id'], enable=False)
                
                if success:
                    self.root.after(0, lambda: self.update_status(f"Disabled {camera['name']}", "green"))
                    self.root.after(0, lambda: self.show_notification("Camera Disabled", f"✗ {camera['name']} is now disabled", "info"))
                    self.root.after(2000, self.refresh_cameras)  # Refresh after 2 seconds
                else:
                    self.root.after(0, lambda: self.update_status("Failed to disable camera - use 🔍 Diagnose for details", "red"))
                    self.root.after(0, lambda: self.show_notification("Disable Failed", f"✗ Could not disable {camera['name']}", "warning"))
                    
            except Exception as e:
                self.root.after(0, lambda: self.update_status(f"Error: {str(e)}", "red"))
        
        threading.Thread(target=disable_thread, daemon=True).start()
    
    def change_device_state(self, instance_id, enable=True):
        """Enable or disable a device using multiple methods"""
        action = "Enable" if enable else "Disable"
        
        # Method 1: Try PnpDevice cmdlet with proper escaping
        success = self.try_pnp_device_method(instance_id, enable)
        if success:
            return True
        
        # Method 2: Try DevCon-style approach
        success = self.try_devcon_method(instance_id, enable)
        if success:
            return True
        
        # Method 3: Try WMI method
        success = self.try_wmi_method(instance_id, enable)
        if success:
            return True
        
        return False
    
    def try_pnp_device_method(self, instance_id, enable=True):
        """Try using PnpDevice PowerShell cmdlet"""
        try:
            action = "Enable" if enable else "Disable"
            
            # Clean and escape the instance ID
            clean_id = instance_id.strip().replace("'", "''")
            
            ps_script = f'''
            $ErrorActionPreference = "Stop"
            try {{
                $deviceId = '{clean_id}'
                Write-Host "Attempting to {action.lower()} device: $deviceId"
                
                # First check if device exists
                $device = Get-PnpDevice -InstanceId $deviceId -ErrorAction SilentlyContinue
                if (-not $device) {{
                    Write-Error "Device not found: $deviceId"
                    exit 1
                }}
                
                Write-Host "Device found: $($device.FriendlyName)"
                Write-Host "Current status: $($device.Status)"
                
                # Perform the action
                {action}-PnpDevice -InstanceId $deviceId -Confirm:$false
                
                # Verify the change
                Start-Sleep -Seconds 1
                $updatedDevice = Get-PnpDevice -InstanceId $deviceId
                Write-Host "New status: $($updatedDevice.Status)"
                
                Write-Output "SUCCESS"
            }} catch {{
                Write-Host "Error: $($_.Exception.Message)"
                Write-Error $_.Exception.Message
                exit 1
            }}
            '''
            
            result = subprocess.run(["powershell", "-ExecutionPolicy", "Bypass", "-Command", ps_script], 
                                  capture_output=True, text=True, timeout=30)
            
            print(f"PnpDevice method - Return code: {result.returncode}")
            print(f"PnpDevice method - Output: {result.stdout}")
            print(f"PnpDevice method - Error: {result.stderr}")
            
            return result.returncode == 0 and "SUCCESS" in result.stdout
            
        except Exception as e:
            print(f"PnpDevice method failed: {e}")
            return False
    
    def try_devcon_method(self, instance_id, enable=True):
        """Try using devcon-style commands"""
        try:
            action = "enable" if enable else "disable"
            
            # Extract hardware ID pattern from instance ID
            # Instance IDs often look like: USB\VID_1234&PID_5678\SerialNumber
            parts = instance_id.split('\\')
            if len(parts) >= 2:
                hardware_pattern = f"{parts[0]}\\{parts[1]}"
            else:
                hardware_pattern = instance_id
            
            ps_script = f'''
            $ErrorActionPreference = "Stop"
            try {{
                $pattern = "{hardware_pattern}"
                Write-Host "Trying devcon-style method with pattern: $pattern"
                
                # Find devices matching the pattern
                $devices = Get-PnpDevice | Where-Object {{ $_.InstanceId -like "*$pattern*" }}
                
                if ($devices) {{
                    foreach ($device in $devices) {{
                        Write-Host "Found device: $($device.FriendlyName) - $($device.InstanceId)"
                        try {{
                            if ("{action}" -eq "enable") {{
                                Enable-PnpDevice -InstanceId $device.InstanceId -Confirm:$false
                            }} else {{
                                Disable-PnpDevice -InstanceId $device.InstanceId -Confirm:$false
                            }}
                            Write-Host "Successfully {action}d: $($device.FriendlyName)"
                        }} catch {{
                            Write-Host "Failed to {action}: $($_.Exception.Message)"
                        }}
                    }}
                    Write-Output "SUCCESS"
                }} else {{
                    Write-Error "No devices found matching pattern: $pattern"
                    exit 1
                }}
            }} catch {{
                Write-Error $_.Exception.Message
                exit 1
            }}
            '''
            
            result = subprocess.run(["powershell", "-ExecutionPolicy", "Bypass", "-Command", ps_script], 
                                  capture_output=True, text=True, timeout=30)
            
            print(f"DevCon method - Return code: {result.returncode}")
            print(f"DevCon method - Output: {result.stdout}")
            
            return result.returncode == 0 and "SUCCESS" in result.stdout
            
        except Exception as e:
            print(f"DevCon method failed: {e}")
            return False
    
    def try_wmi_method(self, instance_id, enable=True):
        """Try using WMI method"""
        try:
            action_code = "1" if enable else "0"  # 1 = enable, 0 = disable
            
            ps_script = f'''
            $ErrorActionPreference = "Stop"
            try {{
                $instanceId = "{instance_id}"
                Write-Host "Trying WMI method for: $instanceId"
                
                # Find the device using WMI
                $device = Get-WmiObject -Class Win32_PnPEntity | Where-Object {{ $_.DeviceID -eq $instanceId }}
                
                if ($device) {{
                    Write-Host "Found WMI device: $($device.Name)"
                    
                    # Try to enable/disable using WMI
                    if ("{enable}".ToLower() -eq "true") {{
                        $result = $device.Enable()
                    }} else {{
                        $result = $device.Disable()
                    }}
                    
                    if ($result.ReturnValue -eq 0) {{
                        Write-Host "WMI operation successful"
                        Write-Output "SUCCESS"
                    }} else {{
                        Write-Error "WMI operation failed with code: $($result.ReturnValue)"
                        exit 1
                    }}
                }} else {{
                    Write-Error "Device not found in WMI: $instanceId"
                    exit 1
                }}
            }} catch {{
                Write-Error $_.Exception.Message
                exit 1
            }}
            '''
            
            result = subprocess.run(["powershell", "-ExecutionPolicy", "Bypass", "-Command", ps_script], 
                                  capture_output=True, text=True, timeout=30)
            
            print(f"WMI method - Return code: {result.returncode}")
            print(f"WMI method - Output: {result.stdout}")
            
            return result.returncode == 0 and "SUCCESS" in result.stdout
            
        except Exception as e:
            print(f"WMI method failed: {e}")
            return False
    
    def test_camera(self):
        """Test the selected camera"""
        camera = self.get_selected_camera()
        if not camera:
            messagebox.showwarning("Selection", "Please select a camera to test")
            return
        
        if camera['status'] != "OK":
            if messagebox.askyesno("Camera Issue", 
                                 f"Camera '{camera['name']}' status is '{camera['status']}'.\n\n"
                                 "Try to enable it first?"):
                self.enable_camera()
                return
        
        # Show test options
        test_window = tk.Toplevel(self.root)
        test_window.title("Test Camera")
        test_window.geometry("300x200")
        test_window.transient(self.root)
        test_window.grab_set()
        
        ttk.Label(test_window, text=f"Test: {camera['name']}", 
                 font=("Arial", 12, "bold")).pack(pady=20)
        
        ttk.Button(test_window, text="📷 Windows Camera App", width=25,
                  command=lambda: self.open_camera_app(test_window)).pack(pady=5)
        ttk.Button(test_window, text="🌐 Browser Camera Test", width=25,
                  command=lambda: self.open_browser_test(test_window)).pack(pady=5)
        ttk.Button(test_window, text="⚙️ Camera Settings", width=25,
                  command=lambda: self.open_camera_settings(test_window)).pack(pady=5)
        
        ttk.Button(test_window, text="Close", 
                  command=test_window.destroy).pack(pady=20)
    
    def open_camera_app(self, parent_window=None):
        """Open Windows Camera app"""
        try:
            subprocess.run(["start", "microsoft.windows.camera:"], shell=True)
            self.update_status("Opened Camera app", "green")
            if parent_window:
                parent_window.destroy()
        except Exception as e:
            messagebox.showerror("Error", f"Could not open Camera app:\n{str(e)}")
    
    def open_browser_test(self, parent_window=None):
        """Open browser camera test"""
        try:
            import webbrowser
            webbrowser.open("https://webcamtests.com/")
            self.update_status("Opened browser camera test", "green")
            if parent_window:
                parent_window.destroy()
        except Exception as e:
            messagebox.showerror("Error", f"Could not open browser test:\n{str(e)}")
    
    def enable_all_cameras(self):
        """Enable all detected cameras"""
        if not self.cameras:
            messagebox.showinfo("No Cameras", "No cameras detected. Try refreshing the list.")
            return
        
        disabled_cameras = [c for c in self.cameras if c['status'] != 'OK']
        if not disabled_cameras:
            messagebox.showinfo("All Enabled", "All cameras are already enabled.")
            return
        
        if not messagebox.askyesno("Confirm", f"Enable {len(disabled_cameras)} disabled camera(s)?"):
            return
        
        self.update_status("Enabling all cameras...", "blue")
        
        def enable_all_thread():
            success_count = 0
            for camera in disabled_cameras:
                try:
                    if self.change_device_state(camera['instance_id'], enable=True):
                        success_count += 1
                    time.sleep(1)  # Delay between operations
                except:
                    continue
            
            self.root.after(0, lambda: self.update_status(f"Enabled {success_count}/{len(disabled_cameras)} cameras", "green"))
            self.root.after(0, lambda: self.show_notification("Bulk Enable Complete", f"✓ Enabled {success_count}/{len(disabled_cameras)} cameras", "info"))
            self.root.after(2000, self.refresh_cameras)
        
        threading.Thread(target=enable_all_thread, daemon=True).start()
    
    def disable_all_cameras(self):
        """Disable all detected cameras"""
        if not self.cameras:
            messagebox.showinfo("No Cameras", "No cameras detected. Try refreshing the list.")
            return
        
        # Check admin privileges
        if not self.is_admin:
            if messagebox.askyesno("Administrator Required", 
                                 "Disabling cameras requires Administrator privileges.\n\n"
                                 "Would you like to restart the application as Administrator?"):
                self.restart_as_admin()
            return
        
        enabled_cameras = [c for c in self.cameras if c['status'] == 'OK']
        if not enabled_cameras:
            messagebox.showinfo("All Disabled", "All cameras are already disabled.")
            return
        
        if not messagebox.askyesno("Confirm", 
                                 f"Disable {len(enabled_cameras)} enabled camera(s)?\n\n"
                                 "This will make all cameras unavailable to applications."):
            return
        
        self.update_status("Disabling all cameras...", "blue")
        
        def disable_all_thread():
            success_count = 0
            for camera in enabled_cameras:
                try:
                    if self.change_device_state(camera['instance_id'], enable=False):
                        success_count += 1
                    time.sleep(1)  # Delay between operations
                except:
                    continue
            
            self.root.after(0, lambda: self.update_status(f"Disabled {success_count}/{len(enabled_cameras)} cameras", "green"))
            self.root.after(0, lambda: self.show_notification("Bulk Disable Complete", f"✗ Disabled {success_count}/{len(enabled_cameras)} cameras", "info"))
            self.root.after(2000, self.refresh_cameras)
        
        threading.Thread(target=disable_all_thread, daemon=True).start()
    
    def open_device_manager(self):
        """Open Windows Device Manager"""
        try:
            subprocess.run(["devmgmt.msc"], shell=True)
            self.update_status("Opened Device Manager", "green")
        except Exception as e:
            messagebox.showerror("Error", f"Could not open Device Manager:\n{str(e)}")
    
    def open_camera_settings(self, parent_window=None):
        """Open Windows Camera settings"""
        try:
            # Try camera privacy settings first
            subprocess.run(["start", "ms-settings:privacy-webcam"], shell=True)
            self.update_status("Opened Camera privacy settings", "green")
            if parent_window:
                parent_window.destroy()
        except Exception as e:
            try:
                # Fallback: Open general privacy settings
                subprocess.run(["start", "ms-settings:privacy"], shell=True)
                self.update_status("Opened Privacy settings", "green")
                if parent_window:
                    parent_window.destroy()
            except:
                messagebox.showerror("Error", f"Could not open Camera settings:\n{str(e)}")
    
    def diagnose_camera(self):
        """Diagnose the selected camera for troubleshooting"""
        camera = self.get_selected_camera()
        if not camera:
            messagebox.showwarning("Selection", "Please select a camera to diagnose")
            return
        
        # Create diagnosis window
        diag_window = tk.Toplevel(self.root)
        diag_window.title(f"Camera Diagnosis - {camera['name']}")
        diag_window.geometry("600x400")
        diag_window.transient(self.root)
        
        # Diagnosis text area
        text_frame = ttk.Frame(diag_window)
        text_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        diag_text = tk.Text(text_frame, wrap=tk.WORD, font=("Consolas", 9))
        scrollbar = ttk.Scrollbar(text_frame, orient=tk.VERTICAL, command=diag_text.yview)
        diag_text.configure(yscrollcommand=scrollbar.set)
        
        diag_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Close button
        ttk.Button(diag_window, text="Close", command=diag_window.destroy).pack(pady=10)
        
        # Run diagnosis in thread
        def run_diagnosis():
            diag_text.insert(tk.END, f"🔍 CAMERA DIAGNOSIS REPORT\n")
            diag_text.insert(tk.END, f"{'='*50}\n\n")
            diag_text.insert(tk.END, f"Device: {camera['name']}\n")
            diag_text.insert(tk.END, f"Status: {camera['status']}\n")
            diag_text.insert(tk.END, f"Instance ID: {camera['instance_id']}\n\n")
            diag_text.see(tk.END)
            diag_window.update()
            
            # Test 1: Check device existence
            diag_text.insert(tk.END, "TEST 1: Device Existence Check\n")
            diag_text.insert(tk.END, "-" * 30 + "\n")
            
            ps_script = f'''
            $deviceId = '{camera['instance_id']}'
            try {{
                $device = Get-PnpDevice -InstanceId $deviceId -ErrorAction Stop
                Write-Output "✓ Device found: $($device.FriendlyName)"
                Write-Output "  Status: $($device.Status)"
                Write-Output "  Class: $($device.Class)"
                Write-Output "  Present: $($device.Present)"
                Write-Output "  Problem Code: $($device.ProblemCode)"
            }} catch {{
                Write-Output "✗ Device not found or error: $($_.Exception.Message)"
            }}
            '''
            
            try:
                result = subprocess.run(["powershell", "-Command", ps_script], 
                                      capture_output=True, text=True, timeout=10)
                diag_text.insert(tk.END, result.stdout + "\n")
            except Exception as e:
                diag_text.insert(tk.END, f"✗ Error running test: {e}\n")
            
            diag_text.see(tk.END)
            diag_window.update()
            
            # Test 2: Check permissions
            diag_text.insert(tk.END, "\nTEST 2: Permission Check\n")
            diag_text.insert(tk.END, "-" * 30 + "\n")
            
            ps_script2 = '''
            try {
                $isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")
                if ($isAdmin) {
                    Write-Output "✓ Running with Administrator privileges"
                } else {
                    Write-Output "✗ NOT running with Administrator privileges"
                }
                
                # Test if we can access PnP cmdlets
                $testDevice = Get-PnpDevice | Select-Object -First 1
                Write-Output "✓ PnP cmdlets accessible"
                
            } catch {
                Write-Output "✗ Error checking permissions: $($_.Exception.Message)"
            }
            '''
            
            try:
                result = subprocess.run(["powershell", "-Command", ps_script2], 
                                      capture_output=True, text=True, timeout=10)
                diag_text.insert(tk.END, result.stdout + "\n")
            except Exception as e:
                diag_text.insert(tk.END, f"✗ Error running permission test: {e}\n")
            
            diag_text.see(tk.END)
            diag_window.update()
            
            # Test 3: Try disable operation (dry run)
            diag_text.insert(tk.END, "\nTEST 3: Disable Operation Test\n")
            diag_text.insert(tk.END, "-" * 30 + "\n")
            
            ps_script3 = f'''
            $deviceId = '{camera['instance_id']}'
            try {{
                Write-Output "Attempting to disable device (test)..."
                
                # Check current status first
                $device = Get-PnpDevice -InstanceId $deviceId
                Write-Output "Current device status: $($device.Status)"
                
                if ($device.Status -eq "OK") {{
                    Write-Output "Device is currently enabled - disable operation should work"
                }} elseif ($device.Status -eq "Error") {{
                    Write-Output "Device is already disabled"
                }} else {{
                    Write-Output "Device status is: $($device.Status)"
                }}
                
                # Check if device can be disabled (some devices are protected)
                $deviceClass = $device.Class
                Write-Output "Device class: $deviceClass"
                
                if ($deviceClass -eq "Camera") {{
                    Write-Output "✓ Standard camera device - should be controllable"
                }} else {{
                    Write-Output "⚠ Non-standard device class - may have restrictions"
                }}
                
            }} catch {{
                Write-Output "✗ Error in disable test: $($_.Exception.Message)"
            }}
            '''
            
            try:
                result = subprocess.run(["powershell", "-Command", ps_script3], 
                                      capture_output=True, text=True, timeout=10)
                diag_text.insert(tk.END, result.stdout + "\n")
            except Exception as e:
                diag_text.insert(tk.END, f"✗ Error running disable test: {e}\n")
            
            diag_text.insert(tk.END, "\n" + "="*50 + "\n")
            diag_text.insert(tk.END, "DIAGNOSIS COMPLETE\n")
            diag_text.insert(tk.END, "\nIf disable still fails, the device may be:\n")
            diag_text.insert(tk.END, "• Protected by Windows security policy\n")
            diag_text.insert(tk.END, "• In use by another application\n")
            diag_text.insert(tk.END, "• A built-in device with hardware restrictions\n")
            diag_text.insert(tk.END, "• Managed by enterprise policy\n")
            
            diag_text.see(tk.END)
        
        threading.Thread(target=run_diagnosis, daemon=True).start()
    
    def add_tooltip(self, widget, text):
        """Add tooltip to widget"""
        def on_enter(event):
            tooltip = tk.Toplevel()
            tooltip.wm_overrideredirect(True)
            tooltip.wm_geometry(f"+{event.x_root+10}+{event.y_root+10}")
            label = tk.Label(tooltip, text=text, background="lightyellow", 
                           relief="solid", borderwidth=1, font=("Arial", 8))
            label.pack()
            widget.tooltip = tooltip
        
        def on_leave(event):
            if hasattr(widget, 'tooltip'):
                widget.tooltip.destroy()
                del widget.tooltip
        
        widget.bind("<Enter>", on_enter)
        widget.bind("<Leave>", on_leave)
    
    def restart_as_admin(self):
        """Restart the application as administrator"""
        try:
            script_path = os.path.abspath(sys.argv[0])
            
            # Use ShellExecute for more reliable elevation
            ctypes.windll.shell32.ShellExecuteW(
                None, 
                "runas", 
                sys.executable, 
                f'"{script_path}"', 
                None, 
                1
            )
            self.root.quit()
        except Exception as e:
            # Fallback to PowerShell method
            try:
                subprocess.run([
                    "powershell", "-Command", 
                    f"Start-Process python -ArgumentList '{script_path}' -Verb RunAs"
                ])
                self.root.quit()
            except:
                messagebox.showerror("Error", f"Could not restart as administrator:\n{str(e)}")

def is_admin():
    """Check if the current process has admin privileges"""
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

def run_as_admin():
    """Re-run the current script with admin privileges"""
    try:
        if is_admin():
            # Already running as admin
            return True
        else:
            # Re-run as admin
            script_path = os.path.abspath(sys.argv[0])
            params = ' '.join([f'"{arg}"' for arg in sys.argv[1:]])
            
            # Use ShellExecute to run as admin
            ctypes.windll.shell32.ShellExecuteW(
                None, 
                "runas", 
                sys.executable, 
                f'"{script_path}" {params}', 
                None, 
                1
            )
            return False
    except Exception as e:
        print(f"Failed to elevate privileges: {e}")
        return False

def show_elevation_dialog():
    """Show a dialog explaining why admin privileges are needed"""
    root = tk.Tk()
    root.withdraw()  # Hide the main window
    
    result = messagebox.askyesno(
        "Administrator Privileges Required",
        "Camera Controller needs Administrator privileges to:\n\n"
        "• Enable and disable camera devices\n"
        "• Access device management functions\n"
        "• Control hardware device states\n\n"
        "Click 'Yes' to restart with Administrator privileges\n"
        "Click 'No' to continue in limited mode (view-only)",
        icon='question'
    )
    
    root.destroy()
    return result

def main():
    # Check if we need to elevate privileges
    if not is_admin():
        # Check if user wants to run as admin
        if "--skip-admin" not in sys.argv:
            if show_elevation_dialog():
                # User chose to run as admin
                if run_as_admin():
                    # Successfully elevated, continue
                    pass
                else:
                    # Elevation initiated, exit this instance
                    return
            else:
                # User chose to continue in limited mode
                pass
    
    # Continue with normal execution
    root = tk.Tk()
    app = AdvancedCameraController(root)
    
    # Center window on screen
    root.update_idletasks()
    x = (root.winfo_screenwidth() // 2) - (root.winfo_width() // 2)
    y = (root.winfo_screenheight() // 2) - (root.winfo_height() // 2)
    root.geometry(f"+{x}+{y}")
    
    # Save settings on close
    def on_closing():
        app.save_settings()
        root.destroy()
    
    root.protocol("WM_DELETE_WINDOW", on_closing)
    root.mainloop()

if __name__ == "__main__":
    main()