# Audio & Camera Device Controllers

Python GUI applications to control microphone volume and camera devices on Windows systems.

## Problems Solved

### Microphone Issues
- Windows automatically changes microphone volume during calls
- Need manual control over microphone volume levels
- Want to lock volume at a specific level during important calls

### Camera Issues
- Integrated cameras sometimes get automatically disabled
- Need to manually enable/disable cameras for privacy or troubleshooting
- Want quick access to camera controls without diving into Device Manager

## Features

### Microphone Controller
- **Volume Control**: Set microphone volume from 0-100%
- **Quick Presets**: Buttons for 25%, 50%, 75%, 100% volume
- **Mute/Unmute**: Instant mute functionality
- **Volume Lock**: Prevents automatic volume changes by Windows
- **Real-time Monitoring**: Maintains your preferred volume level

### Camera Controller
- **Enable/Disable Cameras**: Control camera device states
- **Auto-refresh**: Automatically scan for camera changes
- **Device Details**: View detailed camera information
- **Bulk Operations**: Enable or disable all cameras at once
- **Testing Tools**: Quick access to camera testing applications
- **Privacy Controls**: Easy access to Windows camera settings

## Applications Included

### Microphone Controllers

#### 1. `advanced_mic_controller.py`
- Enhanced version with better audio control
- Uses pycaw library for precise volume control
- Real-time volume monitoring and device detection
- Progress dialog with real-time feedback
- External microphone detection and switching

### Camera Controllers

#### 2. `camera_controller.py`
- Basic camera enable/disable functionality
- Simple GUI for camera management
- PowerShell-based device control
- No additional dependencies required

#### 3. `advanced_camera_controller.py` ⭐ **Recommended**
- **Complete camera management solution**
- Enable/disable individual or all cameras
- Auto-refresh camera list every 5 seconds
- Detailed device information panel
- Multiple testing options (Windows Camera app, browser test)
- Settings persistence
- Integration with Windows Device Manager and Privacy Settings
- Show all imaging devices option

## Quick Start

### For Camera Control (No Dependencies)
```bash
# Basic camera controller
python camera_controller.py

# Advanced camera controller (recommended)
python advanced_camera_controller.py
```

### For Microphone Control
1. Install Python dependencies:
   ```bash
   pip install pycaw comtypes
   ```
2. Run the microphone controller:
   ```bash
   python advanced_mic_controller.py
   ```

## How to Use

### Camera Controller
1. **Launch**: Run `python advanced_camera_controller.py`
2. **View Cameras**: The app automatically scans and lists all connected cameras
3. **Enable/Disable**: Select a camera and click "Enable Selected" or "Disable Selected"
4. **Test Camera**: Use "Test Camera" to verify functionality with multiple testing options
5. **Bulk Operations**: Use "Enable All" or "Disable All" for quick privacy control
6. **Auto-refresh**: Enable auto-refresh to monitor camera status changes

### Microphone Controller
1. **Set Volume**: Use the slider or preset buttons to set your desired volume
2. **Lock Volume**: Check "Lock volume" to prevent Windows from changing it
3. **Detect External Microphones**: 
   - Click "Detect Devices" to scan for external microphones (headsets, earphones)
   - Watch the progress dialog show real-time detection results
   - Select your preferred microphone from the list
4. **During Calls**: 
   - Set your preferred volume (usually 75-100%)
   - Enable volume lock
   - The app will maintain your volume level automatically

## Tips for Best Results

- **Run as Administrator**: For better system access and control
- **Use During Calls**: Enable volume lock before starting important calls
- **Test First**: Try different volume levels to find what works best
- **Keep Running**: Leave the app running in background during calls

## Troubleshooting

### Volume Control Not Working
- Try running as administrator
- Use the simple version if advanced version fails
- Check if your microphone is set as default recording device

### Dependencies Issues
- Install Python from https://python.org
- Use the simple version which requires no extra packages
- Try: `pip install --user pycaw comtypes`

### Volume Still Changes Automatically
- Make sure "Lock volume" is enabled
- Check Windows sound settings for automatic gain control
- Disable "Allow applications to take exclusive control" in sound properties

## Windows Sound Settings

To complement this app, also check these Windows settings:
1. Right-click sound icon → "Open Sound settings"
2. Go to "Sound Control Panel" → "Recording"
3. Right-click your microphone → "Properties"
4. In "Advanced" tab, uncheck "Allow applications to take exclusive control"
5. In "Levels" tab, set your preferred level
6. In "Enhancements" tab, disable all enhancements

## System Requirements
- Windows 10/11
- Python 3.6+
- Administrator privileges (recommended)

## License
Free to use and modify for personal and commercial purposes.