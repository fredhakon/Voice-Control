# Voice Control Platform

A Windows voice control platform built with Python that allows you to control your system using voice commands.

## Features

- **Voice Recognition**: Uses Google Speech Recognition for accurate voice detection
- **System Volume Control**: Control volume with voice commands ("volume up", "set volume 50")
- **Media Controls**: Play, pause, next/previous track commands
- **Custom Commands**: Add your own voice commands to open files and applications
- **Toggle Listening**: Button and keyboard shortcut to enable/disable listening
- **Settings**: Configure microphone input, keyboard shortcuts, and recognition settings

## Built-in Voice Commands

### Volume Control
- `"Volume up"` - Increase volume by 10%
- `"Volume down"` - Decrease volume by 10%
- `"Set volume [0-100]"` - Set volume to a specific level
- `"Mute"` - Mute system audio
- `"Unmute"` - Unmute system audio
- `"Toggle mute"` - Toggle mute state

### Media Control
- `"Play"` / `"Pause"` / `"Play pause"` - Toggle playback
- `"Next track"` - Skip to next track
- `"Previous track"` - Go to previous track
- `"Stop"` - Stop playback

## Installation

### Prerequisites
- Python 3.9 or higher
- Windows 10/11
- Microphone

### Setup

1. Install the required dependencies:
   ```powershell
   pip install -r requirements.txt
   ```

2. Run the application:
   ```powershell
   python main.py
   ```

## Usage

1. **Start Listening**: Click the "🎤 Start Listening" button or press the keyboard shortcut (default: `Ctrl+Shift+V`)
2. **Speak Commands**: Say any of the built-in commands or your custom commands
3. **Stop Listening**: Click the button again or use the shortcut

### Adding Custom Commands

1. Go to the "Custom Commands" tab
2. Click "➕ Add Command"
3. Enter a voice phrase (e.g., "open notepad")
4. Select the file or application to open
5. Click "Add"

### Configuring Settings

1. Go to File → Settings
2. **Microphone**: Select your preferred microphone input
3. **Shortcut**: Record a new keyboard shortcut for toggling listening
4. **Energy Threshold**: Adjust sensitivity for voice detection

## Project Structure

```
Voice Control/
├── main.py              # Main GUI application
├── voice_control.py     # Core voice recognition and command logic
├── config.json          # User settings and custom commands
├── requirements.txt     # Python dependencies
└── README.md            # This file
```

## Dependencies

- `SpeechRecognition` - Voice recognition
- `PyAudio` - Audio input handling
- `pycaw` - Windows audio control
- `comtypes` - COM interface support
- `pyautogui` - Media key simulation
- `keyboard` - Global hotkey registration
- `Pillow` - Image processing

## Troubleshooting

### Microphone not detected
- Ensure your microphone is connected and working
- Check Windows privacy settings for microphone access
- Try selecting a different microphone in Settings

### Recognition errors
- Speak clearly and at a moderate pace
- Adjust the Energy Threshold in Settings
- Ensure you have an active internet connection (required for Google Speech API)

### Shortcut not working
- Run the application as administrator
- Check if another application is using the same shortcut

## License

GNU Affero General Public License v3.0 (AGPLv3)
