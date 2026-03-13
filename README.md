# Voice Control Platform

A Windows voice control application built with Python and CustomTkinter that allows hands-free PC control through voice commands. Supports multiple speech recognition engines, Spotify integration, custom macros, window management, and user profiles.

## Features

- **Multiple Speech Engines** - Google Speech API (online), Vosk (offline), Faster-Whisper (offline)
- **System Volume Control** - Adjust, mute, and set volume levels via voice
- **Media Controls** - Play, pause, skip, previous, and stop playback
- **Spotify Integration** - Search and play songs, artists, albums, playlists, radio mode, recommendations, and playlist management
- **Custom Commands** - Open files/apps, launch URLs, trigger keyboard/mouse macros, manage windows, and chain multiple actions together
- **Natural Language Timers** - "Set a timer for 5 minutes and 30 seconds"
- **Window Management** - Minimize, maximize, restore, close, and snap windows to screen positions
- **Macro System** - Record keyboard combos and mouse actions with configurable delays
- **Chain Actions** - Sequence multiple actions into a single voice command
- **Profiles** - Multiple user profiles with independent commands and settings
- **Toggle Listening** - Button and keyboard shortcut (default: `Ctrl+Shift+V`)

## Built-in Voice Commands

### Volume Control

| Command | Aliases | Description |
|---------|---------|-------------|
| `volume up` | `increase volume`, `louder` | Increase volume by 10% |
| `volume down` | `decrease volume`, `quieter` | Decrease volume by 10% |
| `set volume [0-100]` | | Set volume to a specific level |
| `mute` | | Mute system audio |
| `unmute` | | Unmute system audio |
| `toggle mute` | | Toggle mute state |

### Media Control

| Command | Aliases | Description |
|---------|---------|-------------|
| `play` / `pause` | `play pause` | Toggle playback |
| `next track` | `skip`, `next song` | Skip to next track |
| `previous track` | `previous song`, `go back` | Go to previous track |
| `stop` | `stop playing` | Stop playback |

### Timer

| Command | Aliases | Description |
|---------|---------|-------------|
| `set a timer for [duration]` | `set timer`, `timer for`, `start timer` | Supports hours, minutes, seconds and word numbers |
| `stop timer` | `cancel timer`, `stop the timer` | Cancel the active timer |

### Spotify Commands

> Requires Spotify API credentials (Client ID and Secret) configured in Settings.

| Command | Aliases | Description |
|---------|---------|-------------|
| `play song [name]` | `play track [name]` | Search and play a song |
| `play artist [name]` | | Play an artist's top tracks |
| `play album [name]` | | Play an album |
| `play playlist [name]` | `play my playlist [name]` | Play a playlist |
| `play recommendations` | `play something similar`, `recommend something` | Play tracks similar to current song |
| `play radio` | `start radio` | Start artist radio based on current track |
| `shuffle on` | `enable shuffle` | Enable shuffle |
| `shuffle off` | `disable shuffle` | Disable shuffle |
| `repeat track` | `repeat on`, `repeat this song` | Repeat current track |
| `repeat all` | `repeat playlist`, `repeat album` | Repeat playlist/album |
| `repeat off` | `no repeat`, `stop repeat` | Disable repeat |
| `add to playlist [name]` | `add this to [name]`, `save to [name]` | Add current track to a playlist |
| `what's playing` | `current song`, `now playing` | Show current track info |
| `spotify volume [0-100]` | `set spotify volume [number]` | Set Spotify volume |

## Custom Command Types

| Type | Description |
|------|-------------|
| **Open File/App** | Launch any file or application |
| **Open Website** | Open a URL in the default browser |
| **Keyboard/Mouse Macro** | Recorded keyboard combos and mouse actions with delays |
| **Window Action** | Minimize, maximize, restore, close, or snap windows (target by title, app name, or focused window) |
| **Chain Actions** | Sequence multiple actions into one command with configurable delays |

## Installation

### Prerequisites

- Python 3.9 or higher
- Windows 10/11
- Microphone

### Setup

1. Clone the repository:
   ```powershell
   git clone https://github.com/YOUR_USERNAME/voice-control.git
   cd voice-control
   ```

2. Create a virtual environment (recommended):
   ```powershell
   python -m venv .venv
   .venv\Scripts\Activate.ps1
   ```

3. Install dependencies:
   ```powershell
   pip install -r requirements.txt
   ```

4. Run the application:
   ```powershell
   python main.py
   ```

### Offline Speech Recognition

- **Vosk** - Download a Vosk model and place it in a `vosk-model/` folder in the project root. The default config uses Vosk.
- **Faster-Whisper** - Uses the "tiny" model, downloaded automatically on first use.

## Usage

1. **Start Listening** - Click the "Start Listening" button or press `Ctrl+Shift+V`
2. **Speak Commands** - Say any built-in command, Spotify command, or your custom commands
3. **Stop Listening** - Click the button again or use the shortcut

### Adding Custom Commands

1. Click **Add Command**
2. Enter one or more voice phrases (one per line)
3. Select the command type (file, URL, macro, window action, or chain)
4. Configure the action and click **Add**

### Configuring Settings

| Setting | Description |
|---------|-------------|
| **Microphone** | Select input device |
| **Recognition Engine** | Choose between Google, Vosk, or Whisper |
| **Sensitivity** | Adjust energy threshold for voice detection (lower = more sensitive) |
| **Pause Duration** | How long to wait for silence before processing (0.5-3.0s) |
| **Audio Output** | Select output device for volume control |
| **Keyboard Shortcut** | Record a custom toggle shortcut |
| **Spotify** | Enter Client ID/Secret to connect |

### Profiles

Create multiple profiles with independent commands and settings. Each profile stores its own custom commands, built-in phrase customizations, disabled categories, and timer settings. Profiles can be exported and imported as JSON files.

## Building from Source

### Create Executable

```powershell
.\build.bat
```

Runs PyInstaller and produces `dist/VoiceControl.exe`.

### Create Installer

```powershell
.\build_installer.bat
```

Requires [Inno Setup 6](https://jrsoftware.org/isinfo.php). Produces `installer/VoiceControlSetup.exe`.

## Project Structure

```
Voice Control/
├── main.py              # GUI application (CustomTkinter)
├── voice_control.py     # Core engine (recognition, volume, media, commands)
├── spotify_control.py   # Spotify API integration (OAuth2, playback, search)
├── launch.py            # Launcher script
├── config.default.json  # Default configuration template
├── requirements.txt     # Python dependencies
├── VoiceControl.spec    # PyInstaller build spec
├── build.bat            # Build executable script
├── build_installer.bat  # Build installer script
├── installer.iss        # Inno Setup installer config
├── vosk-model/          # Offline speech model (not included, download separately)
└── profiles/            # User profiles (created at runtime)
```

## Dependencies

| Package | Purpose |
|---------|---------|
| `SpeechRecognition` | Voice recognition |
| `PyAudio` | Audio input handling |
| `pycaw` | Windows audio control (COM) |
| `comtypes` | COM interface support |
| `pyautogui` | Media key and macro simulation |
| `keyboard` | Global hotkey registration |
| `customtkinter` | Modern themed GUI |
| `spotipy` | Spotify Web API |
| `vosk` | Offline speech recognition |
| `faster-whisper` | Offline speech recognition (Whisper) |
| `Pillow` | Image processing |
| `pystray` | System tray support |

## Troubleshooting

<details>
<summary><strong>Microphone not detected</strong></summary>

- Ensure your microphone is connected and enabled in Windows
- Check Windows privacy settings for microphone access
- Try selecting a different microphone in Settings

</details>

<details>
<summary><strong>Recognition not working well</strong></summary>

- Speak clearly and at a moderate pace
- Adjust the Energy Threshold slider in Settings (lower = more sensitive)
- Try a different recognition engine (Vosk is recommended for offline use)
- Use the Calibrate button in Settings

</details>

<details>
<summary><strong>Keyboard shortcut not working</strong></summary>

- Run the application as administrator
- Check if another application is using the same shortcut

</details>

<details>
<summary><strong>Spotify not connecting</strong></summary>

- Ensure you have valid Spotify API credentials (create an app at [developer.spotify.com](https://developer.spotify.com))
- Set the redirect URI in your Spotify app to `http://127.0.0.1:8888/callback`
- Make sure you have an active Spotify device (desktop app, web player, or mobile app)

</details>

## License

GNU Affero General Public License v3.0 (AGPLv3) - see [LICENSE.txt](LICENSE.txt)
