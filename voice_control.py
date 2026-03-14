"""
Voice Control Platform for Windows
A voice recognition system with system controls and custom commands.
"""

import json
import os
import sys
import threading
import subprocess
import wave
import struct
from pathlib import Path
from typing import Optional, Callable, Dict, Any, List, Tuple
import time

import speech_recognition as sr
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume, EDataFlow, ERole
from comtypes import CLSCTX_ALL, CoInitialize, CoUninitialize

try:
    import vosk
    VOSK_AVAILABLE = True
except ImportError:
    VOSK_AVAILABLE = False

try:
    from faster_whisper import WhisperModel
    WHISPER_AVAILABLE = True
except ImportError:
    WHISPER_AVAILABLE = False
from ctypes import cast, POINTER
import pyautogui
import keyboard

# Import Spotify controller
try:
    from spotify_control import SpotifyController, SPOTIPY_AVAILABLE
except ImportError:
    SPOTIPY_AVAILABLE = False
    SpotifyController = None


# Word-to-number mapping for spoken volume commands
_WORD_NUMBERS = {
    "zero": 0, "one": 1, "two": 2, "three": 3, "four": 4, "five": 5,
    "six": 6, "seven": 7, "eight": 8, "nine": 9, "ten": 10,
    "eleven": 11, "twelve": 12, "thirteen": 13, "fourteen": 14, "fifteen": 15,
    "sixteen": 16, "seventeen": 17, "eighteen": 18, "nineteen": 19, "twenty": 20,
    "thirty": 30, "forty": 40, "fifty": 50, "sixty": 60, "seventy": 70,
    "eighty": 80, "ninety": 90, "hundred": 100,
}

def _parse_spoken_number(text: str) -> Optional[int]:
    """Parse a number from text, supporting both digits and spoken words.
    Handles: '50', 'fifty', 'twenty five', 'twenty-five'."""
    text = text.strip().lower().replace("-", " ")
    if text.isdigit():
        return int(text)
    parts = text.split()
    total = 0
    found = False
    for part in parts:
        if part.isdigit():
            total += int(part)
            found = True
        elif part in _WORD_NUMBERS:
            total += _WORD_NUMBERS[part]
            found = True
    return total if found else None


class VolumeController:
    """Handles system volume control using pycaw."""
    
    def __init__(self, device_id: Optional[str] = None):
        self._device_id = device_id
    
    def _ensure_com_initialized(self) -> None:
        """Ensure COM is initialized for the current thread."""
        try:
            CoInitialize()
        except Exception:
            pass  # Already initialized on this thread
    
    def set_device(self, device_id: Optional[str]) -> None:
        """Set the audio output device by ID."""
        self._device_id = device_id
    
    @staticmethod
    def get_output_devices() -> List[Tuple[str, str]]:
        """Get list of available audio output devices. Returns [(id, name), ...]"""
        devices = []
        try:
            CoInitialize()
        except Exception:
            pass
        
        try:
            all_devices = AudioUtilities.GetAllDevices()
            for device in all_devices:
                try:
                    # Only include output devices (speakers/headphones)
                    if hasattr(device, 'FriendlyName') and device.FriendlyName:
                        # Check if it's a render device (output)
                        if hasattr(device, '_dev'):
                            devices.append((device.id, device.FriendlyName))
                except Exception:
                    continue
        except Exception as e:
            print(f"Error getting output devices: {e}")
        
        return devices
    
    def _get_audio_device(self):
        """Get the audio device object."""
        self._ensure_com_initialized()
        try:
            if self._device_id:
                # Find specific device by ID
                all_devices = AudioUtilities.GetAllDevices()
                for device in all_devices:
                    if device.id == self._device_id:
                        return device
            # Fall back to default speakers
            return AudioUtilities.GetSpeakers()
        except Exception as e:
            print(f"Error getting audio device: {e}")
            return None
    
    def _get_volume_interface(self) -> Optional[IAudioEndpointVolume]:
        """Get the volume interface from the audio device."""
        device = self._get_audio_device()
        if device:
            try:
                # Use the EndpointVolume property instead of Activate
                return device.EndpointVolume
            except Exception as e:
                print(f"Error getting volume interface: {e}")
        return None
    
    def get_volume(self) -> int:
        """Get current volume level (0-100)."""
        interface = self._get_volume_interface()
        if interface:
            try:
                return int(interface.GetMasterVolumeLevelScalar() * 100)
            except Exception as e:
                print(f"Error getting volume: {e}")
        return 0
    
    def set_volume(self, level: int) -> None:
        """Set volume level (0-100)."""
        interface = self._get_volume_interface()
        if interface:
            try:
                level = max(0, min(100, level))
                interface.SetMasterVolumeLevelScalar(level / 100.0, None)
            except Exception as e:
                print(f"Error setting volume: {e}")
    
    def increase_volume(self, step: int = 10) -> None:
        """Increase volume by step amount."""
        current = self.get_volume()
        self.set_volume(current + step)
    
    def decrease_volume(self, step: int = 10) -> None:
        """Decrease volume by step amount."""
        current = self.get_volume()
        self.set_volume(current - step)
    
    def mute(self) -> None:
        """Mute the system volume."""
        interface = self._get_volume_interface()
        if interface:
            try:
                interface.SetMute(1, None)
            except Exception as e:
                print(f"Error muting: {e}")
    
    def unmute(self) -> None:
        """Unmute the system volume."""
        interface = self._get_volume_interface()
        if interface:
            try:
                interface.SetMute(0, None)
            except Exception as e:
                print(f"Error unmuting: {e}")
    
    def toggle_mute(self) -> None:
        """Toggle mute state."""
        interface = self._get_volume_interface()
        if interface:
            try:
                current_mute = interface.GetMute()
                interface.SetMute(not current_mute, None)
            except Exception as e:
                print(f"Error toggling mute: {e}")


class MediaController:
    """Handles media playback controls using keyboard simulation."""
    
    @staticmethod
    def play_pause() -> None:
        """Toggle play/pause."""
        pyautogui.press('playpause')
    
    @staticmethod
    def next_track() -> None:
        """Skip to next track."""
        pyautogui.press('nexttrack')
    
    @staticmethod
    def previous_track() -> None:
        """Go to previous track."""
        pyautogui.press('prevtrack')
    
    @staticmethod
    def stop() -> None:
        """Stop playback."""
        pyautogui.press('stop')


class CommandManager:
    """Manages voice commands and their actions."""
    
    # Default phrases for built-in commands (command_id -> list of phrases)
    DEFAULT_BUILTIN_PHRASES: Dict[str, List[str]] = {
        "set_volume": ["set volume"],
        "volume_up": ["volume up", "increase volume", "louder"],
        "volume_down": ["volume down", "decrease volume", "quieter"],
        "mute": ["mute"],
        "unmute": ["unmute"],
        "toggle_mute": ["toggle mute"],
        "play_pause": ["play", "pause", "play pause"],
        "next_track": ["next track", "skip", "next song"],
        "previous_track": ["previous track", "previous song", "go back"],
        "stop": ["stop", "stop playing"],
        # Timer commands - these trigger natural language parsing
        "set_timer": ["set a timer", "set timer", "timer for", "start a timer", "start timer"],
        "stop_timer": ["stop timer", "stop the timer", "cancel timer", "timer stop"],
    }
    
    # Human-readable names for built-in commands
    BUILTIN_COMMAND_NAMES: Dict[str, str] = {
        "set_volume": "Set Volume",
        "volume_up": "Volume Up",
        "volume_down": "Volume Down",
        "mute": "Mute",
        "unmute": "Unmute",
        "toggle_mute": "Toggle Mute",
        "play_pause": "Play/Pause",
        "next_track": "Next Track",
        "previous_track": "Previous Track",
        "stop": "Stop",
        "set_timer": "Set Timer",
        "stop_timer": "Stop Timer",
    }
    
    # Map command IDs to their categories for enable/disable functionality
    COMMAND_CATEGORIES: Dict[str, str] = {
        "volume_up": "volume",
        "volume_down": "volume",
        "mute": "volume",
        "unmute": "volume",
        "toggle_mute": "volume",
        "set_volume": "volume",
        "play_pause": "media",
        "next_track": "media",
        "previous_track": "media",
        "stop": "media",
        "spotify_play": "spotify",
        "spotify_pause": "spotify",
        "spotify_next": "spotify",
        "spotify_previous": "spotify",
        "spotify_volume_up": "spotify",
        "spotify_volume_down": "spotify",
        "spotify_shuffle": "spotify",
        "spotify_repeat": "spotify",
        "set_timer": "timer",
        "stop_timer": "timer",
    }
    
    @staticmethod
    def _get_config_dir() -> Path:
        """Get the configuration directory.
        Uses AppData on Windows for installed apps to ensure write permissions.
        """
        import sys
        
        # Check if running as a packaged app (frozen)
        if getattr(sys, 'frozen', False):
            # Use AppData/Local for installed apps
            appdata = os.environ.get('LOCALAPPDATA', os.path.expanduser('~'))
            config_dir = Path(appdata) / 'VoiceControl'
            config_dir.mkdir(parents=True, exist_ok=True)
            return config_dir
        else:
            # Development mode - use current directory
            return Path(".")
    
    def __init__(self, config_path: str = "config.json"):
        # Use proper config directory for packaged apps
        config_dir = self._get_config_dir()
        self.config_path = config_dir / "config.json"
        
        self._ensure_config_exists()
        self.commands: Dict[str, str] = {}
        self._audio_device_id = self._load_audio_device_setting()
        self.volume_controller = VolumeController(self._audio_device_id)
        self.media_controller = MediaController()
        
        # Initialize Spotify controller with same config path
        self.spotify: Optional[SpotifyController] = None
        if SpotifyController is not None:
            self.spotify = SpotifyController(str(self.config_path))
        
        # Map command IDs to their action functions
        self._builtin_actions: Dict[str, Callable] = self._setup_builtin_actions()
        
        # Load custom phrases for built-in commands (command_id -> list of phrases)
        self.builtin_phrases: Dict[str, List[str]] = self._load_builtin_phrases()
        
        # Profile management - profiles stored as separate files
        self._profiles_dir = self._get_profiles_dir()
        self.current_profile: str = self._load_current_profile()
        
        # Timer state
        self._active_timer = None
        self._timer_thread = None
        self._timer_stop_event = None
        
        self.load_commands()
    
    # ==================== Profile Management ====================
    
    def _get_profiles_dir(self) -> Path:
        """Get the profiles directory. Creates it if it doesn't exist."""
        profiles_dir = self.config_path.parent / "profiles"
        profiles_dir.mkdir(parents=True, exist_ok=True)
        return profiles_dir
    
    def _get_profile_path(self, profile_name: str) -> Path:
        """Get the file path for a profile."""
        # Sanitize profile name for filename
        safe_name = "".join(c for c in profile_name if c.isalnum() or c in (' ', '-', '_')).strip()
        if not safe_name:
            safe_name = "profile"
        return self._profiles_dir / f"{safe_name}.json"
    
    def get_profiles(self) -> List[str]:
        """Get list of all profile names."""
        profiles = []
        
        # Scan profiles directory for .json files
        if self._profiles_dir.exists():
            for file in self._profiles_dir.glob("*.json"):
                profile_name = file.stem
                if profile_name not in profiles:
                    profiles.append(profile_name)
        
        return sorted(profiles)
    
    def has_profiles(self) -> bool:
        """Check if any profiles exist."""
        return len(self.get_profiles()) > 0
    
    def get_current_profile(self) -> str:
        """Get the name of the current active profile."""
        return self.current_profile
    
    def is_command_enabled(self, cmd_id: str) -> bool:
        """Check if a command is enabled (its category is not disabled)."""
        if not self.current_profile:
            return True
        
        category = self.COMMAND_CATEGORIES.get(cmd_id)
        if not category:
            return True  # Unknown commands are enabled by default
        
        profile_data = self._load_profile_data(self.current_profile)
        disabled_categories = profile_data.get("disabled_categories", [])
        return category not in disabled_categories
    
    def is_category_enabled(self, category_id: str) -> bool:
        """Check if a category is enabled."""
        if not self.current_profile:
            return True
        
        profile_data = self._load_profile_data(self.current_profile)
        disabled_categories = profile_data.get("disabled_categories", [])
        return category_id not in disabled_categories
    
    def _load_current_profile(self) -> str:
        """Load the current profile name from config."""
        config = self._load_full_config()
        saved_profile = config.get("settings", {}).get("current_profile", "")
        
        # Verify the profile exists
        if saved_profile:
            profile_path = self._get_profile_path(saved_profile)
            if profile_path.exists():
                return saved_profile
        
        # If no valid profile, return empty string (no profile selected)
        profiles = self.get_profiles()
        if profiles:
            return profiles[0]  # Use first available profile
        return ""  # No profiles exist
    
    def _save_current_profile_name(self, profile_name: str) -> None:
        """Save the current profile name to config."""
        config = self._load_full_config()
        config.setdefault("settings", {})["current_profile"] = profile_name
        self._save_config(config)
    
    def _load_profile_data(self, profile_name: str) -> Dict[str, Any]:
        """Load profile data from file."""
        if not profile_name:
            # No profile selected
            return {"commands": {}, "builtin_phrases": {}}
        
        profile_path = self._get_profile_path(profile_name)
        if profile_path.exists():
            try:
                with open(profile_path, 'r', encoding='utf-8') as f:
                    return json.load(f)
            except (json.JSONDecodeError, IOError):
                pass
        
        # Return empty profile if file doesn't exist
        return {"commands": {}, "builtin_phrases": {}}
    
    def _save_profile_data(self, profile_name: str, data: Dict[str, Any]) -> bool:
        """Save profile data to file."""
        if not profile_name:
            return False
        
        profile_path = self._get_profile_path(profile_name)
        try:
            with open(profile_path, 'w', encoding='utf-8') as f:
                json.dump(data, f, indent=4)
            return True
        except IOError:
            return False
    
    def switch_profile(self, profile_name: str) -> bool:
        """
        Switch to a different profile.
        Returns True if successful.
        """
        if not profile_name:
            return False
        
        # Check if profile exists
        profile_path = self._get_profile_path(profile_name)
        if not profile_path.exists():
            return False
        
        # Save current profile before switching
        if self.current_profile:
            self._save_current_profile()
        
        # Switch to new profile
        self.current_profile = profile_name
        self._save_current_profile_name(profile_name)
        
        # Load the new profile's data
        profile_data = self._load_profile_data(profile_name)
        self.commands = profile_data.get("commands", {})
        self.builtin_phrases = self._load_builtin_phrases()
        
        # Override with profile-specific builtin phrases if any
        profile_phrases = profile_data.get("builtin_phrases", {})
        for cmd_id, phrases in profile_phrases.items():
            if cmd_id in self.builtin_phrases:
                self.builtin_phrases[cmd_id] = phrases
        
        return True
    
    def _save_current_profile(self) -> None:
        """Save the current commands and phrases to the current profile."""
        if not self.current_profile:
            return
        
        # Load existing profile to preserve disabled_categories
        existing_data = self._load_profile_data(self.current_profile)
        
        data = {
            "commands": self.commands.copy(),
            "builtin_phrases": self.builtin_phrases.copy(),
            "disabled_categories": existing_data.get("disabled_categories", [])
        }
        self._save_profile_data(self.current_profile, data)
    
    def set_disabled_categories(self, disabled: List[str]) -> None:
        """Set the disabled categories for the current profile."""
        if not self.current_profile:
            return
        
        existing_data = self._load_profile_data(self.current_profile)
        existing_data["disabled_categories"] = disabled
        existing_data["commands"] = self.commands.copy()
        existing_data["builtin_phrases"] = self.builtin_phrases.copy()
        self._save_profile_data(self.current_profile, existing_data)
    
    def get_disabled_categories(self) -> List[str]:
        """Get the disabled categories for the current profile."""
        if not self.current_profile:
            return []
        
        profile_data = self._load_profile_data(self.current_profile)
        return profile_data.get("disabled_categories", [])
    
    def get_timer_settings(self) -> Dict[str, bool]:
        """Get timer settings for the current profile."""
        if not self.current_profile:
            return {"alarm_sound": True, "confirm_sound": True}
        
        profile_data = self._load_profile_data(self.current_profile)
        return profile_data.get("timer_settings", {"alarm_sound": True, "confirm_sound": True})
    
    def set_timer_settings(self, settings: Dict[str, bool]) -> None:
        """Set timer settings for the current profile."""
        if not self.current_profile:
            return
        
        existing_data = self._load_profile_data(self.current_profile)
        existing_data["timer_settings"] = settings
        existing_data["commands"] = self.commands.copy()
        existing_data["builtin_phrases"] = self.builtin_phrases.copy()
        self._save_profile_data(self.current_profile, existing_data)
    
    def create_profile(self, profile_name: str) -> bool:
        """
        Create a new empty profile.
        Returns True if successful.
        """
        if not profile_name or profile_name.strip() == "":
            return False
        
        profile_name = profile_name.strip()
        
        # Check if profile already exists
        if self._get_profile_path(profile_name).exists():
            return False
        
        # Create new empty profile (only built-in commands, no custom commands)
        data = {
            "commands": {},
            "builtin_phrases": {}  # Will use defaults
        }
        
        return self._save_profile_data(profile_name, data)
    
    def delete_profile(self, profile_name: str) -> bool:
        """
        Delete a profile.
        Returns True if successful.
        """
        if not profile_name:
            return False
        
        profile_path = self._get_profile_path(profile_name)
        if not profile_path.exists():
            return False
        
        # If deleting current profile, switch to another or clear
        if self.current_profile == profile_name:
            profiles = self.get_profiles()
            other_profiles = [p for p in profiles if p != profile_name]
            if other_profiles:
                self.switch_profile(other_profiles[0])
            else:
                self.current_profile = ""
                self.commands = {}
                self._save_current_profile_name("")
        
        try:
            profile_path.unlink()
            return True
        except IOError:
            return False
    
    def rename_profile(self, old_name: str, new_name: str) -> bool:
        """
        Rename a profile.
        Returns True if successful.
        """
        if not old_name or not new_name or new_name.strip() == "":
            return False
        
        new_name = new_name.strip()
        old_path = self._get_profile_path(old_name)
        new_path = self._get_profile_path(new_name)
        
        if not old_path.exists() or new_path.exists():
            return False
        
        try:
            old_path.rename(new_path)
            
            # Update current profile if it was renamed
            if self.current_profile == old_name:
                self.current_profile = new_name
                self._save_current_profile_name(new_name)
            
            return True
        except IOError:
            return False
    
    def export_profile(self, profile_name: str, export_path: str) -> bool:
        """
        Export a profile to a specified path.
        Returns True if successful.
        """
        profile_data = self._load_profile_data(profile_name)
        try:
            with open(export_path, 'w', encoding='utf-8') as f:
                json.dump(profile_data, f, indent=4)
            return True
        except IOError:
            return False
    
    def import_profile(self, import_path: str, profile_name: str = None) -> Tuple[bool, str]:
        """
        Import a profile from a file.
        Returns (success, message/profile_name).
        """
        try:
            with open(import_path, 'r', encoding='utf-8') as f:
                data = json.load(f)
            
            # Validate structure
            if not isinstance(data, dict):
                return False, "Invalid profile file format"
            
            # Use filename as profile name if not specified
            if not profile_name:
                profile_name = Path(import_path).stem
            
            # Ensure unique name
            base_name = profile_name
            counter = 1
            while profile_name == "Default" or self._get_profile_path(profile_name).exists():
                profile_name = f"{base_name} ({counter})"
                counter += 1
            
            # Ensure proper structure
            profile_data = {
                "commands": data.get("commands", {}),
                "builtin_phrases": data.get("builtin_phrases", {})
            }
            
            if self._save_profile_data(profile_name, profile_data):
                return True, profile_name
            return False, "Failed to save profile"
            
        except json.JSONDecodeError:
            return False, "Invalid JSON file"
        except IOError as e:
            return False, f"Error reading file: {e}"
    
    # ==================== End Profile Management ====================

    def _ensure_config_exists(self) -> None:
        """Ensure config.json exists. If not, copy from config.default.json or create empty."""
        if self.config_path.exists():
            return

        import sys
        
        # Try to find and copy from default config
        default_paths = [
            self.config_path.parent / "config.default.json",
            Path(__file__).parent / "config.default.json",
            Path("config.default.json"),
        ]

        # In PyInstaller one-file builds, bundled data is extracted under _MEIPASS.
        meipass = getattr(sys, "_MEIPASS", None)
        if meipass:
            default_paths.insert(0, Path(meipass) / "config.default.json")
        
        for default_path in default_paths:
            if default_path.exists():
                try:
                    import shutil
                    shutil.copy(default_path, self.config_path)
                    return
                except IOError:
                    pass
        
        # Create empty config if no default found
        default_config = {
            "commands": {},
            "settings": {
                "toggle_shortcut": "ctrl+shift+v",
                "energy_threshold": 4000
            },
            "builtin_phrases": {}
        }
        try:
            with open(self.config_path, 'w') as f:
                json.dump(default_config, f, indent=4)
        except IOError:
            pass  # Will be created on first save
    
    def _setup_builtin_actions(self) -> Dict[str, Callable]:
        """Set up built-in command actions mapped by command ID."""
        return {
            "set_volume": self._handle_set_volume,
            "volume_up": lambda: self.volume_controller.increase_volume(),
            "volume_down": lambda: self.volume_controller.decrease_volume(),
            "mute": lambda: self.volume_controller.mute(),
            "unmute": lambda: self.volume_controller.unmute(),
            "toggle_mute": lambda: self.volume_controller.toggle_mute(),
            "play_pause": lambda: self.media_controller.play_pause(),
            "next_track": lambda: self.media_controller.next_track(),
            "previous_track": lambda: self.media_controller.previous_track(),
            "stop": lambda: self.media_controller.stop(),
        }
    
    def _load_builtin_phrases(self) -> Dict[str, List[str]]:
        """Load custom phrases for built-in commands from config."""
        phrases = {}
        # Start with defaults
        for cmd_id, default_phrases in self.DEFAULT_BUILTIN_PHRASES.items():
            phrases[cmd_id] = default_phrases.copy()
        
        # Override with custom phrases from config
        if self.config_path.exists():
            try:
                with open(self.config_path, 'r') as f:
                    config = json.load(f)
                    custom_phrases = config.get("builtin_phrases", {})
                    for cmd_id, phrase_list in custom_phrases.items():
                        if cmd_id in phrases and isinstance(phrase_list, list):
                            phrases[cmd_id] = [p.lower().strip() for p in phrase_list if p.strip()]
            except (json.JSONDecodeError, IOError):
                pass
        
        return phrases
    
    def save_builtin_phrases(self) -> None:
        """Save custom phrases for built-in commands to config."""
        config = self._load_full_config()
        config["builtin_phrases"] = self.builtin_phrases
        self._save_config(config)
    
    def get_builtin_phrases(self, command_id: str) -> List[str]:
        """Get the phrases for a built-in command."""
        return self.builtin_phrases.get(command_id, [])
    
    def set_builtin_phrases(self, command_id: str, phrases: List[str]) -> None:
        """Set the phrases for a built-in command."""
        if command_id in self._builtin_actions:
            # Clean and validate phrases
            cleaned = [p.lower().strip() for p in phrases if p.strip()]
            if cleaned:
                self.builtin_phrases[command_id] = cleaned
                self.save_builtin_phrases()
    
    def reset_builtin_phrases(self, command_id: str) -> None:
        """Reset a built-in command's phrases to defaults."""
        if command_id in self.DEFAULT_BUILTIN_PHRASES:
            self.builtin_phrases[command_id] = self.DEFAULT_BUILTIN_PHRASES[command_id].copy()
            self.save_builtin_phrases()
    
    def _load_audio_device_setting(self) -> Optional[str]:
        """Load the audio device ID from config."""
        if self.config_path.exists():
            try:
                with open(self.config_path, 'r') as f:
                    config = json.load(f)
                    return config.get("settings", {}).get("audio_device_id")
            except (json.JSONDecodeError, IOError):
                pass
        return None
    
    def set_audio_device(self, device_id: Optional[str]) -> None:
        """Set the audio output device."""
        self._audio_device_id = device_id
        self.volume_controller.set_device(device_id)
        self._save_audio_device_setting(device_id)
    
    def _save_audio_device_setting(self, device_id: Optional[str]) -> None:
        """Save the audio device ID to config."""
        config = {"commands": {}, "settings": {}}
        if self.config_path.exists():
            try:
                with open(self.config_path, 'r') as f:
                    config = json.load(f)
            except (json.JSONDecodeError, IOError):
                pass
        
        config.setdefault("settings", {})
        config["settings"]["audio_device_id"] = device_id
        
        try:
            with open(self.config_path, 'w') as f:
                json.dump(config, f, indent=4)
        except IOError as e:
            print(f"Error saving audio device setting: {e}")
    
    def _handle_set_volume(self, level: Optional[int] = None) -> None:
        """Handle set volume command with level parameter."""
        if level is not None:
            self.volume_controller.set_volume(level)
    
    def load_commands(self) -> None:
        """Load custom commands from current profile."""
        profile_data = self._load_profile_data(self.current_profile)
        self.commands = profile_data.get("commands", {})
    
    def save_commands(self) -> None:
        """Save custom commands to current profile."""
        self._save_current_profile()
    
    def _load_full_config(self) -> Dict[str, Any]:
        """Load the full configuration file."""
        if self.config_path.exists():
            try:
                with open(self.config_path, 'r') as f:
                    return json.load(f)
            except (json.JSONDecodeError, IOError):
                pass
        return {"commands": {}, "settings": {}}
    
    def _save_config(self, config: Dict[str, Any]) -> None:
        """Save configuration to file."""
        try:
            with open(self.config_path, 'w') as f:
                json.dump(config, f, indent=4)
        except IOError as e:
            print(f"Error saving config: {e}")
    
    def validate_phrases(self, phrases: List[str]) -> tuple[bool, str]:
        """
        Validate that phrases don't conflict with reserved commands.
        Returns (is_valid, error_message) tuple.
        """
        # Check for Spotify conflicts if Spotify is connected
        if self.spotify and self.spotify.is_authenticated:
            if SpotifyController is not None:
                for phrase in phrases:
                    phrase = phrase.lower().strip()
                    if not phrase:
                        continue
                    is_reserved, reason = SpotifyController.is_phrase_reserved(phrase)
                    if is_reserved:
                        return False, f"Cannot use this phrase: {reason}"
        
        # Check for conflicts with built-in commands
        for phrase in phrases:
            phrase = phrase.lower().strip()
            if not phrase:
                continue
            for cmd_id, builtin_phrases in self.builtin_phrases.items():
                if phrase in builtin_phrases:
                    cmd_name = self.BUILTIN_COMMAND_NAMES.get(cmd_id, cmd_id)
                    return False, f"'{phrase}' is already used by built-in command '{cmd_name}'"
        
        return True, ""
    
    def add_command(self, phrases: List[str], action_type: str, action_data: Any) -> str:
        """
        Add a custom command.
        
        Args:
            phrases: List of voice phrases that trigger this command
            action_type: Type of action ("open_file", "open_url", or "macro")
            action_data: For "open_file": file path string
                        For "open_url": URL string
                        For "macro": list of {"key": str, "delay": float} dicts
        
        Returns:
            The command ID
        """
        import uuid
        command_id = str(uuid.uuid4())[:8]
        
        self.commands[command_id] = {
            "phrases": [p.lower().strip() for p in phrases if p.strip()],
            "type": action_type,
            "data": action_data
        }
        self.save_commands()
        return command_id
    
    def update_command(self, command_id: str, phrases: List[str] = None, 
                       action_type: str = None, action_data: Any = None) -> bool:
        """Update an existing custom command."""
        if command_id not in self.commands:
            return False
        
        if phrases is not None:
            self.commands[command_id]["phrases"] = [p.lower().strip() for p in phrases if p.strip()]
        if action_type is not None:
            self.commands[command_id]["type"] = action_type
        if action_data is not None:
            self.commands[command_id]["data"] = action_data
        
        self.save_commands()
        return True
    
    def remove_command(self, command_id: str) -> bool:
        """Remove a custom command by ID."""
        if command_id in self.commands:
            del self.commands[command_id]
            self.save_commands()
            return True
        return False
    
    def get_command(self, command_id: str) -> Optional[Dict[str, Any]]:
        """Get a custom command by ID."""
        return self.commands.get(command_id)
    
    def _execute_macro(self, macro_steps: List[Dict[str, Any]]) -> None:
        """Execute a macro (sequence of key presses, mouse actions, and delays)."""
        for step in macro_steps:
            step_type = step.get("type", "key")  # Default to "key" for backwards compatibility
            delay = step.get("delay", 0.1)
            
            if step_type == "key":
                # Keyboard action (legacy format support)
                key = step.get("key", "")
                if key:
                    keyboard.send(key)
            
            elif step_type == "mouse_move":
                # Move mouse to position
                x = step.get("x", 0)
                y = step.get("y", 0)
                duration = step.get("duration", 0.1)
                pyautogui.moveTo(x, y, duration=duration)
            
            elif step_type == "mouse_click":
                # Click at position (or current position if no coords)
                x = step.get("x")
                y = step.get("y")
                button = step.get("button", "left")  # left, right, middle
                clicks = step.get("clicks", 1)  # 1 = single, 2 = double
                if x is not None and y is not None:
                    pyautogui.click(x=x, y=y, button=button, clicks=clicks)
                else:
                    pyautogui.click(button=button, clicks=clicks)
            
            elif step_type == "mouse_scroll":
                # Scroll at current position or specific position
                x = step.get("x")
                y = step.get("y")
                amount = step.get("amount", 3)  # Positive = up, negative = down
                if x is not None and y is not None:
                    pyautogui.scroll(amount, x=x, y=y)
                else:
                    pyautogui.scroll(amount)
            
            if delay > 0:
                time.sleep(delay)
    
    def _handle_spotify_command(self, text: str) -> Optional[tuple[bool, str]]:
        """
        Handle Spotify-specific voice commands.
        Returns (success, message) if a Spotify command was recognized, None otherwise.
        """
        if not self.spotify or not self.spotify.is_authenticated:
            return None
        
        # Play song: "play [song name]" or "play song [song name]"
        play_prefixes = ["play song ", "play track "]
        for prefix in play_prefixes:
            if text.startswith(prefix):
                query = text[len(prefix):].strip()
                if query:
                    return self.spotify.play_song(query)
        
        # Play artist: "play artist [name]"
        if text.startswith("play artist "):
            query = text[len("play artist "):].strip()
            if query:
                return self.spotify.play_artist(query)
        
        # Play album: "play album [name]"
        if text.startswith("play album "):
            query = text[len("play album "):].strip()
            if query:
                return self.spotify.play_album(query)
        
        # Play my playlist: "play my playlist [name]" - plays user's own playlist
        my_playlist_prefixes = ["play my playlist ", "play my ", "open my playlist "]
        for prefix in my_playlist_prefixes:
            if text.startswith(prefix):
                query = text[len(prefix):].strip()
                if query:
                    return self.spotify.play_my_playlist(query)
        
        # Play playlist: "play playlist [name]" - searches all playlists
        if text.startswith("play playlist "):
            query = text[len("play playlist "):].strip()
            if query:
                return self.spotify.play_playlist(query)
        
        # Play recommendations: "play recommendations", "play something similar"
        if text in ["play recommendations", "play something similar", "recommend something", 
                    "play similar songs", "play similar", "recommendations"]:
            return self.spotify.play_recommendations()
        
        # Play radio: "play radio", "start radio"
        if text in ["play radio", "start radio", "radio", "create radio", "song radio"]:
            return self.spotify.play_radio()
        
        # Add to playlist: "add to [playlist name]", "add this to [playlist name]"
        add_prefixes = ["add to playlist ", "add this to playlist ", "add to ", "add this to ", 
                        "save to ", "save this to "]
        for prefix in add_prefixes:
            if text.startswith(prefix):
                playlist_name = text[len(prefix):].strip()
                if playlist_name:
                    return self.spotify.add_to_playlist(playlist_name)
        
        # Shuffle commands
        if text in ["shuffle on", "enable shuffle", "turn on shuffle"]:
            return self.spotify.shuffle(True)
        if text in ["shuffle off", "disable shuffle", "turn off shuffle"]:
            return self.spotify.shuffle(False)
        
        # Repeat commands
        if text in ["repeat on", "repeat track", "repeat this song", "repeat song"]:
            return self.spotify.repeat("track")
        if text in ["repeat all", "repeat playlist", "repeat album"]:
            return self.spotify.repeat("context")
        if text in ["repeat off", "no repeat", "stop repeat"]:
            return self.spotify.repeat("off")
        
        # Spotify volume commands: "spotify volume [number]", "set spotify volume [number]"
        spotify_volume_prefixes = ["spotify volume ", "set spotify volume ", "spotify volume to "]
        for prefix in spotify_volume_prefixes:
            if text.startswith(prefix):
                # Extract the number
                volume_str = text[len(prefix):].strip()
                # Handle "percent" suffix
                volume_str = volume_str.replace("percent", "").replace("%", "").strip()
                volume = _parse_spoken_number(volume_str)
                if volume is not None:
                    volume = max(0, min(100, volume))
                    return self.spotify.set_volume(volume)
        
        # What's playing
        if text in ["what's playing", "whats playing", "what is playing", "current song", 
                    "what song is this", "now playing"]:
            track = self.spotify.get_current_track()
            if track:
                status = "Playing" if track["is_playing"] else "Paused"
                return True, f"{status}: '{track['name']}' by {track['artist']}"
            return False, "Nothing is currently playing"
        
        # Generic "play [something]" - assume it's a song (check last to avoid matching other commands)
        if text.startswith("play ") and not text.startswith("play pause"):
            query = text[len("play "):].strip()
            # Avoid matching single words that might be other commands
            skip_words = ["pause", "music", "next", "previous", "recommendations", 
                          "radio", "similar"]
            if query and len(query) > 2 and query not in skip_words:
                return self.spotify.play_song(query)
        
        return None
    
    def execute_command(self, spoken_text: str) -> tuple[bool, str]:
        """
        Execute a command based on spoken text.
        Returns (success, message) tuple.
        """
        text = spoken_text.lower().strip()
        
        # Check for timer commands (if timer category is enabled)
        if self.is_category_enabled("timer"):
            # Check for stop timer
            if self.is_command_enabled("stop_timer"):
                for phrase in self.builtin_phrases.get("stop_timer", ["stop timer"]):
                    if phrase in text:
                        if self.stop_timer():
                            return True, "Timer stopped"
                        return False, "No active timer"
            
            # Check for set timer with natural language parsing
            if self.is_command_enabled("set_timer"):
                timer_result = self._parse_and_set_timer(text)
                if timer_result:
                    return timer_result
        
        # Check for Spotify commands first (if connected and category enabled)
        if self.spotify and self.spotify.is_authenticated and self.is_category_enabled("spotify"):
            spotify_result = self._handle_spotify_command(text)
            if spotify_result:
                return spotify_result
        
        # Check for set volume with number (special handling for numeric parameter)
        if self.is_command_enabled("set_volume"):
            for phrase in self.builtin_phrases.get("set_volume", ["set volume"]):
                if phrase in text:
                    # Extract number from text (after the matched phrase)
                    remainder = text[text.index(phrase) + len(phrase):].strip()
                    level = _parse_spoken_number(remainder)
                    if level is None:
                        # Try finding a number anywhere in the text
                        words = text.split()
                        for word in words:
                            level = _parse_spoken_number(word)
                            if level is not None:
                                break
                    if level is not None:
                        level = max(0, min(100, level))
                        self._handle_set_volume(level)
                        return True, f"Volume set to {level}%"
                    return False, "Please specify a volume level (0-100)"
        
        # Check built-in commands - find the longest matching phrase first
        # This prevents "play" from matching when "play pause" is said
        matched_cmd_id = None
        matched_phrase = None
        matched_len = 0
        
        for cmd_id, phrases in self.builtin_phrases.items():
            if cmd_id == "set_volume":
                continue  # Already handled above
            # Skip disabled commands
            if not self.is_command_enabled(cmd_id):
                continue
            for phrase in phrases:
                if phrase in text and len(phrase) > matched_len:
                    matched_cmd_id = cmd_id
                    matched_phrase = phrase
                    matched_len = len(phrase)
        
        if matched_cmd_id and matched_cmd_id in self._builtin_actions:
            self._builtin_actions[matched_cmd_id]()
            return True, f"Executed: {matched_phrase}"
        
        # Check custom commands - find the longest matching phrase
        matched_custom_id = None
        matched_custom_phrase = None
        matched_custom_len = 0
        
        for cmd_id, cmd_data in self.commands.items():
            # Handle both old format (string) and new format (dict)
            if isinstance(cmd_data, str):
                # Old format: phrase -> file_path
                if cmd_id in text and len(cmd_id) > matched_custom_len:
                    matched_custom_id = cmd_id
                    matched_custom_phrase = cmd_id
                    matched_custom_len = len(cmd_id)
            else:
                # New format: id -> {phrases, type, data}
                phrases = cmd_data.get("phrases", [])
                for phrase in phrases:
                    if phrase in text and len(phrase) > matched_custom_len:
                        matched_custom_id = cmd_id
                        matched_custom_phrase = phrase
                        matched_custom_len = len(phrase)
        
        if matched_custom_id:
            cmd_data = self.commands[matched_custom_id]
            
            # Handle old format
            if isinstance(cmd_data, str):
                try:
                    os.startfile(cmd_data)
                    return True, f"Opened: {matched_custom_phrase}"
                except Exception as e:
                    return False, f"Error opening: {e}"
            
            # Handle new format
            action_type = cmd_data.get("type", "open_file")
            action_data = cmd_data.get("data")
            
            try:
                if action_type == "open_file":
                    os.startfile(action_data)
                    return True, f"Opened: {matched_custom_phrase}"
                elif action_type == "open_url":
                    import webbrowser
                    webbrowser.open(action_data)
                    return True, f"Opened website: {matched_custom_phrase}"
                elif action_type == "macro":
                    self._execute_macro(action_data)
                    return True, f"Executed macro: {matched_custom_phrase}"
                elif action_type == "window_action":
                    self._execute_window_action(action_data)
                    return True, f"Window action: {action_data}"
                elif action_type == "chain":
                    return self._execute_chain(action_data, matched_custom_phrase)
                elif action_type == "timer":
                    return self._start_timer(action_data, matched_custom_phrase)
                else:
                    return False, f"Unknown action type: {action_type}"
            except Exception as e:
                return False, f"Error executing {matched_custom_phrase}: {e}"
        
        return False, f"Unknown command: {text}"
    
    def _execute_window_action(self, action_data) -> None:
        """Execute a window action. 
        action_data can be:
        - A string (legacy format): action name, operates on focused window
        - A dict: {"action": str, "target": "focused"|"app"|"title", "app": str, "title": str}
        """
        import ctypes
        from ctypes import wintypes
        
        user32 = ctypes.windll.user32
        
        # Handle legacy string format
        if isinstance(action_data, str):
            action = action_data
            target_type = "focused"
            target_app = ""
            target_title = ""
        else:
            action = action_data.get("action", "minimize")
            target_type = action_data.get("target", "focused")
            target_app = action_data.get("app", "")
            target_title = action_data.get("title", "")
        
        # Get the target window handle
        hwnd = self._get_target_window(user32, target_type, target_app, target_title)
        
        if not hwnd:
            print(f"No matching window found for target: {target_type}, app={target_app}, title={target_title}")
            return
        
        SW_MINIMIZE = 6
        SW_MAXIMIZE = 3
        SW_RESTORE = 9
        
        # For snap actions, we need to focus the window first
        if action.startswith("snap_"):
            # Bring window to foreground first
            user32.SetForegroundWindow(hwnd)
            user32.ShowWindow(hwnd, SW_RESTORE)  # Restore if minimized
            time.sleep(0.1)  # Small delay to ensure window is focused
        
        if action == "minimize":
            user32.ShowWindow(hwnd, SW_MINIMIZE)
        elif action == "maximize":
            user32.ShowWindow(hwnd, SW_MAXIMIZE)
        elif action == "restore":
            user32.ShowWindow(hwnd, SW_RESTORE)
        elif action == "focus":
            user32.ShowWindow(hwnd, SW_RESTORE)
            user32.SetForegroundWindow(hwnd)
        elif action == "close":
            WM_CLOSE = 0x0010
            user32.PostMessageW(hwnd, WM_CLOSE, 0, 0)
        elif action == "close_all_app":
            # Close all windows of the specified application
            self._close_all_app_windows(hwnd, user32)
        elif action == "close_all_windows":
            # Close ALL windows except Voice Control
            self._close_all_windows_except_self(user32)
        elif action == "snap_left":
            pyautogui.hotkey('win', 'left')
        elif action == "snap_right":
            pyautogui.hotkey('win', 'right')
        elif action == "snap_top_left":
            pyautogui.hotkey('win', 'left')
            time.sleep(0.1)
            pyautogui.hotkey('win', 'up')
        elif action == "snap_top_right":
            pyautogui.hotkey('win', 'right')
            time.sleep(0.1)
            pyautogui.hotkey('win', 'up')
        elif action == "snap_bottom_left":
            pyautogui.hotkey('win', 'left')
            time.sleep(0.1)
            pyautogui.hotkey('win', 'down')
        elif action == "snap_bottom_right":
            pyautogui.hotkey('win', 'right')
            time.sleep(0.1)
            pyautogui.hotkey('win', 'down')
    
    def _get_target_window(self, user32, target_type: str, target_app: str, target_title: str):
        """Find a window handle based on target criteria."""
        import ctypes
        from ctypes import wintypes
        
        if target_type == "focused":
            return user32.GetForegroundWindow()
        
        our_pid = os.getpid()
        found_hwnd = [None]  # Use list to allow modification in callback
        
        def enum_callback(hwnd, lparam):
            if user32.IsWindowVisible(hwnd):
                length = user32.GetWindowTextLengthW(hwnd)
                if length > 0:
                    # Get window title
                    buff = ctypes.create_unicode_buffer(length + 1)
                    user32.GetWindowTextW(hwnd, buff, length + 1)
                    title = buff.value
                    
                    # Get process ID
                    window_pid = wintypes.DWORD()
                    user32.GetWindowThreadProcessId(hwnd, ctypes.byref(window_pid))
                    
                    # Skip our own app
                    if window_pid.value == our_pid:
                        return True
                    
                    if target_type == "title" and target_title:
                        # Match by window title (case-insensitive, partial match)
                        if target_title.lower() in title.lower():
                            found_hwnd[0] = hwnd
                            return False  # Stop enumeration
                    
                    elif target_type == "app" and target_app:
                        # Match by process name
                        try:
                            import psutil
                            proc = psutil.Process(window_pid.value)
                            proc_name = proc.name().lower().replace('.exe', '')
                            if target_app.lower() in proc_name:
                                found_hwnd[0] = hwnd
                                return False  # Stop enumeration
                        except:
                            pass
            return True
        
        WNDENUMPROC = ctypes.WINFUNCTYPE(ctypes.c_bool, wintypes.HWND, wintypes.LPARAM)
        callback = WNDENUMPROC(enum_callback)
        user32.EnumWindows(callback, 0)
        
        return found_hwnd[0]
    
    @staticmethod
    def get_open_windows() -> List[Dict[str, Any]]:
        """Get list of all open windows with their titles and process names."""
        import ctypes
        from ctypes import wintypes
        
        user32 = ctypes.windll.user32
        our_pid = os.getpid()
        windows = []
        
        def enum_callback(hwnd, lparam):
            if user32.IsWindowVisible(hwnd):
                length = user32.GetWindowTextLengthW(hwnd)
                if length > 0:
                    # Get window title
                    buff = ctypes.create_unicode_buffer(length + 1)
                    user32.GetWindowTextW(hwnd, buff, length + 1)
                    title = buff.value
                    
                    # Get process ID
                    window_pid = wintypes.DWORD()
                    user32.GetWindowThreadProcessId(hwnd, ctypes.byref(window_pid))
                    
                    if window_pid.value != our_pid and title not in ["Program Manager", ""]:
                        try:
                            import psutil
                            proc = psutil.Process(window_pid.value)
                            proc_name = proc.name().replace('.exe', '')
                            windows.append({
                                "hwnd": hwnd,
                                "title": title,
                                "process": proc_name
                            })
                        except:
                            windows.append({
                                "hwnd": hwnd,
                                "title": title,
                                "process": "Unknown"
                            })
            return True
        
        WNDENUMPROC = ctypes.WINFUNCTYPE(ctypes.c_bool, wintypes.HWND, wintypes.LPARAM)
        callback = WNDENUMPROC(enum_callback)
        user32.EnumWindows(callback, 0)
        
        return windows
    
    def _close_all_windows_except_self(self, user32) -> None:
        """Close ALL visible windows except Voice Control."""
        import ctypes
        from ctypes import wintypes
        
        # Get our own process ID to exclude Voice Control windows
        our_pid = os.getpid()
        
        # Callback function to enumerate windows
        windows_to_close = []
        
        def enum_callback(hwnd, lparam):
            # Check if window is visible
            if user32.IsWindowVisible(hwnd):
                # Get window title to filter out system windows
                length = user32.GetWindowTextLengthW(hwnd)
                if length > 0:
                    # Get process ID of this window
                    window_pid = wintypes.DWORD()
                    user32.GetWindowThreadProcessId(hwnd, ctypes.byref(window_pid))
                    
                    # Skip our own process (Voice Control)
                    if window_pid.value != our_pid:
                        # Get window title to check for system windows to skip
                        buff = ctypes.create_unicode_buffer(length + 1)
                        user32.GetWindowTextW(hwnd, buff, length + 1)
                        title = buff.value
                        
                        # Skip certain system windows
                        skip_titles = ["Program Manager", "Windows Shell Experience Host", 
                                      "Windows Input Experience", "Microsoft Text Input Application"]
                        if title and title not in skip_titles:
                            windows_to_close.append(hwnd)
            return True
        
        # Define the callback type
        WNDENUMPROC = ctypes.WINFUNCTYPE(ctypes.c_bool, wintypes.HWND, wintypes.LPARAM)
        callback = WNDENUMPROC(enum_callback)
        
        # Enumerate all windows
        user32.EnumWindows(callback, 0)
        
        # Close all found windows
        WM_CLOSE = 0x0010
        for window in windows_to_close:
            user32.PostMessageW(window, WM_CLOSE, 0, 0)
    
    def _close_all_app_windows(self, hwnd, user32) -> None:
        """Close all windows belonging to the same application as the given window.
        Excludes Voice Control app windows to prevent closing itself.
        """
        import ctypes
        from ctypes import wintypes
        
        # Get the process ID of the current window
        pid = wintypes.DWORD()
        user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
        target_pid = pid.value
        
        # Get our own process ID to exclude Voice Control windows
        our_pid = os.getpid()
        
        # Don't close windows if the target is the Voice Control app itself
        if target_pid == our_pid:
            return
        
        # Callback function to enumerate windows
        windows_to_close = []
        
        def enum_callback(hwnd, lparam):
            # Check if window is visible
            if user32.IsWindowVisible(hwnd):
                # Get process ID of this window
                window_pid = wintypes.DWORD()
                user32.GetWindowThreadProcessId(hwnd, ctypes.byref(window_pid))
                # Only add windows from target app, never our own app
                if window_pid.value == target_pid and window_pid.value != our_pid:
                    windows_to_close.append(hwnd)
            return True
        
        # Define the callback type
        WNDENUMPROC = ctypes.WINFUNCTYPE(ctypes.c_bool, wintypes.HWND, wintypes.LPARAM)
        callback = WNDENUMPROC(enum_callback)
        
        # Enumerate all windows
        user32.EnumWindows(callback, 0)
        
        # Close all found windows
        WM_CLOSE = 0x0010
        for window in windows_to_close:
            user32.PostMessageW(window, WM_CLOSE, 0, 0)
    
    def _execute_chain(self, chain_steps: List, phrase: str) -> tuple[bool, str]:
        """Execute a chain of actions in sequence.
        
        chain_steps can be:
        - List of dicts with inline action data: [{"type": str, "data": any, "display": str}, ...]
        - List of command IDs (legacy format): ["cmd_id1", "cmd_id2", ...]
        """
        executed = 0
        
        for step in chain_steps:
            try:
                # Determine action type and data
                if isinstance(step, dict) and "type" in step:
                    # New inline format
                    action_type = step.get("type")
                    action_data = step.get("data")
                elif isinstance(step, str):
                    # Legacy format: command ID
                    cmd_data = self.commands.get(step)
                    if not cmd_data:
                        continue
                    action_type = cmd_data.get("type", "open_file")
                    action_data = cmd_data.get("data")
                else:
                    continue
                
                # Execute the action
                if action_type == "open_file":
                    os.startfile(action_data)
                elif action_type == "open_url":
                    import webbrowser
                    webbrowser.open(action_data)
                elif action_type == "macro":
                    self._execute_macro(action_data)
                elif action_type == "window_action":
                    self._execute_window_action(action_data)
                elif action_type == "wait":
                    # Wait/delay action
                    seconds = action_data.get("seconds", 0.5) if isinstance(action_data, dict) else 0.5
                    time.sleep(seconds)
                
                executed += 1
                
                # Small delay between non-wait actions
                if action_type != "wait":
                    time.sleep(0.3)
                    
            except Exception as e:
                print(f"Chain step error: {e}")
                continue
        
        return True, f"Executed chain '{phrase}' ({executed} actions)"
    
    def _parse_and_set_timer(self, text: str) -> Optional[tuple[bool, str]]:
        """Parse natural language timer command and start the timer.
        
        Handles phrases like:
        - "set a timer for 5 minutes"
        - "set timer for 30 seconds"
        - "timer for 2 minutes and 30 seconds"
        - "start a timer for 10 minutes"
        - "set a timer for one hour"
        """
        import re
        
        # Check if this is a timer command
        timer_phrases = self.builtin_phrases.get("set_timer", ["set a timer", "set timer", "timer for"])
        is_timer_command = any(phrase in text for phrase in timer_phrases)
        
        if not is_timer_command:
            return None
        
        # Word to number mapping
        word_to_num = {
            "zero": 0, "one": 1, "two": 2, "three": 3, "four": 4, "five": 5,
            "six": 6, "seven": 7, "eight": 8, "nine": 9, "ten": 10,
            "eleven": 11, "twelve": 12, "thirteen": 13, "fourteen": 14, "fifteen": 15,
            "sixteen": 16, "seventeen": 17, "eighteen": 18, "nineteen": 19, "twenty": 20,
            "thirty": 30, "forty": 40, "fifty": 50, "sixty": 60,
            "a": 1, "an": 1,  # "a minute" = 1 minute
        }
        
        # Convert word numbers to digits in text
        words = text.split()
        converted_words = []
        i = 0
        while i < len(words):
            word = words[i].lower()
            
            # Handle compound numbers like "twenty five"
            if word in word_to_num and i + 1 < len(words):
                next_word = words[i + 1].lower()
                if next_word in word_to_num and word_to_num[word] >= 20 and word_to_num[next_word] < 10:
                    converted_words.append(str(word_to_num[word] + word_to_num[next_word]))
                    i += 2
                    continue
            
            if word in word_to_num:
                converted_words.append(str(word_to_num[word]))
            else:
                converted_words.append(word)
            i += 1
        
        converted_text = " ".join(converted_words)
        
        # Parse time components
        hours = 0
        minutes = 0
        seconds = 0
        
        # Match patterns like "5 hours", "10 minutes", "30 seconds"
        hour_match = re.search(r'(\d+)\s*(?:hour|hours|hr|hrs)', converted_text)
        minute_match = re.search(r'(\d+)\s*(?:minute|minutes|min|mins)', converted_text)
        second_match = re.search(r'(\d+)\s*(?:second|seconds|sec|secs)', converted_text)
        
        if hour_match:
            hours = int(hour_match.group(1))
        if minute_match:
            minutes = int(minute_match.group(1))
        if second_match:
            seconds = int(second_match.group(1))
        
        # If no time unit found, try to find a standalone number and assume minutes
        if hours == 0 and minutes == 0 and seconds == 0:
            # Look for a number after "for"
            for_match = re.search(r'for\s+(\d+)', converted_text)
            if for_match:
                minutes = int(for_match.group(1))
        
        # Validate we have some time
        total_seconds = hours * 3600 + minutes * 60 + seconds
        if total_seconds <= 0:
            return False, "Please specify a time (e.g., 'set a timer for 5 minutes')"
        
        # Cap at 24 hours
        if total_seconds > 86400:
            return False, "Timer cannot exceed 24 hours"
        
        # Start the timer
        timer_data = {"minutes": hours * 60 + minutes, "seconds": seconds}
        return self._start_timer(timer_data, text)
    
    def _start_timer(self, timer_data: dict, phrase: str) -> tuple[bool, str]:
        """Start a timer that will play a sound when finished."""
        minutes = timer_data.get("minutes", 0)
        seconds = timer_data.get("seconds", 0)
        total_seconds = minutes * 60 + seconds
        
        # Stop any existing timer
        self.stop_timer()
        
        # Create stop event
        self._timer_stop_event = threading.Event()
        
        # Get timer settings
        timer_settings = self.get_timer_settings()
        
        def timer_thread():
            # Wait for the timer duration
            self._timer_stop_event.wait(total_seconds)
            
            if not self._timer_stop_event.is_set():
                # Timer completed, start playing sound if enabled
                if timer_settings.get("alarm_sound", True):
                    self._play_timer_alarm()
        
        self._timer_thread = threading.Thread(target=timer_thread, daemon=True)
        self._timer_thread.start()
        self._active_timer = {"minutes": minutes, "seconds": seconds, "phrase": phrase}
        
        # Play confirmation sound if enabled
        if timer_settings.get("confirm_sound", True):
            self._play_timer_confirm()
        
        time_str = ""
        if minutes > 0:
            time_str += f"{minutes} minute{'s' if minutes != 1 else ''}"
        if seconds > 0:
            if time_str:
                time_str += " and "
            time_str += f"{seconds} second{'s' if seconds != 1 else ''}"
        
        return True, f"Timer set for {time_str}"
    
    def _play_timer_confirm(self) -> None:
        """Play a short confirmation sound when timer is set."""
        import winsound
        try:
            # Short ascending tones to confirm
            winsound.Beep(600, 100)
            time.sleep(0.05)
            winsound.Beep(800, 100)
        except Exception:
            pass
    
    def _generate_alarm_wav(self, volume_percent: int = 100) -> bytes:
        """Generate a WAV alarm sound with specified volume (10-100%)."""
        import struct
        import math
        
        # Audio parameters
        sample_rate = 44100
        duration = 0.4  # seconds
        frequency = 880  # Hz (A5 note)
        
        # Calculate volume (10-100% mapped to 0.1-1.0 amplitude)
        amplitude = (volume_percent / 100.0) * 32767 * 0.8  # 80% max to avoid clipping
        
        # Generate samples
        num_samples = int(sample_rate * duration)
        samples = []
        
        for i in range(num_samples):
            # Apply envelope for smoother sound
            t = i / sample_rate
            envelope = 1.0
            if t < 0.02:  # Attack
                envelope = t / 0.02
            elif t > duration - 0.05:  # Release
                envelope = (duration - t) / 0.05
            
            # Generate sine wave
            sample = int(amplitude * envelope * math.sin(2 * math.pi * frequency * t))
            samples.append(struct.pack('<h', sample))
        
        # Create WAV header
        data = b''.join(samples)
        wav_header = struct.pack('<4sI4s4sIHHIIHH4sI',
            b'RIFF',
            36 + len(data),
            b'WAVE',
            b'fmt ',
            16,  # Subchunk1Size
            1,   # AudioFormat (PCM)
            1,   # NumChannels
            sample_rate,
            sample_rate * 2,  # ByteRate
            2,   # BlockAlign
            16,  # BitsPerSample
            b'data',
            len(data)
        )
        
        return wav_header + data
    
    def play_timer_alarm_once(self) -> None:
        """Play the timer alarm sound once (for testing)."""
        import winsound
        import tempfile
        import os
        
        timer_settings = self.get_timer_settings()
        volume = timer_settings.get("alarm_volume", 100)
        
        try:
            # Generate WAV with volume
            wav_data = self._generate_alarm_wav(volume)
            
            # Write to temp file and play
            with tempfile.NamedTemporaryFile(suffix='.wav', delete=False) as f:
                f.write(wav_data)
                temp_path = f.name
            
            winsound.PlaySound(temp_path, winsound.SND_FILENAME)
            
            # Clean up
            try:
                os.unlink(temp_path)
            except:
                pass
        except Exception as e:
            # Fallback to system beep
            try:
                winsound.Beep(880, 400)
            except:
                pass
    
    def _play_timer_alarm(self) -> None:
        """Play the timer alarm sound repeatedly until stopped."""
        import winsound
        import tempfile
        import os
        
        timer_settings = self.get_timer_settings()
        volume = timer_settings.get("alarm_volume", 100)
        
        # Generate WAV file once
        temp_path = None
        try:
            wav_data = self._generate_alarm_wav(volume)
            with tempfile.NamedTemporaryFile(suffix='.wav', delete=False) as f:
                f.write(wav_data)
                temp_path = f.name
        except:
            temp_path = None
        
        while self._timer_stop_event and not self._timer_stop_event.is_set():
            try:
                if temp_path:
                    winsound.PlaySound(temp_path, winsound.SND_FILENAME)
                else:
                    # Fallback to system beep
                    winsound.Beep(880, 400)
                time.sleep(0.2)
                
                if temp_path:
                    winsound.PlaySound(temp_path, winsound.SND_FILENAME)
                else:
                    winsound.Beep(880, 400)
                time.sleep(0.5)
            except Exception:
                break
        
        # Clean up temp file
        if temp_path:
            try:
                os.unlink(temp_path)
            except:
                pass
    
    def stop_timer(self) -> bool:
        """Stop the active timer and alarm."""
        if self._timer_stop_event:
            self._timer_stop_event.set()
            self._active_timer = None
            return True
        return False
    
    def has_active_timer(self) -> bool:
        """Check if there's an active timer."""
        return self._active_timer is not None


class VoiceRecognizer:
    """Handles voice recognition with Vosk (offline), Whisper (offline), and Google Speech (online) backends."""
    
    VOSK_SAMPLE_RATE = 16000
    VOSK_CHUNK_SIZE = 4000  # ~0.25s of audio at 16kHz mono 16-bit
    WHISPER_SAMPLE_RATE = 16000
    WHISPER_CHUNK_DURATION = 3  # seconds of silence before processing

    def __init__(self, config_path: str = "config.json"):
        self.recognizer = sr.Recognizer()
        self.config_path = Path(config_path)
        self.microphone_index: Optional[int] = None
        self.engine: str = "vosk" if VOSK_AVAILABLE else "google"
        self.configured_energy_threshold: int = 4000
        self._vosk_model = None
        self._whisper_model = None
        self.is_listening = False
        self._listen_thread: Optional[threading.Thread] = None
        self._stop_event = threading.Event()
        self.on_recognized: Optional[Callable[[str], None]] = None
        self.on_error: Optional[Callable[[str], None]] = None
        self.on_listening_state_changed: Optional[Callable[[bool], None]] = None
        self._load_settings()
    
    def _get_vosk_model_path(self) -> Optional[str]:
        """Resolve the Vosk model directory, checking frozen bundle first."""
        candidates = []
        if getattr(sys, '_MEIPASS', None):
            candidates.append(os.path.join(sys._MEIPASS, 'vosk-model'))
        candidates.append(os.path.join(os.path.dirname(os.path.abspath(__file__)), 'vosk-model'))
        for p in candidates:
            if os.path.isdir(p):
                return p
        return None

    def _ensure_vosk_model(self) -> bool:
        """Load the Vosk model if not already loaded. Returns True on success."""
        if self._vosk_model is not None:
            return True
        if not VOSK_AVAILABLE:
            return False
        model_path = self._get_vosk_model_path()
        if model_path is None:
            if self.on_error:
                self.on_error("Vosk model not found. Place 'vosk-model' folder next to the application.")
            return False
        try:
            vosk.SetLogLevel(-1)
            self._vosk_model = vosk.Model(model_path)
            return True
        except Exception as e:
            if self.on_error:
                self.on_error(f"Failed to load Vosk model: {e}")
            return False

    def _ensure_whisper_model(self) -> bool:
        """Load the Whisper model if not already loaded. Returns True on success."""
        if self._whisper_model is not None:
            return True
        if not WHISPER_AVAILABLE:
            return False
        try:
            self._whisper_model = WhisperModel("tiny", device="cpu", compute_type="int8")
            return True
        except Exception as e:
            if self.on_error:
                self.on_error(f"Failed to load Whisper model: {e}")
            return False

    def _load_settings(self) -> None:
        """Load recognizer settings from config."""
        if self.config_path.exists():
            try:
                with open(self.config_path, 'r') as f:
                    config = json.load(f)
                    settings = config.get("settings", {})
                    self.microphone_index = settings.get("microphone_index")
                    self.configured_energy_threshold = settings.get("energy_threshold", 4000)
                    self.recognizer.energy_threshold = self.configured_energy_threshold
                    self.recognizer.pause_threshold = settings.get("pause_threshold", 0.8)
                    saved_engine = settings.get("recognition_engine", "vosk")
                    if saved_engine == "vosk" and VOSK_AVAILABLE:
                        self.engine = "vosk"
                    elif saved_engine == "whisper" and WHISPER_AVAILABLE:
                        self.engine = "whisper"
                    else:
                        self.engine = "google"
            except (json.JSONDecodeError, IOError):
                pass
    
    @staticmethod
    def get_microphones() -> list[tuple[int, str]]:
        """Get list of available microphones."""
        mics = []
        try:
            for index, name in enumerate(sr.Microphone.list_microphone_names()):
                mics.append((index, name))
        except Exception as e:
            print(f"Error getting microphones: {e}")
        return mics
    
    def set_microphone(self, index: Optional[int]) -> None:
        """Set the microphone to use by index."""
        self.microphone_index = index
        self._save_settings()
    
    def _save_settings(self) -> None:
        """Save recognizer settings to config."""
        config = {"commands": {}, "settings": {}}
        if self.config_path.exists():
            try:
                with open(self.config_path, 'r') as f:
                    config = json.load(f)
            except (json.JSONDecodeError, IOError):
                pass
        
        config["settings"]["microphone_index"] = self.microphone_index
        config["settings"]["energy_threshold"] = self.configured_energy_threshold
        config["settings"]["pause_threshold"] = self.recognizer.pause_threshold
        config["settings"]["recognition_engine"] = self.engine
        
        try:
            with open(self.config_path, 'w') as f:
                json.dump(config, f, indent=4)
        except IOError as e:
            print(f"Error saving settings: {e}")
    
    def start_listening(self) -> None:
        """Start listening for voice commands in a background thread."""
        if self.is_listening:
            return
        
        self._stop_event.clear()
        self.is_listening = True

        if self.engine == "vosk":
            target = self._listen_loop_vosk
        elif self.engine == "whisper":
            target = self._listen_loop_whisper
        else:
            target = self._listen_loop_google

        self._listen_thread = threading.Thread(target=target, daemon=True)
        self._listen_thread.start()
        
        if self.on_listening_state_changed:
            self.on_listening_state_changed(True)
    
    def stop_listening(self) -> None:
        """Stop listening for voice commands."""
        self._stop_event.set()
        self.is_listening = False
        
        if self.on_listening_state_changed:
            self.on_listening_state_changed(False)
    
    def toggle_listening(self) -> bool:
        """Toggle listening state. Returns new state."""
        if self.is_listening:
            self.stop_listening()
        else:
            self.start_listening()
        return self.is_listening
    
    def _listen_loop_google(self) -> None:
        """Listening loop using Google Speech Recognition (online)."""
        try:
            mic_kwargs = {}
            if self.microphone_index is not None:
                mic_kwargs['device_index'] = self.microphone_index
            
            with sr.Microphone(**mic_kwargs) as source:
                self.recognizer.adjust_for_ambient_noise(source, duration=0.5)
                
                while not self._stop_event.is_set():
                    try:
                        audio = self.recognizer.listen(
                            source, 
                            timeout=1, 
                            phrase_time_limit=5
                        )
                        
                        if self._stop_event.is_set():
                            break
                        
                        text = self.recognizer.recognize_google(audio)
                        if self.on_recognized:
                            self.on_recognized(text)
                            
                    except sr.WaitTimeoutError:
                        continue
                    except sr.UnknownValueError:
                        continue
                    except sr.RequestError as e:
                        if self.on_error:
                            self.on_error(f"Speech recognition service error: {e}")
                        
        except Exception as e:
            if self.on_error:
                self.on_error(f"Microphone error: {e}")
            self.is_listening = False
            if self.on_listening_state_changed:
                self.on_listening_state_changed(False)

    def _listen_loop_vosk(self) -> None:
        """Listening loop using Vosk (offline)."""
        import pyaudio

        if not self._ensure_vosk_model():
            self.is_listening = False
            if self.on_listening_state_changed:
                self.on_listening_state_changed(False)
            return

        pa = None
        stream = None
        try:
            rec = vosk.KaldiRecognizer(self._vosk_model, self.VOSK_SAMPLE_RATE)
            pa = pyaudio.PyAudio()

            stream_kwargs = {
                "format": pyaudio.paInt16,
                "channels": 1,
                "rate": self.VOSK_SAMPLE_RATE,
                "input": True,
                "frames_per_buffer": self.VOSK_CHUNK_SIZE,
            }
            if self.microphone_index is not None:
                stream_kwargs["input_device_index"] = self.microphone_index

            stream = pa.open(**stream_kwargs)

            while not self._stop_event.is_set():
                data = stream.read(self.VOSK_CHUNK_SIZE, exception_on_overflow=False)
                if rec.AcceptWaveform(data):
                    result = json.loads(rec.Result())
                    text = result.get("text", "").strip()
                    if text and self.on_recognized:
                        self.on_recognized(text)

        except Exception as e:
            if self.on_error:
                self.on_error(f"Vosk microphone error: {e}")
            self.is_listening = False
            if self.on_listening_state_changed:
                self.on_listening_state_changed(False)
        finally:
            if stream is not None:
                try:
                    stream.stop_stream()
                    stream.close()
                except Exception:
                    pass
            if pa is not None:
                try:
                    pa.terminate()
                except Exception:
                    pass

    def _listen_loop_whisper(self) -> None:
        """Listening loop using Whisper (offline). Uses SpeechRecognition for audio capture, Whisper for transcription."""
        import numpy as np

        if not self._ensure_whisper_model():
            self.is_listening = False
            if self.on_listening_state_changed:
                self.on_listening_state_changed(False)
            return

        try:
            mic_kwargs = {}
            if self.microphone_index is not None:
                mic_kwargs['device_index'] = self.microphone_index

            with sr.Microphone(sample_rate=self.WHISPER_SAMPLE_RATE, **mic_kwargs) as source:
                self.recognizer.adjust_for_ambient_noise(source, duration=0.5)

                while not self._stop_event.is_set():
                    try:
                        audio = self.recognizer.listen(
                            source,
                            timeout=1,
                            phrase_time_limit=10
                        )

                        if self._stop_event.is_set():
                            break

                        # Convert SpeechRecognition audio to numpy float32 for Whisper
                        raw_data = audio.get_raw_data(convert_rate=self.WHISPER_SAMPLE_RATE, convert_width=2)
                        audio_array = np.frombuffer(raw_data, dtype=np.int16).astype(np.float32) / 32768.0

                        segments, _ = self._whisper_model.transcribe(
                            audio_array, beam_size=1, language="en"
                        )
                        text = " ".join(seg.text for seg in segments).strip()
                        if text and self.on_recognized:
                            self.on_recognized(text)

                    except sr.WaitTimeoutError:
                        continue
                    except sr.UnknownValueError:
                        continue
                    except Exception as e:
                        if self.on_error:
                            self.on_error(f"Whisper recognition error: {e}")

        except Exception as e:
            if self.on_error:
                self.on_error(f"Whisper microphone error: {e}")
            self.is_listening = False
            if self.on_listening_state_changed:
                self.on_listening_state_changed(False)
