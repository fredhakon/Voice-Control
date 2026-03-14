"""
Voice Control Platform - CustomTkinter UI Variant
Most modern, sleek UI with rounded corners and modern widgets.
"""

import customtkinter as ctk
from typing import Optional, List
import json
from pathlib import Path
import keyboard
import os
import sys
import tkinter as tk
import threading
import io
import webbrowser

# For album art loading
try:
    from PIL import Image
    import urllib.request
    PIL_AVAILABLE = True
except ImportError:
    PIL_AVAILABLE = False

# For system tray
try:
    import pystray
    from pystray import MenuItem as item
    TRAY_AVAILABLE = True
except ImportError:
    TRAY_AVAILABLE = False

from voice_control import VoiceRecognizer, CommandManager, VolumeController, VOSK_AVAILABLE, WHISPER_AVAILABLE

# Try to import Spotify controller
try:
    from spotify_control import SpotifyController, SPOTIPY_AVAILABLE
except ImportError:
    SPOTIPY_AVAILABLE = False
    SpotifyController = None

# Configure CustomTkinter
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")


def get_config_path() -> Path:
    """Get the path to the config file.
    Uses AppData on Windows for installed apps to ensure write permissions.
    """
    if getattr(sys, 'frozen', False):
        # Use AppData/Local for installed apps
        appdata = os.environ.get('LOCALAPPDATA', os.path.expanduser('~'))
        config_dir = Path(appdata) / 'VoiceControl'
        config_dir.mkdir(parents=True, exist_ok=True)
        return config_dir / "config.json"
    else:
        # Development mode - use current directory
        return Path("config.json")


def get_default_config_path() -> Optional[Path]:
    """Locate config.default.json across all possible locations."""
    candidates = []
    meipass = getattr(sys, '_MEIPASS', None)
    if meipass:
        candidates.append(Path(meipass) / 'config.default.json')
    candidates.append(get_config_path().parent / 'config.default.json')
    candidates.append(Path(__file__).parent / 'config.default.json')
    candidates.append(Path('config.default.json'))
    for p in candidates:
        if p.exists():
            return p
    return None


def is_admin() -> bool:
    """Check if the application is running with administrator privileges."""
    import ctypes
    try:
        return ctypes.windll.shell32.IsUserAnAdmin() != 0
    except Exception:
        return False


def get_app_icon_path() -> Optional[str]:
    """Get the path to the application icon file."""
    possible_paths = [
        Path(__file__).parent / "icon.ico",
        Path("icon.ico"),
    ]
    for path in possible_paths:
        if path.exists():
            return str(path)
    return None


def set_window_icon(window) -> None:
    """Set the application icon for a window."""
    icon_path = get_app_icon_path()
    if icon_path:
        try:
            window.iconbitmap(icon_path)
        except Exception:
            pass


class CTkMessagebox:
    """Simple messagebox replacement for CustomTkinter."""
    
    @staticmethod
    def show(parent, title: str, message: str, icon: str = "info"):
        dialog = ctk.CTkToplevel(parent)
        dialog.title(title)
        dialog.geometry("350x180")
        dialog.resizable(False, False)
        dialog.transient(parent)
        dialog.grab_set()
        
        # Center on parent
        dialog.update_idletasks()
        x = parent.winfo_x() + (parent.winfo_width() - 350) // 2
        y = parent.winfo_y() + (parent.winfo_height() - 180) // 2
        dialog.geometry(f"+{x}+{y}")
        
        frame = ctk.CTkFrame(dialog, fg_color="transparent")
        frame.pack(fill="both", expand=True, padx=25, pady=25)
        
        icons = {"info": "ℹ️", "warning": "⚠️", "error": "❌", "success": "✅"}
        icon_text = icons.get(icon, "ℹ️")
        
        ctk.CTkLabel(frame, text=f"{icon_text} {title}", font=ctk.CTkFont(size=16, weight="bold")).pack(anchor="w")
        ctk.CTkLabel(frame, text=message, wraplength=300).pack(anchor="w", pady=(15, 0))
        
        ctk.CTkButton(frame, text="OK", command=dialog.destroy, width=100).pack(side="right", pady=(20, 0))
        
        dialog.wait_window()
    
    @staticmethod
    def ask_yes_no(parent, title: str, message: str) -> bool:
        result = [False]
        
        dialog = ctk.CTkToplevel(parent)
        dialog.title(title)
        dialog.geometry("350x180")
        dialog.resizable(False, False)
        dialog.transient(parent)
        dialog.grab_set()
        
        dialog.update_idletasks()
        x = parent.winfo_x() + (parent.winfo_width() - 350) // 2
        y = parent.winfo_y() + (parent.winfo_height() - 180) // 2
        dialog.geometry(f"+{x}+{y}")
        
        frame = ctk.CTkFrame(dialog, fg_color="transparent")
        frame.pack(fill="both", expand=True, padx=25, pady=25)
        
        ctk.CTkLabel(frame, text=title, font=ctk.CTkFont(size=16, weight="bold")).pack(anchor="w")
        ctk.CTkLabel(frame, text=message, wraplength=300).pack(anchor="w", pady=(15, 0))
        
        btn_frame = ctk.CTkFrame(frame, fg_color="transparent")
        btn_frame.pack(fill="x", pady=(20, 0))
        
        def yes():
            result[0] = True
            dialog.destroy()
        
        ctk.CTkButton(btn_frame, text="Yes", command=yes, width=80).pack(side="right", padx=(10, 0))
        ctk.CTkButton(btn_frame, text="No", command=dialog.destroy, width=80, 
                      fg_color="transparent", border_width=1).pack(side="right")
        
        dialog.wait_window()
        return result[0]


class SettingsDialog(ctk.CTkToplevel):
    """Modern settings dialog."""
    
    def __init__(self, parent, recognizer: VoiceRecognizer, command_manager: CommandManager):
        super().__init__(parent)
        self.parent = parent
        self.recognizer = recognizer
        self.command_manager = command_manager
        
        self.title("Settings")
        self.geometry("550x680")
        self.resizable(False, False)
        self.transient(parent)
        self.grab_set()
        set_window_icon(self)
        
        self._create_widgets()
        self._center_window()
    
    def _center_window(self) -> None:
        self.update_idletasks()
        x = self.parent.winfo_x() + (self.parent.winfo_width() - self.winfo_width()) // 2
        y = self.parent.winfo_y() + (self.parent.winfo_height() - self.winfo_height()) // 2
        self.geometry(f"+{x}+{y}")
    
    def _create_widgets(self) -> None:
        # Scrollable frame
        scroll_frame = ctk.CTkScrollableFrame(self, fg_color="transparent")
        scroll_frame.pack(fill="both", expand=True, padx=20, pady=20)
        
        # ===== Microphone Section =====
        self._create_section(scroll_frame, "🎤 Microphone")
        
        mic_card = ctk.CTkFrame(scroll_frame, corner_radius=12)
        mic_card.pack(fill="x", pady=(0, 20))
        
        mic_inner = ctk.CTkFrame(mic_card, fg_color="transparent")
        mic_inner.pack(fill="x", padx=20, pady=20)
        
        ctk.CTkLabel(mic_inner, text="Input Device").pack(anchor="w")
        self.mic_combo = ctk.CTkComboBox(mic_inner, state="readonly", width=400)
        self.mic_combo.pack(fill="x", pady=(5, 15))
        self._populate_microphones()
        
        # Recognition engine selector
        ctk.CTkLabel(mic_inner, text="Recognition Engine").pack(anchor="w")
        engine_values = []
        if VOSK_AVAILABLE:
            engine_values.append("Vosk (Offline)")
        if WHISPER_AVAILABLE:
            engine_values.append("Whisper (Offline)")
        engine_values.append("Google Speech (Online)")
        engine_map = {"vosk": "Vosk (Offline)", "whisper": "Whisper (Offline)", "google": "Google Speech (Online)"}
        current_engine = engine_map.get(self.recognizer.engine, "Google Speech (Online)")
        self.engine_var = ctk.StringVar(value=current_engine)
        self.engine_combo = ctk.CTkComboBox(mic_inner, values=engine_values, variable=self.engine_var, state="readonly", width=400)
        self.engine_combo.pack(fill="x", pady=(5, 15))
        
        ctk.CTkLabel(mic_inner, text="Sensitivity").pack(anchor="w")
        self.energy_var = ctk.IntVar(value=self.recognizer.configured_energy_threshold)
        energy_slider = ctk.CTkSlider(mic_inner, from_=1000, to=8000, variable=self.energy_var, width=400)
        energy_slider.pack(fill="x", pady=(5, 15))
        
        pause_label_row = ctk.CTkFrame(mic_inner, fg_color="transparent")
        pause_label_row.pack(fill="x")
        ctk.CTkLabel(pause_label_row, text="Pause Duration").pack(side="left")
        self.pause_value_label = ctk.CTkLabel(pause_label_row, text=f"{self.recognizer.recognizer.pause_threshold:.1f}s", text_color="gray")
        self.pause_value_label.pack(side="right")
        self.pause_var = ctk.DoubleVar(value=self.recognizer.recognizer.pause_threshold)
        pause_slider = ctk.CTkSlider(mic_inner, from_=0.5, to=3.0, variable=self.pause_var, width=400,
                                      command=lambda v: self.pause_value_label.configure(text=f"{float(v):.1f}s"))
        pause_slider.pack(fill="x", pady=(5, 15))
        
        btn_row = ctk.CTkFrame(mic_inner, fg_color="transparent")
        btn_row.pack(fill="x")
        
        self.calibrate_btn = ctk.CTkButton(btn_row, text="Calibrate", command=self._calibrate_microphone, width=120)
        self.calibrate_btn.pack(side="left")
        
        self.calibrate_status = ctk.CTkLabel(btn_row, text="", text_color="gray")
        self.calibrate_status.pack(side="left", padx=(15, 0))
        
        # ===== Audio Output Section =====
        self._create_section(scroll_frame, "🔊 Audio Output")
        
        output_card = ctk.CTkFrame(scroll_frame, corner_radius=12)
        output_card.pack(fill="x", pady=(0, 20))
        
        output_inner = ctk.CTkFrame(output_card, fg_color="transparent")
        output_inner.pack(fill="x", padx=20, pady=20)
        
        ctk.CTkLabel(output_inner, text="Output Device").pack(anchor="w")
        self.output_combo = ctk.CTkComboBox(output_inner, state="readonly", width=400)
        self.output_combo.pack(fill="x", pady=(5, 0))
        self._populate_output_devices()
        
        # ===== Shortcut Section =====
        self._create_section(scroll_frame, "⌨️ Keyboard Shortcut")
        
        shortcut_card = ctk.CTkFrame(scroll_frame, corner_radius=12)
        shortcut_card.pack(fill="x", pady=(0, 20))
        
        shortcut_inner = ctk.CTkFrame(shortcut_card, fg_color="transparent")
        shortcut_inner.pack(fill="x", padx=20, pady=20)
        
        ctk.CTkLabel(shortcut_inner, text="Toggle Listening").pack(anchor="w")
        
        shortcut_row = ctk.CTkFrame(shortcut_inner, fg_color="transparent")
        shortcut_row.pack(fill="x", pady=(5, 0))
        
        self.shortcut_var = ctk.StringVar(value=self._load_shortcut())
        self.shortcut_entry = ctk.CTkEntry(shortcut_row, textvariable=self.shortcut_var, width=300)
        self.shortcut_entry.pack(side="left")
        
        ctk.CTkButton(shortcut_row, text="Record", command=self._record_shortcut, width=100).pack(side="left", padx=(10, 0))
        
        # ===== Spotify Section =====
        if SPOTIPY_AVAILABLE:
            self._create_section(scroll_frame, "🎵 Spotify")
            
            spotify_card = ctk.CTkFrame(scroll_frame, corner_radius=12)
            spotify_card.pack(fill="x", pady=(0, 20))
            
            spotify_inner = ctk.CTkFrame(spotify_card, fg_color="transparent")
            spotify_inner.pack(fill="x", padx=20, pady=20)
            
            self.spotify_status_var = ctk.StringVar()
            self._update_spotify_status()
            
            ctk.CTkLabel(spotify_inner, textvariable=self.spotify_status_var, 
                         font=ctk.CTkFont(size=13)).pack(anchor="w", pady=(0, 15))
            
            ctk.CTkLabel(spotify_inner, text="Client ID").pack(anchor="w")
            self.spotify_client_id_var = ctk.StringVar()
            ctk.CTkEntry(spotify_inner, textvariable=self.spotify_client_id_var, width=400).pack(fill="x", pady=(5, 15))
            
            ctk.CTkLabel(spotify_inner, text="Client Secret").pack(anchor="w")
            self.spotify_client_secret_var = ctk.StringVar()
            ctk.CTkEntry(spotify_inner, textvariable=self.spotify_client_secret_var, show="●", width=400).pack(fill="x", pady=(5, 15))
            
            self._load_spotify_credentials()
            
            btn_row = ctk.CTkFrame(spotify_inner, fg_color="transparent")
            btn_row.pack(fill="x")
            
            ctk.CTkButton(btn_row, text="Connect", command=self._connect_spotify, width=100).pack(side="left", padx=(0, 10))
            ctk.CTkButton(btn_row, text="Disconnect", command=self._disconnect_spotify, width=100,
                          fg_color="transparent", border_width=1).pack(side="left", padx=(0, 10))
            ctk.CTkButton(btn_row, text="Commands", command=self._show_spotify_commands, width=100,
                          fg_color="transparent", border_width=1).pack(side="left")
        
        # ===== Bottom Buttons =====
        btn_frame = ctk.CTkFrame(scroll_frame, fg_color="transparent")
        btn_frame.pack(fill="x", pady=(10, 0))
        
        ctk.CTkButton(btn_frame, text="Save", command=self._save_settings, width=100).pack(side="right", padx=(10, 0))
        ctk.CTkButton(btn_frame, text="Cancel", command=self.destroy, width=100,
                      fg_color="transparent", border_width=1).pack(side="right")
        ctk.CTkButton(btn_frame, text="Reset to Default", command=self._reset_to_default, width=140,
                      fg_color="#cc3333", hover_color="#aa2222", border_width=1).pack(side="left")
    
    def _create_section(self, parent, title: str) -> None:
        ctk.CTkLabel(parent, text=title, font=ctk.CTkFont(size=16, weight="bold")).pack(anchor="w", pady=(0, 10))
    
    def _populate_microphones(self) -> None:
        mics = self.recognizer.get_microphones()
        mic_names = ["Default"] + [name for _, name in mics]
        self.mic_combo.configure(values=mic_names)
        self.mic_combo.set(mic_names[0])
    
    def _populate_output_devices(self) -> None:
        self.output_devices = VolumeController.get_output_devices()
        device_names = ["Default"] + [name for _, name in self.output_devices]
        self.output_combo.configure(values=device_names)
        self.output_combo.set(device_names[0])
    
    def _load_shortcut(self) -> str:
        config_path = get_config_path()
        if config_path.exists():
            try:
                with open(config_path, 'r') as f:
                    return json.load(f).get("settings", {}).get("toggle_shortcut", "ctrl+shift+v")
            except:
                pass
        return "ctrl+shift+v"
    
    def _record_shortcut(self) -> None:
        self.shortcut_entry.delete(0, "end")
        self.shortcut_entry.insert(0, "Press keys...")
        
        def on_key_event(event):
            if event.event_type == "down":
                parts = []
                if keyboard.is_pressed("ctrl"): parts.append("ctrl")
                if keyboard.is_pressed("alt"): parts.append("alt")
                if keyboard.is_pressed("shift"): parts.append("shift")
                
                key_name = event.name.lower()
                if key_name not in ["ctrl", "alt", "shift", "left ctrl", "right ctrl", 
                                    "left alt", "right alt", "left shift", "right shift"]:
                    parts.append(key_name)
                
                if parts and parts[-1] not in ["ctrl", "alt", "shift"]:
                    self.shortcut_var.set("+".join(parts))
                    keyboard.unhook(on_key_event)
        
        keyboard.hook(on_key_event)
    
    def _calibrate_microphone(self) -> None:
        import threading
        import speech_recognition as sr
        
        self.calibrate_btn.configure(state="disabled")
        self.calibrate_status.configure(text="Calibrating...")
        
        def do_calibration():
            try:
                mic_index = None
                selection = self.mic_combo.get()
                mics = self.recognizer.get_microphones()
                mic_names = ["Default"] + [name for _, name in mics]
                
                if selection != "Default":
                    idx = mic_names.index(selection) - 1
                    if idx >= 0 and idx < len(mics):
                        mic_index = mics[idx][0]
                
                mic_kwargs = {}
                if mic_index is not None:
                    mic_kwargs['device_index'] = mic_index
                
                with sr.Microphone(**mic_kwargs) as source:
                    self.recognizer.recognizer.adjust_for_ambient_noise(source, duration=2)
                
                new_threshold = int(self.recognizer.recognizer.energy_threshold)
                
                def update_ui():
                    self.energy_var.set(new_threshold)
                    self.calibrate_status.configure(text=f"Done! ({new_threshold})")
                    self.calibrate_btn.configure(state="normal")
                
                self.after(0, update_ui)
                
            except Exception as e:
                def show_error():
                    self.calibrate_status.configure(text="Error!")
                    self.calibrate_btn.configure(state="normal")
                self.after(0, show_error)
        
        threading.Thread(target=do_calibration, daemon=True).start()
    
    def _update_spotify_status(self) -> None:
        if self.command_manager.spotify and self.command_manager.spotify.is_authenticated:
            self.spotify_status_var.set("✅ Connected to Spotify")
        else:
            self.spotify_status_var.set("Not connected")
    
    def _load_spotify_credentials(self) -> None:
        if self.command_manager.spotify:
            client_id, client_secret = self.command_manager.spotify.get_credentials()
            self.spotify_client_id_var.set(client_id)
            self.spotify_client_secret_var.set(client_secret)
    
    def _connect_spotify(self) -> None:
        client_id = self.spotify_client_id_var.get().strip()
        client_secret = self.spotify_client_secret_var.get().strip()
        
        if not client_id or not client_secret:
            CTkMessagebox.show(self, "Warning", "Please enter both Client ID and Client Secret.", "warning")
            return
        
        # Show EULA/Privacy acceptance dialog before connecting
        if not self._show_spotify_eula_dialog():
            return
        
        def do_connect():
            success, message = self.command_manager.spotify.authenticate(client_id, client_secret)
            def update_ui():
                if success:
                    self.spotify_status_var.set(f"✅ {message}")
                    CTkMessagebox.show(self, "Success", message, "success")
                else:
                    self.spotify_status_var.set("Connection failed")
                    CTkMessagebox.show(self, "Error", message, "error")
            self.after(0, update_ui)
        
        import threading
        threading.Thread(target=do_connect, daemon=True).start()
    
    def _show_spotify_eula_dialog(self) -> bool:
        """Show EULA/Privacy acceptance dialog. Returns True if accepted, False if declined."""
        accepted = {"value": False}
        
        dialog = ctk.CTkToplevel(self)
        dialog.title("Terms & Privacy Policy")
        dialog.geometry("520x480")
        dialog.resizable(False, False)
        dialog.transient(self)
        dialog.grab_set()
        set_window_icon(dialog)
        
        frame = ctk.CTkFrame(dialog, fg_color="transparent")
        frame.pack(fill="both", expand=True, padx=25, pady=20)
        
        ctk.CTkLabel(frame, text="Terms & Privacy Policy", 
                     font=ctk.CTkFont(size=18, weight="bold")).pack(anchor="w", pady=(0, 10))
        
        ctk.CTkLabel(frame, text="Please read and accept before connecting to Spotify:",
                     text_color="gray").pack(anchor="w", pady=(0, 10))
        
        text = ctk.CTkTextbox(frame, height=280, font=ctk.CTkFont(size=12))
        text.pack(fill="both", expand=True, pady=(0, 15))
        
        summary = (
            "END USER LICENSE AGREEMENT (Summary)\n"
            "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n\n"
            "• This software is provided \"as is\" without warranty.\n"
            "• Spotify features are subject to Spotify's Terms of Use.\n"
            "• Spotify is a third-party beneficiary of this agreement.\n"
            "• You may not reverse-engineer or create derivative works\n"
            "  of the Spotify Platform, Service, or Content.\n"
            "• No warranties are made on behalf of Spotify.\n\n"
            "PRIVACY POLICY (Summary)\n"
            "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n\n"
            "By connecting to Spotify, this app will access:\n\n"
            "• Your current playback state (track, artist, album art)\n"
            "• Your playlists and liked songs\n"
            "• Playback control (play, pause, skip, search)\n\n"
            "Data handling:\n\n"
            "• Spotify OAuth tokens are stored locally on your device.\n"
            "• Your Spotify credentials (Client ID/Secret) are stored\n"
            "  locally and never sent to us or any third party.\n"
            "• No personal data is sold, shared, or sent to external\n"
            "  servers (except Spotify's API for playback features).\n"
            "• You can disconnect and delete all Spotify data at any time.\n\n"
            "Full documents: See EULA.md and PRIVACY.md in the app folder."
        )
        text.insert("1.0", summary)
        text.configure(state="disabled")
        
        btn_frame = ctk.CTkFrame(frame, fg_color="transparent")
        btn_frame.pack(fill="x")
        
        def on_decline():
            accepted["value"] = False
            dialog.destroy()
        
        def on_accept():
            accepted["value"] = True
            dialog.destroy()
        
        ctk.CTkButton(btn_frame, text="Decline", command=on_decline, width=100,
                      fg_color="transparent", border_width=1, text_color="gray").pack(side="left")
        ctk.CTkButton(btn_frame, text="I Accept", command=on_accept, width=100).pack(side="right")
        
        # Center on parent
        dialog.update_idletasks()
        x = self.winfo_x() + (self.winfo_width() - 520) // 2
        y = self.winfo_y() + (self.winfo_height() - 480) // 2
        dialog.geometry(f"+{x}+{y}")
        
        dialog.wait_window()
        return accepted["value"]
    
    def _disconnect_spotify(self) -> None:
        if self.command_manager.spotify:
            self.command_manager.spotify.disconnect()
            # Clear credentials from UI
            self.spotify_client_id_var.set("")
            self.spotify_client_secret_var.set("")
            # Clear credentials from config
            self.command_manager.spotify.clear_credentials()
            config = self.command_manager._load_full_config()
            config["spotify"] = {"client_id": "", "client_secret": ""}
            self.command_manager._save_config(config)
            self._update_spotify_status()
            CTkMessagebox.show(self, "Disconnected", "Spotify account unlinked. All tokens and credentials have been removed.", "info")
    
    def _show_spotify_commands(self) -> None:
        dialog = ctk.CTkToplevel(self)
        dialog.title("Spotify Commands")
        dialog.geometry("500x550")
        dialog.transient(self)
        set_window_icon(dialog)
        
        frame = ctk.CTkFrame(dialog, fg_color="transparent")
        frame.pack(fill="both", expand=True, padx=20, pady=20)
        
        ctk.CTkLabel(frame, text="🎵 Spotify Voice Commands", 
                     font=ctk.CTkFont(size=18, weight="bold")).pack(anchor="w", pady=(0, 15))
        
        if SpotifyController:
            help_text = SpotifyController.get_voice_commands_help()
        else:
            help_text = "Spotify not available."
        
        text = ctk.CTkTextbox(frame, font=ctk.CTkFont(size=13), corner_radius=10)
        text.pack(fill="both", expand=True, pady=(0, 15))
        text.insert("1.0", help_text)
        text.configure(state="disabled")
        
        ctk.CTkButton(frame, text="Close", command=dialog.destroy, width=100).pack(side="right")
        
        dialog.update_idletasks()
        x = self.winfo_x() + (self.winfo_width() - 500) // 2
        y = self.winfo_y() + (self.winfo_height() - 550) // 2
        dialog.geometry(f"+{x}+{y}")
    
    def _reset_to_default(self) -> None:
        """Reset config.json to the contents of config.default.json."""
        if not CTkMessagebox.ask_yes_no(self, "Reset Config",
                "This will replace your current settings with the defaults.\n"
                "Profiles are kept. Continue?"):
            return

        import shutil
        default_path = get_default_config_path()
        if default_path is None:
            CTkMessagebox.show(self, "Error", "Could not locate config.default.json.", icon="error")
            return

        config_path = get_config_path()
        try:
            shutil.copy(default_path, config_path)
        except IOError as e:
            CTkMessagebox.show(self, "Error", f"Failed to reset config: {e}", icon="error")
            return

        # Reload in-memory state
        self.command_manager._ensure_config_exists()
        self.command_manager.builtin_phrases = self.command_manager._load_builtin_phrases()
        self.command_manager.load_commands()
        self.recognizer.configured_energy_threshold = 4000
        self.recognizer.recognizer.energy_threshold = 4000

        CTkMessagebox.show(self, "Reset Complete",
                "Settings have been reset to defaults.\nPlease restart the app for full effect.",
                icon="success")
        self.destroy()

    def _save_settings(self) -> None:
        # Apply all recognizer values FIRST before any saves
        self.recognizer.configured_energy_threshold = self.energy_var.get()
        self.recognizer.recognizer.energy_threshold = self.recognizer.configured_energy_threshold
        self.recognizer.recognizer.pause_threshold = round(self.pause_var.get(), 1)
        engine_selection = self.engine_var.get()
        if "Vosk" in engine_selection:
            self.recognizer.engine = "vosk"
        elif "Whisper" in engine_selection:
            self.recognizer.engine = "whisper"
        else:
            self.recognizer.engine = "google"

        # Save microphone (this calls recognizer._save_settings internally)
        selection = self.mic_combo.get()
        if selection == "Default":
            self.recognizer.set_microphone(None)
        else:
            mics = self.recognizer.get_microphones()
            mic_names = ["Default"] + [name for _, name in mics]
            idx = mic_names.index(selection) - 1
            if idx >= 0 and idx < len(mics):
                self.recognizer.set_microphone(mics[idx][0])
        
        # Save output device
        output_selection = self.output_combo.get()
        if output_selection == "Default":
            self.command_manager.set_audio_device(None)
        else:
            device_names = ["Default"] + [name for _, name in self.output_devices]
            idx = device_names.index(output_selection) - 1
            if idx >= 0 and idx < len(self.output_devices):
                self.command_manager.set_audio_device(self.output_devices[idx][0])
        
        # Save shortcut (read-modify-write to preserve other settings)
        shortcut = self.shortcut_var.get()
        config = {}
        config_path = get_config_path()
        if config_path.exists():
            try:
                with open(config_path, 'r') as f:
                    config = json.load(f)
            except:
                pass
        config.setdefault("settings", {})["toggle_shortcut"] = shortcut
        with open(config_path, 'w') as f:
            json.dump(config, f, indent=4)
        
        if hasattr(self.parent, 'update_shortcut'):
            self.parent.update_shortcut(shortcut)
        
        self.destroy()


class AddCommandDialog(ctk.CTkToplevel):
    """Dialog for adding/editing commands."""
    
    def __init__(self, parent, command_manager: CommandManager, edit_command_id: str = None):
        super().__init__(parent)
        self.parent = parent
        self.command_manager = command_manager
        self.edit_command_id = edit_command_id
        self.result = False
        self.macro_steps = []
        
        title = "Edit Command" if edit_command_id else "Add Command"
        self.title(title)
        self.geometry("550x650")
        self.resizable(True, True)
        self.minsize(500, 550)
        self.transient(parent)
        self.grab_set()
        set_window_icon(self)
        
        self._create_widgets()
        self._center_window()
        
        if edit_command_id:
            self._load_command_data()
    
    def _center_window(self) -> None:
        self.update_idletasks()
        x = self.parent.winfo_x() + (self.parent.winfo_width() - self.winfo_width()) // 2
        y = self.parent.winfo_y() + (self.parent.winfo_height() - self.winfo_height()) // 2
        self.geometry(f"+{x}+{y}")
    
    def _create_widgets(self) -> None:
        main_frame = ctk.CTkFrame(self, fg_color="transparent")
        main_frame.pack(fill="both", expand=True, padx=25, pady=25)
        
        # Buttons at BOTTOM - pack first to ensure always visible
        btn_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        btn_frame.pack(fill="x", side="bottom")
        
        save_text = "Save" if self.edit_command_id else "Add"
        ctk.CTkButton(btn_frame, text=save_text, command=self._save_command, width=100).pack(side="right", padx=(10, 0))
        ctk.CTkButton(btn_frame, text="Cancel", command=self.destroy, width=100,
                      fg_color="transparent", border_width=1).pack(side="right")
        
        # Phrases
        ctk.CTkLabel(main_frame, text="Voice Phrases", font=ctk.CTkFont(size=14, weight="bold")).pack(anchor="w", side="top", pady=(0, 8))
        
        self.phrases_text = ctk.CTkTextbox(main_frame, height=70, corner_radius=10)
        self.phrases_text.pack(fill="x", side="top", pady=(0, 5))
        
        help_text = "One phrase per line"
        if self.command_manager.spotify and self.command_manager.spotify.is_authenticated:
            help_text += " • Spotify phrases reserved"
        ctk.CTkLabel(main_frame, text=help_text, text_color="gray").pack(anchor="w", side="top", pady=(0, 15))
        
        # Type
        ctk.CTkLabel(main_frame, text="Command Type", font=ctk.CTkFont(size=14, weight="bold")).pack(anchor="w", side="top", pady=(0, 8))
        
        self.command_type = ctk.StringVar(value="open_file")
        
        # Type selection - use multiple rows for better layout
        type_frame1 = ctk.CTkFrame(main_frame, fg_color="transparent")
        type_frame1.pack(fill="x", side="top", pady=(0, 5))
        
        ctk.CTkRadioButton(type_frame1, text="Open File/App", variable=self.command_type, 
                           value="open_file", command=self._on_type_changed).pack(side="left", padx=(0, 20))
        ctk.CTkRadioButton(type_frame1, text="Open Folder", variable=self.command_type,
                           value="open_folder", command=self._on_type_changed).pack(side="left", padx=(0, 20))
        ctk.CTkRadioButton(type_frame1, text="Open Website", variable=self.command_type,
                           value="open_url", command=self._on_type_changed).pack(side="left", padx=(0, 20))
        ctk.CTkRadioButton(type_frame1, text="Keyboard Macro", variable=self.command_type,
                           value="macro", command=self._on_type_changed).pack(side="left")
        
        type_frame2 = ctk.CTkFrame(main_frame, fg_color="transparent")
        type_frame2.pack(fill="x", side="top", pady=(0, 15))
        
        ctk.CTkRadioButton(type_frame2, text="Window Action", variable=self.command_type,
                           value="window_action", command=self._on_type_changed).pack(side="left", padx=(0, 20))
        ctk.CTkRadioButton(type_frame2, text="Chain Actions", variable=self.command_type,
                           value="chain", command=self._on_type_changed).pack(side="left")
        
        # Content - takes remaining space
        self.content_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        self.content_frame.pack(fill="both", expand=True, side="top", pady=(0, 15))
        
        self._create_open_file_content()
    
    def _create_open_file_content(self) -> None:
        for w in self.content_frame.winfo_children():
            w.destroy()
        
        ctk.CTkLabel(self.content_frame, text="File Path").pack(anchor="w", pady=(0, 8))
        
        path_frame = ctk.CTkFrame(self.content_frame, fg_color="transparent")
        path_frame.pack(fill="x")
        
        self.path_var = ctk.StringVar()
        ctk.CTkEntry(path_frame, textvariable=self.path_var, width=320).pack(side="left")
        ctk.CTkButton(path_frame, text="Browse", command=self._browse_file, width=100).pack(side="left", padx=(10, 0))
    
    def _create_macro_content(self) -> None:
        for w in self.content_frame.winfo_children():
            w.destroy()
        
        # Initialize recording state
        self.is_recording = False
        self.held_keys = set()
        self.last_key_time = None
        self._keyboard_hook = None
        
        # Control buttons at BOTTOM (pack first with side="bottom" so they're always visible)
        btn_frame = ctk.CTkFrame(self.content_frame, fg_color="transparent")
        btn_frame.pack(fill="x", side="bottom", pady=(10, 0))
        
        # Row 1: Key and mouse actions
        btn_row1 = ctk.CTkFrame(btn_frame, fg_color="transparent")
        btn_row1.pack(fill="x", pady=(0, 6))
        
        ctk.CTkButton(btn_row1, text="+ Key", command=self._add_macro_step, width=70).pack(side="left", padx=(0, 4))
        ctk.CTkButton(btn_row1, text="+ Mouse", command=self._add_mouse_action, width=80,
                      fg_color="#2d5a2d", hover_color="#1e3d1e").pack(side="left", padx=(0, 4))
        ctk.CTkButton(btn_row1, text="Remove", command=self._remove_macro_step, width=70,
                      fg_color="transparent", border_width=1).pack(side="left", padx=(0, 4))
        ctk.CTkButton(btn_row1, text="↑", command=self._move_step_up, width=30,
                      fg_color="transparent", border_width=1).pack(side="left", padx=(0, 2))
        ctk.CTkButton(btn_row1, text="↓", command=self._move_step_down, width=30,
                      fg_color="transparent", border_width=1).pack(side="left", padx=(0, 4))
        ctk.CTkButton(btn_row1, text="Clear All", command=self._clear_macro_steps, width=75,
                      fg_color="#aa3333", hover_color="#882222").pack(side="right")
        
        # Recording section at TOP
        record_frame = ctk.CTkFrame(self.content_frame, corner_radius=10)
        record_frame.pack(fill="x", side="top", pady=(0, 8))
        
        record_inner = ctk.CTkFrame(record_frame, fg_color="transparent")
        record_inner.pack(fill="x", padx=12, pady=10)
        
        record_top = ctk.CTkFrame(record_inner, fg_color="transparent")
        record_top.pack(fill="x")
        
        self.record_btn_text = ctk.StringVar(value="🔴 Record")
        self.record_btn = ctk.CTkButton(record_top, textvariable=self.record_btn_text,
                                         command=self._toggle_recording, width=100)
        self.record_btn.pack(side="left", padx=(0, 10))
        
        self.record_status = ctk.CTkLabel(record_top, text="Press keys while recording", 
                                           text_color="gray", font=ctk.CTkFont(size=12))
        self.record_status.pack(side="left")
        
        # Default delay settings
        delay_frame = ctk.CTkFrame(record_inner, fg_color="transparent")
        delay_frame.pack(fill="x", pady=(8, 0))
        
        ctk.CTkLabel(delay_frame, text="Delay:", font=ctk.CTkFont(size=12)).pack(side="left")
        self.default_delay_var = ctk.StringVar(value="0.1")
        ctk.CTkEntry(delay_frame, textvariable=self.default_delay_var, width=50).pack(side="left", padx=(5, 3))
        ctk.CTkLabel(delay_frame, text="s", font=ctk.CTkFont(size=12)).pack(side="left")
        
        ctk.CTkButton(delay_frame, text="Apply to All", command=self._apply_default_delay,
                      width=90, height=28, fg_color="transparent", border_width=1).pack(side="left", padx=(10, 0))
        
        # Macro steps label
        ctk.CTkLabel(self.content_frame, text="Macro Steps (double-click to edit)", 
                     font=ctk.CTkFont(size=12)).pack(anchor="w", side="top", pady=(0, 5))
        
        # Listbox frame - takes remaining space
        list_frame = ctk.CTkFrame(self.content_frame, corner_radius=10)
        list_frame.pack(fill="both", expand=True, side="top")
        
        self.macro_listbox = tk.Listbox(list_frame, font=("Segoe UI", 11), height=4,
                                         bg="#2b2b2b", fg="white", selectbackground="#1f6aa5",
                                         highlightthickness=0, bd=0)
        scrollbar = ctk.CTkScrollbar(list_frame, command=self.macro_listbox.yview)
        self.macro_listbox.configure(yscrollcommand=scrollbar.set)
        
        self.macro_listbox.pack(side="left", fill="both", expand=True, padx=5, pady=5)
        scrollbar.pack(side="right", fill="y", padx=(0, 5), pady=5)
        
        # Bind double-click to edit step
        self.macro_listbox.bind("<Double-1>", self._edit_macro_step)
    
    def _toggle_recording(self) -> None:
        """Toggle keyboard recording mode."""
        if self.is_recording:
            self._stop_recording()
        else:
            self._start_recording()
    
    def _start_recording(self) -> None:
        """Start recording keyboard inputs."""
        self.is_recording = True
        self.held_keys = set()
        self.last_key_time = None
        self.record_btn_text.set("⏹️ Stop Recording")
        self.record_status.configure(text="Recording... Press keys now!", text_color="#ff4444")
        
        import time as time_module
        
        def on_key_event(event):
            if not self.is_recording:
                return
            
            key_name = event.name.lower() if event.name else ""
            scan_code = event.scan_code
            
            if event.event_type == "down":
                # Check if this is a new key press (not auto-repeat)
                key_id = (key_name, scan_code)
                if key_id in self.held_keys:
                    return  # Skip auto-repeat
                self.held_keys.add(key_id)
                
                current_time = time_module.time()
                
                # Calculate delay from last key
                if self.last_key_time is not None:
                    delay = round(current_time - self.last_key_time, 2)
                    delay = min(delay, 5.0)  # Cap at 5 seconds
                else:
                    try:
                        delay = float(self.default_delay_var.get())
                    except ValueError:
                        delay = 0.1
                
                self.last_key_time = current_time
                
                # Modifier keys to skip as standalone
                modifier_keys = [
                    "ctrl", "left ctrl", "right ctrl",
                    "alt", "left alt", "right alt", "altgr", "alt gr",
                    "shift", "left shift", "right shift",
                    "left windows", "right windows"
                ]
                
                if key_name in modifier_keys:
                    return
                
                # Build key string with modifiers
                key_parts = []
                
                if any(k[0] in ["ctrl", "left ctrl", "right ctrl"] for k in self.held_keys):
                    key_parts.append("ctrl")
                if any(k[0] in ["alt", "left alt"] for k in self.held_keys):
                    key_parts.append("alt")
                if any(k[0] in ["right alt", "altgr", "alt gr"] for k in self.held_keys):
                    key_parts.append("altgr")
                if any(k[0] in ["shift", "left shift", "right shift"] for k in self.held_keys):
                    key_parts.append("shift")
                if any(k[0] in ["left windows", "right windows"] for k in self.held_keys):
                    key_parts.append("win")
                
                key_parts.append(key_name)
                key_str = "+".join(key_parts)
                
                # Add to macro steps with type
                self.macro_steps.append({"type": "key", "key": key_str, "delay": delay})
                
                # Update UI on main thread
                self.after(0, self._refresh_macro_list)
            
            elif event.event_type == "up":
                key_id = (key_name, scan_code)
                self.held_keys.discard(key_id)
        
        self._keyboard_hook = keyboard.hook(on_key_event)
    
    def _stop_recording(self) -> None:
        """Stop recording keyboard inputs."""
        self.is_recording = False
        self.record_btn_text.set("🔴 Start Recording")
        self.record_status.configure(text="Recording stopped", text_color="#44aa44")
        
        if self._keyboard_hook:
            keyboard.unhook(self._keyboard_hook)
            self._keyboard_hook = None
        
        # Set last step's delay to default
        if self.macro_steps:
            try:
                self.macro_steps[-1]["delay"] = float(self.default_delay_var.get())
            except ValueError:
                self.macro_steps[-1]["delay"] = 0.1
            self._refresh_macro_list()
    
    def _apply_default_delay(self) -> None:
        """Apply default delay to all steps."""
        try:
            delay = float(self.default_delay_var.get())
            if delay < 0:
                raise ValueError()
        except ValueError:
            CTkMessagebox.show(self, "Warning", "Enter a valid delay (0 or greater).", "warning")
            return
        
        for step in self.macro_steps:
            step["delay"] = delay
        self._refresh_macro_list()
    
    def _edit_step_delay(self, event) -> None:
        """Edit the delay of a step when double-clicked."""
        selection = self.macro_listbox.curselection()
        if not selection:
            return
        
        idx = selection[0]
        current_delay = self.macro_steps[idx]["delay"]
        
        dialog = ctk.CTkToplevel(self)
        dialog.title("Edit Delay")
        dialog.geometry("250x150")
        dialog.transient(self)
        dialog.grab_set()
        
        frame = ctk.CTkFrame(dialog, fg_color="transparent")
        frame.pack(fill="both", expand=True, padx=20, pady=20)
        
        ctk.CTkLabel(frame, text="Delay (seconds):").pack(anchor="w")
        delay_var = ctk.StringVar(value=str(current_delay))
        entry = ctk.CTkEntry(frame, textvariable=delay_var, width=200)
        entry.pack(fill="x", pady=(5, 20))
        entry.focus_set()
        entry.select_range(0, tk.END)
        
        def save():
            try:
                new_delay = float(delay_var.get())
                if new_delay >= 0:
                    self.macro_steps[idx]["delay"] = new_delay
                    self._refresh_macro_list()
            except ValueError:
                pass
            dialog.destroy()
        
        def on_enter(e):
            save()
        
        entry.bind("<Return>", on_enter)
        ctk.CTkButton(frame, text="OK", command=save, width=80).pack(side="right")
        
        dialog.update_idletasks()
        x = self.winfo_x() + (self.winfo_width() - 250) // 2
        y = self.winfo_y() + (self.winfo_height() - 150) // 2
        dialog.geometry(f"+{x}+{y}")
    
    def _refresh_macro_list(self) -> None:
        """Refresh the macro listbox with current steps."""
        self.macro_listbox.delete(0, tk.END)
        for step in self.macro_steps:
            step_type = step.get("type", "key")
            delay = step.get("delay", 0.1)
            
            if step_type == "key":
                # Keyboard step (legacy format support)
                key = step.get("key", "")
                display = f"⌨️ {key} ({delay}s)"
            elif step_type == "mouse_move":
                x, y = step.get("x", 0), step.get("y", 0)
                display = f"🖱️ Move to ({x}, {y}) ({delay}s)"
            elif step_type == "mouse_click":
                x, y = step.get("x"), step.get("y")
                button = step.get("button", "left")
                clicks = step.get("clicks", 1)
                click_type = "Double-click" if clicks == 2 else "Click"
                if x is not None and y is not None:
                    display = f"🖱️ {click_type} {button} at ({x}, {y}) ({delay}s)"
                else:
                    display = f"🖱️ {click_type} {button} (current pos) ({delay}s)"
            elif step_type == "mouse_scroll":
                amount = step.get("amount", 3)
                direction = "up" if amount > 0 else "down"
                display = f"🖱️ Scroll {direction} {abs(amount)} ({delay}s)"
            else:
                display = f"? Unknown step ({delay}s)"
            
            self.macro_listbox.insert(tk.END, display)
    
    def _edit_macro_step(self, event) -> None:
        """Edit a macro step when double-clicked."""
        selection = self.macro_listbox.curselection()
        if not selection:
            return
        
        idx = selection[0]
        step = self.macro_steps[idx]
        step_type = step.get("type", "key")
        
        if step_type == "key":
            self._edit_key_step(idx, step)
        elif step_type in ("mouse_move", "mouse_click"):
            self._edit_mouse_step(idx, step)
        elif step_type == "mouse_scroll":
            self._edit_scroll_step(idx, step)
    
    def _edit_key_step(self, idx: int, step: dict) -> None:
        """Edit a keyboard step."""
        dialog = ctk.CTkToplevel(self)
        dialog.title("Edit Key Step")
        dialog.geometry("300x200")
        dialog.transient(self)
        dialog.grab_set()
        
        frame = ctk.CTkFrame(dialog, fg_color="transparent")
        frame.pack(fill="both", expand=True, padx=20, pady=20)
        
        ctk.CTkLabel(frame, text="Key/Combo:").pack(anchor="w")
        key_var = ctk.StringVar(value=step.get("key", ""))
        key_entry = ctk.CTkEntry(frame, textvariable=key_var, width=260)
        key_entry.pack(fill="x", pady=(5, 15))
        
        ctk.CTkLabel(frame, text="Delay (seconds):").pack(anchor="w")
        delay_var = ctk.StringVar(value=str(step.get("delay", 0.1)))
        ctk.CTkEntry(frame, textvariable=delay_var, width=260).pack(fill="x", pady=(5, 20))
        
        def save():
            key = key_var.get().strip()
            try:
                delay = float(delay_var.get())
            except ValueError:
                delay = 0.1
            if key:
                self.macro_steps[idx] = {"type": "key", "key": key, "delay": delay}
                self._refresh_macro_list()
            dialog.destroy()
        
        ctk.CTkButton(frame, text="Save", command=save, width=100).pack(side="right")
        
        dialog.update_idletasks()
        x = self.winfo_x() + (self.winfo_width() - 300) // 2
        y = self.winfo_y() + (self.winfo_height() - 200) // 2
        dialog.geometry(f"+{x}+{y}")
    
    def _edit_mouse_step(self, idx: int, step: dict) -> None:
        """Edit a mouse move or click step."""
        step_type = step.get("type", "mouse_click")
        
        dialog = ctk.CTkToplevel(self)
        dialog.title("Edit Mouse Step")
        dialog.geometry("350x350")
        dialog.transient(self)
        dialog.grab_set()
        
        frame = ctk.CTkFrame(dialog, fg_color="transparent")
        frame.pack(fill="both", expand=True, padx=20, pady=20)
        
        # Action type
        ctk.CTkLabel(frame, text="Action:").pack(anchor="w")
        action_var = ctk.StringVar(value=step_type)
        action_menu = ctk.CTkOptionMenu(frame, variable=action_var,
                                         values=["mouse_move", "mouse_click"],
                                         width=200)
        action_menu.pack(anchor="w", pady=(5, 15))
        
        # Coordinates
        coord_frame = ctk.CTkFrame(frame, fg_color="transparent")
        coord_frame.pack(fill="x", pady=(0, 10))
        
        ctk.CTkLabel(coord_frame, text="X:").pack(side="left")
        x_var = ctk.StringVar(value=str(step.get("x", "")))
        ctk.CTkEntry(coord_frame, textvariable=x_var, width=80).pack(side="left", padx=(5, 15))
        
        ctk.CTkLabel(coord_frame, text="Y:").pack(side="left")
        y_var = ctk.StringVar(value=str(step.get("y", "")))
        ctk.CTkEntry(coord_frame, textvariable=y_var, width=80).pack(side="left", padx=5)
        
        # Record button
        def record_position():
            dialog.withdraw()
            self.withdraw()
            self.after(200, lambda: self._capture_mouse_position(dialog, x_var, y_var))
        
        ctk.CTkButton(frame, text="📍 Record Position", command=record_position,
                      width=150).pack(anchor="w", pady=(0, 15))
        
        # Click options (only for mouse_click)
        click_frame = ctk.CTkFrame(frame, fg_color="transparent")
        click_frame.pack(fill="x", pady=(0, 10))
        
        ctk.CTkLabel(click_frame, text="Button:").pack(side="left")
        button_var = ctk.StringVar(value=step.get("button", "left"))
        ctk.CTkOptionMenu(click_frame, variable=button_var,
                          values=["left", "right", "middle"], width=100).pack(side="left", padx=(5, 15))
        
        ctk.CTkLabel(click_frame, text="Clicks:").pack(side="left")
        clicks_var = ctk.StringVar(value=str(step.get("clicks", 1)))
        ctk.CTkOptionMenu(click_frame, variable=clicks_var,
                          values=["1", "2"], width=60).pack(side="left", padx=5)
        
        # Delay
        ctk.CTkLabel(frame, text="Delay (seconds):").pack(anchor="w")
        delay_var = ctk.StringVar(value=str(step.get("delay", 0.1)))
        ctk.CTkEntry(frame, textvariable=delay_var, width=100).pack(anchor="w", pady=(5, 15))
        
        def save():
            try:
                x = int(x_var.get()) if x_var.get().strip() else None
                y = int(y_var.get()) if y_var.get().strip() else None
                delay = float(delay_var.get())
            except ValueError:
                CTkMessagebox.show(dialog, "Error", "Invalid coordinates or delay value.", "warning")
                return
            
            new_step = {
                "type": action_var.get(),
                "delay": delay
            }
            
            if x is not None and y is not None:
                new_step["x"] = x
                new_step["y"] = y
            
            if action_var.get() == "mouse_click":
                new_step["button"] = button_var.get()
                new_step["clicks"] = int(clicks_var.get())
            
            self.macro_steps[idx] = new_step
            self._refresh_macro_list()
            dialog.destroy()
        
        ctk.CTkButton(frame, text="Save", command=save, width=100).pack(side="right")
        
        dialog.update_idletasks()
        x = self.winfo_x() + (self.winfo_width() - 350) // 2
        y = self.winfo_y() + (self.winfo_height() - 350) // 2
        dialog.geometry(f"+{x}+{y}")
    
    def _edit_scroll_step(self, idx: int, step: dict) -> None:
        """Edit a mouse scroll step."""
        dialog = ctk.CTkToplevel(self)
        dialog.title("Edit Scroll Step")
        dialog.geometry("300x200")
        dialog.transient(self)
        dialog.grab_set()
        
        frame = ctk.CTkFrame(dialog, fg_color="transparent")
        frame.pack(fill="both", expand=True, padx=20, pady=20)
        
        ctk.CTkLabel(frame, text="Scroll amount (positive=up, negative=down):").pack(anchor="w")
        amount_var = ctk.StringVar(value=str(step.get("amount", 3)))
        ctk.CTkEntry(frame, textvariable=amount_var, width=100).pack(anchor="w", pady=(5, 15))
        
        ctk.CTkLabel(frame, text="Delay (seconds):").pack(anchor="w")
        delay_var = ctk.StringVar(value=str(step.get("delay", 0.1)))
        ctk.CTkEntry(frame, textvariable=delay_var, width=100).pack(anchor="w", pady=(5, 20))
        
        def save():
            try:
                amount = int(amount_var.get())
                delay = float(delay_var.get())
            except ValueError:
                CTkMessagebox.show(dialog, "Error", "Invalid values.", "warning")
                return
            
            self.macro_steps[idx] = {
                "type": "mouse_scroll",
                "amount": amount,
                "delay": delay
            }
            self._refresh_macro_list()
            dialog.destroy()
        
        ctk.CTkButton(frame, text="Save", command=save, width=100).pack(side="right")
        
        dialog.update_idletasks()
        x = self.winfo_x() + (self.winfo_width() - 300) // 2
        y = self.winfo_y() + (self.winfo_height() - 200) // 2
        dialog.geometry(f"+{x}+{y}")
    
    def _add_mouse_action(self) -> None:
        """Add a mouse action to the macro."""
        dialog = ctk.CTkToplevel(self)
        dialog.title("Add Mouse Action")
        dialog.geometry("380x460")
        dialog.transient(self)
        dialog.grab_set()
        set_window_icon(dialog)
        
        frame = ctk.CTkFrame(dialog, fg_color="transparent")
        frame.pack(fill="both", expand=True, padx=20, pady=20)

        # Pack buttons FIRST with side="bottom" so they always stay visible
        btn_frame = ctk.CTkFrame(frame, fg_color="transparent")
        btn_frame.pack(fill="x", side="bottom", pady=(8, 0))
        
        # Action type
        ctk.CTkLabel(frame, text="Action Type", font=ctk.CTkFont(size=13, weight="bold")).pack(anchor="w")
        action_var = ctk.StringVar(value="mouse_click")
        
        action_frame = ctk.CTkFrame(frame, fg_color="transparent")
        action_frame.pack(fill="x", pady=(5, 15))
        
        ctk.CTkRadioButton(action_frame, text="Move", variable=action_var,
                           value="mouse_move").pack(side="left", padx=(0, 15))
        ctk.CTkRadioButton(action_frame, text="Click", variable=action_var,
                           value="mouse_click").pack(side="left", padx=(0, 15))
        ctk.CTkRadioButton(action_frame, text="Scroll", variable=action_var,
                           value="mouse_scroll").pack(side="left")
        
        # Position section
        pos_frame = ctk.CTkFrame(frame, corner_radius=10)
        pos_frame.pack(fill="x", pady=(0, 15))
        
        pos_inner = ctk.CTkFrame(pos_frame, fg_color="transparent")
        pos_inner.pack(fill="x", padx=15, pady=15)
        
        ctk.CTkLabel(pos_inner, text="Position", font=ctk.CTkFont(size=12, weight="bold")).pack(anchor="w")
        
        coord_frame = ctk.CTkFrame(pos_inner, fg_color="transparent")
        coord_frame.pack(fill="x", pady=(8, 10))
        
        ctk.CTkLabel(coord_frame, text="X:").pack(side="left")
        x_var = ctk.StringVar()
        x_entry = ctk.CTkEntry(coord_frame, textvariable=x_var, width=80, placeholder_text="0")
        x_entry.pack(side="left", padx=(5, 20))
        
        ctk.CTkLabel(coord_frame, text="Y:").pack(side="left")
        y_var = ctk.StringVar()
        y_entry = ctk.CTkEntry(coord_frame, textvariable=y_var, width=80, placeholder_text="0")
        y_entry.pack(side="left", padx=5)
        
        # Record button with explanation
        def record_position():
            dialog.withdraw()
            self.withdraw()
            # Small delay to let windows hide
            self.after(200, lambda: self._capture_mouse_position(dialog, x_var, y_var))
        
        record_frame = ctk.CTkFrame(pos_inner, fg_color="transparent")
        record_frame.pack(fill="x", pady=(5, 0))
        
        ctk.CTkButton(record_frame, text="📍 Record Position", command=record_position,
                      width=140, fg_color="#2d5a2d", hover_color="#1e3d1e").pack(side="left")
        ctk.CTkLabel(record_frame, text="Click to capture, Esc to cancel", 
                     text_color="gray", font=ctk.CTkFont(size=11)).pack(side="left", padx=(10, 0))
        
        # Click options
        click_frame = ctk.CTkFrame(frame, corner_radius=10)
        click_frame.pack(fill="x", pady=(0, 15))
        
        click_inner = ctk.CTkFrame(click_frame, fg_color="transparent")
        click_inner.pack(fill="x", padx=15, pady=15)
        
        ctk.CTkLabel(click_inner, text="Click Options", font=ctk.CTkFont(size=12, weight="bold")).pack(anchor="w")
        
        click_opts = ctk.CTkFrame(click_inner, fg_color="transparent")
        click_opts.pack(fill="x", pady=(8, 0))
        
        ctk.CTkLabel(click_opts, text="Button:").pack(side="left")
        button_var = ctk.StringVar(value="left")
        ctk.CTkOptionMenu(click_opts, variable=button_var,
                          values=["left", "right", "middle"], width=90).pack(side="left", padx=(5, 20))
        
        ctk.CTkLabel(click_opts, text="Clicks:").pack(side="left")
        clicks_var = ctk.StringVar(value="1")
        ctk.CTkOptionMenu(click_opts, variable=clicks_var,
                          values=["1", "2"], width=60).pack(side="left", padx=5)
        
        # Scroll options
        scroll_frame = ctk.CTkFrame(click_inner, fg_color="transparent")
        scroll_frame.pack(fill="x", pady=(10, 0))
        
        ctk.CTkLabel(scroll_frame, text="Scroll:").pack(side="left")
        scroll_var = ctk.StringVar(value="3")
        ctk.CTkEntry(scroll_frame, textvariable=scroll_var, width=60).pack(side="left", padx=(5, 5))
        ctk.CTkLabel(scroll_frame, text="(+up, -down)", text_color="gray",
                     font=ctk.CTkFont(size=11)).pack(side="left")
        
        # Delay
        delay_frame = ctk.CTkFrame(frame, fg_color="transparent")
        delay_frame.pack(fill="x", pady=(0, 15))
        
        ctk.CTkLabel(delay_frame, text="Delay:").pack(side="left")
        delay_var = ctk.StringVar(value="0.1")
        ctk.CTkEntry(delay_frame, textvariable=delay_var, width=60).pack(side="left", padx=(5, 5))
        ctk.CTkLabel(delay_frame, text="seconds").pack(side="left")
        
        def add():
            action = action_var.get()
            try:
                delay = float(delay_var.get())
            except ValueError:
                delay = 0.1
            
            step = {"type": action, "delay": delay}
            
            # Parse coordinates
            x_str, y_str = x_var.get().strip(), y_var.get().strip()
            if x_str and y_str:
                try:
                    step["x"] = int(x_str)
                    step["y"] = int(y_str)
                except ValueError:
                    CTkMessagebox.show(dialog, "Error", "Invalid X or Y coordinate.", "warning")
                    return
            elif action == "mouse_move":
                CTkMessagebox.show(dialog, "Error", "Move action requires coordinates.", "warning")
                return
            
            if action == "mouse_click":
                step["button"] = button_var.get()
                step["clicks"] = int(clicks_var.get())
            elif action == "mouse_scroll":
                try:
                    step["amount"] = int(scroll_var.get())
                except ValueError:
                    step["amount"] = 3
            
            self.macro_steps.append(step)
            self._refresh_macro_list()
            dialog.destroy()
        
        ctk.CTkButton(btn_frame, text="Add", command=add, width=100).pack(side="right")
        ctk.CTkButton(btn_frame, text="Cancel", command=dialog.destroy, width=100,
                      fg_color="transparent", border_width=1).pack(side="right", padx=(0, 10))
        
        dialog.update_idletasks()
        x = self.winfo_x() + (self.winfo_width() - 380) // 2
        y = self.winfo_y() + (self.winfo_height() - 460) // 2
        dialog.geometry(f"+{x}+{y}")
    
    def _capture_mouse_position(self, parent_dialog, x_var, y_var) -> None:
        """Capture mouse position using a fullscreen overlay. Escape cancels."""
        overlay = tk.Toplevel(self)
        self._mouse_capture_overlay = overlay
        overlay.overrideredirect(True)
        overlay.attributes("-topmost", True)

        # Cover the primary screen so clicks are intercepted and do not reach apps beneath.
        screen_w = self.winfo_screenwidth()
        screen_h = self.winfo_screenheight()
        overlay.geometry(f"{screen_w}x{screen_h}+0+0")
        overlay.configure(bg="#000000")

        try:
            overlay.attributes("-alpha", 0.12)
        except Exception:
            pass

        message = tk.Label(
            overlay,
            text="Click to capture position   Esc to cancel",
            fg="white",
            bg="#000000",
            font=("Segoe UI", 13, "bold")
        )
        message.place(relx=0.5, rely=0.5, anchor="center")

        finished = {"done": False}

        def restore(cancelled: bool = False, event=None) -> None:
            if finished["done"]:
                return
            finished["done"] = True

            try:
                overlay.grab_release()
            except Exception:
                pass

            try:
                overlay.destroy()
            except Exception:
                pass

            if getattr(self, "_mouse_capture_overlay", None) is overlay:
                self._mouse_capture_overlay = None

            self.deiconify()
            parent_dialog.deiconify()
            parent_dialog.lift()
            parent_dialog.focus_force()

            if not cancelled and event is not None:
                x_var.set(str(event.x_root))
                y_var.set(str(event.y_root))

        overlay.bind("<Button-1>", lambda e: restore(False, e))
        overlay.bind("<Escape>", lambda e: restore(True))
        overlay.bind("<Button-3>", lambda e: "break")
        overlay.bind("<Button-2>", lambda e: "break")

        overlay.focus_force()
        overlay.grab_set()

        # Safety net: if overlay somehow loses focus/gets orphaned, do not keep input blocked.
        overlay.bind("<Destroy>", lambda e: setattr(self, "_mouse_capture_overlay", None))

    def _move_step_up(self) -> None:
        """Move selected step up."""
        selection = self.macro_listbox.curselection()
        if selection and selection[0] > 0:
            idx = selection[0]
            self.macro_steps[idx], self.macro_steps[idx-1] = self.macro_steps[idx-1], self.macro_steps[idx]
            self._refresh_macro_list()
            self.macro_listbox.selection_set(idx - 1)
    
    def _move_step_down(self) -> None:
        """Move selected step down."""
        selection = self.macro_listbox.curselection()
        if selection and selection[0] < len(self.macro_steps) - 1:
            idx = selection[0]
            self.macro_steps[idx], self.macro_steps[idx+1] = self.macro_steps[idx+1], self.macro_steps[idx]
            self._refresh_macro_list()
            self.macro_listbox.selection_set(idx + 1)
    
    def _clear_macro_steps(self) -> None:
        """Clear all macro steps."""
        self.macro_steps.clear()
        self._refresh_macro_list()
    
    def _on_type_changed(self) -> None:
        cmd_type = self.command_type.get()
        if cmd_type == "open_file":
            self._create_open_file_content()
        elif cmd_type == "open_folder":
            self._create_open_folder_content()
        elif cmd_type == "open_url":
            self._create_open_url_content()
        elif cmd_type == "window_action":
            self._create_window_action_content()
        elif cmd_type == "chain":
            self._create_chain_content()
        else:
            self._create_macro_content()
    
    def _create_open_folder_content(self) -> None:
        for w in self.content_frame.winfo_children():
            w.destroy()
        
        ctk.CTkLabel(self.content_frame, text="Folder Path").pack(anchor="w", pady=(0, 8))
        
        path_frame = ctk.CTkFrame(self.content_frame, fg_color="transparent")
        path_frame.pack(fill="x")
        
        self.folder_var = ctk.StringVar()
        ctk.CTkEntry(path_frame, textvariable=self.folder_var, width=320).pack(side="left")
        ctk.CTkButton(path_frame, text="Browse", command=self._browse_folder, width=100).pack(side="left", padx=(10, 0))

    def _create_open_url_content(self) -> None:
        for w in self.content_frame.winfo_children():
            w.destroy()
        
        ctk.CTkLabel(self.content_frame, text="Website URL").pack(anchor="w", pady=(0, 8))
        
        self.url_var = ctk.StringVar()
        url_entry = ctk.CTkEntry(self.content_frame, textvariable=self.url_var, 
                                  placeholder_text="https://example.com", width=400)
        url_entry.pack(fill="x", pady=(0, 10))
        
        ctk.CTkLabel(self.content_frame, text="💡 Enter the full URL including https://", 
                     text_color="gray", font=ctk.CTkFont(size=11)).pack(anchor="w")
    
    def _create_window_action_content(self) -> None:
        for w in self.content_frame.winfo_children():
            w.destroy()
        
        # Target selection
        ctk.CTkLabel(self.content_frame, text="Target Window", font=ctk.CTkFont(size=13, weight="bold")).pack(anchor="w", pady=(0, 8))
        
        self.window_target_var = ctk.StringVar(value="focused")
        
        target_frame = ctk.CTkFrame(self.content_frame, fg_color="transparent")
        target_frame.pack(fill="x", pady=(0, 5))
        
        ctk.CTkRadioButton(target_frame, text="Currently Focused", variable=self.window_target_var,
                           value="focused", command=self._on_window_target_changed).pack(side="left", padx=(0, 15))
        ctk.CTkRadioButton(target_frame, text="By App Name", variable=self.window_target_var,
                           value="app", command=self._on_window_target_changed).pack(side="left", padx=(0, 15))
        ctk.CTkRadioButton(target_frame, text="By Window Title", variable=self.window_target_var,
                           value="title", command=self._on_window_target_changed).pack(side="left")
        
        # App selection (for app targeting)
        self.app_select_frame = ctk.CTkFrame(self.content_frame, fg_color="transparent")
        self.app_select_frame.pack(fill="x", pady=(0, 10))
        
        self.target_app_var = ctk.StringVar()
        ctk.CTkLabel(self.app_select_frame, text="App:").pack(side="left")
        self.app_entry = ctk.CTkEntry(self.app_select_frame, textvariable=self.target_app_var,
                                       placeholder_text="e.g., chrome, discord, code", width=180)
        self.app_entry.pack(side="left", padx=(5, 10))
        ctk.CTkButton(self.app_select_frame, text="Pick Running App", 
                      command=self._pick_running_app, width=120).pack(side="left")
        self.app_select_frame.pack_forget()  # Hidden by default
        
        # Title selection (for window title targeting)
        self.title_select_frame = ctk.CTkFrame(self.content_frame, fg_color="transparent")
        self.title_select_frame.pack(fill="x", pady=(0, 10))
        
        self.target_title_var = ctk.StringVar()
        ctk.CTkLabel(self.title_select_frame, text="Title contains:").pack(side="left")
        self.title_entry = ctk.CTkEntry(self.title_select_frame, textvariable=self.target_title_var,
                                         placeholder_text="e.g., 'GitHub' or 'VS Code'", width=180)
        self.title_entry.pack(side="left", padx=(5, 10))
        ctk.CTkButton(self.title_select_frame, text="Pick Window", 
                      command=self._pick_window_by_title, width=100).pack(side="left")
        self.title_select_frame.pack_forget()  # Hidden by default
        
        # Action selection
        ctk.CTkLabel(self.content_frame, text="Action", font=ctk.CTkFont(size=13, weight="bold")).pack(anchor="w", pady=(5, 8))
        
        self.window_action_var = ctk.StringVar(value="minimize")
        
        actions = [
            ("Minimize", "minimize"),
            ("Maximize", "maximize"),
            ("Restore", "restore"),
            ("Focus", "focus"),
            ("Close", "close"),
            ("Close All (same app)", "close_all_app"),
            ("Close ALL (everything)", "close_all_windows"),
            ("Snap Left", "snap_left"),
            ("Snap Right", "snap_right"),
            ("Snap Top-Left", "snap_top_left"),
            ("Snap Top-Right", "snap_top_right"),
            ("Snap Bottom-Left", "snap_bottom_left"),
            ("Snap Bottom-Right", "snap_bottom_right"),
        ]
        
        # Use two columns for actions
        action_frame = ctk.CTkFrame(self.content_frame, fg_color="transparent")
        action_frame.pack(fill="x", pady=(0, 5))
        
        left_col = ctk.CTkFrame(action_frame, fg_color="transparent")
        left_col.pack(side="left", fill="y")
        right_col = ctk.CTkFrame(action_frame, fg_color="transparent")
        right_col.pack(side="left", fill="y", padx=(20, 0))
        
        for i, (text, value) in enumerate(actions):
            col = left_col if i < 6 else right_col
            ctk.CTkRadioButton(col, text=text, variable=self.window_action_var,
                               value=value).pack(anchor="w", pady=1)
        
        ctk.CTkLabel(self.content_frame, text="💡 Use 'By Window Title' to snap specific windows (e.g., snap Chrome to left, VS Code to right)",
                     text_color="gray", font=ctk.CTkFont(size=11), wraplength=400).pack(anchor="w", pady=(5, 0))
    
    def _on_window_target_changed(self) -> None:
        """Show/hide app/title selection based on target choice."""
        target = self.window_target_var.get()
        
        # Hide both frames first
        self.app_select_frame.pack_forget()
        self.title_select_frame.pack_forget()
        
        # Show appropriate frame
        if target == "app":
            self.app_select_frame.pack(fill="x", pady=(0, 10), after=self.content_frame.winfo_children()[1])
        elif target == "title":
            self.title_select_frame.pack(fill="x", pady=(0, 10), after=self.content_frame.winfo_children()[1])
    
    def _pick_running_app(self) -> None:
        """Show dialog to pick from currently running apps."""
        import ctypes
        from ctypes import wintypes
        
        user32 = ctypes.windll.user32
        our_pid = os.getpid()
        
        apps = {}
        
        def enum_callback(hwnd, lparam):
            if user32.IsWindowVisible(hwnd):
                length = user32.GetWindowTextLengthW(hwnd)
                if length > 0:
                    window_pid = wintypes.DWORD()
                    user32.GetWindowThreadProcessId(hwnd, ctypes.byref(window_pid))
                    if window_pid.value != our_pid:
                        buff = ctypes.create_unicode_buffer(length + 1)
                        user32.GetWindowTextW(hwnd, buff, length + 1)
                        title = buff.value
                        if title and title not in ["Program Manager"]:
                            # Get process name
                            try:
                                import psutil
                                proc = psutil.Process(window_pid.value)
                                proc_name = proc.name().replace('.exe', '')
                                if proc_name not in apps:
                                    apps[proc_name] = title
                            except:
                                pass
            return True
        
        WNDENUMPROC = ctypes.WINFUNCTYPE(ctypes.c_bool, wintypes.HWND, wintypes.LPARAM)
        user32.EnumWindows(WNDENUMPROC(enum_callback), 0)
        
        if not apps:
            CTkMessagebox.show(self, "Info", "No running apps found.", "info")
            return
        
        # Create selection dialog
        dialog = ctk.CTkToplevel(self)
        dialog.title("Select Running App")
        dialog.geometry("400x350")
        dialog.transient(self)
        dialog.grab_set()
        set_window_icon(dialog)
        
        dialog.update_idletasks()
        x = self.winfo_x() + (self.winfo_width() - 400) // 2
        y = self.winfo_y() + (self.winfo_height() - 350) // 2
        dialog.geometry(f"+{x}+{y}")
        
        frame = ctk.CTkFrame(dialog, fg_color="transparent")
        frame.pack(fill="both", expand=True, padx=15, pady=15)
        
        ctk.CTkLabel(frame, text="Select an app:", font=ctk.CTkFont(size=13)).pack(anchor="w", pady=(0, 10))
        
        list_frame = ctk.CTkScrollableFrame(frame, height=220)
        list_frame.pack(fill="both", expand=True, pady=(0, 15))
        
        selected_app = ctk.StringVar()
        
        for proc_name, window_title in sorted(apps.items()):
            display = f"{proc_name} - {window_title[:40]}..." if len(window_title) > 40 else f"{proc_name} - {window_title}"
            ctk.CTkRadioButton(list_frame, text=display, variable=selected_app, 
                               value=proc_name).pack(anchor="w", pady=2)
        
        def select():
            if selected_app.get():
                self.target_app_var.set(selected_app.get())
            dialog.destroy()
        
        ctk.CTkButton(frame, text="Select", command=select, width=100).pack(side="right")
        ctk.CTkButton(frame, text="Cancel", command=dialog.destroy, width=100,
                      fg_color="transparent", border_width=1).pack(side="right", padx=(0, 10))
    
    def _pick_window_by_title(self) -> None:
        """Show dialog to pick from currently open windows by title."""
        from voice_control import CommandManager
        
        # Get open windows
        windows = CommandManager.get_open_windows()
        
        if not windows:
            CTkMessagebox.show(self, "Info", "No open windows found.", "info")
            return
        
        # Create selection dialog
        dialog = ctk.CTkToplevel(self)
        dialog.title("Select Window")
        dialog.geometry("500x400")
        dialog.transient(self)
        dialog.grab_set()
        set_window_icon(dialog)
        
        dialog.update_idletasks()
        x = self.winfo_x() + (self.winfo_width() - 500) // 2
        y = self.winfo_y() + (self.winfo_height() - 400) // 2
        dialog.geometry(f"+{x}+{y}")
        
        frame = ctk.CTkFrame(dialog, fg_color="transparent")
        frame.pack(fill="both", expand=True, padx=15, pady=15)
        
        ctk.CTkLabel(frame, text="Select a window (title will be used for matching):", 
                     font=ctk.CTkFont(size=13)).pack(anchor="w", pady=(0, 10))
        
        list_frame = ctk.CTkScrollableFrame(frame, height=260)
        list_frame.pack(fill="both", expand=True, pady=(0, 15))
        
        selected_title = ctk.StringVar()
        
        for window in windows:
            title = window["title"]
            process = window["process"]
            # Truncate long titles
            display_title = title[:50] + "..." if len(title) > 50 else title
            display = f"[{process}] {display_title}"
            
            ctk.CTkRadioButton(list_frame, text=display, variable=selected_title, 
                               value=title, font=ctk.CTkFont(size=11)).pack(anchor="w", pady=2)
        
        def select():
            if selected_title.get():
                # Use a short unique part of the title
                title = selected_title.get()
                # Try to extract a meaningful short portion
                if len(title) > 30:
                    self.target_title_var.set(title[:30])
                else:
                    self.target_title_var.set(title)
            dialog.destroy()
        
        ctk.CTkButton(frame, text="Select", command=select, width=100).pack(side="right")
        ctk.CTkButton(frame, text="Cancel", command=dialog.destroy, width=100,
                      fg_color="transparent", border_width=1).pack(side="right", padx=(0, 10))
    
    def _create_chain_content(self) -> None:
        for w in self.content_frame.winfo_children():
            w.destroy()
        
        self.chain_steps = []  # List of {"type": str, "data": any, "display": str}
        self.selected_chain_index = None  # Track selected step for editing/reordering
        
        ctk.CTkLabel(self.content_frame, text="Actions to Execute (in sequence)", 
                     font=ctk.CTkFont(size=13, weight="bold")).pack(anchor="w", pady=(0, 8))
        
        # List of chain steps
        list_frame = ctk.CTkFrame(self.content_frame)
        list_frame.pack(fill="both", expand=True, pady=(0, 10))
        
        self.chain_list = ctk.CTkScrollableFrame(list_frame, fg_color="transparent", height=100)
        self.chain_list.pack(fill="both", expand=True, padx=5, pady=5)
        
        # Buttons
        btn_frame = ctk.CTkFrame(self.content_frame, fg_color="transparent")
        btn_frame.pack(fill="x")
        
        ctk.CTkButton(btn_frame, text="+ Add Action", command=self._add_chain_action, width=100).pack(side="left", padx=(0, 8))
        ctk.CTkButton(btn_frame, text="Edit", command=self._edit_chain_step, width=70,
                      fg_color="transparent", border_width=1).pack(side="left", padx=(0, 8))
        ctk.CTkButton(btn_frame, text="↑", command=self._move_chain_up, width=35,
                      fg_color="transparent", border_width=1).pack(side="left", padx=(0, 3))
        ctk.CTkButton(btn_frame, text="↓", command=self._move_chain_down, width=35,
                      fg_color="transparent", border_width=1).pack(side="left", padx=(0, 8))
        ctk.CTkButton(btn_frame, text="Remove", command=self._remove_chain_step, width=70,
                      fg_color="transparent", border_width=1).pack(side="left")
        
        ctk.CTkLabel(self.content_frame, text="💡 Build a sequence of actions: open apps, snap windows, run macros, etc.",
                     text_color="gray", font=ctk.CTkFont(size=11), wraplength=400).pack(anchor="w", pady=(10, 0))
    
    def _add_chain_action(self) -> None:
        """Add a new action to the chain."""
        self._open_chain_action_dialog()
    
    def _edit_chain_step(self) -> None:
        """Edit selected chain step."""
        if hasattr(self, 'selected_chain_index') and self.selected_chain_index is not None:
            if 0 <= self.selected_chain_index < len(self.chain_steps):
                step = self.chain_steps[self.selected_chain_index]
                self._open_chain_action_dialog(edit_index=self.selected_chain_index, edit_data=step)
    
    def _move_chain_up(self) -> None:
        """Move selected step up."""
        if hasattr(self, 'selected_chain_index') and self.selected_chain_index is not None:
            idx = self.selected_chain_index
            if idx > 0:
                self.chain_steps[idx], self.chain_steps[idx-1] = self.chain_steps[idx-1], self.chain_steps[idx]
                self.selected_chain_index = idx - 1
                self._refresh_chain_list()
    
    def _move_chain_down(self) -> None:
        """Move selected step down."""
        if hasattr(self, 'selected_chain_index') and self.selected_chain_index is not None:
            idx = self.selected_chain_index
            if idx < len(self.chain_steps) - 1:
                self.chain_steps[idx], self.chain_steps[idx+1] = self.chain_steps[idx+1], self.chain_steps[idx]
                self.selected_chain_index = idx + 1
                self._refresh_chain_list()
    
    def _open_chain_action_dialog(self, edit_index: int = None, edit_data: dict = None) -> None:
        """Open dialog to add/edit a chain action."""
        dialog = ctk.CTkToplevel(self)
        dialog.title("Edit Action" if edit_index is not None else "Add Action")
        dialog.geometry("560x540")
        dialog.minsize(560, 300)
        dialog.transient(self)
        dialog.grab_set()
        set_window_icon(dialog)
        
        dialog.update_idletasks()
        x = self.winfo_x() + (self.winfo_width() - 560) // 2
        y = self.winfo_y() + (self.winfo_height() - 540) // 2
        dialog.geometry(f"+{x}+{y}")
        
        # Main frame with scrollable content
        main_frame = ctk.CTkFrame(dialog, fg_color="transparent")
        main_frame.pack(fill="both", expand=True, padx=20, pady=20)
        
        # Action type selection at top (fixed)
        ctk.CTkLabel(main_frame, text="Action Type", font=ctk.CTkFont(size=13, weight="bold")).pack(anchor="w", pady=(0, 8))
        
        action_type_var = ctk.StringVar(value=edit_data.get("type", "open_file") if edit_data else "open_file")
        
        type_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        type_frame.pack(fill="x", pady=(0, 10))
        
        action_types = [
            ("Open File/App", "open_file"),
            ("Open Folder", "open_folder"),
            ("Open Website", "open_url"),
            ("Window Action", "window_action"),
            ("Macro", "macro"),
            ("Wait (delay)", "wait"),
        ]
        
        type_row1 = ctk.CTkFrame(type_frame, fg_color="transparent")
        type_row1.pack(fill="x", pady=(0, 4))
        type_row2 = ctk.CTkFrame(type_frame, fg_color="transparent")
        type_row2.pack(fill="x")

        for i, (text, value) in enumerate(action_types):
            row = type_row1 if i < 3 else type_row2
            ctk.CTkRadioButton(row, text=text, variable=action_type_var, value=value,
                               command=lambda: update_content_frame()).pack(side="left", padx=(0, 12))
        
        # Scrollable content frame for action-specific options
        content_scroll = ctk.CTkScrollableFrame(main_frame, fg_color="transparent")
        content_scroll.pack(fill="both", expand=True, pady=(0, 15))
        
        # Inner content frame
        content_frame = ctk.CTkFrame(content_scroll, fg_color="transparent")
        content_frame.pack(fill="both", expand=True)
        
        # Variables for action data
        file_path_var = ctk.StringVar()
        url_var = ctk.StringVar()
        window_action_var = ctk.StringVar(value="snap_left")
        window_target_var = ctk.StringVar(value="focused")
        target_app_var = ctk.StringVar()
        target_title_var = ctk.StringVar()
        wait_seconds_var = ctk.StringVar(value="0.5")
        chain_macro_steps = []  # inline macro steps for chain
        
        _dialog_sizes = {
            "open_file": (560, 220),
            "open_folder": (560, 220),
            "open_url":  (560, 210),
            "window_action": (560, 490),
            "macro": (560, 710),
            "wait": (560, 220),
        }

        def update_content_frame():
            for w in content_frame.winfo_children():
                w.destroy()
            
            action_type = action_type_var.get()
            
            if action_type == "open_file":
                ctk.CTkLabel(content_frame, text="File/App Path:").pack(anchor="w", pady=(0, 5))
                path_frame = ctk.CTkFrame(content_frame, fg_color="transparent")
                path_frame.pack(fill="x")
                ctk.CTkEntry(path_frame, textvariable=file_path_var, width=280).pack(side="left")
                
                def browse():
                    from tkinter import filedialog
                    filepath = filedialog.askopenfilename()
                    if filepath:
                        file_path_var.set(filepath)
                
                ctk.CTkButton(path_frame, text="Browse", command=browse, width=80).pack(side="left", padx=(10, 0))
            
            elif action_type == "open_folder":
                ctk.CTkLabel(content_frame, text="Folder Path:").pack(anchor="w", pady=(0, 5))
                folder_frame = ctk.CTkFrame(content_frame, fg_color="transparent")
                folder_frame.pack(fill="x")
                ctk.CTkEntry(folder_frame, textvariable=file_path_var, width=280).pack(side="left")
                
                def browse_folder():
                    from tkinter import filedialog
                    folder = filedialog.askdirectory()
                    if folder:
                        file_path_var.set(folder)
                
                ctk.CTkButton(folder_frame, text="Browse", command=browse_folder, width=80).pack(side="left", padx=(10, 0))
            
            elif action_type == "open_url":
                ctk.CTkLabel(content_frame, text="Website URL:").pack(anchor="w", pady=(0, 5))
                ctk.CTkEntry(content_frame, textvariable=url_var, placeholder_text="https://example.com", 
                             width=350).pack(fill="x")
            
            elif action_type == "window_action":
                # Target
                ctk.CTkLabel(content_frame, text="Target:").pack(anchor="w", pady=(0, 5))
                target_frame = ctk.CTkFrame(content_frame, fg_color="transparent")
                target_frame.pack(fill="x", pady=(0, 8))
                
                def on_target_change():
                    app_frame.pack_forget()
                    title_frame.pack_forget()
                    if window_target_var.get() == "app":
                        app_frame.pack(fill="x", pady=(0, 8))
                    elif window_target_var.get() == "title":
                        title_frame.pack(fill="x", pady=(0, 8))
                
                ctk.CTkRadioButton(target_frame, text="Focused", variable=window_target_var, 
                                   value="focused", command=on_target_change).pack(side="left", padx=(0, 10))
                ctk.CTkRadioButton(target_frame, text="By App", variable=window_target_var,
                                   value="app", command=on_target_change).pack(side="left", padx=(0, 10))
                ctk.CTkRadioButton(target_frame, text="By Title", variable=window_target_var,
                                   value="title", command=on_target_change).pack(side="left")
                
                # App input (hidden by default)
                app_frame = ctk.CTkFrame(content_frame, fg_color="transparent")
                ctk.CTkLabel(app_frame, text="App name:").pack(side="left")
                ctk.CTkEntry(app_frame, textvariable=target_app_var, width=150, 
                             placeholder_text="e.g., chrome").pack(side="left", padx=(5, 0))
                
                # Title input (hidden by default)
                title_frame = ctk.CTkFrame(content_frame, fg_color="transparent")
                ctk.CTkLabel(title_frame, text="Title contains:").pack(side="left")
                ctk.CTkEntry(title_frame, textvariable=target_title_var, width=150,
                             placeholder_text="e.g., GitHub").pack(side="left", padx=(5, 0))
                
                # Action
                ctk.CTkLabel(content_frame, text="Action:").pack(anchor="w", pady=(5, 5))
                actions_frame = ctk.CTkFrame(content_frame, fg_color="transparent")
                actions_frame.pack(fill="x")
                
                left_col = ctk.CTkFrame(actions_frame, fg_color="transparent")
                left_col.pack(side="left", fill="y")
                right_col = ctk.CTkFrame(actions_frame, fg_color="transparent")
                right_col.pack(side="left", fill="y", padx=(15, 0))
                
                actions = [
                    ("Snap Left", "snap_left"), ("Snap Right", "snap_right"),
                    ("Snap Top-Left", "snap_top_left"), ("Snap Top-Right", "snap_top_right"),
                    ("Snap Bottom-Left", "snap_bottom_left"), ("Snap Bottom-Right", "snap_bottom_right"),
                    ("Minimize", "minimize"), ("Maximize", "maximize"),
                    ("Restore", "restore"), ("Focus", "focus"), ("Close", "close"),
                ]
                
                for i, (text, value) in enumerate(actions):
                    col = left_col if i < 5 else right_col
                    ctk.CTkRadioButton(col, text=text, variable=window_action_var, value=value,
                                       font=ctk.CTkFont(size=11)).pack(anchor="w", pady=1)
                
                on_target_change()  # Initialize visibility
            
            elif action_type == "macro":
                # Inline macro step editor inside chain dialog
                # --- Recording section ---
                rec_frame = ctk.CTkFrame(content_frame, corner_radius=8)
                rec_frame.pack(fill="x", pady=(0, 6))
                rec_inner = ctk.CTkFrame(rec_frame, fg_color="transparent")
                rec_inner.pack(fill="x", padx=12, pady=8)

                rec_top = ctk.CTkFrame(rec_inner, fg_color="transparent")
                rec_top.pack(fill="x")

                chain_rec_btn_text = ctk.StringVar(value="🔴 Record")
                chain_rec_btn = ctk.CTkButton(rec_top, textvariable=chain_rec_btn_text,
                                              width=110)
                chain_rec_btn.pack(side="left", padx=(0, 10))
                chain_rec_status = ctk.CTkLabel(rec_top, text="Press keys while recording",
                                                text_color="gray", font=ctk.CTkFont(size=12))
                chain_rec_status.pack(side="left")

                rec_delay_row = ctk.CTkFrame(rec_inner, fg_color="transparent")
                rec_delay_row.pack(fill="x", pady=(6, 0))
                ctk.CTkLabel(rec_delay_row, text="Delay:", font=ctk.CTkFont(size=12)).pack(side="left")
                chain_delay_var = ctk.StringVar(value="0.1")
                ctk.CTkEntry(rec_delay_row, textvariable=chain_delay_var, width=50).pack(side="left", padx=(5, 3))
                ctk.CTkLabel(rec_delay_row, text="s", font=ctk.CTkFont(size=12)).pack(side="left")

                def _chain_apply_delay_all():
                    try:
                        delay = float(chain_delay_var.get())
                        if delay < 0:
                            raise ValueError()
                    except ValueError:
                        CTkMessagebox.show(dialog, "Warning", "Enter a valid delay (0 or greater).", "warning")
                        return
                    for step in chain_macro_steps:
                        step["delay"] = delay
                    refresh_chain_macro_lb()

                ctk.CTkButton(rec_delay_row, text="Apply to All", command=_chain_apply_delay_all,
                              width=90, height=28, fg_color="transparent", border_width=1).pack(side="left", padx=(10, 0))

                # Recording state (mutable container so closures can share it)
                _rec_state = {"active": False, "held": set(), "last_time": None, "hook": None}

                # --- Listbox ---
                ctk.CTkLabel(content_frame, text="Macro Steps",
                             font=ctk.CTkFont(size=12, weight="bold")).pack(anchor="w", pady=(0, 5))

                list_frame = ctk.CTkFrame(content_frame, corner_radius=8)
                list_frame.pack(fill="x", pady=(0, 8))

                macro_lb = tk.Listbox(list_frame, font=("Segoe UI", 10), height=5,
                                      bg="#2b2b2b", fg="white", selectbackground="#1f6aa5",
                                      highlightthickness=0, bd=0)
                macro_lb_scroll = ctk.CTkScrollbar(list_frame, command=macro_lb.yview)
                macro_lb.configure(yscrollcommand=macro_lb_scroll.set)
                macro_lb.pack(side="left", fill="both", expand=True, padx=5, pady=5)
                macro_lb_scroll.pack(side="right", fill="y", padx=(0, 5), pady=5)

                def refresh_chain_macro_lb():
                    macro_lb.delete(0, tk.END)
                    for s in chain_macro_steps:
                        st = s.get("type", "key")
                        d = s.get("delay", 0.1)
                        if st == "key":
                            macro_lb.insert(tk.END, f"⌨️ {s.get('key','')} ({d}s)")
                        elif st == "mouse_move":
                            macro_lb.insert(tk.END, f"🖱️ Move to ({s.get('x',0)}, {s.get('y',0)}) ({d}s)")
                        elif st == "mouse_click":
                            macro_lb.insert(tk.END, f"🖱️ Click {s.get('button','left')} at ({s.get('x','cur')}, {s.get('y','cur')}) ({d}s)")
                        elif st == "mouse_scroll":
                            amt = s.get('amount', 3)
                            macro_lb.insert(tk.END, f"🖱️ Scroll {'up' if amt>0 else 'down'} {abs(amt)} ({d}s)")

                # Populate if editing
                refresh_chain_macro_lb()

                # --- Recording logic ---
                def _chain_stop_recording():
                    _rec_state["active"] = False
                    chain_rec_btn_text.set("🔴 Record")
                    chain_rec_status.configure(text="Recording stopped", text_color="#44aa44")
                    if _rec_state["hook"]:
                        keyboard.unhook(_rec_state["hook"])
                        _rec_state["hook"] = None
                    if chain_macro_steps:
                        try:
                            chain_macro_steps[-1]["delay"] = float(chain_delay_var.get())
                        except ValueError:
                            chain_macro_steps[-1]["delay"] = 0.1
                        refresh_chain_macro_lb()

                def _chain_start_recording():
                    import time as _t
                    _rec_state["active"] = True
                    _rec_state["held"] = set()
                    _rec_state["last_time"] = None
                    chain_rec_btn_text.set("⏹️ Stop")
                    chain_rec_status.configure(text="Recording... Press keys now!", text_color="#ff4444")

                    def on_key(event):
                        if not _rec_state["active"]:
                            return
                        kn = (event.name or "").lower()
                        sc = event.scan_code
                        if event.event_type == "down":
                            kid = (kn, sc)
                            if kid in _rec_state["held"]:
                                return
                            _rec_state["held"].add(kid)
                            now = _t.time()
                            if _rec_state["last_time"] is not None:
                                delay = min(round(now - _rec_state["last_time"], 2), 5.0)
                            else:
                                try:
                                    delay = float(chain_delay_var.get())
                                except ValueError:
                                    delay = 0.1
                            _rec_state["last_time"] = now
                            modifiers = [
                                "ctrl", "left ctrl", "right ctrl",
                                "alt", "left alt", "right alt", "altgr", "alt gr",
                                "shift", "left shift", "right shift",
                                "left windows", "right windows",
                            ]
                            if kn in modifiers:
                                return
                            parts = []
                            held = _rec_state["held"]
                            if any(k[0] in ("ctrl", "left ctrl", "right ctrl") for k in held):
                                parts.append("ctrl")
                            if any(k[0] in ("alt", "left alt") for k in held):
                                parts.append("alt")
                            if any(k[0] in ("right alt", "altgr", "alt gr") for k in held):
                                parts.append("altgr")
                            if any(k[0] in ("shift", "left shift", "right shift") for k in held):
                                parts.append("shift")
                            if any(k[0] in ("left windows", "right windows") for k in held):
                                parts.append("win")
                            parts.append(kn)
                            chain_macro_steps.append({"type": "key", "key": "+".join(parts), "delay": delay})
                            self.after(0, refresh_chain_macro_lb)
                        elif event.event_type == "up":
                            _rec_state["held"].discard((kn, sc))

                    _rec_state["hook"] = keyboard.hook(on_key)

                def _chain_toggle_recording():
                    if _rec_state["active"]:
                        _chain_stop_recording()
                    else:
                        _chain_start_recording()

                chain_rec_btn.configure(command=_chain_toggle_recording)

                # Stop recording if dialog is closed mid-recording
                dialog.protocol("WM_DELETE_WINDOW", lambda: (_chain_stop_recording(), dialog.destroy()))

                # --- Step management buttons ---
                mb_btns = ctk.CTkFrame(content_frame, fg_color="transparent")
                mb_btns.pack(fill="x", pady=(0, 5))

                def chain_add_key():
                    self._open_chain_macro_key_dialog(chain_macro_steps, refresh_chain_macro_lb, dialog)

                def chain_add_mouse():
                    self._open_chain_macro_mouse_dialog(chain_macro_steps, refresh_chain_macro_lb, dialog)

                def chain_remove_step():
                    sel = macro_lb.curselection()
                    if sel:
                        del chain_macro_steps[sel[0]]
                        refresh_chain_macro_lb()

                def chain_move_up():
                    sel = macro_lb.curselection()
                    if sel and sel[0] > 0:
                        i = sel[0]
                        chain_macro_steps[i], chain_macro_steps[i-1] = chain_macro_steps[i-1], chain_macro_steps[i]
                        refresh_chain_macro_lb()
                        macro_lb.selection_set(i-1)

                def chain_move_down():
                    sel = macro_lb.curselection()
                    if sel and sel[0] < len(chain_macro_steps)-1:
                        i = sel[0]
                        chain_macro_steps[i], chain_macro_steps[i+1] = chain_macro_steps[i+1], chain_macro_steps[i]
                        refresh_chain_macro_lb()
                        macro_lb.selection_set(i+1)

                ctk.CTkButton(mb_btns, text="+ Key", command=chain_add_key, width=65).pack(side="left", padx=(0,4))
                ctk.CTkButton(mb_btns, text="+ Mouse", command=chain_add_mouse, width=75,
                              fg_color="#2d5a2d", hover_color="#1e3d1e").pack(side="left", padx=(0,4))
                ctk.CTkButton(mb_btns, text="Remove", command=chain_remove_step, width=65,
                              fg_color="transparent", border_width=1).pack(side="left", padx=(0,4))
                ctk.CTkButton(mb_btns, text="↑", command=chain_move_up, width=28,
                              fg_color="transparent", border_width=1).pack(side="left", padx=(0,2))
                ctk.CTkButton(mb_btns, text="↓", command=chain_move_down, width=28,
                              fg_color="transparent", border_width=1).pack(side="left")

            elif action_type == "wait":
                ctk.CTkLabel(content_frame, text="Wait duration (seconds):").pack(anchor="w", pady=(0, 5))
                ctk.CTkEntry(content_frame, textvariable=wait_seconds_var, width=100).pack(anchor="w")
                ctk.CTkLabel(content_frame, text="💡 Add delays between actions if needed", 
                             text_color="gray", font=ctk.CTkFont(size=11)).pack(anchor="w", pady=(10, 0))

            # Resize dialog to fit this action type
            w, h = _dialog_sizes.get(action_type_var.get(), (560, 400))
            cx = dialog.winfo_x() + (dialog.winfo_width() - w) // 2
            cy = dialog.winfo_y() + (dialog.winfo_height() - h) // 2
            dialog.geometry(f"{w}x{h}+{cx}+{cy}")
        
        # Load edit data if editing
        if edit_data:
            action_type = edit_data.get("type", "open_file")
            data = edit_data.get("data", {})
            
            if action_type in ("open_file", "open_folder"):
                file_path_var.set(data if isinstance(data, str) else "")
            elif action_type == "open_url":
                url_var.set(data if isinstance(data, str) else "")
            elif action_type == "window_action":
                if isinstance(data, dict):
                    window_action_var.set(data.get("action", "snap_left"))
                    window_target_var.set(data.get("target", "focused"))
                    target_app_var.set(data.get("app", ""))
                    target_title_var.set(data.get("title", ""))
                elif isinstance(data, str):
                    window_action_var.set(data)
            elif action_type == "macro":
                chain_macro_steps.clear()
                if isinstance(data, list):
                    chain_macro_steps.extend(data)
            elif action_type == "wait":
                wait_seconds_var.set(str(data.get("seconds", 0.5)) if isinstance(data, dict) else "0.5")
        
        update_content_frame()
        
        # Buttons at bottom (fixed, not scrollable)
        btn_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        btn_frame.pack(fill="x")
        
        def save_action():
            action_type = action_type_var.get()
            
            if action_type == "open_file":
                path = file_path_var.get().strip()
                if not path:
                    CTkMessagebox.show(dialog, "Warning", "Please enter a file path.", "warning")
                    return
                data = path
                display = f"📂 Open: {os.path.basename(path)}"
            
            elif action_type == "open_folder":
                path = file_path_var.get().strip()
                if not path:
                    CTkMessagebox.show(dialog, "Warning", "Please enter a folder path.", "warning")
                    return
                data = path
                display = f"📁 Open: {os.path.basename(path) or path}"
            
            elif action_type == "open_url":
                url = url_var.get().strip()
                if not url:
                    CTkMessagebox.show(dialog, "Warning", "Please enter a URL.", "warning")
                    return
                if not url.startswith(("http://", "https://")):
                    url = "https://" + url
                data = url
                display = f"🌐 Open: {url[:30]}..."
            
            elif action_type == "window_action":
                action = window_action_var.get()
                target = window_target_var.get()
                data = {
                    "action": action,
                    "target": target,
                    "app": target_app_var.get().strip() if target == "app" else "",
                    "title": target_title_var.get().strip() if target == "title" else ""
                }
                
                action_names = {
                    "snap_left": "Snap Left", "snap_right": "Snap Right",
                    "snap_top_left": "Snap Top-Left", "snap_top_right": "Snap Top-Right",
                    "snap_bottom_left": "Snap Bottom-Left", "snap_bottom_right": "Snap Bottom-Right",
                    "minimize": "Minimize", "maximize": "Maximize",
                    "restore": "Restore", "focus": "Focus", "close": "Close"
                }
                action_name = action_names.get(action, action)
                
                if target == "app" and data["app"]:
                    display = f"🪟 {action_name}: {data['app']}"
                elif target == "title" and data["title"]:
                    display = f"🪟 {action_name}: '{data['title'][:15]}...'"
                else:
                    display = f"🪟 {action_name} (focused)"
            
            elif action_type == "macro":
                if not chain_macro_steps:
                    CTkMessagebox.show(dialog, "Warning", "Please add at least one macro step.", "warning")
                    return
                data = list(chain_macro_steps)
                display = f"⌨️ Macro ({len(chain_macro_steps)} steps)"

            elif action_type == "wait":
                try:
                    seconds = float(wait_seconds_var.get())
                    if seconds < 0:
                        raise ValueError()
                except ValueError:
                    CTkMessagebox.show(dialog, "Warning", "Please enter a valid delay.", "warning")
                    return
                data = {"seconds": seconds}
                display = f"⏱️ Wait {seconds}s"
            
            else:
                return
            
            step = {"type": action_type, "data": data, "display": display}
            
            if edit_index is not None:
                self.chain_steps[edit_index] = step
            else:
                self.chain_steps.append(step)
            
            self._refresh_chain_list()
            dialog.destroy()
        
        ctk.CTkButton(btn_frame, text="Save" if edit_index is not None else "Add", 
                      command=save_action, width=100).pack(side="right")
        ctk.CTkButton(btn_frame, text="Cancel", command=dialog.destroy, width=100,
                      fg_color="transparent", border_width=1).pack(side="right", padx=(0, 10))
    
    def _remove_chain_step(self) -> None:
        """Remove the selected step from the chain."""
        if hasattr(self, 'selected_chain_index') and self.selected_chain_index is not None:
            if 0 <= self.selected_chain_index < len(self.chain_steps):
                del self.chain_steps[self.selected_chain_index]
                self.selected_chain_index = None
                self._refresh_chain_list()
        elif self.chain_steps:
            # Fallback: remove last if nothing selected
            self.chain_steps.pop()
            self._refresh_chain_list()
    
    def _open_chain_macro_key_dialog(self, steps: list, refresh_cb, parent) -> None:
        """Add a key step to a chain-embedded macro."""
        dialog = ctk.CTkToplevel(parent)
        dialog.title("Add Key Step")
        dialog.geometry("300x200")
        dialog.transient(parent)
        dialog.grab_set()
        set_window_icon(dialog)

        frame = ctk.CTkFrame(dialog, fg_color="transparent")
        frame.pack(fill="both", expand=True, padx=20, pady=20)

        ctk.CTkLabel(frame, text="Key/Combo").pack(anchor="w")
        key_var = ctk.StringVar()
        ctk.CTkEntry(frame, textvariable=key_var, width=260).pack(fill="x", pady=(5, 15))

        ctk.CTkLabel(frame, text="Delay (seconds)").pack(anchor="w")
        delay_var = ctk.StringVar(value="0.1")
        ctk.CTkEntry(frame, textvariable=delay_var, width=260).pack(fill="x", pady=(5, 20))

        def add():
            key = key_var.get().strip()
            try:
                delay = float(delay_var.get())
            except ValueError:
                delay = 0.1
            if key:
                steps.append({"type": "key", "key": key, "delay": delay})
                refresh_cb()
            dialog.destroy()

        ctk.CTkButton(frame, text="Add", command=add, width=100).pack(side="right")
        dialog.update_idletasks()
        x = parent.winfo_x() + (parent.winfo_width() - 300) // 2
        y = parent.winfo_y() + (parent.winfo_height() - 200) // 2
        dialog.geometry(f"+{x}+{y}")

    def _open_chain_macro_mouse_dialog(self, steps: list, refresh_cb, parent) -> None:
        """Add a mouse step to a chain-embedded macro."""
        dialog = ctk.CTkToplevel(parent)
        dialog.title("Add Mouse Step")
        dialog.geometry("380x460")
        dialog.transient(parent)
        dialog.grab_set()
        set_window_icon(dialog)

        frame = ctk.CTkFrame(dialog, fg_color="transparent")
        frame.pack(fill="both", expand=True, padx=20, pady=20)

        # Pack buttons FIRST with side="bottom" so they always stay visible
        btn_row = ctk.CTkFrame(frame, fg_color="transparent")
        btn_row.pack(fill="x", side="bottom", pady=(8, 0))

        ctk.CTkLabel(frame, text="Action Type", font=ctk.CTkFont(size=13, weight="bold")).pack(anchor="w")
        action_var = ctk.StringVar(value="mouse_click")
        action_row = ctk.CTkFrame(frame, fg_color="transparent")
        action_row.pack(fill="x", pady=(5, 15))
        for label, val in [("Move", "mouse_move"), ("Click", "mouse_click"), ("Scroll", "mouse_scroll")]:
            ctk.CTkRadioButton(action_row, text=label, variable=action_var, value=val).pack(side="left", padx=(0, 12))

        pos_frame = ctk.CTkFrame(frame, corner_radius=10)
        pos_frame.pack(fill="x", pady=(0, 12))
        pos_inner = ctk.CTkFrame(pos_frame, fg_color="transparent")
        pos_inner.pack(fill="x", padx=15, pady=12)
        ctk.CTkLabel(pos_inner, text="Position", font=ctk.CTkFont(size=12, weight="bold")).pack(anchor="w")
        coord_row = ctk.CTkFrame(pos_inner, fg_color="transparent")
        coord_row.pack(fill="x", pady=(8, 8))
        ctk.CTkLabel(coord_row, text="X:").pack(side="left")
        x_var = ctk.StringVar()
        ctk.CTkEntry(coord_row, textvariable=x_var, width=75, placeholder_text="0").pack(side="left", padx=(4, 16))
        ctk.CTkLabel(coord_row, text="Y:").pack(side="left")
        y_var = ctk.StringVar()
        ctk.CTkEntry(coord_row, textvariable=y_var, width=75, placeholder_text="0").pack(side="left", padx=4)

        def record_pos():
            dialog.withdraw()
            parent.withdraw()
            self.after(200, lambda: self._capture_mouse_position_chain(parent, dialog, x_var, y_var))

        rec_row = ctk.CTkFrame(pos_inner, fg_color="transparent")
        rec_row.pack(fill="x")
        ctk.CTkButton(rec_row, text="📍 Record Position", command=record_pos,
                      width=140, fg_color="#2d5a2d", hover_color="#1e3d1e").pack(side="left")
        ctk.CTkLabel(rec_row, text="Click to capture, Esc to cancel",
                     text_color="gray", font=ctk.CTkFont(size=11)).pack(side="left", padx=(8, 0))

        opt_frame = ctk.CTkFrame(frame, corner_radius=10)
        opt_frame.pack(fill="x", pady=(0, 12))
        opt_inner = ctk.CTkFrame(opt_frame, fg_color="transparent")
        opt_inner.pack(fill="x", padx=15, pady=12)
        ctk.CTkLabel(opt_inner, text="Click Options", font=ctk.CTkFont(size=12, weight="bold")).pack(anchor="w")
        click_row = ctk.CTkFrame(opt_inner, fg_color="transparent")
        click_row.pack(fill="x", pady=(6, 0))
        ctk.CTkLabel(click_row, text="Button:").pack(side="left")
        button_var = ctk.StringVar(value="left")
        ctk.CTkOptionMenu(click_row, variable=button_var, values=["left", "right", "middle"], width=85).pack(side="left", padx=(4, 16))
        ctk.CTkLabel(click_row, text="Clicks:").pack(side="left")
        clicks_var = ctk.StringVar(value="1")
        ctk.CTkOptionMenu(click_row, variable=clicks_var, values=["1", "2"], width=55).pack(side="left", padx=4)
        scroll_row = ctk.CTkFrame(opt_inner, fg_color="transparent")
        scroll_row.pack(fill="x", pady=(8, 0))
        ctk.CTkLabel(scroll_row, text="Scroll:").pack(side="left")
        scroll_var = ctk.StringVar(value="3")
        ctk.CTkEntry(scroll_row, textvariable=scroll_var, width=55).pack(side="left", padx=(4, 6))
        ctk.CTkLabel(scroll_row, text="(+up, -down)", text_color="gray", font=ctk.CTkFont(size=11)).pack(side="left")

        delay_row = ctk.CTkFrame(frame, fg_color="transparent")
        delay_row.pack(fill="x", pady=(0, 4))
        ctk.CTkLabel(delay_row, text="Delay:").pack(side="left")
        delay_var = ctk.StringVar(value="0.1")
        ctk.CTkEntry(delay_row, textvariable=delay_var, width=55).pack(side="left", padx=(4, 6))
        ctk.CTkLabel(delay_row, text="seconds").pack(side="left")

        def add():
            action = action_var.get()
            try:
                delay = float(delay_var.get())
            except ValueError:
                delay = 0.1
            step = {"type": action, "delay": delay}
            xs, ys = x_var.get().strip(), y_var.get().strip()
            if xs and ys:
                try:
                    step["x"] = int(xs)
                    step["y"] = int(ys)
                except ValueError:
                    CTkMessagebox.show(dialog, "Error", "Invalid X or Y coordinate.", "warning")
                    return
            elif action == "mouse_move":
                CTkMessagebox.show(dialog, "Error", "Move requires coordinates.", "warning")
                return
            if action == "mouse_click":
                step["button"] = button_var.get()
                step["clicks"] = int(clicks_var.get())
            elif action == "mouse_scroll":
                try:
                    step["amount"] = int(scroll_var.get())
                except ValueError:
                    step["amount"] = 3
            steps.append(step)
            refresh_cb()
            dialog.destroy()

        ctk.CTkButton(btn_row, text="Add", command=add, width=100).pack(side="right")
        ctk.CTkButton(btn_row, text="Cancel", command=dialog.destroy, width=100,
                      fg_color="transparent", border_width=1).pack(side="right", padx=(0, 8))
        dialog.update_idletasks()
        x = parent.winfo_x() + (parent.winfo_width() - 380) // 2
        y = parent.winfo_y() + (parent.winfo_height() - 460) // 2
        dialog.geometry(f"+{x}+{y}")

    def _capture_mouse_position_chain(self, chain_parent, mouse_dialog, x_var, y_var) -> None:
        """Variant of _capture_mouse_position that restores both chain parent and mouse dialog."""
        overlay = tk.Toplevel(self)
        overlay.attributes("-fullscreen", True)
        overlay.attributes("-alpha", 0.01)
        overlay.attributes("-topmost", True)
        overlay.overrideredirect(True)
        overlay.configure(bg="black")

        def on_click(event):
            x, y = event.x_root, event.y_root
            overlay.destroy()
            if self.winfo_exists():
                self.deiconify()
            if chain_parent.winfo_exists():
                chain_parent.deiconify()
                chain_parent.lift()
                chain_parent.focus_force()
            if mouse_dialog.winfo_exists():
                mouse_dialog.deiconify()
                mouse_dialog.lift()
                mouse_dialog.focus_force()
            x_var.set(str(x))
            y_var.set(str(y))

        def on_escape(event):
            overlay.destroy()
            if self.winfo_exists():
                self.deiconify()
            if chain_parent.winfo_exists():
                chain_parent.deiconify()
                chain_parent.lift()
            if mouse_dialog.winfo_exists():
                mouse_dialog.deiconify()
                mouse_dialog.lift()
                mouse_dialog.focus_force()

        overlay.bind("<Button-1>", on_click)
        overlay.bind("<Escape>", on_escape)
        overlay.focus_force()
        overlay.grab_set()

    def _refresh_chain_list(self) -> None:
        """Refresh the chain steps display."""
        for w in self.chain_list.winfo_children():
            w.destroy()
        
        self.selected_chain_index = getattr(self, 'selected_chain_index', None)
        
        for i, step in enumerate(self.chain_steps):
            is_selected = (i == self.selected_chain_index)
            fg_color = ("#1f6aa5", "#1f6aa5") if is_selected else ("gray90", "gray17")
            
            row = ctk.CTkFrame(self.chain_list, fg_color=fg_color, corner_radius=6, height=32)
            row.pack(fill="x", pady=2)
            row.pack_propagate(False)
            
            def on_click(event, idx=i):
                self.selected_chain_index = idx
                self._refresh_chain_list()
            
            row.bind("<Button-1>", on_click)
            
            label = ctk.CTkLabel(row, text=f"{i+1}. {step['display']}", anchor="w")
            label.pack(side="left", padx=10, fill="x", expand=True)
            label.bind("<Button-1>", on_click)
    
    def _browse_file(self) -> None:
        from tkinter import filedialog
        filepath = filedialog.askopenfilename()
        if filepath:
            self.path_var.set(filepath)

    def _browse_folder(self) -> None:
        from tkinter import filedialog
        folder = filedialog.askdirectory()
        if folder:
            self.folder_var.set(folder)
    
    def _add_macro_step(self) -> None:
        dialog = ctk.CTkToplevel(self)
        dialog.title("Add Key")
        dialog.geometry("300x200")
        dialog.transient(self)
        dialog.grab_set()
        
        frame = ctk.CTkFrame(dialog, fg_color="transparent")
        frame.pack(fill="both", expand=True, padx=20, pady=20)
        
        ctk.CTkLabel(frame, text="Key/Combo").pack(anchor="w")
        key_var = ctk.StringVar()
        ctk.CTkEntry(frame, textvariable=key_var, width=260).pack(fill="x", pady=(5, 15))
        
        ctk.CTkLabel(frame, text="Delay (seconds)").pack(anchor="w")
        delay_var = ctk.StringVar(value="0.1")
        ctk.CTkEntry(frame, textvariable=delay_var, width=260).pack(fill="x", pady=(5, 20))
        
        def add():
            key = key_var.get().strip()
            try:
                delay = float(delay_var.get())
            except:
                delay = 0.1
            if key:
                self.macro_steps.append({"type": "key", "key": key, "delay": delay})
                self._refresh_macro_list()
            dialog.destroy()
        
        ctk.CTkButton(frame, text="Add", command=add, width=100).pack(side="right")
        
        dialog.update_idletasks()
        x = self.winfo_x() + (self.winfo_width() - 300) // 2
        y = self.winfo_y() + (self.winfo_height() - 200) // 2
        dialog.geometry(f"+{x}+{y}")
    
    def _remove_macro_step(self) -> None:
        selection = self.macro_listbox.curselection()
        if selection:
            idx = selection[0]
            del self.macro_steps[idx]
            self._refresh_macro_list()
    
    def _load_command_data(self) -> None:
        cmd = self.command_manager.get_command(self.edit_command_id)
        if cmd:
            self.phrases_text.insert("1.0", "\n".join(cmd.get("phrases", [])))
            
            cmd_type = cmd.get("type", "open_file")
            self.command_type.set(cmd_type)
            self._on_type_changed()
            
            if cmd_type == "open_file":
                self.path_var.set(cmd.get("data", ""))
            elif cmd_type == "open_folder":
                self.folder_var.set(cmd.get("data", ""))
            elif cmd_type == "open_url":
                self.url_var.set(cmd.get("data", ""))
            elif cmd_type == "window_action":
                data = cmd.get("data", "minimize")
                # Handle both legacy string format and new dict format
                if isinstance(data, str):
                    self.window_action_var.set(data)
                    self.window_target_var.set("focused")
                else:
                    self.window_action_var.set(data.get("action", "minimize"))
                    self.window_target_var.set(data.get("target", "focused"))
                    self.target_app_var.set(data.get("app", ""))
                    self.target_title_var.set(data.get("title", ""))
                    self._on_window_target_changed()  # Show correct input field
            elif cmd_type == "chain":
                # Load chain steps with inline action data
                chain_data = cmd.get("data", [])
                self.chain_steps = []
                for step in chain_data:
                    if isinstance(step, dict) and "type" in step:
                        # New format: inline actions
                        self.chain_steps.append(step.copy())
                    elif isinstance(step, str):
                        # Legacy format: command ID reference - convert to inline
                        cmd_info = self.command_manager.get_command(step)
                        if cmd_info:
                            self.chain_steps.append({
                                "type": cmd_info.get("type", "open_file"),
                                "data": cmd_info.get("data"),
                                "display": f"[Imported] {cmd_info.get('phrases', ['Unknown'])[0]}"
                            })
                self._refresh_chain_list()
            elif cmd_type == "macro":
                self.macro_steps = cmd.get("data", []).copy()
                self._refresh_macro_list()
    
    def _save_command(self) -> None:
        phrases = [l.strip() for l in self.phrases_text.get("1.0", "end").split("\n") if l.strip()]
        
        if not phrases:
            CTkMessagebox.show(self, "Warning", "Please enter at least one phrase.", "warning")
            return
        
        is_valid, error = self.command_manager.validate_phrases(phrases)
        if not is_valid:
            CTkMessagebox.show(self, "Conflict", error, "warning")
            return
        
        cmd_type = self.command_type.get()
        
        if cmd_type == "open_file":
            path = self.path_var.get().strip()
            if not path:
                CTkMessagebox.show(self, "Warning", "Please select a file.", "warning")
                return
            action_data = path
        elif cmd_type == "open_folder":
            path = self.folder_var.get().strip()
            if not path:
                CTkMessagebox.show(self, "Warning", "Please select a folder.", "warning")
                return
            action_data = path
        elif cmd_type == "open_url":
            url = self.url_var.get().strip()
            if not url:
                CTkMessagebox.show(self, "Warning", "Please enter a URL.", "warning")
                return
            # Add https:// if no protocol specified
            if not url.startswith(("http://", "https://")):
                url = "https://" + url
            action_data = url
        elif cmd_type == "window_action":
            target = self.window_target_var.get()
            action_data = {
                "action": self.window_action_var.get(),
                "target": target,
                "app": self.target_app_var.get().strip() if target == "app" else "",
                "title": self.target_title_var.get().strip() if target == "title" else ""
            }
            # Validate target is specified when needed
            if target == "app" and not action_data["app"]:
                CTkMessagebox.show(self, "Warning", "Please specify an app name.", "warning")
                return
            if target == "title" and not action_data["title"]:
                CTkMessagebox.show(self, "Warning", "Please specify a window title.", "warning")
                return
        elif cmd_type == "chain":
            if not self.chain_steps:
                CTkMessagebox.show(self, "Warning", "Please add at least one action to the chain.", "warning")
                return
            # Save the full step data (type, data, display)
            action_data = [{"type": s["type"], "data": s["data"], "display": s["display"]} for s in self.chain_steps]
        elif cmd_type == "macro":
            if not self.macro_steps:
                CTkMessagebox.show(self, "Warning", "Please add at least one step.", "warning")
                return
            action_data = self.macro_steps.copy()
        else:
            CTkMessagebox.show(self, "Error", "Unknown command type.", "error")
            return
        
        if self.edit_command_id:
            self.command_manager.update_command(self.edit_command_id, phrases, cmd_type, action_data)
        else:
            self.command_manager.add_command(phrases, cmd_type, action_data)
        
        self.result = True
        self.destroy()


class ProfileManagerDialog(ctk.CTkToplevel):
    """Dialog for managing command profiles."""
    
    def __init__(self, parent, command_manager: CommandManager, on_change_callback=None):
        super().__init__(parent)
        self.parent = parent
        self.command_manager = command_manager
        self.on_change_callback = on_change_callback
        
        self.title("Profile Manager")
        self.geometry("480x450")
        self.resizable(False, False)
        self.transient(parent)
        self.grab_set()
        set_window_icon(self)
        
        self._create_widgets()
        self._center_window()
    
    def _center_window(self) -> None:
        self.update_idletasks()
        x = self.parent.winfo_x() + (self.parent.winfo_width() - 480) // 2
        y = self.parent.winfo_y() + (self.parent.winfo_height() - 450) // 2
        self.geometry(f"+{x}+{y}")
    
    def _create_widgets(self) -> None:
        # Main frame
        main = ctk.CTkFrame(self, fg_color="transparent")
        main.pack(fill="both", expand=True, padx=20, pady=20)
        
        # Title
        ctk.CTkLabel(main, text="📁 Profiles", font=ctk.CTkFont(size=18, weight="bold")).pack(anchor="w")
        ctk.CTkLabel(main, text="Create and manage command profiles", 
                     text_color="gray").pack(anchor="w", pady=(0, 15))
        
        # Profile list
        list_frame = ctk.CTkFrame(main)
        list_frame.pack(fill="both", expand=True, pady=(0, 15))
        
        # Scrollable list - anchor content to top
        self.profile_list = ctk.CTkScrollableFrame(list_frame, fg_color="transparent")
        self.profile_list.pack(fill="both", expand=True, padx=5, pady=5, anchor="n")
        
        self._refresh_list()
        
        # New profile section
        new_frame = ctk.CTkFrame(main, fg_color="transparent")
        new_frame.pack(fill="x", pady=(0, 10))
        
        ctk.CTkLabel(new_frame, text="New Profile:").pack(side="left")
        self.new_name_entry = ctk.CTkEntry(new_frame, placeholder_text="Profile name", width=180)
        self.new_name_entry.pack(side="left", padx=(10, 10))
        ctk.CTkButton(new_frame, text="Create", command=self._create_profile, width=80).pack(side="left")
        
        # Import/Export section
        io_frame = ctk.CTkFrame(main, fg_color="transparent")
        io_frame.pack(fill="x", pady=(0, 15))
        
        ctk.CTkButton(io_frame, text="📥 Import Profile", command=self._import_profile, 
                      width=130, fg_color="transparent", border_width=1).pack(side="left", padx=(0, 10))
        ctk.CTkButton(io_frame, text="📤 Export Current", command=self._export_current_profile,
                      width=130, fg_color="transparent", border_width=1).pack(side="left")
        
        # Close button
        ctk.CTkButton(main, text="Close", command=self.destroy, width=100).pack(side="right")
    
    def _refresh_list(self) -> None:
        """Refresh the profile list."""
        # Clear existing items
        for widget in self.profile_list.winfo_children():
            widget.destroy()
        
        current_profile = self.command_manager.get_current_profile()
        profiles = self.command_manager.get_profiles()
        
        if not profiles:
            # Show message when no profiles exist
            no_profiles_label = ctk.CTkLabel(self.profile_list, 
                                              text="No profiles yet. Create one below!",
                                              text_color="gray", font=ctk.CTkFont(size=12))
            no_profiles_label.pack(pady=20, anchor="nw")
            return
        
        for profile in profiles:
            item_frame = ctk.CTkFrame(self.profile_list, fg_color="transparent", height=32)
            item_frame.pack(fill="x", pady=2, anchor="nw")
            item_frame.pack_propagate(False)  # Maintain fixed height
            
            # Profile name with indicator if current
            is_current = profile == current_profile
            name_text = f"● {profile}" if is_current else f"   {profile}"
            name_color = ("#1f6aa5", "#5fa8d3") if is_current else ("gray10", "gray90")
            
            name_label = ctk.CTkLabel(item_frame, text=name_text, font=ctk.CTkFont(size=13),
                                       text_color=name_color, anchor="w", width=120)
            name_label.pack(side="left", padx=(5, 10))
            
            # Action buttons
            btn_frame = ctk.CTkFrame(item_frame, fg_color="transparent")
            btn_frame.pack(side="right")
            
            if not is_current:
                ctk.CTkButton(btn_frame, text="Switch", width=65, height=26,
                              command=lambda p=profile: self._switch_to(p)).pack(side="left", padx=2)
            
            ctk.CTkButton(btn_frame, text="Rename", width=65, height=26, fg_color="transparent",
                          border_width=1, command=lambda p=profile: self._rename_profile(p)).pack(side="left", padx=2)
            ctk.CTkButton(btn_frame, text="Delete", width=65, height=26, fg_color="transparent",
                          border_width=1, text_color="#e74c3c", hover_color="#c0392b",
                          command=lambda p=profile: self._delete_profile(p)).pack(side="left", padx=2)
    
    def _switch_to(self, profile_name: str) -> None:
        """Switch to a profile."""
        if self.command_manager.switch_profile(profile_name):
            self._refresh_list()
            if self.on_change_callback:
                self.on_change_callback()
    
    def _create_profile(self) -> None:
        """Create a new profile."""
        name = self.new_name_entry.get().strip()
        if not name:
            CTkMessagebox.show(self, "Warning", "Please enter a profile name.", "warning")
            return
        
        if name in self.command_manager.get_profiles():
            CTkMessagebox.show(self, "Warning", "A profile with this name already exists.", "warning")
            return
        
        if self.command_manager.create_profile(name):
            self.new_name_entry.delete(0, "end")
            self._refresh_list()
            if self.on_change_callback:
                self.on_change_callback()
            CTkMessagebox.show(self, "Success", f"Profile '{name}' created.", "success")
        else:
            CTkMessagebox.show(self, "Error", "Failed to create profile.", "error")
    
    def _rename_profile(self, old_name: str) -> None:
        """Rename a profile."""
        dialog = CTkCenteredInputDialog(self, "Rename Profile", f"New name for '{old_name}':")
        new_name = dialog.get_input()
        
        if new_name and new_name.strip():
            new_name = new_name.strip()
            if new_name == old_name:
                return
            if new_name in self.command_manager.get_profiles():
                CTkMessagebox.show(self, "Warning", "A profile with this name already exists.", "warning")
                return
            
            if self.command_manager.rename_profile(old_name, new_name):
                self._refresh_list()
                if self.on_change_callback:
                    self.on_change_callback()
    
    def _delete_profile(self, profile_name: str) -> None:
        """Delete a profile."""
        # Simple confirmation
        confirm = CTkConfirmDialog(self, "Delete Profile", 
                                    f"Are you sure you want to delete '{profile_name}'?\nThis cannot be undone.")
        if confirm.result:
            if self.command_manager.delete_profile(profile_name):
                self._refresh_list()
                if self.on_change_callback:
                    self.on_change_callback()
    
    def _export_current_profile(self) -> None:
        """Export the current profile to a file."""
        from tkinter import filedialog
        
        current = self.command_manager.get_current_profile()
        default_name = f"{current}.json"
        
        file_path = filedialog.asksaveasfilename(
            parent=self,
            title="Export Profile",
            defaultextension=".json",
            initialfile=default_name,
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")]
        )
        
        if file_path:
            if self.command_manager.export_profile(current, file_path):
                CTkMessagebox.show(self, "Success", f"Profile exported to:\n{file_path}", "success")
            else:
                CTkMessagebox.show(self, "Error", "Failed to export profile.", "error")
    
    def _import_profile(self) -> None:
        """Import a profile from a file."""
        from tkinter import filedialog
        
        file_path = filedialog.askopenfilename(
            parent=self,
            title="Import Profile",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")]
        )
        
        if file_path:
            success, result = self.command_manager.import_profile(file_path)
            if success:
                self._refresh_list()
                if self.on_change_callback:
                    self.on_change_callback()
                CTkMessagebox.show(self, "Success", f"Profile '{result}' imported successfully!", "success")
            else:
                CTkMessagebox.show(self, "Error", f"Failed to import: {result}", "error")


class CommandCenterDialog(ctk.CTkToplevel):
    """Dialog for managing built-in command phrases and settings."""
    
    # Category definitions with their command IDs
    CATEGORIES = {
        "volume": {
            "name": "🔊 Volume Control",
            "commands": ["set_volume", "volume_up", "volume_down", "mute", "unmute", "toggle_mute"]
        },
        "media": {
            "name": "🎵 Media Keys",
            "commands": ["play_pause", "next_track", "previous_track", "stop"]
        },
        "timer": {
            "name": "⏱️ Timers",
            "commands": ["set_timer", "stop_timer"],
            "description": "Say 'set a timer for 5 minutes' or 'timer for 30 seconds'"
        },
        "spotify": {
            "name": "🎧 Spotify",
            "commands": ["spotify_play", "spotify_pause", "spotify_next", "spotify_previous", 
                        "spotify_volume_up", "spotify_volume_down", "spotify_shuffle", "spotify_repeat"]
        }
    }
    
    def __init__(self, parent, command_manager: CommandManager, on_change_callback=None):
        super().__init__(parent)
        self.parent = parent
        self.command_manager = command_manager
        self.on_change_callback = on_change_callback
        self.category_frames = {}  # Store expandable content frames
        self.category_expanded = {}  # Track expanded state
        self.category_switches = {}  # Store switch references
        
        self.title("Command Center")
        self.geometry("580x550")
        self.resizable(False, False)
        self.transient(parent)
        self.grab_set()
        set_window_icon(self)
        
        self._create_widgets()
        self._center_window()
    
    def _center_window(self) -> None:
        self.update_idletasks()
        x = self.parent.winfo_x() + (self.parent.winfo_width() - 580) // 2
        y = self.parent.winfo_y() + (self.parent.winfo_height() - 550) // 2
        self.geometry(f"+{x}+{y}")
    
    def _get_disabled_categories(self) -> List[str]:
        """Get list of disabled category IDs from current profile."""
        return self.command_manager.get_disabled_categories()
    
    def _set_category_enabled(self, category_id: str, enabled: bool) -> None:
        """Enable or disable a category."""
        if not self.command_manager.current_profile:
            return
        
        disabled = self.command_manager.get_disabled_categories()
        
        if enabled and category_id in disabled:
            disabled.remove(category_id)
        elif not enabled and category_id not in disabled:
            disabled.append(category_id)
        
        self.command_manager.set_disabled_categories(disabled)
        
        if self.on_change_callback:
            self.on_change_callback()
    
    def _create_widgets(self) -> None:
        # Main frame
        main = ctk.CTkFrame(self, fg_color="transparent")
        main.pack(fill="both", expand=True, padx=20, pady=20)
        
        # Title
        ctk.CTkLabel(main, text="🎛️ Command Center", font=ctk.CTkFont(size=18, weight="bold")).pack(anchor="w")
        ctk.CTkLabel(main, text="Enable/disable command categories and customize trigger phrases", 
                     text_color="gray").pack(anchor="w", pady=(0, 15))
        
        # Scrollable content
        content = ctk.CTkScrollableFrame(main, fg_color="transparent")
        content.pack(fill="both", expand=True, pady=(0, 15))
        
        disabled_categories = self._get_disabled_categories()
        
        # Create collapsible sections for each category
        for cat_id, cat_info in self.CATEGORIES.items():
            self._create_category_section(content, cat_id, cat_info, cat_id not in disabled_categories)
        
        # Info
        info_frame = ctk.CTkFrame(main, fg_color="transparent")
        info_frame.pack(fill="x", pady=(0, 10))
        ctk.CTkLabel(info_frame, text="💡 Disable categories you don't need (e.g., Spotify for gaming)",
                     text_color="gray", font=ctk.CTkFont(size=11)).pack(anchor="w")
        
        # Close button
        ctk.CTkButton(main, text="Close", command=self.destroy, width=100).pack(side="right")
    
    def _create_category_section(self, parent, cat_id: str, cat_info: dict, enabled: bool) -> None:
        """Create a collapsible category section with enable/disable toggle."""
        # Container for the whole section
        section = ctk.CTkFrame(parent, corner_radius=8)
        section.pack(fill="x", pady=5, padx=2)
        
        # Header row with expand button, title, and enable switch
        header = ctk.CTkFrame(section, fg_color="transparent", height=40)
        header.pack(fill="x", padx=10, pady=8)
        header.pack_propagate(False)
        
        # Expand/collapse button
        self.category_expanded[cat_id] = False
        expand_btn = ctk.CTkButton(header, text="▶", width=28, height=28,
                                    fg_color="transparent", hover_color=("gray80", "gray30"),
                                    command=lambda c=cat_id: self._toggle_expand(c))
        expand_btn.pack(side="left")
        self.category_frames[cat_id] = {"expand_btn": expand_btn}
        
        # Category name
        name_label = ctk.CTkLabel(header, text=cat_info["name"], 
                                   font=ctk.CTkFont(size=14, weight="bold"))
        name_label.pack(side="left", padx=(5, 0))
        
        # Enable/disable switch
        switch_var = ctk.BooleanVar(value=enabled)
        switch = ctk.CTkSwitch(header, text="Enabled" if enabled else "Disabled",
                                variable=switch_var, width=50,
                                command=lambda c=cat_id, v=switch_var: self._on_switch_toggle(c, v))
        switch.pack(side="right", padx=5)
        self.category_switches[cat_id] = {"switch": switch, "var": switch_var}
        
        # Show description if available (e.g., for timer category)
        if "description" in cat_info:
            desc_label = ctk.CTkLabel(section, text=cat_info["description"],
                                       text_color="gray", font=ctk.CTkFont(size=11))
            desc_label.pack(anchor="w", padx=45, pady=(0, 5))
        
        # Expandable content frame (hidden by default)
        content_frame = ctk.CTkFrame(section, fg_color="transparent")
        self.category_frames[cat_id]["content"] = content_frame
        self.category_frames[cat_id]["commands"] = cat_info["commands"]
        
        # Commands will be populated when expanded
    
    def _on_switch_toggle(self, cat_id: str, var: ctk.BooleanVar) -> None:
        """Handle enable/disable switch toggle."""
        enabled = var.get()
        self._set_category_enabled(cat_id, enabled)
        
        # Update switch text
        switch = self.category_switches[cat_id]["switch"]
        switch.configure(text="Enabled" if enabled else "Disabled")
    
    def _toggle_expand(self, cat_id: str) -> None:
        """Toggle expand/collapse of a category."""
        expanded = self.category_expanded[cat_id]
        content_frame = self.category_frames[cat_id]["content"]
        expand_btn = self.category_frames[cat_id]["expand_btn"]
        
        if expanded:
            # Collapse
            content_frame.pack_forget()
            expand_btn.configure(text="▶")
            self.category_expanded[cat_id] = False
        else:
            # Expand
            content_frame.pack(fill="x", padx=15, pady=(0, 10))
            expand_btn.configure(text="▼")
            self.category_expanded[cat_id] = True
            
            # Populate commands if not already done
            if not content_frame.winfo_children():
                self._populate_category_commands(cat_id, content_frame)
    
    def _populate_category_commands(self, cat_id: str, content_frame: ctk.CTkFrame) -> None:
        """Populate the command rows for a category."""
        commands = self.category_frames[cat_id]["commands"]
        
        for cmd_id in commands:
            if cmd_id in self.command_manager.builtin_phrases:
                phrases = self.command_manager.builtin_phrases[cmd_id]
                self._create_command_row(content_frame, cmd_id, phrases)
        
        # Add timer settings section for timer category
        if cat_id == "timer":
            self._create_timer_settings_section(content_frame)
    
    def _create_timer_settings_section(self, parent: ctk.CTkFrame) -> None:
        """Create the timer settings section with checkboxes."""
        # Settings section header
        settings_frame = ctk.CTkFrame(parent, fg_color=("gray85", "gray20"), corner_radius=8)
        settings_frame.pack(fill="x", pady=(10, 5))
        
        inner = ctk.CTkFrame(settings_frame, fg_color="transparent")
        inner.pack(fill="x", padx=12, pady=10)
        
        ctk.CTkLabel(inner, text="⚙️ Timer Settings", 
                     font=ctk.CTkFont(size=12, weight="bold")).pack(anchor="w", pady=(0, 8))
        
        # Load current settings
        timer_settings = self.command_manager.get_timer_settings()
        
        # Confirmation sound checkbox
        confirm_var = ctk.BooleanVar(value=timer_settings.get("confirm_sound", True))
        confirm_cb = ctk.CTkCheckBox(inner, text="Play confirmation sound when timer is set",
                                      variable=confirm_var,
                                      command=lambda: self._on_timer_setting_changed("confirm_sound", confirm_var.get()))
        confirm_cb.pack(anchor="w", pady=2)
        
        # Alarm sound checkbox
        alarm_var = ctk.BooleanVar(value=timer_settings.get("alarm_sound", True))
        alarm_cb = ctk.CTkCheckBox(inner, text="Play alarm sound when timer finishes",
                                    variable=alarm_var,
                                    command=lambda: self._on_timer_setting_changed("alarm_sound", alarm_var.get()))
        alarm_cb.pack(anchor="w", pady=(2, 8))
        
        # Volume slider section
        volume_frame = ctk.CTkFrame(inner, fg_color="transparent")
        volume_frame.pack(fill="x", pady=(5, 0))
        
        ctk.CTkLabel(volume_frame, text="Alarm Volume:", font=ctk.CTkFont(size=12)).pack(side="left")
        
        # Volume value label
        current_volume = timer_settings.get("alarm_volume", 100)
        volume_label = ctk.CTkLabel(volume_frame, text=f"{current_volume}%", width=45,
                                     font=ctk.CTkFont(size=12))
        volume_label.pack(side="right", padx=(10, 0))
        
        # Test button
        def test_alarm():
            # Play test alarm sound at current volume
            self.command_manager.play_timer_alarm_once()
        
        ctk.CTkButton(volume_frame, text="🔊 Test", width=70, height=28,
                      command=test_alarm).pack(side="right", padx=(10, 0))
        
        # Volume slider
        def on_volume_change(value):
            vol = int(value)
            volume_label.configure(text=f"{vol}%")
            self._on_timer_setting_changed("alarm_volume", vol)
        
        volume_slider = ctk.CTkSlider(volume_frame, from_=10, to=100, number_of_steps=9,
                                       width=150, command=on_volume_change)
        volume_slider.set(current_volume)
        volume_slider.pack(side="right", padx=(15, 0))
    
    def _on_timer_setting_changed(self, setting_key: str, value: bool) -> None:
        """Handle timer setting change."""
        settings = self.command_manager.get_timer_settings()
        settings[setting_key] = value
        self.command_manager.set_timer_settings(settings)
        
        if self.on_change_callback:
            self.on_change_callback()
    
    def _create_command_row(self, parent, cmd_id: str, phrases: List[str]) -> None:
        """Create a row for a built-in command."""
        row = ctk.CTkFrame(parent, fg_color=("gray90", "gray17"), corner_radius=6, height=36)
        row.pack(fill="x", pady=3)
        row.pack_propagate(False)
        
        # Command name (formatted)
        display_name = cmd_id.replace("_", " ").title()
        name_label = ctk.CTkLabel(row, text=display_name, width=130, anchor="w",
                                   font=ctk.CTkFont(size=12))
        name_label.pack(side="left", padx=(10, 10))
        
        # Phrases (truncated)
        phrases_text = ", ".join(phrases)
        if len(phrases_text) > 28:
            phrases_text = phrases_text[:25] + "..."
        
        phrases_label = ctk.CTkLabel(row, text=f'"{phrases_text}"', text_color="gray",
                                      font=ctk.CTkFont(size=11), anchor="w", width=180)
        phrases_label.pack(side="left", padx=(0, 10))
        
        # Edit button
        ctk.CTkButton(row, text="Edit", width=55, height=24,
                      fg_color="transparent", border_width=1,
                      command=lambda c=cmd_id: self._edit_phrases(c)).pack(side="right", padx=8)
    
    def _edit_phrases(self, cmd_id: str) -> None:
        """Edit the phrases for a built-in command."""
        current_phrases = self.command_manager.builtin_phrases.get(cmd_id, [])
        display_name = cmd_id.replace("_", " ").title()
        
        dialog = EditPhrasesDialog(self, display_name, current_phrases)
        
        if dialog.result is not None:
            self.command_manager.builtin_phrases[cmd_id] = dialog.result
            self.command_manager._save_current_profile()
            
            # Refresh the command row
            for cat_id, frame_info in self.category_frames.items():
                if cmd_id in frame_info.get("commands", []):
                    content_frame = frame_info["content"]
                    if content_frame.winfo_children():
                        # Clear and repopulate
                        for widget in content_frame.winfo_children():
                            widget.destroy()
                        self._populate_category_commands(cat_id, content_frame)
                    break
            
            if self.on_change_callback:
                self.on_change_callback()


class EditPhrasesDialog(ctk.CTkToplevel):
    """Dialog to edit trigger phrases for a command."""
    
    def __init__(self, parent, command_name: str, phrases: List[str]):
        super().__init__(parent)
        self.result = None
        
        self.title(f"Edit Phrases - {command_name}")
        self.geometry("400x300")
        self.resizable(False, False)
        self.transient(parent)
        self.grab_set()
        set_window_icon(self)
        
        # Center on parent
        self.update_idletasks()
        x = parent.winfo_x() + (parent.winfo_width() - 400) // 2
        y = parent.winfo_y() + (parent.winfo_height() - 300) // 2
        self.geometry(f"+{x}+{y}")
        
        frame = ctk.CTkFrame(self, fg_color="transparent")
        frame.pack(fill="both", expand=True, padx=20, pady=20)
        
        ctk.CTkLabel(frame, text=f"Trigger phrases for '{command_name}':",
                     font=ctk.CTkFont(size=13)).pack(anchor="w", pady=(0, 5))
        ctk.CTkLabel(frame, text="One phrase per line", text_color="gray",
                     font=ctk.CTkFont(size=11)).pack(anchor="w", pady=(0, 10))
        
        self.text = ctk.CTkTextbox(frame, height=150, font=ctk.CTkFont(size=12))
        self.text.pack(fill="both", expand=True, pady=(0, 15))
        self.text.insert("1.0", "\n".join(phrases))
        
        btn_frame = ctk.CTkFrame(frame, fg_color="transparent")
        btn_frame.pack(fill="x")
        
        ctk.CTkButton(btn_frame, text="Cancel", command=self.destroy, width=100,
                      fg_color="transparent", border_width=1).pack(side="left", padx=(0, 10))
        ctk.CTkButton(btn_frame, text="Save", command=self._save, width=100).pack(side="right")
        
        self.wait_window()
    
    def _save(self) -> None:
        text = self.text.get("1.0", "end-1c")
        phrases = [p.strip() for p in text.split("\n") if p.strip()]
        self.result = phrases
        self.destroy()


class CTkCenteredInputDialog(ctk.CTkToplevel):
    """Input dialog that centers on parent window."""
    
    def __init__(self, parent, title: str, text: str):
        super().__init__(parent)
        self.result = None
        
        self.title(title)
        self.geometry("400x180")
        self.resizable(False, False)
        self.transient(parent)
        self.grab_set()
        set_window_icon(self)
        
        # Center on parent
        self.update_idletasks()
        x = parent.winfo_x() + (parent.winfo_width() - 400) // 2
        y = parent.winfo_y() + (parent.winfo_height() - 180) // 2
        self.geometry(f"+{x}+{y}")
        
        frame = ctk.CTkFrame(self, fg_color="transparent")
        frame.pack(fill="both", expand=True, padx=25, pady=25)
        
        ctk.CTkLabel(frame, text=text, wraplength=350).pack(pady=(0, 15))
        
        self.entry = ctk.CTkEntry(frame, width=300)
        self.entry.pack(pady=(0, 20))
        self.entry.focus_set()
        self.entry.bind("<Return>", lambda e: self._ok())
        
        btn_frame = ctk.CTkFrame(frame, fg_color="transparent")
        btn_frame.pack(fill="x")
        
        ctk.CTkButton(btn_frame, text="Cancel", command=self.destroy, width=100,
                      fg_color="transparent", border_width=1).pack(side="left", expand=True, padx=5)
        ctk.CTkButton(btn_frame, text="OK", command=self._ok, width=100).pack(side="right", expand=True, padx=5)
        
        self.wait_window()
    
    def _ok(self) -> None:
        self.result = self.entry.get()
        self.destroy()
    
    def get_input(self) -> str:
        return self.result


class CTkConfirmDialog(ctk.CTkToplevel):
    """Simple confirmation dialog."""
    
    def __init__(self, parent, title: str, message: str):
        super().__init__(parent)
        self.result = False
        
        self.title(title)
        self.geometry("350x150")
        self.resizable(False, False)
        self.transient(parent)
        self.grab_set()
        set_window_icon(self)
        
        # Center on parent
        self.update_idletasks()
        x = parent.winfo_x() + (parent.winfo_width() - 350) // 2
        y = parent.winfo_y() + (parent.winfo_height() - 150) // 2
        self.geometry(f"+{x}+{y}")
        
        frame = ctk.CTkFrame(self, fg_color="transparent")
        frame.pack(fill="both", expand=True, padx=25, pady=25)
        
        ctk.CTkLabel(frame, text=message, wraplength=300).pack(pady=(0, 20))
        
        btn_frame = ctk.CTkFrame(frame, fg_color="transparent")
        btn_frame.pack(fill="x")
        
        ctk.CTkButton(btn_frame, text="Cancel", command=self.destroy, width=100,
                      fg_color="transparent", border_width=1).pack(side="left", expand=True, padx=5)
        ctk.CTkButton(btn_frame, text="Delete", command=self._confirm, width=100,
                      fg_color="#e74c3c", hover_color="#c0392b").pack(side="right", expand=True, padx=5)
        
        self.wait_window()
    
    def _confirm(self) -> None:
        self.result = True
        self.destroy()


class VoiceControlApp(ctk.CTk):
    """Main application with CustomTkinter modern UI."""
    
    def __init__(self):
        super().__init__()
        
        self.title("Voice Control")
        self.geometry("800x720")
        self.minsize(700, 680)
        set_window_icon(self)
        
        # Initialize
        self.recognizer = VoiceRecognizer()
        self.command_manager = CommandManager()
        
        self.recognizer.on_recognized = self._on_speech_recognized
        self.recognizer.on_error = self._on_error
        self.recognizer.on_listening_state_changed = self._on_listening_state_changed
        self._executing_voice_command = False
        self._command_lock = threading.Lock()
        
        self.current_shortcut = self._load_shortcut()
        
        # System tray
        self.tray_icon = None
        self._setup_tray()
        
        self._create_widgets()
        self._register_shortcut()
        self._center_window()
        
        self.protocol("WM_DELETE_WINDOW", self._on_close)
        self.bind("<Unmap>", self._on_minimize)
    
    def _center_window(self) -> None:
        self.update_idletasks()
        x = (self.winfo_screenwidth() - self.winfo_width()) // 2
        y = (self.winfo_screenheight() - self.winfo_height()) // 2
        self.geometry(f"+{x}+{y}")
    
    def _load_shortcut(self) -> str:
        config_path = get_config_path()
        if config_path.exists():
            try:
                with open(config_path, 'r') as f:
                    return json.load(f).get("settings", {}).get("toggle_shortcut", "ctrl+shift+v")
            except:
                pass
        return "ctrl+shift+v"
    
    def _create_widgets(self) -> None:
        # Sidebar
        sidebar = ctk.CTkFrame(self, width=220, corner_radius=0)
        sidebar.pack(side="left", fill="y")
        sidebar.pack_propagate(False)
        
        # Logo/Title in sidebar
        ctk.CTkLabel(sidebar, text="🎤", font=ctk.CTkFont(size=40)).pack(pady=(30, 5))
        ctk.CTkLabel(sidebar, text="Voice Control", font=ctk.CTkFont(size=18, weight="bold")).pack(pady=(0, 20))
        
        # Shortcut label (above button)
        self.shortcut_label = ctk.CTkLabel(sidebar, text=f"Shortcut: {self.current_shortcut}",
                                            font=ctk.CTkFont(size=11), text_color="gray")
        self.shortcut_label.pack(pady=(0, 8))
        
        # Toggle button
        self.toggle_btn = ctk.CTkButton(sidebar, text="▶  Start Listening", command=self._toggle_listening,
                                         height=45, font=ctk.CTkFont(size=14), corner_radius=10)
        self.toggle_btn.pack(padx=20, fill="x")
        
        # Status
        self.status_var = ctk.StringVar(value="Not listening")
        self.status_label = ctk.CTkLabel(sidebar, textvariable=self.status_var, text_color="gray")
        self.status_label.pack(pady=(10, 20))
        
        # Profile selector
        profile_frame = ctk.CTkFrame(sidebar, fg_color="transparent")
        profile_frame.pack(fill="x", padx=15, pady=(0, 10))
        
        ctk.CTkLabel(profile_frame, text="Profile:", font=ctk.CTkFont(size=12)).pack(side="left")
        
        profiles = self.command_manager.get_profiles()
        current_profile = self.command_manager.get_current_profile()
        self.profile_var = ctk.StringVar(value=current_profile if current_profile else "(No profiles)")
        self.profile_combo = ctk.CTkComboBox(profile_frame, variable=self.profile_var,
                                              values=profiles if profiles else ["(No profiles)"],
                                              command=self._on_profile_changed,
                                              width=100, height=28)
        self.profile_combo.pack(side="left", padx=(8, 5))
        
        ctk.CTkButton(profile_frame, text="⋮", command=self._open_profile_manager,
                      width=28, height=28, fg_color="transparent", border_width=1).pack(side="left")
        
        # Navigation buttons
        ctk.CTkButton(sidebar, text="⚙️  Settings", command=self._open_settings,
                      fg_color="transparent", anchor="w", height=40).pack(padx=15, fill="x", pady=2)
        
        # Now Playing card (Spotify)
        self._create_now_playing_card(sidebar)
        
        # Admin warning (if not running as admin)
        if not is_admin():
            admin_warning = ctk.CTkLabel(sidebar, text="⚠️ Problems? Run as Administrator",
                                          font=ctk.CTkFont(size=10), text_color="#d4a017",
                                          cursor="hand2")
            admin_warning.pack(side="bottom", pady=(0, 5))
            admin_warning.bind("<Button-1>", lambda e: self._show_admin_info())
        
        # Legal links at bottom of sidebar
        legal_frame = ctk.CTkFrame(sidebar, fg_color="transparent")
        legal_frame.pack(side="bottom", fill="x", padx=20, pady=(0, 5))
        
        eula_link = ctk.CTkLabel(legal_frame, text="EULA", font=ctk.CTkFont(size=10),
                                  text_color="gray", cursor="hand2")
        eula_link.pack(side="left")
        eula_link.bind("<Button-1>", lambda e: self._open_legal_doc("EULA.md"))
        
        ctk.CTkLabel(legal_frame, text=" · ", font=ctk.CTkFont(size=10), text_color="gray").pack(side="left")
        
        privacy_link = ctk.CTkLabel(legal_frame, text="Privacy", font=ctk.CTkFont(size=10),
                                     text_color="gray", cursor="hand2")
        privacy_link.pack(side="left")
        privacy_link.bind("<Button-1>", lambda e: self._open_legal_doc("PRIVACY.md"))
        
        # Theme switch at bottom
        theme_frame = ctk.CTkFrame(sidebar, fg_color="transparent")
        theme_frame.pack(side="bottom", fill="x", padx=20, pady=15)
        
        ctk.CTkLabel(theme_frame, text="Theme").pack(side="left")
        self.theme_switch = ctk.CTkSwitch(theme_frame, text="", command=self._toggle_theme, 
                                           onvalue="light", offvalue="dark", width=40)
        self.theme_switch.pack(side="right")
        
        # Main content
        main = ctk.CTkFrame(self, fg_color="transparent")
        main.pack(side="right", fill="both", expand=True, padx=25, pady=25)
        
        # Tabview
        self.tabview = ctk.CTkTabview(main, corner_radius=15)
        self.tabview.pack(fill="both", expand=True)
        
        self.tabview.add("Commands")
        self.tabview.add("Activity")
        
        self._create_commands_tab(self.tabview.tab("Commands"))
        self._create_log_tab(self.tabview.tab("Activity"))
    
    def _create_commands_tab(self, parent) -> None:
        # Toolbar
        toolbar = ctk.CTkFrame(parent, fg_color="transparent")
        toolbar.pack(fill="x", pady=(0, 15))
        
        ctk.CTkButton(toolbar, text="+ Add Command", command=self._add_command, width=130).pack(side="left", padx=(0, 10))
        ctk.CTkButton(toolbar, text="Edit", command=self._edit_command, width=80,
                      fg_color="transparent", border_width=1).pack(side="left", padx=(0, 10))
        ctk.CTkButton(toolbar, text="Remove", command=self._remove_command, width=80,
                      fg_color="transparent", border_width=1).pack(side="left")
        
        # Commands frame
        self.commands_frame = ctk.CTkScrollableFrame(parent, corner_radius=10)
        self.commands_frame.pack(fill="both", expand=True)
        
        # Bottom toolbar with Command Center button
        bottom_bar = ctk.CTkFrame(parent, fg_color="transparent")
        bottom_bar.pack(fill="x", pady=(10, 0))
        
        ctk.CTkButton(bottom_bar, text="⚙️ Command Center", command=self._open_command_center,
                      width=150, fg_color="#6b46c1", hover_color="#553c9a").pack(side="right")
        
        self._refresh_commands()
    
    def _create_log_tab(self, parent) -> None:
        self.log_text = ctk.CTkTextbox(parent, font=ctk.CTkFont(family="Consolas", size=12), corner_radius=10)
        self.log_text.pack(fill="both", expand=True, pady=(0, 10))
        
        ctk.CTkButton(parent, text="Clear Log", command=self._clear_log, width=100,
                      fg_color="transparent", border_width=1).pack(anchor="e")
    
    def _toggle_theme(self) -> None:
        if self.theme_switch.get() == "light":
            ctk.set_appearance_mode("light")
        else:
            ctk.set_appearance_mode("dark")
    
    def _show_admin_info(self) -> None:
        """Show information about running as administrator."""
        CTkMessagebox.show(
            self, 
            "Administrator Mode",
            "Some features work better with administrator privileges:\n\n"
            "• Global keyboard shortcuts in all apps\n"
            "• Window management in elevated apps\n"
            "• Media key simulation\n\n"
            "Right-click the app and select 'Run as administrator'.",
            "info"
        )
    
    def _open_legal_doc(self, filename: str) -> None:
        """Open a legal document (EULA or Privacy Policy)."""
        doc_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), filename)
        if os.path.exists(doc_path):
            os.startfile(doc_path)
        else:
            CTkMessagebox.show(self, "Not Found", f"Could not find {filename}.", "warning")
    
    def _refresh_commands(self) -> None:
        for widget in self.commands_frame.winfo_children():
            widget.destroy()

        # Reset selection state because previous card widgets are destroyed.
        self.selected_card = None
        self.selected_command = None
        
        if not self.command_manager.commands:
            ctk.CTkLabel(self.commands_frame, text="No custom commands yet.\nClick '+ Add Command' to create one.",
                         text_color="gray").pack(pady=40)
            return
        
        for cmd_id, cmd_data in self.command_manager.commands.items():
            card = ctk.CTkFrame(self.commands_frame, corner_radius=10, height=70)
            card.pack(fill="x", pady=(0, 8))
            card.pack_propagate(False)
            
            inner = ctk.CTkFrame(card, fg_color="transparent")
            inner.pack(fill="both", expand=True, padx=15, pady=10)
            
            if isinstance(cmd_data, dict):
                phrases = ", ".join(cmd_data.get("phrases", [])[:2])
                if len(cmd_data.get("phrases", [])) > 2:
                    phrases += "..."
                cmd_type = cmd_data.get("type", "open_file")
                cmd_action_data = cmd_data.get("data")
                
                type_labels = {
                    "open_file": "Open File/App",
                    "open_folder": "Open Folder",
                    "open_url": "Open Website",
                    "macro": "Keyboard Macro",
                    "chain": "Chain Action",
                    "window_action": "Window Action",
                }
                type_label = type_labels.get(cmd_type, cmd_type)
                
                # Add detail for window actions
                if cmd_type == "window_action" and cmd_action_data:
                    action_names = {
                        "minimize": "Minimize", "maximize": "Maximize",
                        "restore": "Restore", "focus": "Focus", "close": "Close",
                        "close_all_app": "Close All (App)", "close_all_windows": "Close All",
                        "snap_left": "Snap Left", "snap_right": "Snap Right",
                        "snap_top_left": "Snap Top-Left", "snap_top_right": "Snap Top-Right",
                        "snap_bottom_left": "Snap Bottom-Left", "snap_bottom_right": "Snap Bottom-Right",
                    }
                    if isinstance(cmd_action_data, dict):
                        action = cmd_action_data.get("action", "")
                        target = cmd_action_data.get("target", "focused")
                        target_name = cmd_action_data.get("app", "") or cmd_action_data.get("title", "")
                    else:
                        action = str(cmd_action_data)
                        target = "focused"
                        target_name = ""
                    action_name = action_names.get(action, action)
                    if target_name:
                        type_label = f"Window Action > {action_name} > {target_name}"
                    else:
                        type_label = f"Window Action > {action_name}"
                elif cmd_type == "open_file" and cmd_action_data:
                    # Show the filename/app name
                    display_path = os.path.basename(str(cmd_action_data))
                    type_label = f"Open File/App > {display_path}"
                elif cmd_type == "open_folder" and cmd_action_data:
                    display_path = str(cmd_action_data)
                    if len(display_path) > 40:
                        display_path = "..." + display_path[-37:]
                    type_label = f"Open Folder > {display_path}"
                elif cmd_type == "open_url" and cmd_action_data:
                    # Show the URL (trimmed)
                    url_display = str(cmd_action_data).replace("https://", "").replace("http://", "")
                    if len(url_display) > 40:
                        url_display = url_display[:37] + "..."
                    type_label = f"Open Website > {url_display}"
                elif cmd_type == "chain" and isinstance(cmd_action_data, list):
                    type_label = f"Chain Action > {len(cmd_action_data)} step{'s' if len(cmd_action_data) != 1 else ''}"
            else:
                phrases = cmd_id
                type_label = "Open File/App"
            
            left = ctk.CTkFrame(inner, fg_color="transparent")
            left.pack(side="left", fill="y")
            
            label = ctk.CTkLabel(left, text=phrases, font=ctk.CTkFont(size=13, weight="bold"),
                         anchor="w")
            label.pack(anchor="w")
            
            sub_label = ctk.CTkLabel(left, text=type_label, font=ctk.CTkFont(size=11),
                         text_color="gray", anchor="w")
            sub_label.pack(anchor="w")
            
            # Store cmd_id for selection
            card.cmd_id = cmd_id
            
            # Bind click to card and all children so clicking anywhere selects
            select_handler = lambda e, c=card: self._select_command(c)
            card.bind("<Button-1>", select_handler)
            inner.bind("<Button-1>", select_handler)
            left.bind("<Button-1>", select_handler)
            label.bind("<Button-1>", select_handler)
            sub_label.bind("<Button-1>", select_handler)
            
        self.selected_command = None
    
    def _select_command(self, card) -> None:
        # Deselect previous
        if hasattr(self, 'selected_card') and self.selected_card:
            try:
                if self.selected_card.winfo_exists():
                    self.selected_card.configure(border_width=0)
            except Exception:
                pass
        
        # Select new
        card.configure(border_width=2, border_color="#1f6aa5")
        self.selected_card = card
        self.selected_command = card.cmd_id
    
    def _add_command(self) -> None:
        # Check if a profile exists
        if not self.command_manager.has_profiles():
            self._prompt_create_first_profile()
            return
        
        dialog = AddCommandDialog(self, self.command_manager)
        self.wait_window(dialog)
        if dialog.result:
            self._refresh_commands()
            self._log("Command added")
    
    def _prompt_create_first_profile(self) -> None:
        """Prompt user to create a profile when none exist."""
        dialog = CTkCenteredInputDialog(self, "Create Profile", 
                                         "No profiles exist. Enter a name to create your first profile:")
        name = dialog.get_input()
        
        if name and name.strip():
            if self.command_manager.create_profile(name.strip()):
                self.command_manager.switch_profile(name.strip())
                self._refresh_profiles()
                self._log(f"Created profile: {name.strip()}")
                CTkMessagebox.show(self, "Success", f"Profile '{name.strip()}' created. You can now add commands.", "success")
            else:
                CTkMessagebox.show(self, "Error", "Failed to create profile.", "error")
    
    def _edit_command(self) -> None:
        if not hasattr(self, 'selected_command') or not self.selected_command:
            CTkMessagebox.show(self, "Warning", "Select a command to edit.", "warning")
            return
        
        dialog = AddCommandDialog(self, self.command_manager, edit_command_id=self.selected_command)
        self.wait_window(dialog)
        if dialog.result:
            self._refresh_commands()
            self._log("Command updated")
    
    def _remove_command(self) -> None:
        if not hasattr(self, 'selected_command') or not self.selected_command:
            CTkMessagebox.show(self, "Warning", "Select a command to remove.", "warning")
            return
        
        if CTkMessagebox.ask_yes_no(self, "Confirm", "Remove this command?"):
            self.command_manager.remove_command(self.selected_command)
            self.selected_command = None
            self._refresh_commands()
            self._log("Command removed")
    
    def _toggle_listening(self) -> None:
        self.recognizer.toggle_listening()
    
    def _on_listening_state_changed(self, is_listening: bool) -> None:
        def update():
            if is_listening:
                self.status_var.set("● Listening...")
                self.status_label.configure(text_color="#2ecc71")
                self.toggle_btn.configure(text="⏹  Stop Listening", fg_color="#e74c3c", hover_color="#c0392b")
            else:
                self.status_var.set("Not listening")
                self.status_label.configure(text_color="gray")
                self.toggle_btn.configure(text="▶  Start Listening", fg_color="#1f6aa5", hover_color="#144870")
        self.after(0, update)
    
    def _on_speech_recognized(self, text: str) -> None:
        self.after(0, lambda: self._log(f"Heard: {text}"))

        # Run command execution off the UI thread so cards/buttons stay clickable.
        def run_command() -> None:
            with self._command_lock:
                if self._executing_voice_command:
                    return
                self._executing_voice_command = True

            try:
                success, message = self.command_manager.execute_command(text)
                self.after(0, lambda: self._log(f"{'✅' if success else '❌'} {message}"))
                # Brief cooldown after successful commands to prevent
                # the recognizer from picking up audio playback or
                # residual speech and triggering a follow-up command
                # (e.g. "play X" starts music, then "play" is heard
                # again and toggles pause).
                if success:
                    import time
                    time.sleep(1.5)
            finally:
                with self._command_lock:
                    self._executing_voice_command = False

        threading.Thread(target=run_command, daemon=True).start()
    
    def _on_error(self, error: str) -> None:
        self.after(0, lambda: self._log(f"❌ Error: {error}"))
    
    def _log(self, message: str) -> None:
        import datetime
        timestamp = datetime.datetime.now().strftime("%H:%M:%S")
        self.log_text.insert("end", f"[{timestamp}] {message}\n")
        self.log_text.see("end")
    
    def _clear_log(self) -> None:
        self.log_text.delete("1.0", "end")
    
    def _create_now_playing_card(self, parent) -> None:
        """Create the Now Playing card in the sidebar."""
        # Now Playing frame
        self.now_playing_frame = ctk.CTkFrame(parent, corner_radius=12, height=235)
        self.now_playing_frame.pack(padx=15, pady=(20, 10), fill="x")
        self.now_playing_frame.pack_propagate(False)
        
        # Inner padding
        inner = ctk.CTkFrame(self.now_playing_frame, fg_color="transparent")
        inner.pack(fill="both", expand=True, padx=12, pady=10)
        
        # Title row with Spotify attribution
        title_row = ctk.CTkFrame(inner, fg_color="transparent")
        title_row.pack(fill="x", anchor="w")
        ctk.CTkLabel(title_row, text="Now Playing", font=ctk.CTkFont(size=11), 
                     text_color="gray").pack(side="left")
        ctk.CTkLabel(title_row, text="· Spotify", font=ctk.CTkFont(size=11), 
                     text_color="#1DB954").pack(side="left", padx=(3, 0))
        
        # Album art placeholder (clickable)
        self.album_art_label = ctk.CTkLabel(inner, text="🎵", font=ctk.CTkFont(size=40),
                                             width=80, height=80, cursor="hand2")
        self.album_art_label.pack(pady=(5, 5))
        self.album_art_label.bind("<Button-1>", lambda e: self._open_spotify_track())
        self._current_album_art = None
        self._current_track_url = None
        
        # Track name (clickable)
        self.np_track_var = ctk.StringVar(value="Not playing")
        self.np_track_label = ctk.CTkLabel(inner, textvariable=self.np_track_var, 
                                            font=ctk.CTkFont(size=12, weight="bold"),
                                            wraplength=170, cursor="hand2")
        self.np_track_label.pack(anchor="w")
        self.np_track_label.bind("<Button-1>", lambda e: self._open_spotify_track())
        
        # Artist
        self.np_artist_var = ctk.StringVar(value="")
        self.np_artist_label = ctk.CTkLabel(inner, textvariable=self.np_artist_var,
                                             font=ctk.CTkFont(size=11), text_color="gray",
                                             wraplength=170)
        self.np_artist_label.pack(anchor="w")
        
        # Progress bar
        self.np_progress = ctk.CTkProgressBar(inner, height=4, corner_radius=2)
        self.np_progress.pack(fill="x", pady=(10, 5))
        self.np_progress.set(0)
        
        # Time remaining
        self.np_time_var = ctk.StringVar(value="")
        ctk.CTkLabel(inner, textvariable=self.np_time_var, 
                     font=ctk.CTkFont(size=10), text_color="gray").pack(anchor="e")
        
        # Start update loop
        self._update_now_playing()
    
    def _update_now_playing(self) -> None:
        """Update the Now Playing card with current track info."""
        if self.command_manager.spotify and self.command_manager.spotify.is_authenticated:
            try:
                track = self.command_manager.spotify.get_current_track()
                if track:
                    # Update track info
                    name = track["name"]
                    if len(name) > 25:
                        name = name[:22] + "..."
                    self.np_track_var.set(name)
                    
                    artist = track["artist"]
                    if len(artist) > 28:
                        artist = artist[:25] + "..."
                    self.np_artist_var.set(artist)
                    
                    # Update progress
                    progress = track.get("progress_percent", 0) / 100
                    self.np_progress.set(progress)
                    
                    remaining = track.get("remaining", "0:00")
                    status = "▶" if track["is_playing"] else "⏸"
                    self.np_time_var.set(f"{status} -{remaining}")
                    
                    # Store track URL for link-back
                    self._current_track_url = track.get("track_url", "")
                    
                    # Update album art
                    self._load_album_art(track.get("album_art_url"))
                else:
                    self._reset_now_playing(connected=True)
            except Exception:
                self._reset_now_playing(connected=True)
        else:
            self._reset_now_playing(connected=False)
        
        # Update every 2 seconds
        self.after(2000, self._update_now_playing)
    
    def _reset_now_playing(self, connected: bool = False) -> None:
        """Reset Now Playing to default state."""
        self.np_track_var.set("Not playing")
        if connected:
            self.np_artist_var.set("🟢 Connected")
        else:
            self.np_artist_var.set("🔴 Not connected")
        self.np_progress.set(0)
        self.np_time_var.set("")
        self.album_art_label.configure(image=None, text="🎵")
        self._current_album_art = None
        self._current_track_url = None
    
    def _open_spotify_track(self) -> None:
        """Open the current track in Spotify."""
        if self._current_track_url:
            webbrowser.open(self._current_track_url)
    
    def _load_album_art(self, url: str) -> None:
        """Load album art from URL in background."""
        if not url or not PIL_AVAILABLE:
            return
        
        # Skip if same URL
        if hasattr(self, '_last_art_url') and self._last_art_url == url:
            return
        self._last_art_url = url
        
        def load():
            try:
                with urllib.request.urlopen(url, timeout=5) as response:
                    image_data = response.read()
                
                image = Image.open(io.BytesIO(image_data))
                image = image.resize((75, 75), Image.LANCZOS)
                
                # Update on main thread
                self.after(0, lambda: self._set_album_art(image))
            except Exception:
                pass
        
        threading.Thread(target=load, daemon=True).start()
    
    def _set_album_art(self, image: "Image.Image") -> None:
        """Set the album art image on the label."""
        try:
            ctk_image = ctk.CTkImage(light_image=image, dark_image=image, size=(75, 75))
            self._current_album_art = ctk_image  # Keep reference
            self.album_art_label.configure(image=ctk_image, text="")
        except Exception:
            pass
    
    def _open_settings(self) -> None:
        SettingsDialog(self, self.recognizer, self.command_manager)
    
    def _register_shortcut(self) -> None:
        try:
            keyboard.add_hotkey(self.current_shortcut, self._toggle_listening)
        except Exception as e:
            self._log(f"Shortcut error: {e}")
    
    def update_shortcut(self, new_shortcut: str) -> None:
        try:
            keyboard.remove_hotkey(self.current_shortcut)
        except:
            pass
        self.current_shortcut = new_shortcut
        self._register_shortcut()
        self.shortcut_label.configure(text=f"Shortcut: {new_shortcut}")
    
    def _setup_tray(self) -> None:
        """Setup the system tray icon."""
        if not TRAY_AVAILABLE or not PIL_AVAILABLE:
            return
        
        # Create tray icon image
        icon_path = get_app_icon_path()
        if icon_path:
            try:
                tray_image = Image.open(icon_path)
                tray_image = tray_image.resize((64, 64), Image.LANCZOS)
            except:
                # Create a simple colored icon if loading fails
                tray_image = Image.new('RGB', (64, 64), color=(30, 144, 255))
        else:
            # Create a simple colored icon
            tray_image = Image.new('RGB', (64, 64), color=(30, 144, 255))
        
        # Create menu
        menu = pystray.Menu(
            item('Show', self._show_window, default=True),
            item('Toggle Listening', self._tray_toggle_listening),
            pystray.Menu.SEPARATOR,
            item('Exit', self._quit_app)
        )
        
        self.tray_icon = pystray.Icon("VoiceControl", tray_image, "Voice Control", menu)
    
    def _on_minimize(self, event) -> None:
        """Handle window minimize - hide to tray."""
        if event.widget == self and self.state() == 'iconic':
            if TRAY_AVAILABLE and self.tray_icon:
                self.withdraw()  # Hide the window
                if not self.tray_icon.visible:
                    # Run tray icon in background thread
                    threading.Thread(target=self.tray_icon.run, daemon=True).start()
    
    def _show_window(self, icon=None, item=None) -> None:
        """Show the window from tray."""
        self.after(0, self._restore_window)
    
    def _restore_window(self) -> None:
        """Restore the window."""
        self.deiconify()
        self.state('normal')
        self.lift()
        self.focus_force()
    
    def _tray_toggle_listening(self, icon=None, item=None) -> None:
        """Toggle listening from tray."""
        self.after(0, self._toggle_listening)
    
    def _quit_app(self, icon=None, item=None) -> None:
        """Quit the application from tray."""
        if self.tray_icon:
            self.tray_icon.stop()
        self.after(0, self._on_close)
    
    def _on_profile_changed(self, profile_name: str) -> None:
        """Handle profile selection change."""
        # Ignore placeholder text
        if profile_name == "(No profiles)":
            return
        
        if profile_name != self.command_manager.get_current_profile():
            if self.command_manager.switch_profile(profile_name):
                self._refresh_commands()
                self._log(f"Switched to profile: {profile_name}")
            else:
                # Revert combo box if switch failed
                current = self.command_manager.get_current_profile()
                self.profile_var.set(current if current else "(No profiles)")
    
    def _open_profile_manager(self) -> None:
        """Open the profile manager dialog."""
        ProfileManagerDialog(self, self.command_manager, self._refresh_profiles)
    
    def _open_command_center(self) -> None:
        """Open the command center settings dialog."""
        CommandCenterDialog(self, self.command_manager, self._refresh_commands)
    
    def _refresh_profiles(self) -> None:
        """Refresh the profile combo box."""
        profiles = self.command_manager.get_profiles()
        if profiles:
            self.profile_combo.configure(values=profiles)
            self.profile_var.set(self.command_manager.get_current_profile())
        else:
            self.profile_combo.configure(values=["(No profiles)"])
            self.profile_var.set("(No profiles)")
        self._refresh_commands()
    
    def _on_close(self) -> None:
        if self.recognizer.is_listening:
            self.recognizer.stop_listening()
        try:
            keyboard.unhook_all()
        except:
            pass
        if self.tray_icon and self.tray_icon.visible:
            self.tray_icon.stop()
        self.destroy()


def main():
    app = VoiceControlApp()
    app.mainloop()


if __name__ == "__main__":
    main()
