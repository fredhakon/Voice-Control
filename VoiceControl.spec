# -*- mode: python ; coding: utf-8 -*-
"""
PyInstaller spec file for Voice Control Platform
"""

import sys
from pathlib import Path
import importlib

block_cipher = None

# Get the directory containing this spec file
spec_dir = Path(SPECPATH)

# Locate Vosk native DLLs
vosk_dir = str(Path(importlib.import_module('vosk').__file__).parent)

a = Analysis(
    ['main.py'],
    pathex=[str(spec_dir)],
    binaries=[
        (vosk_dir + '/*.dll', 'vosk'),
    ],
    datas=[
        ('icon.ico', '.'),
        ('config.default.json', '.'),  # Use clean default config, not user's personal config
        ('vosk-model', 'vosk-model'),  # Vosk offline speech recognition model
        ('EULA.md', '.'),
        ('PRIVACY.md', '.'),
    ],
    hiddenimports=[
        'pycaw',
        'pycaw.pycaw',
        'comtypes',
        'comtypes.client',
        'speech_recognition',
        'pyaudio',
        'keyboard',
        'pyautogui',
        'PIL',
        'PIL.Image',
        'PIL.ImageDraw',
        'spotipy',
        'spotipy.oauth2',
        'spotify_control',
        'customtkinter',
        'darkdetect',
        'pystray',
        'pystray._win32',
        'vosk',
        'faster_whisper',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

# Strip user-specific / secret files that should never ship in a release build
_exclude = {'.spotify_cache', 'config.json'}
a.datas = [d for d in a.datas
           if d[0] not in _exclude and not d[0].startswith('profiles/') and not d[0].startswith('profiles\\')]

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='VoiceControl',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,  # No console window - GUI app
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='icon.ico',  # Application icon
    uac_admin=False,  # Set to True if you need admin privileges for some features
)
