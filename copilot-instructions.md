<!-- Use this file to provide workspace-specific custom instructions to Copilot. For more details, visit https://code.visualstudio.com/docs/copilot/copilot-customization#_use-a-githubcopilotinstructionsmd-file -->

# Voice Control Platform

This is a Windows voice control platform built with Python.

## Technologies Used
- **Speech Recognition**: Using the `SpeechRecognition` library with Google Speech API
- **Audio**: PyAudio for microphone input
- **System Volume Control**: pycaw library for Windows audio control
- **Media Keys**: pyautogui for simulating media key presses
- **Keyboard Shortcuts**: keyboard library for global hotkey registration
- **GUI**: Tkinter for the user interface

## Key Features
- Voice command recognition with toggle listening
- Pre-built system commands (volume, media controls)
- Custom commands for opening files/applications
- Settings menu for microphone and shortcut configuration
- Global keyboard shortcut to toggle listening

## Code Guidelines
- Use type hints for function parameters and return values
- Follow PEP 8 style guidelines
- Handle audio device errors gracefully
- Use threading for voice recognition to keep UI responsive
