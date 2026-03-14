# Privacy Policy

**Voice Control Platform**
**Last Updated: March 13, 2026**

This Privacy Policy describes how Voice Control Platform ("the Software") collects, uses, and handles your information. By using the Software, you agree to the practices described below.

## 1. Information We Collect

### 1.1 Voice Audio Data

- **Google Speech Recognition**: If you select Google as your recognition engine, your voice audio is sent to Google's servers for transcription. We do not store your audio recordings. Google's handling of this data is governed by [Google's Privacy Policy](https://policies.google.com/privacy).
- **Vosk (Offline)**: Voice audio is processed entirely on your local device. No audio data leaves your computer.
- **Faster-Whisper (Offline)**: Voice audio is processed entirely on your local device. No audio data leaves your computer.

### 1.2 Application Settings

The Software stores your configuration locally on your device in a `config.json` file, including:

- Speech recognition preferences (engine choice, energy threshold, pause duration)
- Keyboard shortcut preferences
- Custom voice command profiles (phrases and associated actions)
- Audio output device selection

### 1.3 Spotify Data

If you choose to connect your Spotify account, the Software accesses the following data through the Spotify API:

- **Playback state**: Current track name, artist, album, album artwork, and playback status
- **Playback control**: Ability to play, pause, skip, and search for music
- **Playlists**: Ability to read your playlists and liked songs for voice-controlled playback
- **Authentication tokens**: OAuth access and refresh tokens are stored locally in a `.spotify_cache` file on your device

The Spotify API scopes used by the Software are:

| Scope | Purpose |
|-------|---------|
| `user-read-playback-state` | Read what is currently playing |
| `user-modify-playback-state` | Control playback (play, pause, skip) |
| `user-read-currently-playing` | Display current track information |
| `playlist-read-private` | Access your private playlists for voice playback |
| `playlist-read-collaborative` | Access collaborative playlists for voice playback |
| `user-library-read` | Access your liked songs for voice playback |

### 1.4 Spotify Credentials

Your Spotify Client ID and Client Secret are stored locally in your `config.json` file. These credentials are provided by you from your own Spotify Developer account and are never transmitted to us or any third party.

## 2. How We Use Your Information

All data collected is used solely to provide the functionality of the Software:

- Voice audio is used for real-time speech-to-text recognition to execute voice commands.
- Application settings are used to configure the Software to your preferences.
- Spotify data is used to display now-playing information and control playback via voice commands.

## 3. Data Storage

All data is stored **locally on your device**. The Software does not use remote servers, cloud storage, or external databases. Specifically:

- `config.json` — Settings and Spotify credentials
- `.spotify_cache` — Spotify OAuth tokens
- `profiles/` — Voice command profiles

No data is transmitted to us or any third party, with the exception of:

- **Google Speech Recognition**: Voice audio is sent to Google if you select Google as your engine.
- **Spotify API**: Playback commands and data requests are sent to Spotify's servers when you use Spotify features.

## 4. Data Sharing

We do not sell, trade, rent, or share your personal data with any third party, advertising network, data broker, or analytics service.

## 5. Cookies and Tracking

The Software is a desktop application and does not use cookies or web-based tracking technologies.

## 6. Children's Privacy

The Software is not directed at children under the age of 13. We do not knowingly collect personal information from children under 13. If you are under 13, do not use the Software.

## 7. Disconnecting Spotify

You can disconnect your Spotify account at any time by:

1. Removing your Client ID and Client Secret from the Software's settings
2. Deleting the `.spotify_cache` file from the application directory

Upon disconnection, the Software will no longer access your Spotify account or data. You can also revoke access from your [Spotify Account Settings](https://www.spotify.com/account/apps/).

## 8. Data Deletion

Since all data is stored locally on your device, you have full control over your data:

- **Uninstalling** the Software and deleting its directory removes all stored data.
- **Deleting** `config.json` removes all settings and Spotify credentials.
- **Deleting** `.spotify_cache` removes Spotify authentication tokens.
- **Deleting** files in `profiles/` removes your saved voice command profiles.

## 9. Security

Your data is stored locally on your device using standard file system permissions. Spotify authentication tokens and credentials are stored in plaintext in local files. We recommend keeping your device secure and not sharing these files.

## 10. Changes to This Privacy Policy

We may update this Privacy Policy from time to time. The "Last Updated" date at the top of this document will reflect the most recent revision. Continued use of the Software after changes constitutes acceptance of the updated policy.

## 11. Contact

If you have questions about this Privacy Policy or your data, please open an issue on the project's GitHub repository.
