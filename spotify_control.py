"""
Spotify Integration for Voice Control Platform
Provides voice-controlled Spotify playback using the Spotify Web API.
"""

import json
import os
import sys
import webbrowser
from pathlib import Path
from typing import Optional, Dict, Any, List
import threading
import time
import socket
import base64
from http.server import HTTPServer, BaseHTTPRequestHandler
from urllib.parse import urlparse, parse_qs

try:
    import requests
    from requests.exceptions import ConnectionError as RequestsConnectionError
    REQUESTS_AVAILABLE = True
except ImportError:
    REQUESTS_AVAILABLE = False
    RequestsConnectionError = OSError  # fallback so except clauses don't break

try:
    import spotipy
    from spotipy.oauth2 import SpotifyOAuth
    SPOTIPY_AVAILABLE = True
except ImportError:
    SPOTIPY_AVAILABLE = False

# Ensure SSL certificates are found inside PyInstaller one-file builds
if getattr(sys, 'frozen', False):
    try:
        import certifi
        os.environ.setdefault('SSL_CERT_FILE', certifi.where())
        os.environ.setdefault('REQUESTS_CA_BUNDLE', certifi.where())
    except ImportError:
        pass


class OAuthCallbackHandler(BaseHTTPRequestHandler):
    """HTTP handler to capture OAuth callback."""
    
    auth_code = None
    error = None
    
    def do_GET(self):
        """Handle the OAuth callback GET request."""
        parsed = urlparse(self.path)
        params = parse_qs(parsed.query)
        
        if 'code' in params:
            OAuthCallbackHandler.auth_code = params['code'][0]
            self.send_response(200)
            self.send_header('Content-type', 'text/html')
            self.end_headers()
            success_html = """
            <html>
            <head><title>Spotify Connected</title></head>
            <body style="font-family: Arial, sans-serif; display: flex; justify-content: center; align-items: center; height: 100vh; margin: 0; background: linear-gradient(135deg, #1DB954, #191414);">
                <div style="text-align: center; color: white; padding: 40px; background: rgba(0,0,0,0.5); border-radius: 20px;">
                    <h1>✅ Successfully Connected!</h1>
                    <p>You can close this tab and return to the Voice Control app.</p>
                </div>
            </body>
            </html>
            """
            self.wfile.write(success_html.encode())
        elif 'error' in params:
            OAuthCallbackHandler.error = params['error'][0]
            self.send_response(200)
            self.send_header('Content-type', 'text/html')
            self.end_headers()
            error_html = f"""
            <html>
            <head><title>Spotify Connection Failed</title></head>
            <body style="font-family: Arial, sans-serif; display: flex; justify-content: center; align-items: center; height: 100vh; margin: 0; background: linear-gradient(135deg, #e74c3c, #191414);">
                <div style="text-align: center; color: white; padding: 40px; background: rgba(0,0,0,0.5); border-radius: 20px;">
                    <h1>❌ Connection Failed</h1>
                    <p>Error: {OAuthCallbackHandler.error}</p>
                    <p>Please close this tab and try again.</p>
                </div>
            </body>
            </html>
            """
            self.wfile.write(error_html.encode())
        else:
            self.send_response(404)
            self.end_headers()
    
    def log_message(self, format, *args):
        """Suppress HTTP server logging."""
        pass


class SpotifyController:
    """Controls Spotify playback via the Spotify Web API."""
    
    SCOPES = [
        "user-read-playback-state",
        "user-modify-playback-state", 
        "user-read-currently-playing",
        "playlist-read-private",
        "playlist-read-collaborative",
        "user-library-read",
    ]
    
    def __init__(self, config_path: str = "config.json"):
        self.config_path = Path(config_path)
        self.spotify: Optional[spotipy.Spotify] = None
        self.is_authenticated = False
        self._credentials = self._load_credentials()
        
        # Try to authenticate if credentials exist
        if self._credentials.get("client_id") and self._credentials.get("client_secret"):
            self._try_authenticate()
    
    def _get_cache_path(self) -> Path:
        """Get the path for the Spotify token cache file.
        Uses AppData on Windows for proper permissions in installed apps.
        """
        import os
        import sys
        
        # Check if running as a packaged app (frozen)
        if getattr(sys, 'frozen', False):
            # Use AppData/Local for installed apps
            appdata = os.environ.get('LOCALAPPDATA', os.path.expanduser('~'))
            cache_dir = Path(appdata) / 'VoiceControl'
            cache_dir.mkdir(parents=True, exist_ok=True)
            return cache_dir / '.spotify_cache'
        else:
            # Development mode - use project directory
            return self.config_path.parent / ".spotify_cache"
    
    def _load_credentials(self) -> Dict[str, str]:
        """Load Spotify credentials from config."""
        if self.config_path.exists():
            try:
                with open(self.config_path, 'r') as f:
                    config = json.load(f)
                    return config.get("spotify", {})
            except (json.JSONDecodeError, IOError):
                pass
        return {}
    
    def _save_credentials(self, client_id: str, client_secret: str) -> None:
        """Save Spotify credentials to config."""
        config = {}
        if self.config_path.exists():
            try:
                with open(self.config_path, 'r') as f:
                    config = json.load(f)
            except (json.JSONDecodeError, IOError):
                pass
        
        config["spotify"] = {
            "client_id": client_id,
            "client_secret": client_secret
        }
        
        try:
            with open(self.config_path, 'w') as f:
                json.dump(config, f, indent=4)
        except IOError:
            pass
    
    def _try_authenticate(self) -> bool:
        """Try to authenticate with existing credentials."""
        if not SPOTIPY_AVAILABLE:
            return False
            
        try:
            cache_path = self._get_cache_path()
            
            auth_manager = SpotifyOAuth(
                client_id=self._credentials.get("client_id"),
                client_secret=self._credentials.get("client_secret"),
                redirect_uri="http://127.0.0.1:8888/callback",
                scope=" ".join(self.SCOPES),
                cache_path=str(cache_path),
                open_browser=False
            )
            
            # Check if we have a valid cached token
            token_info = auth_manager.get_cached_token()
            if token_info:
                self.spotify = spotipy.Spotify(auth_manager=auth_manager)
                # Test the connection
                self.spotify.current_user()
                self.is_authenticated = True
                return True
        except Exception:
            pass
        
        return False
    
    def authenticate(self, client_id: str, client_secret: str) -> tuple[bool, str]:
        """
        Authenticate with Spotify using OAuth.
        Uses a local HTTP server to capture the callback.
        Manually handles token exchange to avoid stdin issues in packaged apps.
        Returns (success, message).
        """
        if not SPOTIPY_AVAILABLE:
            return False, "Spotipy library not installed. Run: pip install spotipy"
        
        if not REQUESTS_AVAILABLE:
            return False, "Requests library not installed. Run: pip install requests"
        
        try:
            # Save credentials
            self._save_credentials(client_id, client_secret)
            self._credentials = {"client_id": client_id, "client_secret": client_secret}
            
            cache_path = self._get_cache_path()
            redirect_uri = "http://127.0.0.1:8888/callback"
            
            # Clear any previous auth state
            OAuthCallbackHandler.auth_code = None
            OAuthCallbackHandler.error = None
            
            # Check if we already have a valid cached token
            if cache_path.exists():
                try:
                    with open(cache_path, 'r') as f:
                        token_info = json.load(f)
                    
                    # Check if token is still valid or can be refreshed
                    if token_info.get('access_token'):
                        # Try to use existing token
                        auth_manager = SpotifyOAuth(
                            client_id=client_id,
                            client_secret=client_secret,
                            redirect_uri=redirect_uri,
                            scope=" ".join(self.SCOPES),
                            cache_path=str(cache_path),
                            open_browser=False
                        )
                        self.spotify = spotipy.Spotify(auth_manager=auth_manager)
                        try:
                            user = self.spotify.current_user()
                            self.is_authenticated = True
                            return True, f"Successfully connected as {user['display_name']}"
                        except Exception:
                            # Token invalid, continue with new auth
                            pass
                except Exception:
                    pass
            
            # Start local server to capture callback
            server = HTTPServer(('127.0.0.1', 8888), OAuthCallbackHandler)
            server.timeout = 120  # 2 minute timeout
            
            # Build auth URL manually
            auth_url = (
                "https://accounts.spotify.com/authorize"
                f"?client_id={client_id}"
                f"&response_type=code"
                f"&redirect_uri={redirect_uri}"
                f"&scope={'+'.join(self.SCOPES)}"
            )
            webbrowser.open(auth_url)
            
            # Wait for callback (handle one request)
            server.handle_request()
            server.server_close()
            
            if OAuthCallbackHandler.error:
                return False, f"Authorization denied: {OAuthCallbackHandler.error}"
            
            if not OAuthCallbackHandler.auth_code:
                return False, "No authorization code received. Please try again."
            
            # Manually exchange code for token (avoids stdin issues)
            auth_header = base64.b64encode(f"{client_id}:{client_secret}".encode()).decode()
            token_response = requests.post(
                "https://accounts.spotify.com/api/token",
                data={
                    "grant_type": "authorization_code",
                    "code": OAuthCallbackHandler.auth_code,
                    "redirect_uri": redirect_uri
                },
                headers={
                    "Authorization": f"Basic {auth_header}",
                    "Content-Type": "application/x-www-form-urlencoded"
                }
            )
            
            if token_response.status_code != 200:
                return False, f"Token exchange failed: {token_response.text}"
            
            token_info = token_response.json()
            
            # Add expires_at for Spotipy compatibility
            token_info['expires_at'] = int(time.time()) + token_info.get('expires_in', 3600)
            
            # Save token to cache file
            with open(cache_path, 'w') as f:
                json.dump(token_info, f)
            
            # Create Spotify client with the new token
            auth_manager = SpotifyOAuth(
                client_id=client_id,
                client_secret=client_secret,
                redirect_uri=redirect_uri,
                scope=" ".join(self.SCOPES),
                cache_path=str(cache_path),
                open_browser=False
            )
            self.spotify = spotipy.Spotify(auth_manager=auth_manager)
            
            # Test the connection
            user = self.spotify.current_user()
            self.is_authenticated = True
            
            return True, f"Successfully connected as {user['display_name']}"
            
        except spotipy.SpotifyException as e:
            return False, f"Spotify error: {str(e)}"
        except OSError as e:
            if "address already in use" in str(e).lower() or e.errno == 10048:
                return False, "Port 8888 is in use. Please close any other apps using it and try again."
            return False, f"Server error: {str(e)}"
        except Exception as e:
            return False, f"Authentication failed: {str(e)}"
    
    def disconnect(self) -> None:
        """Disconnect from Spotify and clear cached tokens."""
        self.spotify = None
        self.is_authenticated = False
        
        # Remove cached token
        cache_path = self._get_cache_path()
        if cache_path.exists():
            try:
                cache_path.unlink()
            except IOError:
                pass
    
    def get_credentials(self) -> tuple[str, str]:
        """Get stored credentials (client_id, client_secret)."""
        return (
            self._credentials.get("client_id", ""),
            self._credentials.get("client_secret", "")
        )
    
    def _spotify_api_call(self, func, *args, **kwargs):
        """Execute a Spotify API call with automatic retry and re-authentication.
        
        On connection/token errors: retries once after attempting to refresh
        the auth session.  Returns the result of `func` on success or
        re-raises the last exception on failure.
        """
        last_exc = None
        for attempt in range(2):
            try:
                return func(*args, **kwargs)
            except (RequestsConnectionError, OSError) as e:
                last_exc = e
                if attempt == 0:
                    # Network blip – try to re-establish the session
                    self._try_authenticate()
            except spotipy.SpotifyException as e:
                last_exc = e
                # 401 Unauthorized → token expired / revoked
                if e.http_status == 401 and attempt == 0:
                    self._try_authenticate()
                else:
                    raise
        # If we get here the retry also failed
        raise last_exc

    def _ensure_active_device(self) -> bool:
        """Ensure there's an active Spotify device."""
        if not self.is_authenticated or not self.spotify:
            return False
        
        try:
            devices = self._spotify_api_call(self.spotify.devices)
            if not devices.get("devices"):
                return False
            
            # Check if any device is active
            active = any(d.get("is_active") for d in devices["devices"])
            if not active:
                # Activate the first available device
                first_device = devices["devices"][0]
                self._spotify_api_call(self.spotify.transfer_playback, first_device["id"], force_play=False)
            
            return True
        except Exception:
            return False
    
    def play_song(self, query: str) -> tuple[bool, str]:
        """
        Search for and play a song.
        Supports "song name by artist" format for better accuracy.
        Automatically excludes remixes unless "remix" is in the query.
        Returns (success, message).
        """
        if not self.is_authenticated or not self.spotify:
            return False, "Not connected to Spotify"
        
        try:
            # Check if user explicitly wants a remix/edit version
            query_lower = query.lower()
            wants_remix = any(word in query_lower for word in ["remix", "edit", "mix", "version"])
            
            # Parse "by" to separate track and artist for better search accuracy
            search_query = query
            if " by " in query_lower:
                parts = query_lower.split(" by ", 1)
                track_name = parts[0].strip()
                artist_name = parts[1].strip()
                # Use Spotify's search filters for more accurate results
                search_query = f"track:{track_name} artist:{artist_name}"
            
            # Search for the track (get more results to filter)
            results = self._spotify_api_call(self.spotify.search, q=search_query, type="track", limit=10)
            tracks = results.get("tracks", {}).get("items", [])
            
            if not tracks:
                # If filtered search fails, try the original query
                if search_query != query:
                    results = self._spotify_api_call(self.spotify.search, q=query, type="track", limit=10)
                    tracks = results.get("tracks", {}).get("items", [])
                
                if not tracks:
                    return False, f"No song found for '{query}'"
            
            # Filter out remixes/edits if user didn't ask for one
            if not wants_remix and len(tracks) > 1:
                remix_keywords = ["remix", "edit", "mix", "bootleg", "rework", "vip", "extended", "radio edit"]
                original_tracks = [
                    t for t in tracks 
                    if not any(kw in t["name"].lower() for kw in remix_keywords)
                ]
                # Only use filtered list if we found originals
                if original_tracks:
                    tracks = original_tracks
            
            # If we parsed "by", try to find the best match by artist
            track = tracks[0]
            if " by " in query_lower and len(tracks) > 1:
                parts = query_lower.split(" by ", 1)
                artist_search = parts[1].strip().lower()
                # Look for a track where the artist name matches more closely
                for t in tracks:
                    for artist in t.get("artists", []):
                        if artist_search in artist["name"].lower():
                            track = t
                            break
            
            track_name = track["name"]
            artist_name = track["artists"][0]["name"] if track["artists"] else "Unknown"
            track_uri = track["uri"]
            
            # Ensure we have an active device
            if not self._ensure_active_device():
                return False, "No active Spotify device. Open Spotify on any device first."
            
            # Play the track
            self._spotify_api_call(self.spotify.start_playback, uris=[track_uri])
            
            return True, f"Playing '{track_name}' by {artist_name}"
            
        except spotipy.SpotifyException as e:
            if "No active device" in str(e):
                return False, "No active Spotify device. Open Spotify on any device first."
            return False, f"Spotify error: {str(e)}"
        except Exception as e:
            return False, f"Error: {str(e)}"
    
    def play_artist(self, query: str) -> tuple[bool, str]:
        """
        Search for and play an artist's top tracks.
        Returns (success, message).
        """
        if not self.is_authenticated or not self.spotify:
            return False, "Not connected to Spotify"
        
        try:
            # Search for the artist
            results = self._spotify_api_call(self.spotify.search, q=query, type="artist", limit=1)
            artists = results.get("artists", {}).get("items", [])
            
            if not artists:
                return False, f"No artist found for '{query}'"
            
            artist = artists[0]
            artist_name = artist["name"]
            artist_uri = artist["uri"]
            
            # Ensure we have an active device
            if not self._ensure_active_device():
                return False, "No active Spotify device. Open Spotify on any device first."
            
            # Play the artist
            self._spotify_api_call(self.spotify.start_playback, context_uri=artist_uri)
            
            return True, f"Playing {artist_name}"
            
        except spotipy.SpotifyException as e:
            if "No active device" in str(e):
                return False, "No active Spotify device. Open Spotify on any device first."
            return False, f"Spotify error: {str(e)}"
        except Exception as e:
            return False, f"Error: {str(e)}"
    
    def play_album(self, query: str) -> tuple[bool, str]:
        """
        Search for and play an album.
        Returns (success, message).
        """
        if not self.is_authenticated or not self.spotify:
            return False, "Not connected to Spotify"
        
        try:
            # Search for the album
            results = self._spotify_api_call(self.spotify.search, q=query, type="album", limit=1)
            albums = results.get("albums", {}).get("items", [])
            
            if not albums:
                return False, f"No album found for '{query}'"
            
            album = albums[0]
            album_name = album["name"]
            artist_name = album["artists"][0]["name"] if album["artists"] else "Unknown"
            album_uri = album["uri"]
            
            # Ensure we have an active device
            if not self._ensure_active_device():
                return False, "No active Spotify device. Open Spotify on any device first."
            
            # Play the album
            self._spotify_api_call(self.spotify.start_playback, context_uri=album_uri)
            
            return True, f"Playing '{album_name}' by {artist_name}"
            
        except spotipy.SpotifyException as e:
            if "No active device" in str(e):
                return False, "No active Spotify device. Open Spotify on any device first."
            return False, f"Spotify error: {str(e)}"
        except Exception as e:
            return False, f"Error: {str(e)}"
    
    def play_playlist(self, query: str) -> tuple[bool, str]:
        """
        Search for and play a playlist.
        Returns (success, message).
        """
        if not self.is_authenticated or not self.spotify:
            return False, "Not connected to Spotify"
        
        try:
            # Search for the playlist
            results = self._spotify_api_call(self.spotify.search, q=query, type="playlist", limit=1)
            playlists = results.get("playlists", {}).get("items", [])
            
            if not playlists:
                return False, f"No playlist found for '{query}'"
            
            playlist = playlists[0]
            playlist_name = playlist["name"]
            playlist_uri = playlist["uri"]
            
            # Ensure we have an active device
            if not self._ensure_active_device():
                return False, "No active Spotify device. Open Spotify on any device first."
            
            # Play the playlist
            self._spotify_api_call(self.spotify.start_playback, context_uri=playlist_uri)
            
            return True, f"Playing playlist '{playlist_name}'"
            
        except spotipy.SpotifyException as e:
            if "No active device" in str(e):
                return False, "No active Spotify device. Open Spotify on any device first."
            return False, f"Spotify error: {str(e)}"
        except Exception as e:
            return False, f"Error: {str(e)}"
    
    def pause(self) -> tuple[bool, str]:
        """Pause playback."""
        if not self.is_authenticated or not self.spotify:
            return False, "Not connected to Spotify"
        
        try:
            self._spotify_api_call(self.spotify.pause_playback)
            return True, "Paused"
        except Exception as e:
            return False, f"Error: {str(e)}"
    
    def resume(self) -> tuple[bool, str]:
        """Resume playback."""
        if not self.is_authenticated or not self.spotify:
            return False, "Not connected to Spotify"
        
        try:
            self._spotify_api_call(self.spotify.start_playback)
            return True, "Resumed"
        except Exception as e:
            return False, f"Error: {str(e)}"
    
    def next_track(self) -> tuple[bool, str]:
        """Skip to next track."""
        if not self.is_authenticated or not self.spotify:
            return False, "Not connected to Spotify"
        
        try:
            self._spotify_api_call(self.spotify.next_track)
            return True, "Skipped to next track"
        except Exception as e:
            return False, f"Error: {str(e)}"
    
    def previous_track(self) -> tuple[bool, str]:
        """Go to previous track."""
        if not self.is_authenticated or not self.spotify:
            return False, "Not connected to Spotify"
        
        try:
            self._spotify_api_call(self.spotify.previous_track)
            return True, "Previous track"
        except Exception as e:
            return False, f"Error: {str(e)}"
    
    def set_volume(self, volume: int) -> tuple[bool, str]:
        """Set Spotify volume (0-100)."""
        if not self.is_authenticated or not self.spotify:
            return False, "Not connected to Spotify"
        
        try:
            volume = max(0, min(100, volume))
            self._spotify_api_call(self.spotify.volume, volume)
            return True, f"Spotify volume set to {volume}%"
        except Exception as e:
            return False, f"Error: {str(e)}"
    
    def shuffle(self, state: bool) -> tuple[bool, str]:
        """Set shuffle state."""
        if not self.is_authenticated or not self.spotify:
            return False, "Not connected to Spotify"
        
        try:
            self._spotify_api_call(self.spotify.shuffle, state)
            return True, f"Shuffle {'on' if state else 'off'}"
        except Exception as e:
            return False, f"Error: {str(e)}"
    
    def repeat(self, state: str) -> tuple[bool, str]:
        """Set repeat mode: 'track', 'context' (playlist/album), or 'off'."""
        if not self.is_authenticated or not self.spotify:
            return False, "Not connected to Spotify"
        
        try:
            self._spotify_api_call(self.spotify.repeat, state)
            if state == "track":
                return True, "Repeat current track"
            elif state == "context":
                return True, "Repeat playlist/album"
            else:
                return True, "Repeat off"
        except Exception as e:
            return False, f"Error: {str(e)}"
    
    def play_my_playlist(self, playlist_name: str) -> tuple[bool, str]:
        """Play one of the user's own playlists."""
        if not self.is_authenticated or not self.spotify:
            return False, "Not connected to Spotify"
        
        try:
            # Get user's playlists
            playlists = self._spotify_api_call(self.spotify.current_user_playlists, limit=50)
            
            # Find matching playlist (case-insensitive partial match)
            playlist_name_lower = playlist_name.lower()
            best_match = None
            
            for playlist in playlists.get("items", []):
                if playlist_name_lower in playlist["name"].lower():
                    best_match = playlist
                    break
            
            if not best_match:
                # List available playlists in error message
                available = [p["name"] for p in playlists.get("items", [])[:5]]
                return False, f"Playlist '{playlist_name}' not found. Try: {', '.join(available)}"
            
            # Ensure we have an active device
            if not self._ensure_active_device():
                return False, "No active Spotify device. Open Spotify on any device first."
            
            # Play the playlist
            self._spotify_api_call(self.spotify.start_playback, context_uri=best_match["uri"])
            
            return True, f"Playing your playlist '{best_match['name']}'"
            
        except spotipy.SpotifyException as e:
            if "No active device" in str(e):
                return False, "No active Spotify device. Open Spotify on any device first."
            return False, f"Spotify error: {str(e)}"
        except Exception as e:
            return False, f"Error: {str(e)}"
    
    def add_to_playlist(self, playlist_name: str) -> tuple[bool, str]:
        """Add the currently playing track to a playlist."""
        if not self.is_authenticated or not self.spotify:
            return False, "Not connected to Spotify"
        
        try:
            # Get current track
            current = self._spotify_api_call(self.spotify.current_playback)
            if not current or not current.get("item"):
                return False, "No track currently playing"
            
            track_uri = current["item"]["uri"]
            track_name = current["item"]["name"]
            
            # Get user's playlists
            playlists = self._spotify_api_call(self.spotify.current_user_playlists, limit=50)
            
            # Find matching playlist
            playlist_name_lower = playlist_name.lower()
            best_match = None
            
            for playlist in playlists.get("items", []):
                if playlist_name_lower in playlist["name"].lower():
                    best_match = playlist
                    break
            
            if not best_match:
                available = [p["name"] for p in playlists.get("items", [])[:5]]
                return False, f"Playlist '{playlist_name}' not found. Try: {', '.join(available)}"
            
            # Add track to playlist
            self._spotify_api_call(self.spotify.playlist_add_items, best_match["id"], [track_uri])
            
            return True, f"Added '{track_name}' to '{best_match['name']}'"
            
        except spotipy.SpotifyException as e:
            return False, f"Spotify error: {str(e)}"
        except Exception as e:
            return False, f"Error: {str(e)}"
    
    def play_recommendations(self) -> tuple[bool, str]:
        """Play recommendations based on currently playing track using related artists."""
        if not self.is_authenticated or not self.spotify:
            return False, "Not connected to Spotify"
        
        try:
            # Get current track
            current = self._spotify_api_call(self.spotify.current_playback)
            if not current or not current.get("item"):
                return False, "No track currently playing to base recommendations on"
            
            artist_id = current["item"]["artists"][0]["id"] if current["item"]["artists"] else None
            artist_name = current["item"]["artists"][0]["name"] if current["item"]["artists"] else "Unknown"
            
            if not artist_id:
                return False, "Couldn't identify current artist"
            
            # Get related artists
            related = self._spotify_api_call(self.spotify.artist_related_artists, artist_id)
            related_artists = related.get("artists", [])[:5]  # Get top 5 related artists
            
            if not related_artists:
                return False, "Couldn't find related artists"
            
            # Collect top tracks from related artists
            track_uris = []
            for artist in related_artists:
                try:
                    top_tracks = self._spotify_api_call(self.spotify.artist_top_tracks, artist["id"])
                    for track in top_tracks.get("tracks", [])[:4]:  # 4 tracks per artist
                        track_uris.append(track["uri"])
                except:
                    continue
            
            if not track_uris:
                return False, "Couldn't get recommendations"
            
            # Shuffle the tracks for variety
            import random
            random.shuffle(track_uris)
            
            # Ensure we have an active device
            if not self._ensure_active_device():
                return False, "No active Spotify device. Open Spotify on any device first."
            
            # Play the recommendations
            self._spotify_api_call(self.spotify.start_playback, uris=track_uris[:20])  # Limit to 20 tracks
            
            return True, f"Playing music similar to {artist_name}"
            
        except spotipy.SpotifyException as e:
            if "No active device" in str(e):
                return False, "No active Spotify device. Open Spotify on any device first."
            return False, f"Spotify error: {str(e)}"
        except Exception as e:
            return False, f"Error: {str(e)}"
    
    def play_radio(self) -> tuple[bool, str]:
        """Play a radio station based on currently playing track/artist."""
        if not self.is_authenticated or not self.spotify:
            return False, "Not connected to Spotify"
        
        try:
            # Get current track
            current = self._spotify_api_call(self.spotify.current_playback)
            if not current or not current.get("item"):
                return False, "No track currently playing to create radio from"
            
            track_name = current["item"]["name"]
            artist_id = current["item"]["artists"][0]["id"] if current["item"]["artists"] else None
            artist_name = current["item"]["artists"][0]["name"] if current["item"]["artists"] else "Unknown"
            
            if not artist_id:
                return False, "Couldn't identify current artist"
            
            # Get related artists
            related = self._spotify_api_call(self.spotify.artist_related_artists, artist_id)
            related_artists = related.get("artists", [])[:8]  # More artists for radio
            
            # Collect top tracks from current artist and related artists
            track_uris = []
            
            # Add current artist's top tracks
            try:
                top_tracks = self._spotify_api_call(self.spotify.artist_top_tracks, artist_id)
                for track in top_tracks.get("tracks", []):
                    track_uris.append(track["uri"])
            except:
                pass
            
            # Add related artists' top tracks
            for artist in related_artists:
                try:
                    top_tracks = self._spotify_api_call(self.spotify.artist_top_tracks, artist["id"])
                    for track in top_tracks.get("tracks", [])[:5]:
                        track_uris.append(track["uri"])
                except:
                    continue
            
            if not track_uris:
                return False, "Couldn't create radio station"
            
            # Shuffle for radio feel
            import random
            random.shuffle(track_uris)
            
            # Ensure we have an active device
            if not self._ensure_active_device():
                return False, "No active Spotify device. Open Spotify on any device first."
            
            # Enable shuffle for radio feel
            try:
                self._spotify_api_call(self.spotify.shuffle, True)
            except:
                pass  # Shuffle might not be available
            
            # Play the radio
            self._spotify_api_call(self.spotify.start_playback, uris=track_uris[:50])
            
            return True, f"Playing radio based on '{track_name}' by {artist_name}"
            
        except spotipy.SpotifyException as e:
            if "No active device" in str(e):
                return False, "No active Spotify device. Open Spotify on any device first."
            return False, f"Spotify error: {str(e)}"
        except Exception as e:
            return False, f"Error: {str(e)}"
    
    def get_current_track(self) -> Optional[Dict[str, Any]]:
        """Get currently playing track info including progress and album art."""
        if not self.is_authenticated or not self.spotify:
            return None
        
        try:
            current = self._spotify_api_call(self.spotify.current_playback)
            if current and current.get("item"):
                track = current["item"]
                album = track.get("album", {})
                images = album.get("images", [])
                
                # Get the smallest image (usually 64x64) for the card, or medium if available
                album_art_url = None
                if images:
                    # Sort by size and get a reasonable one (not too big, not too small)
                    sorted_images = sorted(images, key=lambda x: x.get("width", 0))
                    # Prefer medium size (~300px) or fallback to largest
                    for img in sorted_images:
                        if img.get("width", 0) >= 100:
                            album_art_url = img.get("url")
                            break
                    if not album_art_url and sorted_images:
                        album_art_url = sorted_images[-1].get("url")
                
                # Calculate remaining time
                progress_ms = current.get("progress_ms", 0)
                duration_ms = track.get("duration_ms", 0)
                remaining_ms = max(0, duration_ms - progress_ms)
                remaining_sec = remaining_ms // 1000
                remaining_min = remaining_sec // 60
                remaining_sec = remaining_sec % 60
                
                return {
                    "name": track.get("name", "Unknown"),
                    "artist": track["artists"][0]["name"] if track.get("artists") else "Unknown",
                    "album": album.get("name", "Unknown"),
                    "is_playing": current.get("is_playing", False),
                    "album_art_url": album_art_url,
                    "progress_ms": progress_ms,
                    "duration_ms": duration_ms,
                    "remaining": f"{remaining_min}:{remaining_sec:02d}",
                    "progress_percent": (progress_ms / duration_ms * 100) if duration_ms > 0 else 0
                }
        except Exception:
            pass
        
        return None
    
    @staticmethod
    def get_reserved_phrases() -> tuple[list[str], list[str]]:
        """
        Return lists of reserved Spotify phrases and prefixes.
        Returns (exact_phrases, prefixes) where:
        - exact_phrases: phrases that must match exactly
        - prefixes: phrase prefixes that reserve all phrases starting with them
        """
        # Exact phrases that are reserved
        exact_phrases = [
            # Shuffle
            "shuffle on", "enable shuffle", "turn on shuffle",
            "shuffle off", "disable shuffle", "turn off shuffle",
            # Repeat
            "repeat on", "repeat track", "repeat this song", "repeat song",
            "repeat all", "repeat playlist", "repeat album",
            "repeat off", "no repeat", "stop repeat",
            # What's playing
            "what's playing", "whats playing", "what is playing", 
            "current song", "what song is this", "now playing",
            # Discovery
            "play recommendations", "play something similar", "recommend something",
            "play similar songs", "play similar", "recommendations",
            "play radio", "start radio", "radio", "create radio", "song radio",
        ]
        
        # Prefixes - any phrase starting with these is reserved
        prefixes = [
            "play song ", "play track ",
            "play artist ",
            "play album ",
            "play my playlist ", "play my ", "open my playlist ",
            "play playlist ",
            "add to playlist ", "add this to playlist ", "add to ", "add this to ",
            "save to ", "save this to ",
            "spotify volume ", "set spotify volume ", "spotify volume to ",
            "play ",  # Generic play - catches "play [song name]"
        ]
        
        return exact_phrases, prefixes
    
    @staticmethod
    def is_phrase_reserved(phrase: str) -> tuple[bool, str]:
        """
        Check if a phrase conflicts with Spotify commands.
        Returns (is_reserved, reason) tuple.
        """
        phrase = phrase.lower().strip()
        exact_phrases, prefixes = SpotifyController.get_reserved_phrases()
        
        # Check exact matches
        if phrase in exact_phrases:
            return True, f"'{phrase}' is a Spotify command"
        
        # Check if phrase matches a prefix exactly or starts with a prefix
        for prefix in prefixes:
            if phrase == prefix.strip():
                return True, f"'{phrase}' is a Spotify command prefix"
            if phrase.startswith(prefix):
                return True, f"'{phrase}' conflicts with Spotify command '{prefix}...'"
        
        # Check if any exact phrase starts with this phrase (would cause conflicts)
        for exact in exact_phrases:
            if exact.startswith(phrase) and exact != phrase:
                return True, f"'{phrase}' would conflict with Spotify command '{exact}'"
        
        return False, ""

    @staticmethod
    def get_voice_commands_help() -> str:
        """Return a formatted string of all available Spotify voice commands."""
        return """
🎵 SPOTIFY VOICE COMMANDS

▶️ PLAY MUSIC:
  • "Play [song name]" - Play a song
  • "Play song [name]" - Play a specific song
  • "Play artist [name]" - Play an artist
  • "Play album [name]" - Play an album
  • "Play playlist [name]" - Search & play a playlist
  • "Play my playlist [name]" - Play your own playlist

🔀 PLAYBACK CONTROLS:
  • "Shuffle on/off" - Toggle shuffle
  • "Repeat on" - Repeat current track
  • "Repeat all" - Repeat playlist/album
  • "Repeat off" - Turn off repeat

� VOLUME:
  • "Spotify volume [0-100]" - Set Spotify volume
  • "Set Spotify volume [number]" - Set Spotify volume

�📻 DISCOVERY:
  • "Play recommendations" - Play similar songs
  • "Play radio" - Start radio from current song
  • "Play something similar" - Same as recommendations

📝 PLAYLIST MANAGEMENT:
  • "Add to [playlist name]" - Add current song to playlist
  • "Add this to [playlist]" - Add current song to playlist

ℹ️ INFO:
  • "What's playing" - Show current track
  • "Current song" - Show current track

⏯️ STANDARD CONTROLS:
  • "Play" / "Pause" - Toggle playback
  • "Next" / "Skip" - Next track
  • "Previous" - Previous track
"""
