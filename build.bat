@echo off
echo ========================================
echo Building Voice Control Platform
echo ========================================
echo.

REM Activate virtual environment if it exists
if exist ".venv\Scripts\activate.bat" (
    call .venv\Scripts\activate.bat
)

REM Clean previous builds
echo Cleaning previous builds...
if exist "build" rmdir /s /q build
if exist "dist" rmdir /s /q dist

REM Move sensitive files out of reach so PyInstaller cannot bundle them
echo Removing sensitive files from build tree...
if exist ".spotify_cache" (
    move ".spotify_cache" ".spotify_cache.bak" >nul 2>&1
)
if exist "config.json" (
    move "config.json" "config.json.bak" >nul 2>&1
)

REM Build the executable
echo.
echo Building executable...
pyinstaller VoiceControl.spec --clean

REM Restore sensitive files
echo Restoring user files...
if exist ".spotify_cache.bak" (
    move ".spotify_cache.bak" ".spotify_cache" >nul 2>&1
)
if exist "config.json.bak" (
    move "config.json.bak" "config.json" >nul 2>&1
)

echo.
if exist "dist\VoiceControl.exe" (
    echo ========================================
    echo Build successful!
    echo Executable: dist\VoiceControl.exe
    echo ========================================
) else (
    echo ========================================
    echo Build failed! Check the output above.
    echo ========================================
)

pause
