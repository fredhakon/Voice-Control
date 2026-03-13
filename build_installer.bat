@echo off
echo ========================================
echo Building Voice Control Installer
echo ========================================
echo.

REM Check if Inno Setup is installed
set "ISCC="

REM Check common installation paths
if exist "C:\Program Files (x86)\Inno Setup 6\ISCC.exe" (
    set "ISCC=C:\Program Files (x86)\Inno Setup 6\ISCC.exe"
) else if exist "C:\Program Files\Inno Setup 6\ISCC.exe" (
    set "ISCC=C:\Program Files\Inno Setup 6\ISCC.exe"
) else if exist "%LOCALAPPDATA%\Programs\Inno Setup 6\ISCC.exe" (
    set "ISCC=%LOCALAPPDATA%\Programs\Inno Setup 6\ISCC.exe"
)

if "%ISCC%"=="" (
    echo ERROR: Inno Setup 6 not found!
    echo.
    echo Please install Inno Setup 6 from:
    echo https://jrsoftware.org/isdl.php
    echo.
    echo After installation, run this script again.
    pause
    exit /b 1
)

echo Found Inno Setup at: %ISCC%
echo.

REM Check if the executable exists
if not exist "dist\VoiceControl.exe" (
    echo ERROR: VoiceControl.exe not found in dist folder!
    echo.
    echo Please run build.bat first to create the executable.
    pause
    exit /b 1
)

REM Build the installer
echo Building installer...
"%ISCC%" installer.iss

echo.
if exist "installer\VoiceControlSetup.exe" (
    echo ========================================
    echo Installer created successfully!
    echo Location: installer\VoiceControlSetup.exe
    echo ========================================
) else (
    echo ========================================
    echo Installer build failed!
    echo Check the output above for errors.
    echo ========================================
)

pause
