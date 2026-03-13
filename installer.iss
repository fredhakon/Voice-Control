; Voice Control Platform Installer Script
; Created with Inno Setup

#define MyAppName "Voice Control Platform"
#define MyAppVersion "1.0.0"
#define MyAppPublisher "Voice Control"
#define MyAppURL "https://github.com/yourusername/voice-control"
#define MyAppExeName "VoiceControl.exe"

[Setup]
; Basic installer settings
AppId={{8F3B9A72-5E4D-4C1A-B8E7-2A9F6D3C5E1B}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppVerName={#MyAppName} {#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppURL}
AppSupportURL={#MyAppURL}
AppUpdatesURL={#MyAppURL}

; Installation directory
DefaultDirName={autopf}\{#MyAppName}
DefaultGroupName={#MyAppName}
DisableProgramGroupPage=yes

; Output settings
OutputDir=installer
OutputBaseFilename=VoiceControlSetup
SetupIconFile=icon.ico
UninstallDisplayIcon={app}\{#MyAppExeName}

; Compression
Compression=lzma2/ultra64
SolidCompression=yes
LZMAUseSeparateProcess=yes

; Modern look
WizardStyle=modern
WizardSizePercent=100

; Privileges (request admin for global shortcuts)
PrivilegesRequired=admin
PrivilegesRequiredOverridesAllowed=dialog

; Minimum Windows version (Windows 10)
MinVersion=10.0

; License and info pages
LicenseFile=LICENSE.txt
InfoBeforeFile=INSTALL_README.txt

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked
Name: "startupicon"; Description: "Start with Windows"; GroupDescription: "Startup:"; Flags: unchecked

[Files]
; Main executable
Source: "dist\{#MyAppExeName}"; DestDir: "{app}"; Flags: ignoreversion

; Icon file (for shortcuts)
Source: "icon.ico"; DestDir: "{app}"; Flags: ignoreversion

; Config file - use clean default config, don't overwrite user's existing config
Source: "config.default.json"; DestDir: "{app}"; DestName: "config.json"; Flags: ignoreversion onlyifdoesntexist

; README
Source: "README.md"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
; Start Menu
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; IconFilename: "{app}\icon.ico"
Name: "{group}\{cm:UninstallProgram,{#MyAppName}}"; Filename: "{uninstallexe}"

; Desktop (optional)
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; IconFilename: "{app}\icon.ico"; Tasks: desktopicon

; Startup (optional)
Name: "{userstartup}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; IconFilename: "{app}\icon.ico"; Tasks: startupicon

[Run]
; Option to run after install
Filename: "{app}\{#MyAppExeName}"; Description: "{cm:LaunchProgram,{#StringChange(MyAppName, '&', '&&')}}"; Flags: nowait postinstall skipifsilent

[UninstallDelete]
; Clean up generated files on uninstall (install directory)
Type: files; Name: "{app}\config.json"
Type: dirifempty; Name: "{app}"

; Clean up AppData user data (tokens, config, cache)
Type: files; Name: "{localappdata}\VoiceControl\config.json"
Type: files; Name: "{localappdata}\VoiceControl\.spotify_cache"
Type: filesandordirs; Name: "{localappdata}\VoiceControl\profiles"
Type: dirifempty; Name: "{localappdata}\VoiceControl"

[Code]
// Custom code for additional functionality

function InitializeSetup(): Boolean;
begin
  Result := True;
end;

procedure CurStepChanged(CurStep: TSetupStep);
begin
  if CurStep = ssPostInstall then
  begin
    // Post-installation tasks can be added here
  end;
end;

procedure CurUninstallStepChanged(CurUninstallStep: TUninstallStep);
var
  AppDataDir: String;
begin
  if CurUninstallStep = usUninstall then
  begin
    AppDataDir := ExpandConstant('{localappdata}\VoiceControl');
    if DirExists(AppDataDir) then
    begin
      if MsgBox('Do you want to remove all user data (settings, Spotify tokens, profiles)?'
        + #13#10 + #13#10 + 'Location: ' + AppDataDir,
        mbConfirmation, MB_YESNO) = IDYES then
      begin
        DelTree(AppDataDir, True, True, True);
      end;
    end;
  end;
end;
