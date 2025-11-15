; Inno Setup Script for BBGRL Slides Windows Installer
; Builds an installer that places a Start Menu and Desktop shortcut
; for the packaged Flask app (PyInstaller-built single EXE).

#define MyAppName "BBGRL Slides"
#define MyAppPublisher "BBGRL"
#define MyAppURL "https://github.com/anthonybone/bbgrl-slides"
#define MyAppVersion "1.0.0"
#define MyAppExeName "BBGRL Slides App.exe"

[Setup]
AppId={{B7C7E0D1-9E8B-4DB8-9FD2-5B0F3C7D1C20}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppURL}
DefaultDirName={autopf}\{#MyAppName}
DefaultGroupName={#MyAppName}
DisableDirPage=yes
DisableReadyMemo=yes
DisableProgramGroupPage=yes
OutputBaseFilename=BBGRL-Slides-Setup
Compression=lzma
SolidCompression=yes
WizardStyle=modern

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "Create a &desktop shortcut"; GroupDescription: "Additional icons:"; Flags: unchecked

[Files]
; Include the PyInstaller-built EXE
Source: "..\..\dist\{#MyAppExeName}"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
; Start menu shortcut
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
; Desktop shortcut (optional task)
Name: "{commondesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
; Optionally launch app after install
Filename: "{app}\{#MyAppExeName}"; Description: "Launch {#MyAppName}"; Flags: nowait postinstall skipifsilent
