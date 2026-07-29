; ============================================================
; UPS Bid Analyzer - Inno Setup installer
;
; File location:
;     packaging\windows\installer_windows.iss
;
; Expected PyInstaller output:
;     packaging\windows\Bid_Analyzer\Bid_Analyzer.exe
; ============================================================

#define MyAppName "UPS Bid Analyzer"
#define MyAppVersion "2.0.0"
#define MyAppPublisher "Jerome Leluc"
#define MyAppExeName "Bid_Analyzer.exe"

; These paths are relative to packaging\windows.
#define MyAppBuildDir "Bid_Analyzer"
#define MyAppIconFile "..\..\src\bid_analyzer\resources\app_icon.ico"

[Setup]
; Keep this AppId unchanged between versions so installing a newer
; version upgrades the existing installation.
AppId={{76B20B8D-26DD-4EAB-A271-2DF664EAD470}

AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppVerName={#MyAppName} {#MyAppVersion}
AppPublisher={#MyAppPublisher}

; The application always receives its own dedicated installation folder.
DefaultDirName={localappdata}\Programs\{#MyAppName}

; Do not allow the user to select a different directory. This makes the
; full-directory cleanup in [UninstallDelete] safe and predictable.
DisableDirPage=yes

DefaultGroupName={#MyAppName}
DisableProgramGroupPage=yes

; Per-user installation. No administrator prompt is required.
PrivilegesRequired=lowest

; This assumes the application is built with 64-bit Python.
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible

; Output:
; packaging\windows\installer\Bid_Analyzer_Setup_<version>.exe
OutputDir=installer
OutputBaseFilename=Bid_Analyzer_Setup_{#MyAppVersion}

SetupIconFile={#MyAppIconFile}
UninstallDisplayIcon={app}\{#MyAppExeName}

Compression=lzma2
SolidCompression=yes
WizardStyle=modern

CloseApplications=yes
RestartApplications=no
UsePreviousAppDir=yes

VersionInfoCompany={#MyAppPublisher}
VersionInfoDescription={#MyAppName} Installer
VersionInfoProductName={#MyAppName}
VersionInfoProductVersion={#MyAppVersion}
VersionInfoVersion={#MyAppVersion}

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; \
    Description: "{cm:CreateDesktopIcon}"; \
    GroupDescription: "{cm:AdditionalIcons}"; \
    Flags: unchecked

[Files]
; Copy the complete PyInstaller --onedir output, including _internal
; and every nested dependency or resource folder.
Source: "{#MyAppBuildDir}\*"; \
    DestDir: "{app}"; \
    Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{autoprograms}\{#MyAppName}"; \
    Filename: "{app}\{#MyAppExeName}"; \
    WorkingDir: "{app}"

Name: "{autodesktop}\{#MyAppName}"; \
    Filename: "{app}\{#MyAppExeName}"; \
    WorkingDir: "{app}"; \
    Tasks: desktopicon

[Run]
Filename: "{app}\{#MyAppExeName}"; \
    Description: "Launch {#MyAppName}"; \
    WorkingDir: "{app}"; \
    Flags: nowait postinstall skipifsilent

[UninstallDelete]
; Inno Setup automatically removes files it installed. This additional
; entry removes anything the application created afterward, including
; configuration files, logs, caches, temporary files, and empty folders.
;
; User-created exports saved elsewhere are not affected.
Type: filesandordirs; Name: "{app}"
