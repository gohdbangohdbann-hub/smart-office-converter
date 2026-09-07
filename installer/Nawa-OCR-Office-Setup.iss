; Compile with Inno Setup on Windows: ISCC.exe Nawa-OCR-Office-Setup.iss
#define MyAppName "Nawa OCR Office"
#define MyAppVersion "1.0.0"
#define MyAppPublisher "Nawa"
#define MyAppExeName "Nawa-OCR-Office-Setup.ps1"

[Setup]
AppId={{3B29A1D4-F0C2-4E6E-9E2E-7F2C8C8F7A19}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={localappdata}\NawaOCROffice
DisableProgramGroupPage=yes
OutputDir=output
OutputBaseFilename=Nawa-OCR-Office-Setup
Compression=lzma
SolidCompression=yes
PrivilegesRequired=lowest
WizardStyle=modern

[Files]
Source: "..\manifest.xml"; DestDir: "{app}"; Flags: ignoreversion
Source: "Nawa-OCR-Office-Setup.ps1"; DestDir: "{app}"; Flags: ignoreversion
Source: "INSTALL-WINDOWS-AR.md"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{autodesktop}\Nawa OCR Office"; Filename: "powershell.exe"; Parameters: "-NoProfile -ExecutionPolicy Bypass -File \"{app}\Nawa-OCR-Office-Setup.ps1\""; WorkingDir: "{app}"
Name: "{group}\Nawa OCR Office Setup"; Filename: "powershell.exe"; Parameters: "-NoProfile -ExecutionPolicy Bypass -File \"{app}\Nawa-OCR-Office-Setup.ps1\""; WorkingDir: "{app}"

[Run]
Filename: "powershell.exe"; Parameters: "-NoProfile -ExecutionPolicy Bypass -File \"{app}\Nawa-OCR-Office-Setup.ps1\""; WorkingDir: "{app}"; Flags: postinstall nowait skipifsilent
