; Outlook Firma Rehberi v5 - Inno Setup 6

#define AppName "Outlook Firma Rehberi"
#ifndef AppVersion
  #define AppVersion "5.0.0"
#endif
#define AppPublisher "S.SEYMEN"
#define AppExeName "OutlookFirmaRehberi.exe"
#define AppDir "OutlookFirmaRehberi"

[Setup]
AppId={{F3A7C2B1-9D4E-4F81-B6A3-2E8C1D5F7094}
AppName={#AppName}
AppVersion={#AppVersion}
AppVerName={#AppName} v{#AppVersion}
AppPublisher={#AppPublisher}
DefaultDirName          = {localappdata}\Programs\{#AppDir}
DefaultGroupName={#AppName}
AllowNoIcons=yes
OutputDir=dist
OutputBaseFilename=OutlookFirmaRehberi_v{#AppVersion}_Windows_x64_Setup
SetupIconFile=assets\icon.ico
UninstallDisplayIcon={app}\{#AppExeName}
Compression=lzma2/ultra64
SolidCompression=yes
WizardStyle=modern
PrivilegesRequired=lowest
DisableProgramGroupPage=yes
ShowLanguageDialog=no
CloseApplications=yes
ArchitecturesAllowed=x64compatible

[Languages]
Name: "tr"; MessagesFile: "compiler:Languages\Turkish.isl"

[Tasks]
Name: "desktopicon"; Description: "Masaüstü kısayolu oluştur"; GroupDescription: "Kısayollar:"
Name: "startmenu"; Description: "Başlat menüsüne ekle"; GroupDescription: "Kısayollar:"

[Files]
Source: "dist\OutlookFirmaRehberi.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "assets\icon.ico"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\{#AppName}"; Filename: "{app}\{#AppExeName}"; IconFilename: "{app}\icon.ico"
Name: "{group}\Kaldır"; Filename: "{uninstallexe}"
Name: "{autodesktop}\{#AppName}"; Filename: "{app}\{#AppExeName}"; IconFilename: "{app}\icon.ico"; Tasks: desktopicon
Name: "{userstartmenu}\{#AppName}"; Filename: "{app}\{#AppExeName}"; IconFilename: "{app}\icon.ico"; Tasks: startmenu

[Run]
Filename: "{app}\{#AppExeName}"; Description: "Uygulamayı başlat"; Flags: nowait postinstall skipifsilent
