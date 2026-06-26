; Outlook Firma Rehberi v4.0
; Inno Setup 6 Script - Poliport / S.SEYMEN

#define AppName      "Outlook Firma Rehberi"
#define AppVersion   "4.0"
#define AppPublisher "Poliport - S.SEYMEN"
#define AppExeName   "OutlookFirmaRehberi.exe"
#define AppDir       "OutlookFirmaRehberi"

[Setup]
AppId                    = {{F3A7C2B1-9D4E-4F81-B6A3-2E8C1D5F7094}
AppName                  = {#AppName}
AppVersion               = {#AppVersion}
AppVerName               = {#AppName} v{#AppVersion}
AppPublisher             = {#AppPublisher}
DefaultDirName           = {autopf}\{#AppDir}
DefaultGroupName         = {#AppName}
AllowNoIcons             = yes
OutputDir                = dist
OutputBaseFilename       = OutlookFirmaRehberi_v4.0_Setup
Compression              = lzma2/ultra64
SolidCompression         = yes
WizardStyle              = modern
PrivilegesRequired       = lowest
DisableProgramGroupPage  = yes
ShowLanguageDialog       = no
CloseApplications        = yes

[Languages]
Name: "turkish"; MessagesFile: "compiler:Languages\Turkish.isl"

[Tasks]
Name: "desktopicon"; Description: "Masaustu kisayolu"; GroupDescription: "Kisayollar:"; Flags: checked
Name: "startmenu";   Description: "Baslat menusu";     GroupDescription: "Kisayollar:"; Flags: checked

[Files]
Source: "dist\OutlookFirmaRehberi.exe"; DestDir: "{app}"; Flags: ignoreversion

[Dirs]
Name: "{app}"; Permissions: users-modify

[Icons]
Name: "{group}\{#AppName}";           Filename: "{app}\{#AppExeName}"
Name: "{group}\Kaldır";               Filename: "{uninstallexe}"
Name: "{autodesktop}\{#AppName}";     Filename: "{app}\{#AppExeName}"; Tasks: desktopicon
Name: "{userstartmenu}\{#AppName}";   Filename: "{app}\{#AppExeName}"; Tasks: startmenu

[Run]
Filename: "{app}\{#AppExeName}"; Description: "Uygulamayi baslat"; Flags: nowait postinstall skipifsilent

[UninstallDelete]
Type: files;      Name: "{app}\rehber_data.json"
Type: dirifempty; Name: "{app}"

[Code]
procedure CurStepChanged(CurStep: TSetupStep);
var
  ResultCode: Integer;
begin
  if CurStep = ssInstall then
    Exec('taskkill.exe', '/F /IM OutlookFirmaRehberi.exe', '',
         SW_HIDE, ewWaitUntilTerminated, ResultCode);
end;
