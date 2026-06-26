; ============================================================
;  Outlook Firma Rehberi v4.0
;  Inno Setup 6 Script
;  Poliport / S.SEYMEN
; ============================================================

#define AppName      "Outlook Firma Rehberi"
#define AppVersion   "4.0"
#define AppPublisher "Poliport - S.SEYMEN"
#define AppExeName   "OutlookFirmaRehberi.exe"
#define AppDir       "OutlookFirmaRehberi"
#define AppURL       "https://poliport.com"

[Setup]
AppId                    = {{F3A7C2B1-9D4E-4F81-B6A3-2E8C1D5F7094}
AppName                  = {#AppName}
AppVersion               = {#AppVersion}
AppVerName               = {#AppName} v{#AppVersion}
AppPublisher             = {#AppPublisher}
AppPublisherURL          = {#AppURL}
AppSupportURL            = {#AppURL}
DefaultDirName           = {autopf}\{#AppDir}
DefaultGroupName         = {#AppName}
AllowNoIcons             = yes
OutputDir                = dist
OutputBaseFilename       = OutlookFirmaRehberi_v{#AppVersion}_Setup
SetupIconFile            =
UninstallDisplayIcon     = {app}\{#AppExeName}
Compression              = lzma2/ultra64
SolidCompression         = yes
WizardStyle              = modern
WizardResizable          = no
PrivilegesRequired       = lowest
DisableProgramGroupPage  = yes
ShowLanguageDialog       = no
MinVersion               = 10.0
ArchitecturesAllowed     = x64compatible
ArchitecturesInstallIn64BitMode = x64compatible
VersionInfoVersion       = 4.0.0.0
VersionInfoCompany       = Poliport
VersionInfoDescription   = Outlook Firma Rehberi Kurulum
VersionInfoCopyright     = S.SEYMEN - Poliport
CloseApplications        = yes
RestartApplications      = no

[Languages]
Name: "turkish"; MessagesFile: "compiler:Languages\Turkish.isl"

[CustomMessages]
turkish.WelcomeLabel1    = Outlook Firma Rehberi kurulumuna ho�geldiniz
turkish.WelcomeLabel2    = Bu sihirbaz bilgisayar�n�za {#AppName} v{#AppVersion} y�kleyecektir.%n%nDevam etmeden �nce di�er uygulamalar� kapat�n�z.
turkish.FinishedLabel    = {#AppName} ba�ar�yla kuruldu!%n%nUygulama masaüstünden veya Ba�lat menüsünden a��labilir.

[Tasks]
Name: "desktopicon";   Description: "Masaüstüne k�sayol olu�tur";  GroupDescription: "K�sayollar:"; Flags: checked
Name: "startmenu";     Description: "Ba�lat menüsüne ekle";         GroupDescription: "K�sayollar:"; Flags: checked
Name: "quicklaunch";   Description: "Görev çubu�una sabitle";        GroupDescription: "K�sayollar:"; Flags: unchecked

[Files]
; Ana EXE — PyInstaller ile build edilmi�
Source: "dist\OutlookFirmaRehberi.exe"; DestDir: "{app}"; Flags: ignoreversion

[Dirs]
; Veri dizini (rehber_data.json burada saklanacak)
Name: "{app}"; Permissions: users-modify

[Icons]
Name: "{group}\{#AppName}";           Filename: "{app}\{#AppExeName}"; Comment: "Outlook mail adreslerini firma bazl� rehbere aktarır"
Name: "{group}\Kald�r";               Filename: "{uninstallexe}"
Name: "{autodesktop}\{#AppName}";     Filename: "{app}\{#AppExeName}"; Tasks: desktopicon; Comment: "Outlook Firma Rehberi"
Name: "{userstartmenu}\{#AppName}";   Filename: "{app}\{#AppExeName}"; Tasks: startmenu

[Run]
Filename: "{app}\{#AppExeName}"; Description: "Uygulamay� �imdi ba�lat"; Flags: nowait postinstall skipifsilent

[UninstallDelete]
Type: files;    Name: "{app}\rehber_data.json"
Type: dirifempty; Name: "{app}"

[Code]
// ─── Önceki sürüm açıksa kapat ───────────────────────────
function InitializeSetup(): Boolean;
begin
  Result := True;
end;

procedure CurStepChanged(CurStep: TSetupStep);
begin
  if CurStep = ssInstall then
  begin
    // Çalışıyorsa kapat
    Exec('taskkill.exe', '/F /IM OutlookFirmaRehberi.exe', '',
         SW_HIDE, ewWaitUntilTerminated, ResultCode);
  end;
end;

// ─── Kaldırma onayı ──────────────────────────────────────
function InitializeUninstall(): Boolean;
begin
  Result := MsgBox(
    '{#AppName} kaldırılacak.' + #13#10 +
    'rehber_data.json (kaydedilmiş veriler) de silinecek.' + #13#10 + #13#10 +
    'Devam etmek istiyor musunuz?',
    mbConfirmation, MB_YESNO
  ) = IDYES;
end;
