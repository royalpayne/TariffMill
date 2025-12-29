; Inno Setup Script for TariffMill
; Build with: iscc installer.iss

#define MyAppName "TariffMill"
#define MyAppVersion "0.96.8"
#define MyAppPublisher "TariffMill"
#define MyAppExeName "TariffMill.exe"

[Setup]
; Application info
AppId={{8F3B9A2E-5C7D-4E1F-B8A6-9D2C3E4F5A6B}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={autopf}\{#MyAppName}
DefaultGroupName={#MyAppName}
DisableProgramGroupPage=yes
OutputDir=dist
OutputBaseFilename=TariffMill_Setup_{#MyAppVersion}
SetupIconFile=Resources\icon.ico
Compression=lzma2
SolidCompression=yes
WizardStyle=modern

; Privileges - no admin needed (Python 3.12 doesn't require VC++ install)
PrivilegesRequired=lowest
PrivilegesRequiredOverridesAllowed=dialog

; Uninstall info
UninstallDisplayIcon={app}\Resources\icon.ico
UninstallDisplayName={#MyAppName}

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked

[Files]
; Main executable
Source: "dist\{#MyAppExeName}"; DestDir: "{app}"; Flags: ignoreversion
; Resources directory (includes databases, icons, etc.)
Source: "Resources\*"; DestDir: "{app}\Resources"; Flags: ignoreversion recursesubdirs createallsubdirs; Excludes: "vc_redist.x64.exe"
; Templates directory (invoice templates)
Source: "templates\*"; DestDir: "{app}\templates"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
; Always recreate shortcuts to update icons
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; IconFilename: "{app}\Resources\icon.ico"
Name: "{group}\Uninstall {#MyAppName}"; Filename: "{uninstallexe}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; IconFilename: "{app}\Resources\icon.ico"; Tasks: desktopicon

[Run]
; Launch app after install
Filename: "{app}\{#MyAppExeName}"; Description: "{cm:LaunchProgram,{#StringChange(MyAppName, '&', '&&')}}"; Flags: nowait postinstall skipifsilent

[InstallDelete]
; Delete old shortcuts to force recreation with new icons
Type: files; Name: "{group}\{#MyAppName}.lnk"
Type: files; Name: "{autodesktop}\{#MyAppName}.lnk"
; Delete icon cache files to force refresh
Type: files; Name: "{localappdata}\Microsoft\Windows\Explorer\iconcache*.db"
Type: files; Name: "{localappdata}\IconCache.db"

[Code]
procedure RefreshIconCache;
var
  ResultCode: Integer;
begin
  // Notify Windows shell that icons have changed
  Exec('cmd.exe', '/c taskkill /f /im explorer.exe & start explorer.exe', '', SW_HIDE, ewWaitUntilTerminated, ResultCode);
end;

procedure CurStepChanged(CurStep: TSetupStep);
begin
  if CurStep = ssPostInstall then
  begin
    // Refresh icon cache after installation
    RefreshIconCache;
  end;
end;
