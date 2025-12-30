; Inno Setup Script for TariffMill
; Build with: iscc installer.iss

#define MyAppName "TariffMill"
#define MyAppVersion "0.96.12"
#define MyAppPublisher "Process Logic Labs, LLC"
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
; Use faster compression for quicker installer startup
; lzma2/max gives best compression but slow startup
; lzma/fast gives good compression with faster startup
Compression=lzma/fast
SolidCompression=no
WizardStyle=modern

; Privileges - no admin needed
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

[Code]
var
  KeepDatabaseCheckbox: TCheckBox;
  KeepTemplatesCheckbox: TCheckBox;

procedure InitializeUninstallProgressForm();
var
  UninstallLabel: TLabel;
  OptionsPanel: TPanel;
begin
  // Create a panel for options
  OptionsPanel := TPanel.Create(UninstallProgressForm);
  OptionsPanel.Parent := UninstallProgressForm;
  OptionsPanel.Left := ScaleX(20);
  OptionsPanel.Top := ScaleY(10);
  OptionsPanel.Width := UninstallProgressForm.ClientWidth - ScaleX(40);
  OptionsPanel.Height := ScaleY(110);
  OptionsPanel.BevelOuter := bvNone;
  OptionsPanel.Caption := '';

  // Add explanatory label
  UninstallLabel := TLabel.Create(UninstallProgressForm);
  UninstallLabel.Parent := OptionsPanel;
  UninstallLabel.Caption := 'Select which user data to keep after uninstalling:';
  UninstallLabel.Left := ScaleX(0);
  UninstallLabel.Top := ScaleY(5);
  UninstallLabel.AutoSize := True;
  UninstallLabel.Font.Style := [fsBold];

  // Checkbox to keep database
  KeepDatabaseCheckbox := TCheckBox.Create(UninstallProgressForm);
  KeepDatabaseCheckbox.Parent := OptionsPanel;
  KeepDatabaseCheckbox.Caption := 'Keep database (tariffmill.db) - Contains your parts, MIDs, and settings';
  KeepDatabaseCheckbox.Left := ScaleX(0);
  KeepDatabaseCheckbox.Top := ScaleY(35);
  KeepDatabaseCheckbox.Width := OptionsPanel.Width;
  KeepDatabaseCheckbox.Checked := True;  // Default to keeping database

  // Checkbox to keep OCRmill templates
  KeepTemplatesCheckbox := TCheckBox.Create(UninstallProgressForm);
  KeepTemplatesCheckbox.Parent := OptionsPanel;
  KeepTemplatesCheckbox.Caption := 'Keep OCRmill templates - Your custom invoice parsing templates';
  KeepTemplatesCheckbox.Left := ScaleX(0);
  KeepTemplatesCheckbox.Top := ScaleY(60);
  KeepTemplatesCheckbox.Width := OptionsPanel.Width;
  KeepTemplatesCheckbox.Checked := True;  // Default to keeping templates

  // Move the progress bar down to make room
  UninstallProgressForm.ProgressBar.Top := UninstallProgressForm.ProgressBar.Top + ScaleY(120);
  UninstallProgressForm.StatusLabel.Top := UninstallProgressForm.StatusLabel.Top + ScaleY(120);

  // Increase form height to accommodate new controls
  UninstallProgressForm.ClientHeight := UninstallProgressForm.ClientHeight + ScaleY(120);
end;

procedure CurUninstallStepChanged(CurUninstallStep: TUninstallStep);
var
  AppDir: String;
  TemplatesDir: String;
  DatabasePath: String;
begin
  if CurUninstallStep = usPostUninstall then
  begin
    AppDir := ExpandConstant('{app}');
    DatabasePath := AppDir + '\Resources\tariffmill.db';
    TemplatesDir := AppDir + '\templates';

    // Handle database removal
    if not KeepDatabaseCheckbox.Checked then
    begin
      if FileExists(DatabasePath) then
        DeleteFile(DatabasePath);
    end;

    // Handle templates removal
    if not KeepTemplatesCheckbox.Checked then
    begin
      if DirExists(TemplatesDir) then
        DelTree(TemplatesDir, True, True, True);
    end;

    // Clean up empty directories if both were removed
    if (not KeepDatabaseCheckbox.Checked) and (not KeepTemplatesCheckbox.Checked) then
    begin
      // Try to remove Resources dir if empty
      RemoveDir(AppDir + '\Resources');
      // Try to remove app dir if empty
      RemoveDir(AppDir);
    end;
  end;
end;
