#define AppName "Haronex Radiology Agent for Word 2013"
#define AppVersion "11.12.12"
[Setup]
AppId={{F29657A1-86DE-4565-B391-E9E4D35A8219}
AppName={#AppName}
AppVersion={#AppVersion}
DefaultDirName={localappdata}\RadiologyAgent\Word2013Safe
DefaultGroupName=Haronex
OutputDir=output
OutputBaseFilename=Haronex_Word2013_V11_12_12_MultiModel_Setup
Compression=lzma2
SolidCompression=yes
PrivilegesRequired=lowest
WizardStyle=modern
CloseApplications=yes
DisableProgramGroupPage=yes
UninstallDisplayName=Haronex (Word 2013 x86)

[Files]
Source: "bin\HaronexFiles.sha256"; DestDir: "{app}"; Flags: ignoreversion
Source: "bin\NDP48-x86-x64-AllOS-ENU.exe"; Flags: dontcopy
Source: "bin\RadiologyAgent.NativeShim.dll"; DestDir: "{app}"; Flags: ignoreversion regserver
Source: "bin\RadiologyAgent.ReviewHost.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "bin\HaronexGroqWorker\HaronexGroqWorker.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "bin\HaronexGroqWorker\_HaronexRuntime\*"; DestDir: "{app}\_HaronexRuntime"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "bin\RadiologyAgent.KeySetup.exe"; Flags: dontcopy
Source: "bin\RadiologyAgent.KeySetup.exe"; DestDir: "{app}"; Flags: ignoreversion

[Registry]
Root: HKCU; Subkey: "Software\Microsoft\Office\Word\Addins\RadiologyAgent.WordAddin"; Flags: deletekey
Root: HKCU; Subkey: "Software\Microsoft\Office\15.0\Word\Addins\RadiologyAgent.WordAddin"; Flags: deletekey
Root: HKCU; Subkey: "Software\Classes\RadiologyAgent.WordAddin"; Flags: deletekey
Root: HKCU; Subkey: "Software\Classes\CLSID\{{D1FBEFD5-3B43-4516-9B7B-CDEB47AB0FA1}"; Flags: deletekey

[Code]
var
  KeyPage: TInputQueryWizardPage;
  ConsentPage: TInputOptionWizardPage;
  MandatoryPage: TWizardPage;
  MandatoryCheck: TNewCheckBox;
  RecoveryCheck: TNewCheckBox;

// Embedded official Microsoft .NET Framework 4.8 offline redistributable.
// No Visual Studio or VC++ redistributable required: native DLL uses /MT.
// The installer remains per-user; only a missing .NET prerequisite is elevated.
function HasDotNet48: Boolean;
var
  ReleaseNumber: Cardinal;
begin
  Result := RegQueryDWordValue(HKLM, 'SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full',
    'Release', ReleaseNumber) and (ReleaseNumber >= 528040);
end;

function InitializeSetup: Boolean;
begin
  Result := True;
end;

function PrepareToInstall(var NeedsRestart: Boolean): String;
var
  ExitCode: Integer;
  Installer: String;
  Launched: Boolean;
begin
  Result := '';
  NeedsRestart := False;
  if HasDotNet48 then begin
    ExtractTemporaryFile('RadiologyAgent.KeySetup.exe');
    if not Exec(ExpandConstant('{tmp}\RadiologyAgent.KeySetup.exe'), '--prepare-user-install', '', SW_HIDE, ewWaitUntilTerminated, ExitCode) or (ExitCode <> 0) then
      Result := 'Close Word before installing or repairing Haronex.';
    Exit;
  end;
  Log('Haronex: .NET Framework 4.8 missing; extracting bundled Microsoft offline installer.');
  try
    ExtractTemporaryFile('NDP48-x86-x64-AllOS-ENU.exe');
  except
    Result := 'Unable to extract the bundled .NET Framework 4.8 prerequisite.';
    Exit;
  end;
  Installer := ExpandConstant('{tmp}\NDP48-x86-x64-AllOS-ENU.exe');
  if WizardSilent then
    Launched := ShellExec('runas', Installer, '/q /norestart', '',
      SW_HIDE, ewWaitUntilTerminated, ExitCode)
  else
    Launched := ShellExec('runas', Installer, '/passive /norestart', '',
      SW_SHOW, ewWaitUntilTerminated, ExitCode);
  if not Launched then begin
    Result := '.NET Framework 4.8 setup was cancelled or could not be started. Administrator approval is required only when .NET is missing.';
    Exit;
  end;
  Log('Haronex: bundled .NET setup exit code ' + IntToStr(ExitCode));
  if (ExitCode = 3010) or (ExitCode = 1641) then begin
    NeedsRestart := True;
    Result := '.NET Framework was installed, but Windows requires a restart. Restart Windows and rerun Haronex setup.';
    Exit;
  end;
  if (ExitCode <> 0) and (not HasDotNet48) then begin
    Result := '.NET Framework 4.8 installation failed (exit code ' + IntToStr(ExitCode) + ').';
    Exit;
  end;
  if not HasDotNet48 then begin
    Result := '.NET Framework 4.8 is not detected after prerequisite installation. Restart Windows and rerun setup.';
    Exit;
  end;
  ExtractTemporaryFile('RadiologyAgent.KeySetup.exe');
  if not Exec(ExpandConstant('{tmp}\RadiologyAgent.KeySetup.exe'), '--prepare-user-install', '', SW_HIDE, ewWaitUntilTerminated, ExitCode) or (ExitCode <> 0) then
    Result := 'Close Word before installing or repairing Haronex.';
end;

function HasExistingHaronexAuthorization: Boolean;
var
  Allowed: Cardinal;
begin
  Result := FileExists(ExpandConstant('{localappdata}\RadiologyAgent\groq.key.dpapi'))
    and RegQueryDWordValue(HKCU, 'Software\Haronex', 'WebTransferAllowed', Allowed)
    and (Allowed = 1);
end;

function ShouldSkipPage(PageID: Integer): Boolean;
begin
  // An in-place upgrade does not request or overwrite the user's existing
  // DPAPI-protected key and previously granted patient-data authorization.
  Result := HasExistingHaronexAuthorization and
    ((PageID = KeyPage.ID) or (PageID = ConsentPage.ID));
end;

function InitialUserOption(Name, ParamName: String): Boolean;
var Value: Cardinal; ParamValue: String;
begin
  Result := RegQueryDWordValue(HKCU, 'Software\Haronex', Name, Value) and (Value = 1);
  ParamValue := ExpandConstant('{param:' + ParamName + '|}');
  if ParamValue = '1' then Result := True;
  if ParamValue = '0' then Result := False;
end;

procedure InitializeWizard;
begin
  KeyPage := CreateInputQueryPage(wpSelectDir,
    'Groq API Configuration', 'Enter your Groq API Key',
    'Enter an existing API key from console.groq.com/keys. The key is not embedded in the installer and will be protected locally for this Windows account with DPAPI.');
  KeyPage.Add('Groq API Key:', True);

  ConsentPage := CreateInputOptionPage(KeyPage.ID,
    'External Report Processing Consent', 'Authorization required',
    'Haronex sends the selected Word report text directly from this PC to the Groq cloud API when the user clicks Review Text. Patient identifiers may be in the text. Use only if authorized by your medical organization. No report is sent in the background. Clinical verification is mandatory.',
    False, False);
  ConsentPage.Add('I am authorized to send these medical report texts to Groq for proofreading.');
  ConsentPage.Values[0] := False;

  MandatoryPage := CreateCustomPage(ConsentPage.ID,
    'Mandatory Review',
    'Choose whether saving and printing require a completed review.');
  MandatoryCheck := TNewCheckBox.Create(MandatoryPage);
  MandatoryCheck.Parent := MandatoryPage.Surface;
  MandatoryCheck.Left := ScaleX(0);
  MandatoryCheck.Top := ScaleY(12);
  MandatoryCheck.Width := MandatoryPage.SurfaceWidth;
  MandatoryCheck.Height := ScaleY(44);
  MandatoryCheck.Caption :=
    'Enable Mandatory Review';
  MandatoryCheck.Checked := InitialUserOption('MandatoryReviewEnabled', 'MANDATORYREVIEW');
  RecoveryCheck := TNewCheckBox.Create(MandatoryPage);
  RecoveryCheck.Parent := MandatoryPage.Surface;
  RecoveryCheck.Left := ScaleX(0);
  RecoveryCheck.Top := ScaleY(65);
  RecoveryCheck.Width := MandatoryPage.SurfaceWidth;
  RecoveryCheck.Height := ScaleY(30);
  RecoveryCheck.Caption := 'Disable AutoRecover (Word 2013; unsaved work may be lost)';
  RecoveryCheck.Checked := InitialUserOption('DisableAutoRecover', 'DISABLEAUTORECOVER');
end;

function NextButtonClick(CurPageID: Integer): Boolean;
var
  KeyText: String;
begin
  Result := True;
  if WizardSilent then Exit;
  if CurPageID = KeyPage.ID then begin
    KeyText := Trim(KeyPage.Values[0]);
    if (Length(KeyText) < 20) or (Copy(KeyText, 1, 4) <> 'gsk_') then begin
      MsgBox('Please enter a valid Groq API key beginning with gsk_.', mbError, MB_OK);
      Result := False;
    end;
  end;
  if (CurPageID = ConsentPage.ID) and (not ConsentPage.Values[0]) then begin
    MsgBox('External processing authorization is required to enable review.', mbError, MB_OK);
    Result := False;
  end;
end;

procedure CurStepChanged(CurStep: TSetupStep);
var
  StagingPath: String;
  ResultCode: Integer;
begin
  if CurStep <> ssPostInstall then Exit;
  if not Exec(ExpandConstant('{app}\RadiologyAgent.KeySetup.exe'), '--verify-install', '', SW_HIDE, ewWaitUntilTerminated, ResultCode) or (ResultCode <> 0) then
    RaiseException('Haronex installation integrity check failed. Re-run Setup and contact IT if files are blocked.');
  if not Exec(ExpandConstant('{app}\RadiologyAgent.ReviewHost.exe'), '--worker-health', '', SW_HIDE, ewWaitUntilTerminated, ResultCode) or (ResultCode <> 0) then
    RaiseException('Haronex local worker handshake failed. Check diagnostics.log.');
  // Default OFF, including silent installs. When enabled, the native Word
  // event sink blocks Save, Save As and Print until Review Text + Clear Both.
  if not RegWriteDWordValue(HKCU, 'Software\Haronex',
       'MandatoryReviewEnabled', Ord(MandatoryCheck.Checked)) then
    RaiseException('Could not save Mandatory Review installation setting.');
  if not RegWriteDWordValue(HKCU, 'Software\Haronex',
       'DisableAutoRecover', Ord(RecoveryCheck.Checked)) then
    RaiseException('Could not save AutoRecover preference.');
  if WizardSilent or HasExistingHaronexAuthorization then Exit;
  StagingPath := ExpandConstant('{localappdata}\RadiologyAgent\pending-groq.key');
  ForceDirectories(ExtractFileDir(StagingPath));
  try
    if not SaveStringToFile(StagingPath, Trim(KeyPage.Values[0]), False) then
      RaiseException('Could not stage Groq configuration');
    if not Exec(ExpandConstant('{app}\RadiologyAgent.KeySetup.exe'),
        '--import "' + StagingPath + '"', '', SW_HIDE, ewWaitUntilTerminated, ResultCode)
       or (ResultCode <> 0) then
      RaiseException('Could not protect Groq key. Installation incomplete.');
    RegWriteDWordValue(HKCU, 'Software\Haronex', 'WebTransferAllowed', 1);
    RegWriteDWordValue(HKCU, 'Software\Haronex', 'LiveAutoReviewAllowed', 0);
  finally
    DeleteFile(StagingPath);
  end;
end;
