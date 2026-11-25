#define AppName "Haronex Enterprise for Word 2013"
#define AppVersion "11.12.15"
[Setup]
AppId={{56B8B1AF-99B7-4580-821E-276FF00C79B6}
AppName={#AppName}
AppVersion={#AppVersion}
DefaultDirName={pf32}\Haronex
DefaultGroupName=Haronex Enterprise
OutputDir=output
OutputBaseFilename=Haronex_Word2013_V11_12_15_Enterprise_Setup
Compression=lzma2
SolidCompression=yes
PrivilegesRequired=admin
DisableDirPage=yes
UsePreviousAppDir=no
DisableProgramGroupPage=yes
CloseApplications=yes
WizardStyle=modern
UninstallDisplayName=Haronex Enterprise (Word 2013 x86)
UninstallDisplayIcon={app}\RadiologyAgent.ReviewHost.exe

[Files]
Source: "bin\HaronexFiles.sha256"; DestDir: "{app}"; Flags: ignoreversion
Source: "bin\NDP48-x86-x64-AllOS-ENU.exe"; Flags: dontcopy
Source: "bin\RadiologyAgent.KeySetup.exe"; Flags: dontcopy
Source: "bin\RadiologyAgent.KeySetup.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "bin\RadiologyAgent.NativeShim.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "bin\RadiologyAgent.ReviewHost.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "bin\HaronexGroqWorker\HaronexGroqWorker.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "bin\HaronexGroqWorker\_HaronexRuntime\*"; DestDir: "{app}\_HaronexRuntime"; Flags: ignoreversion recursesubdirs createallsubdirs

[Registry]
Root: HKLM32; Subkey: "Software\Haronex\Enterprise"; ValueType: string; ValueName: "InstallPath"; ValueData: "{app}"; Flags: uninsdeletekey
Root: HKLM32; Subkey: "Software\Haronex\Enterprise"; ValueType: dword; ValueName: "MandatoryReviewEnabled"; ValueData: "{code:PolicyMandatory}"
Root: HKLM32; Subkey: "Software\Haronex\Enterprise"; ValueType: dword; ValueName: "DisableAutoRecover"; ValueData: "{code:PolicyRecovery}"
Root: HKLM32; Subkey: "Software\Classes\RadiologyAgent.Word2013.Safe"; ValueType: string; ValueData: "Radiology Agent for Word 2013"; Flags: uninsdeletekey
Root: HKLM32; Subkey: "Software\Classes\RadiologyAgent.Word2013.Safe\CLSID"; ValueType: string; ValueData: "{{7B918AD1-1C42-4B73-A4BB-9FC83743DD6E}"
Root: HKLM32; Subkey: "Software\Classes\CLSID\{{7B918AD1-1C42-4B73-A4BB-9FC83743DD6E}"; ValueType: string; ValueData: "RadiologyAgent.Word2013.Safe"; Flags: uninsdeletekey
Root: HKLM32; Subkey: "Software\Classes\CLSID\{{7B918AD1-1C42-4B73-A4BB-9FC83743DD6E}\InprocServer32"; ValueType: string; ValueData: "{app}\RadiologyAgent.NativeShim.dll"
Root: HKLM32; Subkey: "Software\Classes\CLSID\{{7B918AD1-1C42-4B73-A4BB-9FC83743DD6E}\InprocServer32"; ValueType: string; ValueName: "ThreadingModel"; ValueData: "Apartment"
Root: HKLM32; Subkey: "Software\Classes\CLSID\{{7B918AD1-1C42-4B73-A4BB-9FC83743DD6E}\ProgID"; ValueType: string; ValueData: "RadiologyAgent.Word2013.Safe"
Root: HKLM32; Subkey: "Software\Microsoft\Office\Word\Addins\RadiologyAgent.Word2013.Safe"; ValueType: string; ValueName: "FriendlyName"; ValueData: "Radiology Agent"; Flags: uninsdeletekey
Root: HKLM32; Subkey: "Software\Microsoft\Office\Word\Addins\RadiologyAgent.Word2013.Safe"; ValueType: string; ValueName: "Description"; ValueData: "AI report proofreading. Clinical verification required."
Root: HKLM32; Subkey: "Software\Microsoft\Office\Word\Addins\RadiologyAgent.Word2013.Safe"; ValueType: dword; ValueName: "LoadBehavior"; ValueData: "3"

[Icons]
Name: "{commonstartup}\Haronex Enterprise profile migration"; Filename: "{app}\RadiologyAgent.KeySetup.exe"; Parameters: "--migrate-legacy"; Flags: runminimized
Name: "{commonprograms}\Haronex Enterprise\Configure my account"; Filename: "{app}\RadiologyAgent.KeySetup.exe"; Parameters: "--configure"
Name: "{commonprograms}\Haronex Enterprise\Uninstall"; Filename: "{uninstallexe}"

[Code]
var
  PolicyPage: TWizardPage;
  MandatoryCheck: TNewCheckBox;
  RecoveryCheck: TNewCheckBox;

function InitialPolicy(Name, ParamName: String): Boolean;
var Value: Cardinal; ParamValue: String;
begin
  Result := RegQueryDWordValue(HKLM32, 'Software\Haronex\Enterprise', Name, Value) and (Value = 1);
  ParamValue := ExpandConstant('{param:' + ParamName + '|}');
  if ParamValue = '1' then Result := True;
  if ParamValue = '0' then Result := False;
end;

procedure InitializeWizard;
begin
  PolicyPage := CreateCustomPage(wpWelcome, 'Word Review Settings',
    'Change these options on every installation, update or repair. They apply to all users; existing Groq keys are preserved.');
  MandatoryCheck := TNewCheckBox.Create(PolicyPage);
  MandatoryCheck.Parent := PolicyPage.Surface;
  MandatoryCheck.Left := 0;
  MandatoryCheck.Top := ScaleY(12);
  MandatoryCheck.Width := PolicyPage.SurfaceWidth;
  MandatoryCheck.Height := ScaleY(48);
  MandatoryCheck.Caption := 'Enable Mandatory Review (Review Text + Clear Both if findings before Save / Print)';
  MandatoryCheck.Checked := InitialPolicy('MandatoryReviewEnabled', 'MANDATORYREVIEW');
  RecoveryCheck := TNewCheckBox.Create(PolicyPage);
  RecoveryCheck.Parent := PolicyPage.Surface;
  RecoveryCheck.Left := 0;
  RecoveryCheck.Top := ScaleY(75);
  RecoveryCheck.Width := PolicyPage.SurfaceWidth;
  RecoveryCheck.Height := ScaleY(40);
  RecoveryCheck.Caption := 'Disable AutoRecover (unsaved work may be lost)';
  RecoveryCheck.Checked := InitialPolicy('DisableAutoRecover', 'DISABLEAUTORECOVER');
end;

function ShouldSkipPage(PageID: Integer): Boolean;
begin
  Result := False;
end;

function PolicyMandatory(Param: String): String;
begin
  Result := IntToStr(Ord(MandatoryCheck.Checked));
end;

function PolicyRecovery(Param: String): String;
begin
  Result := IntToStr(Ord(RecoveryCheck.Checked));
end;

function HasDotNet48: Boolean;
var ReleaseNumber: Cardinal;
begin
  Result := RegQueryDWordValue(HKLM, 'SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full', 'Release', ReleaseNumber) and (ReleaseNumber >= 528040);
end;

function PrepareToInstall(var NeedsRestart: Boolean): String;
var ExitCode: Integer;
begin
  Result := '';
  NeedsRestart := False;
  // Machine location must not be redirected into a user-writable directory.
  if CompareText(ExpandConstant('{app}'), ExpandConstant('{pf32}\Haronex')) <> 0 then begin
    Result := 'Haronex Enterprise must be installed in Program Files (x86)\Haronex.';
    Exit;
  end;
  if not HasDotNet48 then begin
    ExtractTemporaryFile('NDP48-x86-x64-AllOS-ENU.exe');
    if not Exec(ExpandConstant('{tmp}\NDP48-x86-x64-AllOS-ENU.exe'), '/q /norestart', '', SW_HIDE, ewWaitUntilTerminated, ExitCode) then begin
      Result := 'Could not start .NET Framework 4.8 setup.';
      Exit;
    end;
    if (ExitCode = 3010) or (ExitCode = 1641) then begin
      NeedsRestart := True;
      Result := 'Restart Windows and rerun Haronex Enterprise setup.';
      Exit;
    end;
    if not HasDotNet48 then begin
      Result := '.NET Framework 4.8 prerequisite not detected.';
      Exit;
    end;
  end;
  ExtractTemporaryFile('RadiologyAgent.KeySetup.exe');
  if not Exec(ExpandConstant('{tmp}\RadiologyAgent.KeySetup.exe'), '--check-domain', '', SW_HIDE, ewWaitUntilTerminated, ExitCode) then begin
    Result := 'Unable to verify the joined domain. Contact IT.';
    Exit;
  end;
  if ExitCode <> 0 then begin
    Result := 'Haronex Enterprise requires a computer joined to dscan.corp and domain connectivity.';
    Exit;
  end;
  if not Exec(ExpandConstant('{tmp}\RadiologyAgent.KeySetup.exe'), '--prepare-machine-install', '', SW_HIDE, ewWaitUntilTerminated, ExitCode) or (ExitCode <> 0) then begin
    Result := 'Close Word in all sessions before installing or repairing Haronex. Check with IT if a review process cannot be stopped.';
    Exit;
  end;
end;

procedure CurStepChanged(CurStep: TSetupStep);
var ExitCode: Integer;
begin
  if CurStep <> ssPostInstall then Exit;
  if not Exec(ExpandConstant('{app}\RadiologyAgent.KeySetup.exe'), '--verify-install', '', SW_HIDE, ewWaitUntilTerminated, ExitCode) or (ExitCode <> 0) then
    RaiseException('Haronex installation integrity check failed. Re-run the approved installer; check security software logs with IT.');
  if not Exec(ExpandConstant('{app}\RadiologyAgent.ReviewHost.exe'), '--worker-health', '', SW_HIDE, ewWaitUntilTerminated, ExitCode) or (ExitCode <> 0) then
    RaiseException('Haronex local worker handshake failed. Check diagnostics.log; installation was not verified.');
  if not Exec(ExpandConstant('{app}\RadiologyAgent.KeySetup.exe'), '--migrate-loaded-users', '', SW_HIDE, ewWaitUntilTerminated, ExitCode) or (ExitCode <> 0) then
    RaiseException('Haronex user profile migration failed. Close Word in all sessions and rerun Setup.');
end;

function InitializeUninstall: Boolean;
var ExitCode: Integer;
begin
  Result := Exec(ExpandConstant('{app}\RadiologyAgent.KeySetup.exe'), '--authorize-uninstall', '', SW_HIDE, ewWaitUntilTerminated, ExitCode);
  if Result then Result := ExitCode = 0;
  if not Result then begin
    Log('Haronex Enterprise: uninstall authorization denied.');
    if not UninstallSilent then
      MsgBox('Uninstallation requires an elevated DSCAN Domain Admin account and domain connectivity. Contact your IT Administrator.', mbError, MB_OK);
  end;
end;
