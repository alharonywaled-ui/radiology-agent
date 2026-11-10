#define AppName "Haronex Radiology Agent for Word 2013"
#define AppVersion "7.2.0"
[Setup]
AppId={{F29657A1-86DE-4565-B391-E9E4D35A8219}
AppName={#AppName}
AppVersion={#AppVersion}
DefaultDirName={localappdata}\RadiologyAgent\Word2013Safe
DefaultGroupName=Haronex
OutputDir=output
OutputBaseFilename=Haronex_Word2013_Groq_Qwen38_V7_2_Setup
Compression=lzma2
SolidCompression=yes
PrivilegesRequired=lowest
WizardStyle=modern
CloseApplications=yes
DisableProgramGroupPage=yes
UninstallDisplayName=Haronex (Word 2013 x86)

[Files]
Source: "bin\RadiologyAgent.NativeShim.dll"; DestDir: "{app}"; Flags: ignoreversion regserver
Source: "bin\RadiologyAgent.ReviewHost.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "bin\HaronexGroqWorker.exe"; DestDir: "{app}"; Flags: ignoreversion
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

function InitializeSetup: Boolean;
var
  ReleaseNumber: Cardinal;
begin
  Result := RegQueryDWordValue(HKLM32, 'SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full', 'Release', ReleaseNumber)
    and (ReleaseNumber >= 528040);
  if not Result then
    MsgBox('Haronex requires Microsoft .NET Framework 4.8 or newer.', mbError, MB_OK);
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
  if WizardSilent then Exit;
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
