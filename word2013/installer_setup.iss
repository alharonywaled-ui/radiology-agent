#define AppName "Radiology Agent for Word 2013"
#define AppVersion "2.2.0"
[Setup]
AppId={{F29657A1-86DE-4565-B391-E9E4D35A8219}
AppName={#AppName}
AppVersion={#AppVersion}
DefaultDirName={localappdata}\RadiologyAgent\Word2013Safe
DefaultGroupName=Radiology Agent
OutputDir=output
OutputBaseFilename=RadiologyAgent_Word2013_x86_Safe_Setup
Compression=lzma2
SolidCompression=yes
PrivilegesRequired=lowest
WizardStyle=modern
CloseApplications=yes
DisableProgramGroupPage=yes
UninstallDisplayName=Radiology Agent (Word 2013 x86)

[Files]
; Native 32-bit COM DLL: does not use managed COM activation through mscoree.dll.
Source: "bin\RadiologyAgent.NativeShim.dll"; DestDir: "{app}"; Flags: ignoreversion regserver
Source: "bin\RadiologyAgent.ReviewHost.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "bin\RadiologyAgent.KeySetup.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "bin\RadiologyAgent.KeySetup.exe"; DestDir: "{tmp}"; DestName: "KeyCheck.exe"; Flags: dontcopy

[Registry]
; Remove ONLY the known older Radiology Agent registrations. Does not touch Nitro/other add-ins.
Root: HKCU; Subkey: "Software\Microsoft\Office\Word\Addins\RadiologyAgent.WordAddin"; Flags: deletekey
Root: HKCU; Subkey: "Software\Microsoft\Office\15.0\Word\Addins\RadiologyAgent.WordAddin"; Flags: deletekey
Root: HKCU; Subkey: "Software\Classes\RadiologyAgent.WordAddin"; Flags: deletekey
Root: HKCU; Subkey: "Software\Classes\CLSID\{{D1FBEFD5-3B43-4516-9B7B-CDEB47AB0FA1}"; Flags: deletekey

[Code]
var
  KeyPage: TInputQueryWizardPage;
  NeedsKey: Boolean;
  ApiKey: String;

function InitializeSetup: Boolean;
var
  ReleaseNumber: Cardinal;
begin
  Result := RegQueryDWordValue(HKLM32, 'SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full', 'Release', ReleaseNumber)
    and (ReleaseNumber >= 528040);
  if not Result then
    MsgBox('Radiology Agent requires Microsoft .NET Framework 4.8 or newer. Install it before running this setup.', mbError, MB_OK);
end;

procedure InitializeWizard;
var
  CheckExitCode: Integer;
begin
  NeedsKey := True;
  if FileExists(ExpandConstant('{localappdata}\RadiologyAgent\key.dpapi')) then begin
    ExtractTemporaryFile('KeyCheck.exe');
    if Exec(ExpandConstant('{tmp}\KeyCheck.exe'), '--check', '',
      SW_HIDE, ewWaitUntilTerminated, CheckExitCode) then
      NeedsKey := CheckExitCode <> 0;
  end;
  KeyPage := CreateInputQueryPage(wpSelectDir, 'Haronex API Key',
     'One-time Haronex API Key',
     'Enter your Haronex API Key once. It will be encrypted for your Windows account and reused automatically on later installs.');
  KeyPage.Add('Haronex API Key:', True);
end;

function ShouldSkipPage(PageID: Integer): Boolean;
begin
  Result := (PageID = KeyPage.ID) and not NeedsKey;
end;

function NextButtonClick(CurPageID: Integer): Boolean;
begin
  Result := True;
  if CurPageID = KeyPage.ID then begin
    ApiKey := Trim(KeyPage.Values[0]);
    if (Length(ApiKey) < 30) or ((Copy(ApiKey,1,4) <> 'AIza') and (Copy(ApiKey,1,3) <> 'AQ.')) or (Pos(' ',ApiKey) > 0) then begin
      MsgBox('Enter a valid Haronex API Key (AQ. or AIza format).', mbError, MB_OK);
      Result := False;
    end;
  end;
end;

procedure CurStepChanged(CurStep: TSetupStep);
var
  DirectoryName, TempPath: String;
  ExitCode: Integer;
begin
  if CurStep = ssPostInstall then begin
    DirectoryName := ExpandConstant('{localappdata}\RadiologyAgent');
    ForceDirectories(DirectoryName);
    if NeedsKey then begin
      TempPath := DirectoryName + '\pending.key';
      if not SaveStringToFile(TempPath, ApiKey, False) then
        RaiseException('Failed to create temporary key file.');
      if not Exec(ExpandConstant('{app}\RadiologyAgent.KeySetup.exe'), '',
        ExpandConstant('{app}'), SW_HIDE, ewWaitUntilTerminated, ExitCode) then
        RaiseException('Failed to launch key encryption helper.');
      if ExitCode <> 0 then
        RaiseException('Unable to encrypt Haronex API Key. Exit: ' + IntToStr(ExitCode));
    end;
    ApiKey := '';
  end;
end;
