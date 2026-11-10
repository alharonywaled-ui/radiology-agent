#define AppName "Radiology Agent for Word 2013"
#define AppVersion "3.1.0"
[Setup]
AppId={{F29657A1-86DE-4565-B391-E9E4D35A8219}
AppName={#AppName}
AppVersion={#AppVersion}
DefaultDirName={localappdata}\RadiologyAgent\Word2013Safe
DefaultGroupName=Radiology Agent
OutputDir=output
OutputBaseFilename=Haronex_Word2013_AutoWeb_Setup
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
Source: "bin\HaronexWebWorker.exe"; DestDir: "{app}"; Flags: ignoreversion

[Registry]
; Remove ONLY the known older Radiology Agent registrations. Does not touch Nitro/other add-ins.
Root: HKCU; Subkey: "Software\Microsoft\Office\Word\Addins\RadiologyAgent.WordAddin"; Flags: deletekey
Root: HKCU; Subkey: "Software\Microsoft\Office\15.0\Word\Addins\RadiologyAgent.WordAddin"; Flags: deletekey
Root: HKCU; Subkey: "Software\Classes\RadiologyAgent.WordAddin"; Flags: deletekey
Root: HKCU; Subkey: "Software\Classes\CLSID\{{D1FBEFD5-3B43-4516-9B7B-CDEB47AB0FA1}"; Flags: deletekey

[Code]
var
  AuthorizationPage: TInputOptionWizardPage;

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
  AuthorizationPage := CreateInputOptionPage(wpSelectDir,
    'Haronex Web Data Transfer Permission', 'Authorize external processing once',
    'Haronex can automatically send the selected Word text, including clinical details or patient identifiers, to the external Gemini website. You are responsible for removing identifiers and obtaining organizational approval. This is an experimental web integration, not a guaranteed background API.',
    False, False);
  AuthorizationPage.Add('I confirm I am authorized to send my document text to this external service automatically.');
  AuthorizationPage.Values[0] := False;
end;

function NextButtonClick(CurPageID: Integer): Boolean;
begin
  Result := True;
  if CurPageID = AuthorizationPage.ID then begin
    if not AuthorizationPage.Values[0] then begin
      MsgBox('Please grant authorization, or cancel installation. Automatic transfer will remain disabled.', mbInformation, MB_OK);
      Result := False;
    end;
  end;
end;

procedure CurStepChanged(CurStep: TSetupStep);
begin
  if CurStep = ssPostInstall then begin
    if WizardSilent then Exit;
    if AuthorizationPage.Values[0] then
      RegWriteDWordValue(HKCU, 'Software\Haronex', 'WebTransferAllowed', 1);
  end;
end;
