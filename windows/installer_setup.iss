#define AppName "Radiology Agent"
#define AppVersion "1.0.0"
[Setup]
AppId={{13D71976-C125-4B84-9762-E3278B06D6C3}
AppName={#AppName}
AppVersion={#AppVersion}
DefaultDirName={localappdata}\RadiologyAgent
DefaultGroupName=Radiology Agent
OutputDir=output
OutputBaseFilename=RadiologyAgent_Setup_1.0.0
Compression=lzma2
SolidCompression=yes
PrivilegesRequired=lowest
WizardStyle=modern
DisableProgramGroupPage=yes
CloseApplications=yes

[Files]
Source: "..\manifest.xml"; DestDir: "{app}\Catalog"; Flags: ignoreversion
Source: "..\manifest.xml"; DestDir: "{localappdata}\Microsoft\Office\16.0\Wef"; DestName: "RadiologyAgent.manifest.xml"; Flags: ignoreversion
Source: "agent_bridge.ps1"; DestDir: "{app}"; Flags: ignoreversion
Source: "start_bridge.vbs"; DestDir: "{app}"; Flags: ignoreversion

[Registry]
; Sideload registration only; some Word versions require separate admin deployment or UI activation.
Root: HKCU; Subkey: "Software\Microsoft\Office\16.0\WEF\Developer"; ValueType: string; ValueName: "63d395aa-953b-4d8f-bf56-7a45327622fb"; ValueData: "{app}\Catalog\manifest.xml"; Flags: uninsdeletevalue
Root: HKCU; Subkey: "Software\Microsoft\Windows\CurrentVersion\Run"; ValueType: string; ValueName: "RadiologyAgentBridge"; ValueData: """{sys}\wscript.exe"" ""{app}\start_bridge.vbs"""; Flags: uninsdeletevalue

[Run]
Filename: "{sys}\wscript.exe"; Parameters: """{app}\start_bridge.vbs"""; Flags: nowait postinstall skipifsilent; Description: "Start Radiology Agent Bridge"

[Code]
var
 KeyPage: TInputQueryWizardPage;
 ApiKey: string;
 Configured: Boolean;

function JsonEscape(const S: string): string;
var
 I: Integer;
 C: string;
begin
 Result := '';
 for I := 1 to Length(S) do begin
  C := Copy(S, I, 1);
  if C = '\' then Result := Result + '\\'
  else if C = '"' then Result := Result + '\"'
  else if Ord(S[I]) < 32 then Result := Result + ' '
  else Result := Result + C;
 end;
end;

procedure InitializeWizard;
begin
 Configured := FileExists(ExpandConstant('{localappdata}\RadiologyAgent\config.json'));
 KeyPage := CreateInputQueryPage(wpSelectDir, 'Gemini API Configuration',
   'One-time Gemini API Key', 'Enter the API Key on first installation. Reinstalls reuse the local key.');
 KeyPage.Add('Gemini API Key:', True);
end;

function ShouldSkipPage(PageID: Integer): Boolean;
begin
 Result := Configured and (PageID = KeyPage.ID);
end;

function NextButtonClick(CurPageID: Integer): Boolean;
begin
 Result := True;
 if CurPageID = KeyPage.ID then begin
  ApiKey := Trim(KeyPage.Values[0]);
  if (Length(ApiKey) < 12) or (Pos(' ', ApiKey) > 0) then begin
   MsgBox('Enter a valid Gemini API Key.', mbError, MB_OK);
   Result := False;
  end;
 end;
end;

procedure CurStepChanged(CurStep: TSetupStep);
var
 ConfigPath: string;
 Contents: string;
begin
 if CurStep = ssPostInstall then begin
  ConfigPath := ExpandConstant('{app}\config.json');
  if not FileExists(ConfigPath) then begin
   Contents := '{"gemini_api_key":"' + JsonEscape(ApiKey) + '"}';
   if not SaveStringToFile(ConfigPath, Contents, False) then
    RaiseException('Failed to save config.json');
  end;
  ApiKey := '';
 end;
end;
