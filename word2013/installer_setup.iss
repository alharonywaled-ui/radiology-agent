#define AppName "Radiology Agent for Word 2013"
#define AppVersion "2.0.0"
#define Clsid "{{D1FBEFD5-3B43-4516-9B7B-CDEB47AB0FA1}"
#define ProgId "RadiologyAgent.WordAddin"

[Setup]
AppId={{30BC199A-18A5-4B07-8184-2899E30E2CD3}
AppName={#AppName}
AppVersion={#AppVersion}
DefaultDirName={localappdata}\RadiologyAgent\Word2013
DefaultGroupName=Radiology Agent
OutputDir=output
OutputBaseFilename=RadiologyAgent_Word2013_x86_Setup
Compression=lzma2
SolidCompression=yes
PrivilegesRequired=lowest
WizardStyle=modern
CloseApplications=yes
UninstallDisplayName=Radiology Agent (Word 2013 x86)
; This per-user add-in is for 32-bit Word only and needs a properly licensed Word installation.

[Files]
Source: "bin\RadiologyAgent.WordAddin.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "bin\RadiologyAgent.KeySetup.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "bin\RadiologyAgent.KeySetup.exe"; DestDir: "{tmp}"; DestName: "KeyCheck.exe"; Flags: dontcopy

[Registry]
; Per-user .NET COM class activation: 32-bit Word reads the 32-bit registry view.
Root: HKCU; Subkey: "Software\Classes\{#ProgId}"; ValueType: string; ValueName: ""; ValueData: "Radiology Agent Word Add-in"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Classes\{#ProgId}\CLSID"; ValueType: string; ValueName: ""; ValueData: "{#Clsid}"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Classes\CLSID\{#Clsid}"; ValueType: string; ValueName: ""; ValueData: "RadiologyAgent.WordAddin"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Classes\CLSID\{#Clsid}\InprocServer32"; ValueType: string; ValueName: ""; ValueData: "mscoree.dll"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Classes\CLSID\{#Clsid}\InprocServer32"; ValueType: string; ValueName: "ThreadingModel"; ValueData: "Both"
Root: HKCU; Subkey: "Software\Classes\CLSID\{#Clsid}\InprocServer32"; ValueType: string; ValueName: "Class"; ValueData: "RadiologyAgent.WordAddin"
Root: HKCU; Subkey: "Software\Classes\CLSID\{#Clsid}\InprocServer32"; ValueType: string; ValueName: "Assembly"; ValueData: "RadiologyAgent.WordAddin, Version=0.0.0.0, Culture=neutral, PublicKeyToken=null"
Root: HKCU; Subkey: "Software\Classes\CLSID\{#Clsid}\InprocServer32"; ValueType: string; ValueName: "RuntimeVersion"; ValueData: "v4.0.30319"
Root: HKCU; Subkey: "Software\Classes\CLSID\{#Clsid}\InprocServer32"; ValueType: string; ValueName: "CodeBase"; ValueData: "{code:ManagedCodeBase}"
Root: HKCU; Subkey: "Software\Classes\CLSID\{#Clsid}\InprocServer32\0.0.0.0"; ValueType: string; ValueName: "Class"; ValueData: "RadiologyAgent.WordAddin"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Classes\CLSID\{#Clsid}\InprocServer32\0.0.0.0"; ValueType: string; ValueName: "Assembly"; ValueData: "RadiologyAgent.WordAddin, Version=0.0.0.0, Culture=neutral, PublicKeyToken=null"
Root: HKCU; Subkey: "Software\Classes\CLSID\{#Clsid}\InprocServer32\0.0.0.0"; ValueType: string; ValueName: "RuntimeVersion"; ValueData: "v4.0.30319"
Root: HKCU; Subkey: "Software\Classes\CLSID\{#Clsid}\InprocServer32\0.0.0.0"; ValueType: string; ValueName: "CodeBase"; ValueData: "{code:ManagedCodeBase}"
Root: HKCU; Subkey: "Software\Classes\CLSID\{#Clsid}\ProgId"; ValueType: string; ValueName: ""; ValueData: "{#ProgId}"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Microsoft\Office\Word\Addins\{#ProgId}"; ValueType: string; ValueName: "FriendlyName"; ValueData: "Radiology Agent"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Microsoft\Office\Word\Addins\{#ProgId}"; ValueType: string; ValueName: "Description"; ValueData: "AI radiology proofreading — verify all suggestions"
Root: HKCU; Subkey: "Software\Microsoft\Office\Word\Addins\{#ProgId}"; ValueType: dword; ValueName: "LoadBehavior"; ValueData: "3"

[Code]
var
  KeyPage: TInputQueryWizardPage;
  NeedsKey: Boolean;
  ApiKey: String;

function ManagedCodeBase(Param: String): String;
var
  P: String;
begin
  P := ExpandConstant('{app}\RadiologyAgent.WordAddin.dll');
  StringChangeEx(P, '\', '/', True);
  StringChangeEx(P, ' ', '%20', True);
  Result := 'file:///' + P;
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
  KeyPage := CreateInputQueryPage(wpSelectDir,'Gemini API Key',
     'Only required on first installation',
     'Enter the key once. It is encrypted for your Windows account using DPAPI. Reinstalls reuse the saved key.');
  KeyPage.Add('Gemini API Key:', True);
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
    if (Length(ApiKey) < 30) or (Copy(ApiKey, 1, 4) <> 'AIza') or (Pos(' ', ApiKey) > 0) then begin
      MsgBox('Enter a valid Gemini API Key.', mbError, MB_OK);
      Result := False;
    end;
  end;
end;

function InitializeSetup: Boolean;
begin
  Result := True;
  if not RegKeyExists(HKLM, 'SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full') then begin
    MsgBox('Microsoft .NET Framework 4.5 or newer is required. Install it before this add-in.', mbError, MB_OK);
    Result := False;
  end;
  if Result and not FileExists(ExpandConstant('{sys}\mscoree.dll')) then begin
    MsgBox('The 32-bit .NET Framework COM runtime (mscoree.dll) is missing. Repair or enable .NET Framework 4.8 before installing. Installation has been cancelled to prevent an unloaded Word add-in.', mbError, MB_OK);
    Result := False;
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
        RaiseException('Cannot write temporary API key');
      if not Exec(ExpandConstant('{app}\RadiologyAgent.KeySetup.exe'),
          '', ExpandConstant('{app}'), SW_HIDE, ewWaitUntilTerminated, ExitCode) then
        RaiseException('Failed to launch secure credential setup');
      if ExitCode <> 0 then RaiseException('Could not encrypt API key (exit ' + IntToStr(ExitCode) + ')');
    end;
    ApiKey := '';
  end;
end;
