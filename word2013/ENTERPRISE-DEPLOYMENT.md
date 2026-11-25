# Haronex Enterprise 11.12.10 (DSCAN)

This is a separate machine-wide Word 2013 x86 package. Keep the existing personal installer for recovery/testing only. Install the enterprise EXE from an elevated administrator session on a computer joined to dscan.corp. Close Word and Haronex before upgrading. Files inherit Program Files ACLs. Do not grant standard users Modify or Full Control there.

For existing per-user installations, while Word is closed, deploy the following as a **user logon script** via GPO (not as SYSTEM or as the administrator's user):
```powershell
$exe = Join-Path ${env:ProgramFiles(x86)} 'Haronex\RadiologyAgent.KeySetup.exe'
if (!(Test-Path $exe)) { $exe = Join-Path $env:ProgramFiles 'Haronex\RadiologyAgent.KeySetup.exe' }
if (Test-Path $exe) {
    & $exe --migrate-legacy
    if ($LASTEXITCODE -ne 0) { Write-Warning "Haronex migration failed: $LASTEXITCODE" }
}
```
Run once per affected user, before opening Word. It removes only old per-user COM registrations and retains DPAPI key, consent, settings and old files. No automatic data deletion. A legacy uninstall entry may remain; removing it does not remove Enterprise. For new users, open Start > Haronex Enterprise > Configure my account, enter the user's Groq key and authorize external processing if approved by the organization. Existing configured users keep their settings.

The official Enterprise uninstaller checks the elevated process token for the actual joined domain's Domain Admins SID (RID 512); it refuses removal if verification fails. A local administrator can ultimately override local ACLs/registry or manually remove software. Remove employee local admin rights with Group Policy. This package does not forcibly re-enable an Office add-in or prevent Office safe mode. For enforced Office add-in loading, evaluate the Office managed add-ins GPO separately. Domain connectivity is required for installation and authorized uninstall. Test one joined PC with a standard user, a local admin, and a Domain Admin before rollout. Preserve an authorized IT recovery procedure.
