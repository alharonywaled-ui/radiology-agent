# Haronex Enterprise 11.12.13 (DSCAN)

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

## Setup checkboxes
Every interactive installer shows both options. Enterprise saves them in HKLM (32-bit); the machine setting takes precedence over HKCU, even when disabled. Configure my account displays managed options read-only. Ordinary installer uses HKCU. Upgrades preselect existing values; new installs default OFF. Silent deployment accepts /MANDATORYREVIEW=1 or 0 and /DISABLEAUTORECOVER=1 or 0; omitted flags preserve existing settings. AutoRecover changes are applied when Word connects or Review Text is clicked. Turning it off restores a previously recorded nonzero interval, if available; earlier versions did not record it, so manual Word configuration may be required. Close Word before installing/updating. The controls require the add-in to be loaded; they do not prohibit Word safe mode or all manual administrative changes.

## Recovery and installation integrity
Setup verifies the manifest hashes of the complete application and packaged runtime, then performs a local Host/Worker protocol handshake without any Groq request. Close Word in all sessions before Enterprise installation/repair. Setup migrates loaded user profiles automatically; a visible common Startup migration shortcut handles other users at their next logon (before opening Word). If Word is already open, close it and run the migration shortcut or sign out and back in. The ordinary installer repairs its own user installation. If protected files are missing, IT must rerun the approved installer; the user-side add-in does not bypass administrator rights or security software. No silent key sharing or consent changes. If security software quarantined a file, review the detection through IT security management before repair. Run installed KeySetup --verify-install for hash diagnostics and ReviewHost --worker-health for the offline protocol probe. A successful installer probe does not validate live Groq access, domain-admin removal, or real clinical/Word behavior.

Both interactive installers show Word Review Settings before credential pages on every setup run. Existing key files are preserved byte-for-byte, irrespective of whether consent is present; existing consent is not changed. Use Configure my account explicitly to repair an invalid key or change consent. Silent installs have no UI; use /MANDATORYREVIEW=1 or 0 and /DISABLEAUTORECOVER=1 or 0, or omit to preserve existing choices. Enterprise machine policy remains editable by authorized installation administrators, not ordinary employees.
