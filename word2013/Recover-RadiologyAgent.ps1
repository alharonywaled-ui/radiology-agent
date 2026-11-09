# Repair tool for Word 2013 crashes: disable only Radiology Agent COM auto-loading.
# Leaves the saved Gemini DPAPI key and all unrelated Office add-ins intact.
$ErrorActionPreference='Stop'
if (Get-Process WINWORD -ErrorAction SilentlyContinue) { throw 'Close all Word windows first.' }
$stamp=Get-Date -Format 'yyyyMMdd_HHmmss'
$dir=Join-Path ([Environment]::GetFolderPath('Desktop')) "RadiologyAgent_Recovery_$stamp"
New-Item -ItemType Directory -Path $dir -Force | Out-Null
$roots=@('HKCU:\Software\Microsoft\Office\Word\Addins','HKCU:\Software\Microsoft\Office\15.0\Word\Addins')
foreach($root in $roots) {
 if(-not(Test-Path $root)){continue}
 foreach($item in (Get-ChildItem $root)) {
  if($item.PSChildName -match '^radiology[. _-]*agent') {
   $native='HKCU\'+$item.Name.Substring('HKEY_CURRENT_USER\'.Length)
   $safe=($item.PSChildName -replace '[^a-zA-Z0-9._-]','_')
   & reg.exe export $native (Join-Path $dir "$safe.reg") /y | Out-Null
   New-ItemProperty -LiteralPath $item.PSPath -Name LoadBehavior -PropertyType DWord -Value 0 -Force | Out-Null
   Write-Host "Disabled: $native"
  }
 }
}
Write-Host "Backups: $dir"
Write-Host 'Gemini key was not touched. Reopen Word to confirm normal startup.'