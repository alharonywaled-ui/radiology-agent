# Radiology Agent for Microsoft Word 2013 (32-bit)

## Product / operating requirements

- Native **C# managed COM add-in**, targeting x86 and **.NET Framework 4.8**. It implements `IDTExtensibility2` and the Office 2007+ Ribbon XML interface.
- Designed for **Microsoft Word 2013 32-bit**; **does not require Office.js, GitHub Pages, or WebView2**.
- New ribbon **Radiology Agent → Review Text** opens a review dialog.
- Select text or leave selection empty to review the whole document; maximum 18,000 characters. Gemini reviews English-language wording and potential *internal* contradictions. Qualified clinicians must validate every suggestion.
- Gemini output is matched to exact text in the report; duplicate or unmatched quotes are skipped. Only uniquely matched excerpts get yellow (language) or red (potential internal inconsistency) highlighting, and Word comments. No automatic rewriting.
- The installation wizard asks for the Gemini API key on first installation only. The per-user key is DPAPI-protected under `%LOCALAPPDATA%\RadiologyAgent\key.dpapi`; reinstall uses the same key. A temporary plaintext file (`pending.key`) exists briefly during installation and is deleted by the key-encryption helper. Do not sync or back up that temporary file.

## Build the Windows installer

Open [GitHub Actions](https://github.com/alharonywaled-ui/radiology-agent/actions/workflows/build-word2013.yml) and run **Build Word 2013 x86 COM Installer**, or push to the `main` branch. It uses Windows MSBuild, the Inno Setup compiler, and performs a silent-install COM activation smoke test. Download `RadiologyAgent_Word2013_x86_Setup.exe` from the `RadiologyAgent_Word2013_x86_Setup` artifact.

## Install

1. Close all Word windows. Ensure that Word 2013 is **legitimately activated** and .NET Framework 4.8 is installed on Windows.
2. Run the EXE under the same Windows user who uses Word. This is a **per-user** installation and does not require elevated administrator privileges.
3. Enter a Gemini API key. Only the user account can decrypt the protected key. The installer does not ask again while the DPAPI key file is present.
4. Start Word. A new **Radiology Agent** tab should appear, containing **Review Text**. This is a native COM add-in, not listed in **Insert → My Apps**.
5. Select a report section and choose **Review Text → Review By Agent**. Test first on synthetic/de-identified documents.

## Troubleshooting

If the tab is missing, open **File → Options → Add-ins**, select **COM Add-ins → Go**, and inspect `Radiology Agent` (including its disabled state). Check the per-user registry key:

```powershell
Get-ItemProperty 'HKCU:\Software\Microsoft\Office\Word\Addins\RadiologyAgent.WordAddin'
```

If Word says it could not load the add-in, inspect `HKCU:\Software\Classes\CLSID\{D1FBEFD5-3B43-4516-9B7B-CDEB47AB0FA1}\InprocServer32`, check the .NET Framework installation and Office bitness. Do not enable untrusted add-ins or bypass Office security policies. This deployment uses a manual per-user .NET COM registration and should be tested on the intended Word 2013 workstation.

## Security and medical review

**Important:** Clicking Review sends text to the Google Gemini API. Do **not** use identifiable patient reports unless your organization has approved the provider, contracts, patient-data handling, retention, and international transfer conditions. This system does not automatically remove patient identifiers. Clinical statements are not verified against imaging. Review all annotations yourself. Office 2013 is end-of-support; Microsoft no longer provides its regular security updates. A valid Office license is necessary; this installer does not activate Office.

## Legacy project

The root-level `manifest.xml`, `taskpane.*`, and `windows/` folder are the older Office.js prototype for newer Office versions. **Do not use the old installer for Word 2013**. The native add-in source is in `word2013/`.
