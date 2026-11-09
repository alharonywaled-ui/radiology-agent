# ⚠️ Word 2013 native add-in: deployment blocked after a real Word crash

**9 Nov 2026:** The Windows CI COM activation test passed, but the add-in caused Microsoft Word 2013 (32-bit) to report serious Data Execution Prevention compatibility failures for `radiologyagent.wordaddin` and `radiology agent`. **Do not install or re-enable the current native EXE while the incident is unresolved.** CI does not run Word 2013 and therefore cannot certify Office compatibility.

To recover Word safely, close Word and use [Recover-RadiologyAgent.ps1](word2013/Recover-RadiologyAgent.ps1). This disables Radiology Agent Word startup without disabling DEP, touching other add-ins, or deleting the Gemini key. Next engineering step: debug against a real Word 2013 x86 test environment with crash dumps and Office event logs, then certify an isolated new COM registration before release.

---

# Radiology Agent for Microsoft Word 2013 (32-bit)

**Current installer:** [Native COM x86 Windows build](https://github.com/alharonywaled-ui/radiology-agent/actions/runs/42646949868) → **Artifacts** → `RadiologyAgent_Word2013_x86_TestedSetup`.

The **Word 2013-compatible edition** is stored in `word2013/`. It has a native 32-bit C++ COM Ribbon loader, a .NET Framework 4.8 review application, and a per-user DPAPI key helper. The installer displays **Gemini API Key** on a first setup when no valid saved key exists. Reinstalls reuse the valid saved key without prompting. The saved key is located in `%LOCALAPPDATA%\RadiologyAgent\key.dpapi` (encrypted to the current Windows account).

## Installation

1. Close Word completely.
2. Run `RadiologyAgent_Word2013_x86_Setup.exe` under your normal Windows user account.
3. Enter your Gemini API key **once** if requested (Google AI Studio key begins with `AIza`).
4. Start Word 2013 (32-bit). The **Radiology Agent** Ribbon tab contains **Review Text**. **Do not look inside Insert → My Apps**; that's for Office web apps, not COM add-ins.
5. Select report text and click **Review Text → Review By Agent**. Review all AI-generated comments before accepting changes.

## Checks performed

The Windows build compiled the C# review host and key helper, compiled the native 32-bit C++ COM DLL, installed the EXE silently, verified DPAPI read/write, verified Word's LoadBehavior=3 registration, and activated the COM DLL through `CoCreateInstance` requesting the Office Ribbon interface and retrieving the Ribbon XML. **These automated checks passed on November 9, 2026.** The workflow does **not** contain Word 2013 itself and cannot verify Ribbon display in that product; test that separately on your Windows machine.

## Limitations and safety

- Microsoft .NET Framework 4.8+ is required for the review window. Windows 32-bit Word is required for this installer.
- Office 2013 is no longer supported by Microsoft. Proper Office activation may be required for reliable use.
- Text submitted for review is sent to Google's Gemini API. **Do not process patient-identifiable data** unless your organization has authorized this transfer with appropriate contractual, security and privacy controls. The tool does not automatically de-identify reports.
- Language errors are highlighted yellow and potential internal logical contradictions red. Comments are AI suggestions, not clinical validation.
- If the tab is missing, see **Word → File → Options → Add-ins → COM Add-ins → Go**, not My Apps. Check whether **Radiology Agent** is loaded.

## Source structure

- `word2013/NativeShim.cpp`: native COM Ribbon integration
- `word2013/ReviewHost.cs`, `word2013/RadiologyAgent.WordAddin.cs`: review UI, Word automation, Gemini API requests
- `word2013/KeySetup.cs`: Windows DPAPI encryption and saved key verification
- `word2013/installer_setup.iss`: Inno Setup installation
- `.github/workflows/build-word2013.yml`: Windows compilation and native COM smoke test

The root `manifest.xml`, `taskpane.*`, and `windows/` files are a **legacy Office.js prototype** that does not work for this Word 2013 scenario. Do not use that old EXE.
