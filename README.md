# Haronex Live Review v6.0 — Word 2013 x86

**One-click apply, background pre-review:** The native Word Ribbon add-in starts a single persistent, invisible .NET Framework 4.8 STA monitor at Word startup. The monitor detects opening and switching active documents and polls the main-document body every 400 ms; Word 2013 classic COM does not supply a general keystroke event. It waits **1200 ms after the most recent detected text change** (debounce) before preparing the next review. Gemini Web Guest Mode processes report chunks in a hidden experimental Edge browser worker; the first request prioritizes internally supported clinical contradictions, and a second independent audit searches for additional findings.

While text is changing, in-flight work for superseded revisions is canceled, and mismatched results are discarded. Completed responses are validated (guest-only, matching version, two passes, exact quotes) and DPAPI-cached by report text. **The monitor does not insert Word comments before the user clicks Review Text.** Clicking the button applies the ready matching results immediately; if the current text is still being processed, the click queues the apply until review succeeds or fails. Logic annotations (red) precede language annotations (yellow), and existing comment positions are not duplicated. The only success/error popups are **تمت المراجعة** or **رجاءً أعد المحاولة** on click-related completion. The Ribbon icon is green for ready and red for failure.

**Explicit data-transfer authorization:** Because this version automatically sends Word report text to an external consumer site as soon as a document is opened/edited, the installer demands a *new* explicit approval. The old v5 WebTransferAllowed flag alone does NOT enable this feature. The silent installer never grants it. Do not enable automatic transfer for documents with identifiable patient data unless approved by the organization and covered by appropriate safeguards.

**Limitations:** The active main body of the Word document is monitored, not headers, text boxes or every StoryRange. Polling long documents can consume resources and a first-time uncached analysis still takes the external service's response time. Guest availability and headless Gemini Web behavior are not guaranteed; the website is not an official integration. Switching active documents cancels outdated work; no results for older text are silently applied. A completed response does not certify clinical correctness, and a person must verify possible internal contradictions.

**Validation:** Windows CI compiles x86 COM and the standalone worker, runs offline result-validation/debounce checks, checks JSON parsing and installer activation. It does not test a real logged-in/guest Gemini Web session or live Word 2013 editing; end-to-end user testing is required.

---

# Haronex Comprehensive Review v5.0 — Word 2013 x86

**One-click, Guest-only, hidden-browser experimental review.** For every report chunk Haronex runs two separate Gemini Web reviews:

1. **Structured first pass:** explicit sentence-by-sentence language/grammar review, medical terminology consistency, internally supported logic contradictions, and final completeness scan.
2. **Independent second pass:** audits the same chunk while knowing the earlier quotes, to seek additional distinct issues. The worker validates quoted phrases against the source, merges findings deterministically, and removes duplicates before Word annotations.

No manual steps or new buttons are introduced. Errors still produce the generic **رجاءً أعد المحاولة** popup; completed workflows produce **تمت المراجعة**. The Ribbon status icon is green/red. Every chunk must complete *both* passes; partial web responses are rejected rather than quietly marked successful. Cached earlier versions are isolated by an algorithm-version key; unchanged documents reuse an encrypted review for speed. This version typically performs **2 web requests per chunk**, so new uncached reviews can take longer and may hit Guest Mode availability/rate limits. It cannot guarantee every possible radiological error or full clinical consistency across separate chunks; medical review remains necessary.

Avoid transmitting identifiable patient data unless expressly approved by the organization. The program does not use an API key or Google login. Site automation is not an official Gemini integration and requires end-to-end testing in the intended environment.

---

# Haronex Guest-Only Web Mode (Word 2013 x86)

This version uses only **signed-out Guest Mode** in a disposable, headless Microsoft Edge browser profile. It requires no Google account, password, API key, or browser login. It never reuses a previously signed-in Haronex/Edge profile. Word's **Radiology Agent → Review Text** triggers review automatically, then applies yellow language highlights, red potential-internal-contradiction highlights, and Word comments if a complete, valid JSON response is obtained.

Guest access to Gemini text chats is not guaranteed by Google for all devices, regions, or sessions; if it is unavailable or the site asks for sign-in, Haronex fails clearly rather than circumventing the restriction. Selenium site automation itself remains experimental and cannot be guaranteed to work with future site updates. First-run installer authorization for external text transfer remains mandatory. **Never send identifiable patient reports without explicit organizational approval.** Repeated unchanged-text reviews use a DPAPI-encrypted 30-day cache, separated from earlier signed-in-mode caches. The old `%LOCALAPPDATA%\\Haronex\\BrowserProfile` directory from prior versions is not used and can be removed separately by an administrator after confirming it is no longer needed.

CI validates Windows packaging, the offline JSON parser, installer, and COM activation; it does not execute live Gemini guest-mode tests or clinical validation.


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
