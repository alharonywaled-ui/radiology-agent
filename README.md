# Radiology Agent

Microsoft Word Office.js add-in for radiology report language and internal-consistency review with Gemini through a local Windows bridge. Run the GitHub Actions installer workflow to build an Inno Setup EXE. Never commit API keys or patient information.

**Hosting:** GitHub Pages must be enabled manually in repository Settings → Pages after making the repository public (or ensuring the plan supports private-repository Pages). The installer is not a signed or medically validated product.

**Install:** Download the `RadiologyAgent_Setup_Windows` artifact from Actions, unzip and run the EXE. Enter Gemini API key on first installation only. On later runs the per-user config persists.

**Important:** Microsoft Office web-add-in sideloading is not equivalent to a trusted COM/VSTO add-in; deployment or initial activation in Word may still be required. This installer does not bypass Office trust policies.
