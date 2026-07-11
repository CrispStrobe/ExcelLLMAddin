# Installers

Native installers for the two editions. Both are **pointer-light**: the macOS
`.pkg` mechanics are automatable/CI-able; only the payload differs.

## `mac-xlam/` — offline VBA `.xlam` (recommended for offline)

Installs the self-contained `ExcelLLMAddin.xlam` into Excel-for-Mac's add-ins
folder. The `.xlam` runs **fully offline** (VBA + local Ollama) — no hosting, no
server. After install, enable it once in Excel (Tools ▸ Excel Add-ins) or
double-click the `.xlam`.

```bash
# Prereq: a built ExcelLLMAddin.xlam at the repo root (see "Building the .xlam").
installer/mac-xlam/build-xlam-pkg.sh          # unsigned (local test)

SIGN_IDENTITY="Developer ID Installer: Name (TEAM)" \
AC_KEY_ID=.. AC_ISSUER_ID=.. AC_KEY_PATH=AuthKey_XXXX.p8 \
installer/mac-xlam/build-xlam-pkg.sh          # signed + notarized
```

CI: `.github/workflows/installer-mac.yml` builds, signs, and notarizes it on a
`macos-latest` runner from your Apple secrets.

## `mac/` — Office.js manifest (online, cross-platform)

Installs the Office.js manifest into the sideload (`wef`) folder. The add-in loads
from the hosted origin (GitHub Pages), so it needs internet. Same `.pkg` tooling.

## Building the `.xlam` (the one part that isn't automatable on Mac)

A `.xlam` is compiled VBA; producing it **requires Excel**, and Excel for Mac has
no VBA-project automation. So the `.xlam` cannot be built on Mac or on
GitHub-hosted CI. Options:

- **Windows + Excel:** `pwsh tools/Build-Addin.ps1` (imports the modules and saves
  the `.xlam`). Repeatable; can run on a self-hosted Windows+Excel runner.
- **One-time manual (Mac or Windows):** open Excel's VBA editor, File ▸ Import
  each `.bas`/`.cls` (incl. `vendor/*`, `modTasks.bas`, `modAgent.bas`), then
  Save As `.xlam`. Run `RunAllTests` to verify.

Once the `.xlam` exists at the repo root, the signing/notarization is CI-automated.

## Windows `.msi`

Not built yet. A WiX `.msi` (or PowerShell + `.reg`) would register the manifest
via a trusted add-in catalog, or drop the `.xlam` into `%APPDATA%\Microsoft\AddIns`.
