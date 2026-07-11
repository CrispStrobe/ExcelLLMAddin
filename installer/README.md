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

CI: `.github/workflows/installer-mac.yml` builds it on a `macos-latest` runner. It
produces an **unsigned** `.pkg` by default, and **auto-signs + notarizes** if a
Developer ID Installer cert secret is present (see Signing below).

## Signing status

Two levels, pick per audience:

- **Unsigned** (default, no cert needed) — installs fine for personal/offline use;
  Gatekeeper's first-open needs a right-click ▸ Open (or `xattr -dr
  com.apple.quarantine ExcelLLMAddin-Offline.pkg`). The build here and in CI
  produce this today.
- **Signed + notarized** (for public download) — needs a **Developer ID Installer**
  certificate. Note: the Mac App Store certs (`Apple Distribution`, `3rd Party Mac
  Developer Installer`) do **not** work for direct distribution. Create a Developer
  ID Installer cert in the Apple Developer portal (an Account-Holder decision;
  Developer ID slots are limited and account-wide), export it as a `.p12`, and add
  these repo secrets:

  | Secret | Value |
  |---|---|
  | `APPLE_INSTALLER_CERT_P12` | base64 of the `.p12` |
  | `APPLE_CERT_PASSWORD` | its password |
  | `APPLE_SIGN_IDENTITY` | e.g. `Developer ID Installer: Name (TEAMID)` |
  | `AC_KEY_ID`, `AC_ISSUER_ID`, `AC_KEY_P8` | App Store Connect API key (for notarization) — **already set** |

  Once the cert secret is present, the CI signs + notarizes automatically.

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
