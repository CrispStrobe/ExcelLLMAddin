#!/bin/bash
# Build (and optionally sign + notarize) a macOS installer that drops the
# self-contained, offline ExcelLLMAddin.xlam into Excel-for-Mac's add-ins folder.
#
# Prereq: a built XLAM at repo root (ExcelLLMAddin.xlam). Building the .xlam from
# the .bas requires Excel (Windows via tools/Build-Addin.ps1, or a manual Mac
# import) -- it CANNOT be produced on Mac or GitHub-hosted CI. This script only
# wraps + signs the already-built .xlam.
#
# Unsigned (local test):   ./build-xlam-pkg.sh
# Signed + notarized:      SIGN_IDENTITY="Developer ID Installer: Name (TEAM)" \
#                          AC_KEY_ID=.. AC_ISSUER_ID=.. AC_KEY_PATH=..p8 ./build-xlam-pkg.sh
set -euo pipefail

HERE="$(cd "$(dirname "$0")" && pwd)"
REPO="$(cd "$HERE/../.." && pwd)"
VERSION="${VERSION:-1.0.0}"
IDENTIFIER="com.crispstrobe.excelllmaddin.xlam"
OUT="$HERE/ExcelLLMAddin-Offline.pkg"
COMPONENT="$HERE/component.pkg"

XLAM="${XLAM:-$REPO/ExcelLLMAddin.xlam}"
if [ ! -f "$XLAM" ]; then
  echo "ERROR: $XLAM not found. Build the .xlam first (Windows: tools/Build-Addin.ps1,"
  echo "       or a one-time manual import in Excel for Mac's VBA editor)."
  exit 1
fi

STAGE="$(mktemp -d)"
mkdir -p "$STAGE/payload"
cp "$XLAM" "$STAGE/payload/ExcelLLMAddin.xlam"

chmod +x "$HERE/scripts/postinstall"
pkgbuild \
  --root "$STAGE/payload" \
  --install-location "/usr/local/share/excel-llm-addin" \
  --scripts "$HERE/scripts" \
  --identifier "$IDENTIFIER" \
  --version "$VERSION" \
  "$COMPONENT"

if [ -n "${SIGN_IDENTITY:-}" ]; then
  productbuild --package "$COMPONENT" --sign "$SIGN_IDENTITY" "$OUT"
else
  echo "NOTE: SIGN_IDENTITY not set -> UNSIGNED .pkg (Gatekeeper will warn)."
  productbuild --package "$COMPONENT" "$OUT"
fi
rm -f "$COMPONENT"
rm -rf "$STAGE"

if [ -n "${SIGN_IDENTITY:-}" ] && [ -n "${AC_KEY_ID:-}" ] && [ -n "${AC_ISSUER_ID:-}" ] && [ -n "${AC_KEY_PATH:-}" ]; then
  echo "Notarizing $OUT ..."
  xcrun notarytool submit "$OUT" --key "$AC_KEY_PATH" --key-id "$AC_KEY_ID" --issuer "$AC_ISSUER_ID" --wait
  xcrun stapler staple "$OUT"
else
  echo "NOTE: not signed or no notary creds -> skipping notarization (unsigned .pkg is fine for personal/offline use; right-click > Open)."
fi

echo "Built: $OUT"
