#!/bin/bash
# Build (and optionally sign + notarize) a macOS installer that registers the
# Excel LLM Add-in manifest into Excel's sideload folder.
#
# Unsigned (local test):
#   ./build-pkg.sh
# Signed + notarized (CI / release), using an App Store Connect API key:
#   SIGN_IDENTITY="Developer ID Installer: Name (TEAMID)" \
#   AC_KEY_ID=XXXX AC_ISSUER_ID=xxxx-... AC_KEY_PATH=/path/AuthKey_XXXX.p8 \
#   ./build-pkg.sh
#
# The manifest is taken from officejs/dist/manifest.xml (run `npm run build`
# first, or this script builds it). That manifest points at the production
# hosted origin, so the installed add-in loads from there.
set -euo pipefail

HERE="$(cd "$(dirname "$0")" && pwd)"
REPO="$(cd "$HERE/../.." && pwd)"
VERSION="${VERSION:-1.0.0}"
IDENTIFIER="com.crispstrobe.excelllmaddin.installer"
OUT="$HERE/ExcelLLMAddin.pkg"
COMPONENT="$HERE/ExcelLLMAddin-component.pkg"

# 1. Ensure a production manifest exists.
MANIFEST="$REPO/officejs/dist/manifest.xml"
if [ ! -f "$MANIFEST" ]; then
  echo "Building officejs to produce dist/manifest.xml..."
  (cd "$REPO/officejs" && npm run build >/dev/null)
fi

# 2. Stage the payload.
STAGE="$(mktemp -d)"
mkdir -p "$STAGE/payload"
cp "$MANIFEST" "$STAGE/payload/manifest.xml"

# 3. Build the component package (payload + postinstall script).
chmod +x "$HERE/scripts/postinstall"
pkgbuild \
  --root "$STAGE/payload" \
  --install-location "/usr/local/share/excel-llm-addin" \
  --scripts "$HERE/scripts" \
  --identifier "$IDENTIFIER" \
  --version "$VERSION" \
  "$COMPONENT"

# 4. Distribution package (+ sign if an identity is given).
if [ -n "${SIGN_IDENTITY:-}" ]; then
  productbuild --package "$COMPONENT" --sign "$SIGN_IDENTITY" "$OUT"
else
  echo "NOTE: SIGN_IDENTITY not set — producing an UNSIGNED .pkg (Gatekeeper will warn)."
  productbuild --package "$COMPONENT" "$OUT"
fi
rm -f "$COMPONENT"
rm -rf "$STAGE"

# 5. Notarize + staple if App Store Connect API creds are provided.
if [ -n "${SIGN_IDENTITY:-}" ] && [ -n "${AC_KEY_ID:-}" ] && [ -n "${AC_ISSUER_ID:-}" ] && [ -n "${AC_KEY_PATH:-}" ]; then
  echo "Notarizing $OUT ..."
  xcrun notarytool submit "$OUT" \
    --key "$AC_KEY_PATH" --key-id "$AC_KEY_ID" --issuer "$AC_ISSUER_ID" --wait
  xcrun stapler staple "$OUT"
  echo "Notarized + stapled."
else
  echo "NOTE: notarization creds (AC_KEY_ID/AC_ISSUER_ID/AC_KEY_PATH) not set — skipping notarization."
fi

echo "Built: $OUT"
