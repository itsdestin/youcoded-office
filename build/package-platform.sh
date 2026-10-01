#!/usr/bin/env bash
# Builds dist/youcoded-office-<version>-<platform>.tar.gz for darwin-x64, darwin-arm64 or win32-x64,
# from the Linux bundle build-linux.sh made (work/bundle) plus that platform's converter.
# Usage: build/package-platform.sh <darwin-x64|darwin-arm64|win32-x64>
# Needs: 7-Zip 22+ (7zz, or 7z) — the macOS disk images are APFS, which older 7-Zip cannot open.
#
# WHY reuse the Linux bundle's editors/: the editor UI is plain web files, the same on every
# platform, so it is built once (build-linux.sh, from source) and only converter/ differs.
# WHY the converter comes from euro-office-lite's own installers for the SAME pinned tag, not its
# moving `dependencies` release: those installers are what that tag shipped and tested with, and
# a release asset never changes under a pinned tag.
set -euo pipefail
ROOT="$(cd "$(dirname "$0")/.." && pwd)"
PLATFORM="${1:?usage: package-platform.sh <darwin-x64|darwin-arm64|win32-x64>}"
VERSION=$(node -p "require('$ROOT/PIN.json').version")
TAG=$(node -p "require('$ROOT/PIN.json').euroOfficeLite.tag")
REPO=$(node -p "require('$ROOT/PIN.json').euroOfficeLite.repo")
V="${TAG#v}"
WORK="$ROOT/work"; BASE="$WORK/bundle"; OUT="$WORK/bundle-$PLATFORM"; X="$WORK/x-$PLATFORM"
SEVENZIP="${SEVENZIP:-$(command -v 7zz || command -v 7z || true)}"
[ -n "$SEVENZIP" ] || { echo "7-Zip (7zz or 7z) is required" >&2; exit 1; }
[ -f "$BASE/manifest.json" ] || { echo "no Linux bundle at $BASE — run build/build-linux.sh first" >&2; exit 1; }

case "$PLATFORM" in
  darwin-x64)   ASSET="Euro-Office-Lite_${V}_x86_64.dmg" ;;
  darwin-arm64) ASSET="Euro-Office-Lite_${V}_aarch64.dmg" ;;
  win32-x64)    ASSET="Euro-Office-Lite_${V}_x64-setup.exe" ;;
  *) echo "unknown platform $PLATFORM" >&2; exit 1 ;;
esac
[ -f "$WORK/$ASSET" ] || curl -fL -o "$WORK/$ASSET" "https://github.com/$REPO/releases/download/$TAG/$ASSET"
rm -rf "$X" "$OUT"; mkdir -p "$X" "$OUT" "$ROOT/dist"
# WHY -y and a quiet log: the dmg holds an "Applications" symlink 7-Zip would otherwise ask about.
"$SEVENZIP" x -y -o"$X" "$WORK/$ASSET" > "$X.log"

# Everything but converter/ (and the manifest) is shared with the Linux bundle.
cp -r "$BASE/editors" "$BASE/templates" "$OUT/"
cp "$ROOT/LICENSE" "$ROOT/NOTICE" "$OUT/"

case "$PLATFORM" in
  darwin-*)
    APP=$(find "$X" -maxdepth 2 -name '*.app' -type d | head -n1)
    [ -n "$APP" ] || { echo "no .app inside $ASSET" >&2; exit 1; }
    # WHY Resources/binaries as one flat folder: x2t finds its dylibs through @rpath, and its
    # rpaths include @executable_path, so the libraries beside it are found wherever the folder
    # sits. Its DoctRenderer.config reads ../editors (the same layout as the Linux bundle).
    # The upstream Developer ID signatures do not carry over: YouCoded's own macOS build re-signs
    # x2t and every dylib (electron-builder signs Mach-O files under Contents/Resources).
    cp -r "$APP/Contents/Resources/binaries" "$OUT/converter"
    chmod +x "$OUT/converter/x2t"
    # WHY (as build-linux.sh, v0.1.12): x2t draws PDFs with editors/sdkjs/common/Native/native.js,
    # and the one built from source is too new for this x2t. Take the one this release's own
    # Mac app ships beside the same x2t.
    cp "$APP/Contents/Resources/editors/sdkjs/common/Native/native.js" "$OUT/editors/sdkjs/common/Native/native.js"
    ;;
  win32-x64)
    # WHY x2t.exe moves into binaries/: the installer keeps it one folder above its DLLs (its app
    # puts that folder on the search path). Windows looks for a program's DLLs in the program's
    # own folder first, so one flat folder needs no search-path setup from YouCoded. Its
    # DoctRenderer.config names files beside it, so it runs from any folder.
    cp -r "$X/binaries" "$OUT/converter"
    cp "$X/x2t.exe" "$OUT/converter/x2t.exe"
    ;;
esac

printf '{ "version": "%s", "euroOfficeLite": "%s", "platform": "%s" }\n' "$VERSION" "$TAG" "$PLATFORM" > "$OUT/manifest.json"
tar -C "$OUT" -czf "$ROOT/dist/youcoded-office-$VERSION-$PLATFORM.tar.gz" .
echo "built dist/youcoded-office-$VERSION-$PLATFORM.tar.gz"
