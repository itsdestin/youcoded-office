#!/usr/bin/env bash
# Builds dist/youcoded-office-<version>-linux-x64.tar.gz. Needs: node 22 on PATH, git, gh or curl, ar, tar.
# WHY Node 22: euro-office-lite's grunt chain calls util.isRegExp, removed in Node 23+.
set -euo pipefail
ROOT="$(cd "$(dirname "$0")/.." && pwd)"
VERSION=$(node -p "require('$ROOT/PIN.json').version")
TAG=$(node -p "require('$ROOT/PIN.json').euroOfficeLite.tag")
REPO=$(node -p "require('$ROOT/PIN.json').euroOfficeLite.repo")
WORK="$ROOT/work"; OUT="$WORK/bundle"; mkdir -p "$WORK" "$ROOT/dist"; rm -rf "$OUT"; mkdir -p "$OUT"

# 1. Editor UI from source, at the pinned tag.
[ -d "$WORK/eol" ] || git clone --depth 1 --branch "$TAG" --recurse-submodules --shallow-submodules "https://github.com/$REPO.git" "$WORK/eol"
cd "$WORK/eol"
# WHY --ignore-scripts then rebuild: npm 11+ blocks install scripts by default, and imagemin's
# binaries only arrive through them. Only gifsicle and optipng-bin are actually pulled in by
# this tag's grunt-contrib-imagemin (via imagemin-gifsicle/imagemin-optipng in
# work/eol/package-lock.json); mozjpeg and pngquant-bin aren't in the tree, so rebuilding them
# would silently no-op. No `|| true`: a real rebuild failure here should stop the build, not
# produce a bundle with a broken imagemin binary.
npm ci --ignore-scripts
npm rebuild gifsicle optipng-bin
# WHY: the sdkjs submodule's grunt build (compile-word/cell/slide via Closure) is its own
# npm package with its own node_modules, not hoisted by the root install; the build script
# fails loudly ("grunt is missing") without this.
(cd src/sdkjs/build && npm ci)
# WHY: the frontend build needs src/sdkjs/common/AllFonts.js and src/fonts/*.ttf, which this
# repo does not commit (font redistribution belongs to prepare-fonts.sh, which fetches
# Liberation + Carlito on demand); its thumbnail-sprite step imports Pillow, which a bare
# Python 3 does not carry.
python3 -c 'import PIL' 2>/dev/null || python3 -m pip install --quiet --disable-pip-version-check --user Pillow \
  || python3 -m pip install --quiet --disable-pip-version-check --user --break-system-packages Pillow
bash scripts/prepare-fonts.sh
node scripts/build-frontend-prod.mjs
cp -r src-dist "$OUT/editors"

# 2. Translator + templates from the same release's .deb.
DEB="$WORK/eol.deb"
V="${TAG#v}"
[ -f "$DEB" ] || curl -fL -o "$DEB" "https://github.com/$REPO/releases/download/$TAG/Euro-Office-Lite_${V}_amd64.deb"
mkdir -p "$WORK/deb" && cd "$WORK/deb" && ar x "$DEB" && tar xf data.tar.*
cp -r "$WORK/deb/usr/lib/Euro-Office-Lite/binaries" "$OUT/converter"
cp -r "$WORK/deb/usr/lib/Euro-Office-Lite/templates" "$OUT/templates"

# 3. YouCoded's patches, licence, notices, manifest.
node "$ROOT/build/patch.mjs" "$OUT/editors"
cp "$ROOT/LICENSE" "$ROOT/NOTICE" "$OUT/"
printf '{ "version": "%s", "euroOfficeLite": "%s", "platform": "linux-x64" }\n' "$VERSION" "$TAG" > "$OUT/manifest.json"
tar -C "$OUT" -czf "$ROOT/dist/youcoded-office-$VERSION-linux-x64.tar.gz" .
cd "$ROOT/dist" && sha256sum youcoded-office-*.tar.gz > SHA256SUMS
echo "built dist/youcoded-office-$VERSION-linux-x64.tar.gz"
