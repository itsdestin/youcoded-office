# youcoded-office

YouCoded's Office add-on: a separate AGPL-3.0 program that runs inside the
YouCoded app to open and edit Word/Excel/PowerPoint-compatible documents.
It bundles [euro-office-lite](https://github.com/delmarguillen/euro-office-lite)
(a fork of OnlyOffice's editors, via Euro-Office) plus two small YouCoded-side
bridge scripts, baked into the editor's `index.html`:

- `bridge/tauri-relay.js` — a fake `window.__TAURI__` so euro-office-lite's
  own `bridge.js` (written for a Tauri desktop shell) can run unmodified
  inside YouCoded's iframe, relaying its calls to the host app via
  `postMessage`.
- `bridge/yc-bridge.js` — YouCoded's theme bridge: applies YouCoded's colour
  theme to the editor frames, hides the editor's own title row, and runs
  the host's toolbar commands and autosave inside the editor.

Why a separate repo, and why AGPL: euro-office-lite (and the OnlyOffice code
it descends from) is AGPL-3.0. Keeping it as its own program with its own
repo keeps that licence obligation contained to this add-on, instead of
spreading it across YouCoded's main app.

## What a release contains

Each GitHub release publishes `youcoded-office-<version>-linux-x64.tar.gz`
and `SHA256SUMS`. The tarball, unpacked, is:

```
manifest.json          # { version, euroOfficeLite: "<tag>", platform: "linux-x64" }
editors/                # the editor UI, built from euro-office-lite source
  index.html
  bridge.js
  tauri-relay.js         # YouCoded's relay (this repo's bridge/tauri-relay.js)
  yc-bridge.js           # YouCoded's theme bridge (this repo's bridge/yc-bridge.js)
  web-apps/  sdkjs/  fonts/  dictionaries/  ...
converter/              # x2t and its shared libraries, from the release .deb
  x2t  *.so  AllFonts.js  fonts/  ...
templates/
  blank.docx  blank.xlsx  blank.pptx
LICENSE
NOTICE
```

## Export and PDF

Save As, Export and Export to PDF write a separate file through the host (YouCoded), which runs
the bundled `x2t`. Two things to know:

- **PDFs use the computer's installed fonts.** A PDF is drawn from real font files, and the host
  has `x2t -create-allfonts` list them once per run — the bundled fonts plus every font installed
  on the computer. A PDF can therefore look slightly different from the editor (which shows only
  the bundled and theme fonts) when a document names a font installed locally.
- **Which export choices are honoured.** A CSV's encoding and delimiter, and a spreadsheet PDF's
  page settings and print range, reach `x2t`. A TXT's encoding does not (`x2t` always writes
  UTF-8), so Word's TXT encoding dialog is answered for the person and never shown.

## Print

Print (File → Print, Ctrl+P) goes through the host too: `bridge.js` hands over the document and
the print panel's choices (a document's or presentation's pages; a workbook's sheets, pages and
print area), the host makes a PDF with `x2t` in its own temp folder and shows the operating
system's print dialog for it. That dialog chooses the printer, copies, two-sided and colour, so
those rows leave the editor's print panel; printing only the selection is not offered (`x2t`
prints from the saved document, which has no selection), nor is quick print. Page breaks follow
`x2t`'s layout, which can differ slightly from the editor's preview (see the fonts note above).

## Building

Linux x64 only, for now:

```bash
bash build/build-linux.sh
```

Needs Node 22 first on `PATH` (euro-office-lite's grunt build chain calls
`util.isRegExp`, which Node 23+ removed), plus `git`, `curl`, `ar` and `tar`.
The script clones euro-office-lite at the tag pinned in `PIN.json`, builds
its frontend from source, pulls `x2t` and the blank templates out of that
same release's `.deb`, applies this repo's two patches
(`build/patch.mjs`), and writes `dist/youcoded-office-<version>-linux-x64.tar.gz`
plus `dist/SHA256SUMS`.

Run the tests against the built bundle:

```bash
node --test test/bundle.test.mjs
```

## Corresponding source

This repo *is* the corresponding source for the AGPL obligation on the
patches it applies. The editor UI and `x2t` converter's own corresponding
source is euro-office-lite's, and OnlyOffice's beneath that — both at the
exact tag this bundle pins, recorded in `PIN.json`
(`euroOfficeLite.tag`, currently `v0.17.21-alpha`) and in every release's
`manifest.json`:

- This repo, at the tag matching the release version.
- https://github.com/delmarguillen/euro-office-lite, at the pinned tag.

See `NOTICE` for the full licence chain (Euro-Office, OnlyOffice,
euro-office-lite, and the bundled fonts).
