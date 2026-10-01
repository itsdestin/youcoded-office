// Makes PowerPoint's standard slide themes for the bundle (v0.1.37).
// Usage: node build/gen-themes.mjs <bundle dir holding editors/ and converter/>
//
// WHY: the Design tab's theme gallery reads three things the editor build never makes —
//   editors/sdkjs/slide/themes/theme<N>/theme.bin (+ media/)  each theme, in the editor's own format
//   editors/sdkjs/slide/themes/themes.js                       their names (AscCommon.g_defaultThemes)
//   editors/sdkjs/common/Images/themes_thumbnail*.png          one picture strip per screen scale
// OnlyOffice makes them with `allthemesgen`; euro-office-lite ships only the source decks
// (themes/src/*.pptx), so the gallery was empty. This repeats allthemesgen's steps with the
// bundle's own x2t and libraries (build/themethumbs.cpp draws the pictures). The result is plain
// files, the same on every platform, so package-platform.sh carries it to Mac and Windows as is.
import { readdir, readFile, writeFile, mkdir, rm, mkdtemp } from 'node:fs/promises';
import { execFile } from 'node:child_process';
import { promisify } from 'node:util';
import os from 'node:os';
import path from 'node:path';

const run = promisify(execFile);
const bundle = process.argv[2];
if (!bundle) { console.error('usage: gen-themes.mjs <bundle dir>'); process.exit(2); }
const conv = path.resolve(bundle, 'converter');
const themes = path.resolve(bundle, 'editors', 'sdkjs', 'slide', 'themes');
const images = path.resolve(bundle, 'editors', 'sdkjs', 'common', 'Images');
const env = { ...process.env, LD_LIBRARY_PATH: conv };
const fail = (msg) => { console.error(`gen-themes: ${msg}`); process.exit(1); };

// The screen scales the presentation editor's stylesheet asks for (themes_thumbnail.png …
// themes_thumbnail@5x.png), and allthemesgen's file naming for each.
const SCALES = [1, 1.25, 1.5, 1.75, 2, 2.5, 3, 3.5, 4, 4.5, 5];
const suffix = (s) => (s === 1 ? '' : `@${s}x`);
const W = 88, H = 40;

const xmlEscape = (s) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');

// WHY sorted by file name: allthemesgen numbers themes in that order, and the gallery's picture
// strip, theme<N> folders and the names list must agree index for index.
const decks = (await readdir(path.join(themes, 'src'))).filter((n) => /\.(pptx|potx|pptm|potm)$/i.test(n)).sort();
if (!decks.length) fail(`no theme decks in ${path.join(themes, 'src')}`);

const tmp = await mkdtemp(path.join(os.tmpdir(), 'yco-themes-'));
const drawer = path.join(tmp, 'themethumbs');
await run('g++', ['-O1', '-o', drawer, path.join(import.meta.dirname, 'themethumbs.cpp'), `-L${conv}`, '-ldoctrenderer', '-lgraphics', '-lkernel',
  // WHY rpath-link: these libraries need the converter's other libraries (PDF, Unicode…) to link.
  `-Wl,-rpath-link,${conv}`])
  .catch((e) => fail(`could not build the picture drawer (needs g++): ${e.stderr || e.message}`));

// WHY a font list of real files (as the app's PDF export, desktop x2t.ts): the bundled
// converter/AllFonts.js names bare file names, and the pictures' sample text then drew as
// missing-glyph boxes. x2t writes AllFonts.js and font_selection.bin naming the bundled fonts' paths.
const fontData = path.join(tmp, 'fontdata');
await mkdir(fontData);
await run(path.join(conv, 'x2t'), ['-create-allfonts', fontData, path.join(conv, 'fonts')], { cwd: conv, env })
  .catch((e) => fail(`x2t could not list the fonts: ${e.message}`));

const names = [];
for (const [i, deck] of decks.entries()) {
  const out = path.join(themes, `theme${i + 1}`);
  await rm(out, { recursive: true, force: true });
  await mkdir(out, { recursive: true });
  // allthemesgen's x2t task, word for word: the deck in the editor's own format (8192),
  // without the extras a document save adds.
  const params = path.join(tmp, `params${i + 1}.xml`);
  await writeFile(params, '<?xml version="1.0" encoding="utf-8"?><TaskQueueDataConvert>' +
    `<m_sFileFrom>${xmlEscape(path.join(themes, 'src', deck))}</m_sFileFrom>` +
    `<m_sFileTo>${xmlEscape(path.join(out, 'theme.bin'))}</m_sFileTo><m_nFormatTo>8192</m_nFormatTo>` +
    '<m_sThemeDir>./</m_sThemeDir><m_bDontSaveAdditional>true</m_bDontSaveAdditional>' +
    `<m_sAllFontsPath>${xmlEscape(path.join(fontData, 'AllFonts.js'))}</m_sAllFontsPath></TaskQueueDataConvert>`);
  await run(path.join(conv, 'x2t'), [params], { cwd: conv, env }).catch((e) => fail(`x2t could not convert ${deck}: ${e.message}`));
  const sizes = SCALES.flatMap((s) => [String(Math.trunc(W * s)), String(Math.trunc(H * s)), path.join(tmp, `t${i + 1}${suffix(s)}.png`)]);
  const { stdout } = await run(drawer, [conv, fontData, out, ...sizes], { cwd: conv, env }).catch((e) => fail(`could not draw ${deck}: ${e.stderr || e.message}`));
  const name = stdout.trim().split('\n').pop();
  if (!name) fail(`no theme name from ${deck}`);
  names.push(name);
}

// WHY a plain assignment: allthemesgen writes exactly this; patch.mjs loads it right after the
// presentation editor's sdk-all-min.js (which defines AscCommon), before the editor starts.
await writeFile(path.join(themes, 'themes.js'), `AscCommon.g_defaultThemes = ${JSON.stringify(names)};`);

// One strip per scale: the themes' pictures stacked top to bottom, in theme order (the gallery
// shows theme N at offset N×40px of the 88px-wide strip). WHY Pillow: the build already needs it
// (euro-office-lite's own thumbnail sprites).
const stack = `
import sys
from PIL import Image
out, parts = sys.argv[1], sys.argv[2:]
ims = [Image.open(p).convert('RGBA') for p in parts]
w, h = ims[0].size
strip = Image.new('RGBA', (w, h * len(ims)))
for i, im in enumerate(ims):
    if im.size != (w, h): sys.exit('picture %s is %s, not %s' % (parts[i], im.size, (w, h)))
    strip.paste(im, (0, i * h))
strip.save(out, optimize=True)
`;
for (const s of SCALES) {
  const parts = decks.map((_, i) => path.join(tmp, `t${i + 1}${suffix(s)}.png`));
  await run('python3', ['-c', stack, path.join(images, `themes_thumbnail${suffix(s)}.png`), ...parts]).catch((e) => fail(`could not stack the ${s}x pictures: ${e.stderr || e.message}`));
}
await rm(tmp, { recursive: true, force: true });
console.log(`gen-themes: ${names.length} themes (${names.join(', ')})`);
