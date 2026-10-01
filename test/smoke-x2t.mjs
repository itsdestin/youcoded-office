// Runs a built bundle's converter the way YouCoded does, on whatever OS this runs on.
// Usage: node test/smoke-x2t.mjs <extracted bundle folder>
//
// WHY a script and not part of bundle.test.mjs: the Mac and Windows bundles can only be run on a
// Mac or Windows machine, so CI unpacks each one on its own runner and runs this there. It
// checks what the app relies on: a document opens (docx/xlsx/pptx → the editor's form), saves
// back, and prints to PDF with real text drawn (with a font list x2t makes itself, as the app's
// pdfFontData does). Exits non-zero, naming the step, on the first failure.
import { execFile } from 'node:child_process';
import { mkdir, mkdtemp, readFile, rm, stat, writeFile } from 'node:fs/promises';
import { inflateSync } from 'node:zlib';
import { promisify } from 'node:util';
import os from 'node:os';
import path from 'node:path';

const run = promisify(execFile);
const B = path.resolve(process.argv[2] ?? path.join(import.meta.dirname, '..', 'work', 'bundle'));
const conv = path.join(B, 'converter');
const exe = path.join(conv, process.platform === 'win32' ? 'x2t.exe' : 'x2t');
const env = { ...process.env };
if (process.platform === 'linux') env.LD_LIBRARY_PATH = conv;
if (process.platform === 'darwin') env.DYLD_LIBRARY_PATH = conv;
const FORMAT = { bin: 8192, docx: 65, xlsx: 257, pptx: 129, pdf: 513 };

const tmp = await mkdtemp(path.join(os.tmpdir(), 'yco-smoke-'));
const x2t = async (args) => {
  try {
    await run(exe, args, { cwd: conv, env, timeout: 120000, maxBuffer: 64 * 1024 * 1024 });
  } catch (e) {
    throw new Error(`x2t ${args.join(' ')} failed: ${e.code ?? e.signal} ${e.stderr ?? ''}`);
  }
};
let n = 0;
const job = async (from, to, fmt, allFonts = path.join(conv, 'AllFonts.js')) => {
  const jt = await mkdtemp(path.join(tmp, 'job-'));
  const xml = `<?xml version="1.0" encoding="utf-8"?><TaskQueueDataConvert><m_sFileFrom>${from}</m_sFileFrom><m_sFileTo>${to}</m_sFileTo><m_nFormatTo>${fmt}</m_nFormatTo><m_sTempDir>${jt}</m_sTempDir><m_sFontDir>${path.join(conv, 'fonts')}</m_sFontDir><m_sAllFontsPath>${allFonts}</m_sAllFontsPath></TaskQueueDataConvert>`;
  const p = path.join(tmp, `params-${n++}.xml`);
  await writeFile(p, xml);
  await x2t([p]);
  const st = await stat(to).catch(() => null);
  if (!st || st.size === 0) throw new Error(`${path.basename(from)} → ${path.basename(to)}: no output`);
  return st.size;
};
// Each document's editor form goes in its own folder: x2t writes pictures beside it (media/).
const into = async (name) => { const d = path.join(tmp, name); await mkdir(d); return d; };

const memo = path.join(import.meta.dirname, 'fixtures', 'memo.docx');
const docs = [
  ['docx', memo],
  ['xlsx', path.join(B, 'templates', 'blank.xlsx')],
  ['pptx', path.join(B, 'templates', 'blank.pptx')],
];

const fonts = await into('fontdata');
await x2t(['-create-allfonts', fonts, path.join(conv, 'fonts')]);
console.log('ok   font list made');

for (const [kind, file] of docs) {
  const d = await into(kind);
  const bin = path.join(d, 'Editor.bin');
  await job(file, bin, FORMAT.bin);
  const back = await job(bin, path.join(d, `back.${kind}`), FORMAT[kind]);
  console.log(`ok   ${kind} opens and saves back (${back} bytes)`);
  const pdf = path.join(d, 'out.pdf');
  await job(bin, pdf, FORMAT.pdf, path.join(fonts, 'AllFonts.js'));
  const bytes = await readFile(pdf);
  if (bytes.subarray(0, 5).toString('latin1') !== '%PDF-') throw new Error(`${kind} → pdf: not a PDF`);
  if (kind === 'docx') {
    // A PDF whose every character is glyph 0 is a blank page: the memo must draw real glyphs.
    let real = 0;
    for (const m of bytes.toString('latin1').matchAll(/stream\r?\n([\s\S]*?)\r?\nendstream/g)) {
      let t; try { t = inflateSync(Buffer.from(m[1], 'latin1')).toString('latin1'); } catch { continue; }
      for (const g of t.matchAll(/<([0-9A-Fa-f]{4})>/g)) if (g[1] !== '0000') real++;
    }
    if (real === 0) throw new Error('docx → pdf: no text drawn');
  }
  console.log(`ok   ${kind} prints to PDF (${bytes.length} bytes)`);
}
await rm(tmp, { recursive: true, force: true });
console.log(`smoke passed: ${B}`);
