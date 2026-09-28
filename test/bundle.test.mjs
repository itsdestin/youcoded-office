import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFile, stat, mkdtemp, writeFile, rm } from 'node:fs/promises';
import { execFile } from 'node:child_process';
import { promisify } from 'node:util';
import os from 'node:os';
import path from 'node:path';

const B = path.resolve(import.meta.dirname, '..', 'work', 'bundle');
const run = promisify(execFile);

test('index.html loads the relay, then the theme bridge, then bridge.js', async () => {
  const html = await readFile(path.join(B, 'editors', 'index.html'), 'utf8');
  const relay = html.indexOf('tauri-relay.js');
  const yc = html.indexOf('yc-bridge.js');
  // WHY a regex, not indexOf('bridge.js"'): 'yc-bridge.js"' contains the literal substring
  // 'bridge.js"', so a plain indexOf falsely matches inside the yc-bridge tag and never checks
  // the real <script src="bridge.js"> tag at all. Require the character before "bridge" to not
  // be a word char, quote or hyphen, so the "yc-" prefix can't satisfy the match.
  const bridgeMatch = /(?<![\w'"-])src="bridge\.js"/.exec(html);
  assert.ok(relay > 0 && yc > relay, 'relay then yc-bridge');
  assert.ok(bridgeMatch, 'bridge.js script tag present');
  assert.ok(bridgeMatch.index > yc, 'bridge.js after yc-bridge');
});

test('bridge.js serves media from the document origin', async () => {
  const js = await readFile(path.join(B, 'editors', 'bridge.js'), 'utf8');
  assert.match(js, /var ASC_PROTO_BASE = location\.origin \+ '\/asc\/';/);
});

test('bundle carries licence, notice, manifest and templates', async () => {
  for (const f of ['LICENSE', 'NOTICE', 'manifest.json', 'templates/blank.docx', 'templates/blank.xlsx', 'templates/blank.pptx', 'converter/x2t', 'converter/AllFonts.js'])
    await stat(path.join(B, f));
});

test('x2t round-trips a docx through Editor.bin', async () => {
  const tmp = await mkdtemp(path.join(os.tmpdir(), 'yco-'));
  const conv = path.join(B, 'converter');
  const job = async (from, to, fmt) => {
    const jt = await mkdtemp(path.join(tmp, 'job-'));
    const xml = `<?xml version="1.0" encoding="utf-8"?><TaskQueueDataConvert><m_sFileFrom>${from}</m_sFileFrom><m_sFileTo>${to}</m_sFileTo><m_nFormatTo>${fmt}</m_nFormatTo><m_sTempDir>${jt}</m_sTempDir><m_sFontDir>${conv}/fonts</m_sFontDir><m_sAllFontsPath>${conv}/AllFonts.js</m_sAllFontsPath></TaskQueueDataConvert>`;
    const p = path.join(tmp, `p-${fmt}.xml`); await writeFile(p, xml);
    await run(path.join(conv, 'x2t'), [p], { cwd: conv, env: { ...process.env, LD_LIBRARY_PATH: conv }, timeout: 60000 });
  };
  const bin = path.join(tmp, 'Editor.bin'), back = path.join(tmp, 'back.docx');
  await job(path.join(import.meta.dirname, 'fixtures', 'memo.docx'), bin, 8192);
  await job(bin, back, 65);
  assert.ok((await stat(back)).size > 1000);
  await rm(tmp, { recursive: true, force: true });
});
