import { test as nodeTest } from 'node:test';
import assert from 'node:assert/strict';
import { readFile, readdir, stat, mkdtemp, writeFile, rm } from 'node:fs/promises';
import { execFile } from 'node:child_process';
import { promisify } from 'node:util';
import os from 'node:os';
import path from 'node:path';

const B = path.resolve(import.meta.dirname, '..', 'work', 'bundle');
const run = promisify(execFile);

// WHY a freshness gate (v0.1.10): these tests read work/bundle, which only build/build-linux.sh
// makes. A local clone keeps whatever bundle it last built, so after any change to bridge/ or
// build/ they failed on old files and hid real results. Locally they now skip, saying why, when
// the bundle is missing or older than those sources; CI builds the bundle first and sets CI, and
// there they always run.
async function staleReason() {
  if (process.env.CI) return null;
  const built = await stat(path.join(B, 'manifest.json')).catch(() => null);
  if (!built) return 'no local bundle — run build/build-linux.sh (CI always runs these)';
  const root = path.resolve(import.meta.dirname, '..');
  let newest = (await stat(path.join(root, 'PIN.json'))).mtimeMs;
  for (const dir of ['bridge', 'build']) {
    for (const f of await readdir(path.join(root, dir))) newest = Math.max(newest, (await stat(path.join(root, dir, f))).mtimeMs);
  }
  return newest > built.mtimeMs ? 'local bundle is older than bridge/, build/ or PIN.json — run build/build-linux.sh (CI always runs these)' : null;
}
const skip = await staleReason();
const test = (name, fn) => nodeTest(name, skip ? { skip } : {}, fn);

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

test('every editor page runs yc-early.js before any of its own scripts', async () => {
  const { readdir } = await import('node:fs/promises');
  const apps = path.join(B, 'editors', 'web-apps', 'apps');
  let pages = 0;
  for (const app of await readdir(apps)) {
    const main = path.join(apps, app, 'main');
    for (const f of (await readdir(main).catch(() => [])).filter((n) => /^index.*\.html$/.test(n))) {
      const html = await readFile(path.join(main, f), 'utf8');
      const early = html.indexOf('<script src="../../../../yc-early.js"></script>');
      assert.ok(early > 0 && early === html.indexOf('<script'), `${app}/${f}: yc-early.js is the first script`);
      pages++;
    }
  }
  assert.ok(pages >= 3, 'the document, spreadsheet and presentation pages');
  await stat(path.join(B, 'editors', 'yc-early.js'));
});

test('bridge.js serves media from the document origin', async () => {
  const js = await readFile(path.join(B, 'editors', 'bridge.js'), 'utf8');
  assert.match(js, /var ASC_PROTO_BASE = location\.origin \+ '\/asc\/';/);
});

test('the built editor turns macros off', async () => {
  const js = await readFile(path.join(B, 'editors', 'editor-patches.js'), 'utf8');
  assert.match(js, /customization: \{\s*macros: false,\s*macrosMode: 'disable',/);
});

// v0.1.7: no "Suggest a feature" and no printing (the host refuses print_document).
test('the built editor has no Suggest a feature and no printing', async () => {
  const js = await readFile(path.join(B, 'editors', 'editor-patches.js'), 'utf8');
  assert.match(js, /macrosMode: 'disable',\s*suggestFeature: false,/);
  assert.match(js, /print: false/);
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
