import { test } from 'node:test';
import assert from 'node:assert/strict';
import { mkdtemp, mkdir, writeFile, readFile, rm } from 'node:fs/promises';
import { execFile } from 'node:child_process';
import { promisify } from 'node:util';
import os from 'node:os';
import path from 'node:path';

const run = promisify(execFile);

// build/patch.mjs against a tiny stand-in for euro-office-lite's built folder — the same text
// the real files carry at the pinned tag — so its changes are checked without a full build.
async function patchedCopy() {
  const dir = await mkdtemp(path.join(os.tmpdir(), 'yco-patch-'));
  await writeFile(path.join(dir, 'index.html'), '<html><head></head></html>');
  await writeFile(path.join(dir, 'bridge.js'), "var ASC_PROTO_BASE = _isWindows ? 'http://ascdesktop.localhost/' : 'ascdesktop://';\n");
  await writeFile(path.join(dir, 'editor-patches.js'),
    "        editorConfig: {\n          mode: 'edit',\n          customization: {\n            about: false,\n            feedback: false\n          }\n        },\n");
  const main = path.join(dir, 'web-apps', 'apps', 'documenteditor', 'main');
  await mkdir(main, { recursive: true });
  await writeFile(path.join(main, 'index.html'), '<html><head></head></html>');
  await run(process.execPath, [path.resolve(import.meta.dirname, '..', 'build', 'patch.mjs'), dir]);
  return dir;
}

test('the editor opens every document with macros turned off, so a document cannot run its own scripts', async () => {
  const dir = await patchedCopy();
  const js = await readFile(path.join(dir, 'editor-patches.js'), 'utf8');
  const custom = js.slice(js.indexOf('customization: {'), js.indexOf('}', js.indexOf('customization: {')));
  assert.match(custom, /macros: false,/);
  assert.match(custom, /macrosMode: 'disable',/);
  // Euro-Office's own settings are kept beside the new ones.
  assert.match(custom, /about: false,/);
  await rm(dir, { recursive: true, force: true });
});
