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
  await writeFile(path.join(dir, 'bridge.js'), "var ASC_PROTO_BASE = _isWindows ? 'http://ascdesktop.localhost/' : 'ascdesktop://';\n" +
    "            await invoke('save_file_as', { path: savePath });\n            if (pathExt !== 'pdf') {\n              invoke('set_window_title', {});\n            }\n");
  await writeFile(path.join(dir, 'editor-patches.js'),
    "          permissions: {\n            edit: true,\n            download: true,\n            print: true\n          }\n        },\n        editorConfig: {\n          mode: 'edit',\n          customization: {\n            about: false,\n            feedback: false\n          }\n        },\n" +
    "                  if (options && options.advancedOptions &&\n                      typeof options.advancedOptions.asc_getNativeOptions !== 'function') {\n                    options.advancedOptions = undefined;\n                  }\n");
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

test('the File tab has no "Suggest a feature" and no printing (the host cannot print)', async () => {
  const dir = await patchedCopy();
  const js = await readFile(path.join(dir, 'editor-patches.js'), 'utf8');
  assert.match(js, /suggestFeature: false,/);
  assert.match(js, /print: false/);
  assert.doesNotMatch(js, /print: true/);
  await rm(dir, { recursive: true, force: true });
});

// v0.1.12: YouCoded's Save As writes a copy and the document stays on its own file, so the editor
// must not retitle itself (or move its recovery) to the copy's name afterwards.
test('after a Save As the editor keeps its own name (the copy is a separate file)', async () => {
  const dir = await patchedCopy();
  const js = await readFile(path.join(dir, 'bridge.js'), 'utf8');
  assert.doesNotMatch(js, /\n\s*if \(pathExt !== 'pdf'\) \{/);
  assert.match(js, /if \(false && pathExt !== 'pdf'\) \{/);
  await rm(dir, { recursive: true, force: true });
});

// v0.1.14 (Task 2 fix round 1): the editor's export choices reach the host. A CSV's encoding and
// delimiter live in the dialog's text options, which editor-patches.js drops (sdkjs's desktop save
// path cannot take them); they are kept aside first. bridge.js then sends them, and its save
// options (a spreadsheet PDF's print range), with save_file_as.
test('Save As sends the editor\'s export choices with save_file_as', async () => {
  const dir = await patchedCopy();
  const bridge = await readFile(path.join(dir, 'bridge.js'), 'utf8');
  assert.match(bridge, /invoke\('save_file_as', \{ path: savePath, json: jsonOptions \|\| '', text: window\.__ycTextOptions \|\| null \}\)/);
  const patches = await readFile(path.join(dir, 'editor-patches.js'), 'utf8');
  assert.match(patches, /window\.__ycTextOptions = null;/);
  assert.match(patches, /asc_getCodePage/);
  const kept = patches.indexOf('window.__ycTextOptions = {');
  assert.ok(kept > 0 && kept < patches.indexOf('options.advancedOptions = undefined;'), 'kept aside before they are dropped');
  await rm(dir, { recursive: true, force: true });
});
