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
    "            await invoke('save_file_as', { path: savePath });\n            if (pathExt !== 'pdf') {\n              invoke('set_window_title', {});\n            }\n" +
    "    try {\n      window.AscDesktopEditor._isPrinting = true;\n" +
    "      await invoke('write_editor_bin', { data: b64 });\n" +
    "      var pdfPath = await invoke('print_document');\n\n" +
    "      if (printerName) {\n        var printResult = await invoke('plugin:printer|print_pdf', {\n          id: printerName,\n          path: pdfPath,\n          printer: printerName,\n          print_settings: '{}',\n          remove_after_print: true\n        });\n" +
    "      } else {\n        await invoke('open_pdf_viewer', { path: pdfPath });\n      }\n\n" +
    "      if (ref.ew && ref.ew.DesktopOfflineAppDocumentEndSave) {\n        ref.ew.DesktopOfflineAppDocumentEndSave(0);\n      }\n" +
    "    } catch(e) {\n      window._eoLog('[EO] Print: ERROR: ' + (e.message || e));\n" +
    "      if (ref.ew && ref.ew.DesktopOfflineAppDocumentEndSave) {\n        ref.ew.DesktopOfflineAppDocumentEndSave(1);\n      }\n" +
    "    } finally {\n      window.AscDesktopEditor._isPrinting = false;\n    }\n" +
    "  try {\n    var b64data = await invoke('open_file', { path: filePath });\n    var fileName = filePath.replace(/\\\\/g, '/').split('/').pop();\n" +
    "    window._pendingFileData = { data: b64data, path: filePath, name: fileName };\n    if (window._openEditor) {\n");
  await writeFile(path.join(dir, 'editor-patches.js'),
    "          permissions: {\n            edit: true,\n            download: true,\n            print: true\n          }\n        },\n        editorConfig: {\n          mode: 'edit',\n          user: {\n            id: 'local-user',\n            name: _t('user')\n          },\n          customization: {\n            about: false,\n            feedback: false\n          }\n        },\n" +
    "                  if (options && options.advancedOptions &&\n                      typeof options.advancedOptions.asc_getNativeOptions !== 'function') {\n                    options.advancedOptions = undefined;\n                  }\n");
  const main = path.join(dir, 'web-apps', 'apps', 'documenteditor', 'main');
  await mkdir(main, { recursive: true });
  await writeFile(path.join(main, 'index.html'), '<html><head></head></html>');
  // The presentation page as the pinned tag builds it: its SDK loader, then the themes.js guard.
  const slides = path.join(dir, 'web-apps', 'apps', 'presentationeditor', 'main');
  await mkdir(slides, { recursive: true });
  await writeFile(path.join(slides, 'index.html'), '<html><head></head><body>' +
    '<script src="../../../../sdkjs/slide/sdk-all-min.js"></script><script>\n(function() { /* themes.js guard */ })();\n</script></body></html>');
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

test('the File tab has no "Suggest a feature", and printing stays on (the host prints)', async () => {
  const dir = await patchedCopy();
  const js = await readFile(path.join(dir, 'editor-patches.js'), 'utf8');
  assert.match(js, /suggestFeature: false,/);
  assert.match(js, /print: true/);
  assert.doesNotMatch(js, /print: false/);
  await rm(dir, { recursive: true, force: true });
});

// v0.1.18 (Print): the host makes the PDF and shows the system print dialog; bridge.js only hands
// over the document and the print panel's choices.
test('Print hands the host the document and the panel\'s choices, and nothing else', async () => {
  const dir = await patchedCopy();
  const js = await readFile(path.join(dir, 'bridge.js'), 'utf8');
  assert.match(js, /invoke\('print_document', \{ json: window\.__ycPrintJson\(ref\.ew, optionsJson\) \}\)/);
  assert.doesNotMatch(js, /plugin:printer\|print_pdf|open_pdf_viewer|DesktopOfflineAppDocumentEndSave/);
  // Saves work again before the host is asked (the dialog can stay open for minutes).
  assert.ok(js.indexOf('_isPrinting = false;\n      await invoke(\'print_document\'') > 0);
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
  // Sent once and cleared in the same step (fix round 2), run here against a stand-in window.
  const m = /invoke\('save_file_as', (\(function \(\) \{[\s\S]*?\}\)\(\))\)/.exec(bridge);
  assert.ok(m, 'save_file_as sends the choices');
  const window = { __ycTextOptions: { codePage: 44, delimiter: [2] } };
  const args = new Function('window', 'savePath', 'jsonOptions', 'return ' + m[1])(window, 'yc-save/x/a.csv', '{"a":1}');
  assert.deepEqual(args, { path: 'yc-save/x/a.csv', json: '{"a":1}', text: { codePage: 44, delimiter: [2] } });
  assert.equal(window.__ycTextOptions, null, 'a later Save As does not reuse them');
  const again = new Function('window', 'savePath', 'jsonOptions', 'return ' + m[1])(window, 'yc-save/x/b.csv', undefined);
  assert.deepEqual(again, { path: 'yc-save/x/b.csv', json: '', text: null });
  const patches = await readFile(path.join(dir, 'editor-patches.js'), 'utf8');
  assert.match(patches, /window\.__ycTextOptions = null;/);
  assert.match(patches, /asc_getCodePage/);
  const kept = patches.indexOf('window.__ycTextOptions = {');
  assert.ok(kept > 0 && kept < patches.indexOf('options.advancedOptions = undefined;'), 'kept aside before they are dropped');
  await rm(dir, { recursive: true, force: true });
});

// v0.1.21 (finish plan Task 6): the person's own comments are "You", the name YouCoded writes for
// them into Word and Excel files, and the comments bridge loads after the theme bridge.
test('the person is "You" in the editor, and the comments bridge loads after the theme bridge', async () => {
  const dir = await patchedCopy();
  const js = await readFile(path.join(dir, 'editor-patches.js'), 'utf8');
  assert.match(js, /name: 'You'/);
  assert.doesNotMatch(js, /_t\('user'\)/);
  const html = await readFile(path.join(dir, 'index.html'), 'utf8');
  assert.ok(html.indexOf('yc-bridge.js') < html.indexOf('yc-comments.js'));
  await readFile(path.join(dir, 'yc-comments.js'), 'utf8');
  await rm(dir, { recursive: true, force: true });
});

// v0.1.24 (finish plan Task 8): a document opens through the bridge's recovery check, so edits
// YouCoded's host kept for it (a crash, a closed window) are replayed by the editor's own pipeline.
test('opening a document goes through the recovery check', async () => {
  const dir = await patchedCopy();
  const js = await readFile(path.join(dir, 'bridge.js'), 'utf8');
  assert.match(js, /window\._pendingFileData = await window\.__ycOpenFile\(invoke, filePath\);\n    if \(window\._openEditor\)/);
  assert.doesNotMatch(js, /var b64data = await invoke\('open_file'/);
  await rm(dir, { recursive: true, force: true });
});

// v0.1.27: a pattern found twice means upstream changed shape; patching only the first copy would
// leave the second running the editor's own behaviour, so the build stops instead.
test('the patch stops when a pattern it replaces appears more than once', async () => {
  const dir = await mkdtemp(path.join(os.tmpdir(), 'yco-patch-dup-'));
  await writeFile(path.join(dir, 'index.html'), '<html><head></head><head></head></html>');
  await assert.rejects(
    run(process.execPath, [path.resolve(import.meta.dirname, '..', 'build', 'patch.mjs'), dir]),
    (e) => /pattern found more than once in index\.html/.test(e.stderr),
  );
  await rm(dir, { recursive: true, force: true });
});

// v0.1.37: the Design tab's standard themes. Their names (themes.js, which build/gen-themes.mjs
// makes) must be set after sdk-all-min.js defines AscCommon and before the editor starts, which
// reads them once; euro-office-lite's own page never loads the file.
test('the presentation page loads the standard theme names right after its SDK', async () => {
  const dir = await patchedCopy();
  const html = await readFile(path.join(dir, 'web-apps', 'apps', 'presentationeditor', 'main', 'index.html'), 'utf8');
  assert.match(html, /<script src="\.\.\/\.\.\/\.\.\/\.\.\/sdkjs\/slide\/sdk-all-min\.js"><\/script><script src="\.\.\/\.\.\/\.\.\/\.\.\/sdkjs\/slide\/themes\/themes\.js"><\/script>/);
  // The other editors have no slide themes.
  const word = await readFile(path.join(dir, 'web-apps', 'apps', 'documenteditor', 'main', 'index.html'), 'utf8');
  assert.doesNotMatch(word, /themes\.js/);
  await rm(dir, { recursive: true, force: true });
});
