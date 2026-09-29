// Applies YouCoded's changes to euro-office-lite's built editor folder.
// WHY at bundle time: the app then serves plain files; nothing is rewritten per request.
import { readFile, writeFile, copyFile, readdir } from 'node:fs/promises';
import path from 'node:path';

const dir = process.argv[2];
if (!dir) { console.error('usage: patch.mjs <editors dir>'); process.exit(2); }

async function replaceOnce(file, from, to) {
  const p = path.join(dir, file);
  const text = await readFile(p, 'utf8');
  if (!text.includes(from)) { console.error(`patch: pattern not found in ${file}: ${from}`); process.exit(1); }
  await writeFile(p, text.replace(from, to));
}

// 1. Relay first (bridge.js calls __TAURI__ at load), then the theme bridge.
await replaceOnce('index.html', '<head>', '<head><script src="tauri-relay.js"></script><script src="yc-bridge.js"></script>');
// 2. Media and dictionaries under the document's own origin (design §3a), never a shared scheme.
await replaceOnce('bridge.js',
  "var ASC_PROTO_BASE = _isWindows ? 'http://ascdesktop.localhost/' : 'ascdesktop://';",
  "var ASC_PROTO_BASE = location.origin + '/asc/';");
// 3. yc-early.js first in every editor page (v0.1.5). WHY at bundle time: it must patch sdkjs's
//    external-links calls the moment sdkjs defines them, before the editor can open a file; the
//    host page's bridge can only reach the frame later (see bridge/yc-early.js). The pages live at
//    web-apps/apps/<editor>/main/, four levels below editors/.
const apps = path.join(dir, 'web-apps', 'apps');
let pages = 0;
for (const app of await readdir(apps)) {
  const main = path.join(apps, app, 'main');
  const files = await readdir(main).catch(() => []);
  for (const f of files.filter((n) => /^index.*\.html$/.test(n))) {
    await replaceOnce(path.join('web-apps', 'apps', app, 'main', f), '<head>', '<head><script src="../../../../yc-early.js"></script>');
    pages++;
  }
}
if (!pages) { console.error('patch: no editor pages found under web-apps/apps/*/main/'); process.exit(1); }
// 4. No macros (v0.1.6). WHY: Euro-Office's defaults are macros: true, macrosMode: 'warn', so a
//    document's own scripts could run after one click on the warning — scripts that could reach
//    past the editor's other seals. YouCoded's editor has no use for them; turning both off also
//    removes the warning itself.
await replaceOnce('editor-patches.js',
  "customization: {\n            about: false,",
  "customization: {\n            macros: false,\n            macrosMode: 'disable',\n            about: false,");
// 5. File tab (v0.1.7): no "Suggest a feature" (it opens a web page) and no printing. WHY print:
//    printing ends in bridge.js's print_document, a command YouCoded's host refuses, so Print (the
//    File tab item, its toolbar button and Ctrl+P) failed silently. Switching it off in the
//    editor's own config removes all three; yc-bridge.js hides the File tab items no config
//    reaches (see HIDDEN_FILE_ITEMS there).
await replaceOnce('editor-patches.js',
  "customization: {\n            macros: false,",
  "customization: {\n            suggestFeature: false,\n            macros: false,");
await replaceOnce('editor-patches.js', 'print: true', 'print: false');
const here = path.dirname(new URL(import.meta.url).pathname);
await copyFile(path.join(here, '..', 'bridge', 'tauri-relay.js'), path.join(dir, 'tauri-relay.js'));
await copyFile(path.join(here, '..', 'bridge', 'yc-bridge.js'), path.join(dir, 'yc-bridge.js'));
await copyFile(path.join(here, '..', 'bridge', 'yc-early.js'), path.join(dir, 'yc-early.js'));
console.log('patch: ok');
