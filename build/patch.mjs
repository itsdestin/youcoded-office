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
// 5. File tab (v0.1.7): no "Suggest a feature" (it opens a web page). Printing was switched off
//    here too until v0.1.18, when YouCoded's host learned print_document (step 8).
await replaceOnce('editor-patches.js',
  "macrosMode: 'disable',\n",
  "macrosMode: 'disable',\n            suggestFeature: false,\n");
// 6. Save As writes a copy (v0.1.12). WHY: YouCoded's host translates the document into the
//    chosen file and leaves the open document on its own file — like "Save a copy", so autosave
//    keeps writing where the person opened it. bridge.js would then retitle the editor to the
//    copy's name and move its recovery there, both describing a move that did not happen.
await replaceOnce('bridge.js', "if (pathExt !== 'pdf') {", "if (false && pathExt !== 'pdf') {");
// 7. The editor's export choices reach the host (v0.1.14, Task 2 fix round 1). WHY: a CSV's
//    encoding and delimiter are in the TXT/CSV dialog's text options, which editor-patches.js drops
//    before sdkjs's desktop save path (it cannot take them) — so they are kept aside first. bridge.js
//    then sends them, and its save options (a spreadsheet PDF's print range), with save_file_as.
//    The host checks every value before x2t sees it.
await replaceOnce('editor-patches.js',
  "                  if (options && options.advancedOptions &&\n                      typeof options.advancedOptions.asc_getNativeOptions !== 'function') {\n                    options.advancedOptions = undefined;",
  "                  window.__ycTextOptions = null;\n" +
  "                  if (options && options.advancedOptions &&\n                      typeof options.advancedOptions.asc_getNativeOptions !== 'function') {\n" +
  "                    try { var yto = options.advancedOptions; window.__ycTextOptions = { codePage: yto.asc_getCodePage(), delimiter: yto.asc_getDelimiter(), delimiterChar: yto.asc_getDelimiterChar() }; } catch (e) { window.__ycTextOptions = null; }\n" +
  "                    options.advancedOptions = undefined;");
// WHY taken and cleared in one step (fix round 2): the choices belong to this one Save As; a later
// one that never went through the TXT/CSV dialog (Ctrl+Shift+S, Save As) must not reuse them.
await replaceOnce('bridge.js', "invoke('save_file_as', { path: savePath })",
  "invoke('save_file_as', (function () { var t = window.__ycTextOptions || null; window.__ycTextOptions = null; return { path: savePath, json: jsonOptions || '', text: t }; })())");
// 8. Print (v0.1.18, finish plan Task 3). WHY: YouCoded's host answers print_document by making a
//    PDF of the document and showing the operating system's print dialog for it — or, when that
//    dialog can't be shown, by saying why and offering the PDF as a file. So Print sends only the
//    print panel's choices (a page list; a workbook's sheets, pages and print area), and neither
//    bridge.js's printer plugin nor its PDF viewer (both Tauri's, refused by the host). WHY
//    _isPrinting ends before the host is asked: the dialog can stay open for minutes, and
//    LocalFileSave drops every save while it is set — autosave must keep working meanwhile. WHY no
//    DesktopOfflineAppDocumentEndSave: printing never started a save, and ending one moves the
//    editor's "saved" mark. A failure is only logged here; YouCoded's strip shows main's reason.
await replaceOnce('bridge.js',
  "      await invoke('write_editor_bin', { data: b64 });\n" +
  "      var pdfPath = await invoke('print_document');\n" +
  "\n" +
  "      if (printerName) {\n" +
  "        var printResult = await invoke('plugin:printer|print_pdf', {\n" +
  "          id: printerName,\n" +
  "          path: pdfPath,\n" +
  "          printer: printerName,\n" +
  "          print_settings: '{}',\n" +
  "          remove_after_print: true\n" +
  "        });\n" +
  "      } else {\n" +
  "        await invoke('open_pdf_viewer', { path: pdfPath });\n" +
  "      }\n" +
  "\n" +
  "      if (ref.ew && ref.ew.DesktopOfflineAppDocumentEndSave) {\n" +
  "        ref.ew.DesktopOfflineAppDocumentEndSave(0);\n" +
  "      }\n" +
  "    } catch(e) {\n" +
  "      window._eoLog('[EO] Print: ERROR: ' + (e.message || e));\n" +
  "      if (ref.ew && ref.ew.DesktopOfflineAppDocumentEndSave) {\n" +
  "        ref.ew.DesktopOfflineAppDocumentEndSave(1);\n" +
  "      }\n" +
  "    } finally {",
  "      await invoke('write_editor_bin', { data: b64 });\n" +
  "      window.AscDesktopEditor._isPrinting = false;\n" +
  "      await invoke('print_document', { json: window.__ycPrintJson(ref.ew, optionsJson) });\n" +
  "    } catch(e) {\n" +
  "      window._eoLog('[EO] Print: ERROR: ' + (e.message || e));\n" +
  "    } finally {");
const here = path.dirname(new URL(import.meta.url).pathname);
await copyFile(path.join(here, '..', 'bridge', 'tauri-relay.js'), path.join(dir, 'tauri-relay.js'));
await copyFile(path.join(here, '..', 'bridge', 'yc-bridge.js'), path.join(dir, 'yc-bridge.js'));
await copyFile(path.join(here, '..', 'bridge', 'yc-early.js'), path.join(dir, 'yc-early.js'));
console.log('patch: ok');
