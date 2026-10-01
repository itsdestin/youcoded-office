import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import path from 'node:path';
import vm from 'node:vm';

// tauri-relay.js stands in for Tauri's JS API inside YouCoded: every call becomes one postMessage
// to the framing app page, whose answer comes back as a message from window.parent.
async function relayPage() {
  const src = await readFile(path.resolve(import.meta.dirname, '..', 'bridge', 'tauri-relay.js'), 'utf8');
  const posted = [];
  const handlers = [];
  const parent = { postMessage: (m) => posted.push(m) };
  const window = { parent, addEventListener: (type, cb) => { if (type === 'message') handlers.push(cb); } };
  vm.runInContext(src, vm.createContext({ window, Promise }));
  const answer = (id, result) => handlers.forEach((h) => h({ source: parent, data: { yc: 'rpc-result', id, result } }));
  return { tauri: window.__TAURI__, posted, answer };
}

test('dialog.open asks the host for a file dialog and hands back its answer', async () => {
  const { tauri, posted, answer } = await relayPage();
  const filters = [{ name: 'Images', extensions: ['png'] }];
  const pending = tauri.dialog.open({ multiple: true, filters });
  assert.equal(posted.length, 1);
  assert.equal(posted[0].yc, 'rpc');
  assert.equal(posted[0].cmd, 'open_dialog');
  assert.equal(posted[0].args.multiple, true);
  assert.deepEqual(JSON.parse(JSON.stringify(posted[0].args.filters)), filters);
  answer(posted[0].id, ['yc-picked/abc/cat.png']);
  assert.deepEqual(await pending, ['yc-picked/abc/cat.png']);
});

test('dialog.open with no options asks for one file, and a cancel answers null', async () => {
  const { tauri, posted, answer } = await relayPage();
  const pending = tauri.dialog.open();
  assert.equal(posted[0].args.multiple, false);
  assert.deepEqual(JSON.parse(JSON.stringify(posted[0].args.filters)), []);
  answer(posted[0].id, null);
  assert.equal(await pending, null);
});

// v0.1.12 (Save As / Export / PDF): bridge.js's LocalFileSave asks dialog.save for where the file
// goes, then hands the answer to save_file_as. The host shows its own dialog and answers a handle
// (never a folder) that ends in the chosen name; only the filters go across.
test('dialog.save asks the host for a save dialog with the editor\'s filters and hands back its answer', async () => {
  const { tauri, posted, answer } = await relayPage();
  const filters = [{ name: 'PDF', extensions: ['pdf'] }];
  const pending = tauri.dialog.save({ filters, defaultPath: '/somewhere/else.pdf' });
  assert.equal(posted.length, 1);
  assert.equal(posted[0].cmd, 'save_dialog');
  assert.deepEqual(JSON.parse(JSON.stringify(posted[0].args)), { filters });
  answer(posted[0].id, 'yc-save/abc/Report.pdf');
  assert.equal(await pending, 'yc-save/abc/Report.pdf');
});

test('dialog.save with no options sends no filters, and a cancel answers null', async () => {
  const { tauri, posted, answer } = await relayPage();
  const pending = tauri.dialog.save();
  assert.deepEqual(JSON.parse(JSON.stringify(posted[0].args.filters)), []);
  answer(posted[0].id, null);
  assert.equal(await pending, null);
});

test('the print panel\'s printer list is answered here, never asked of the host', async () => {
  const { tauri, posted } = await relayPage();
  const list = JSON.parse(await tauri.core.invoke('plugin:printer|get_printers'));
  assert.equal(list.length, 1);
  assert.equal(list[0].name, 'YouCoded');
  assert.equal(posted.length, 0);
});

// v0.1.31 (perf investigation 2026-10-01): the editor's bridge.js says "modified" about twice per
// typed character, and "not modified" on every caret move of an unchanged document. Each one cost
// the app a message, an IPC call to its main process and a redraw. Only a change of the value goes
// to the host; a repeat is answered here, as the host would (it answers null).
test('set_document_modified reaches the host only when the value changes', async () => {
  const { tauri, posted } = await relayPage();
  const sent = () => posted.filter((m) => m.cmd === 'set_document_modified').map((m) => m.args.modified);
  for (const v of [false, false, true, true, true, false, false, true]) tauri.core.invoke('set_document_modified', { modified: v });
  assert.deepEqual(sent(), [false, true, false, true]);
  // A repeat is answered at once with the host's own answer.
  assert.equal(await tauri.core.invoke('set_document_modified', { modified: true }), null);
  assert.equal(sent().length, 4);
  // Other commands are never held back.
  tauri.core.invoke('save_changes', { changes: ['a'] });
  tauri.core.invoke('save_changes', { changes: ['a'] });
  assert.equal(posted.filter((m) => m.cmd === 'save_changes').length, 2);
});
