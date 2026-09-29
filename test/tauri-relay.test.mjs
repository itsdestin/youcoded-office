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
