import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import path from 'node:path';
import vm from 'node:vm';

// v0.1.19: the editor frame's settings changes (its localStorage writes, which reach the host page
// as 'storage' events) are sent to the host to be remembered; nothing else the page stores is.
const BRIDGE = process.env.BRIDGE || path.resolve(import.meta.dirname, '..', 'bridge', 'yc-bridge.js');

async function load() {
  const src = await readFile(BRIDGE, 'utf8');
  const listeners = {};
  const timers = [];
  const sent = [];
  const ls = { getItem: () => null, setItem() {} };
  const session = {};
  const win = {
    parent: { postMessage() {} },
    location: { origin: 'office://t' },
    localStorage: ls,
    addEventListener: (t, cb) => { (listeners[t] = listeners[t] || []).push(cb); },
    document: { documentElement: {}, querySelectorAll: () => [] },
    __TAURI__: { core: { invoke: (cmd, args) => { sent.push({ cmd, args: JSON.parse(JSON.stringify(args)) }); return Promise.resolve(null); } } },
  };
  const ctx = vm.createContext({
    window: win, document: win.document, location: win.location, localStorage: ls,
    setInterval: () => 0, setTimeout: (fn) => { timers.push(fn); return timers.length; }, clearTimeout() {},
    requestAnimationFrame: () => 0, getComputedStyle: () => ({ visibility: 'visible' }),
    WeakSet, JSON, Object, MutationObserver: class { observe() {} },
  });
  vm.runInContext(src, ctx);
  const store = (key, newValue, storageArea = ls) => (listeners.storage || []).forEach((cb) => cb({ key, newValue, storageArea }));
  const flush = () => { while (timers.length) timers.shift()(); };
  return { store, flush, sent, session };
}

test('settings changes in the editor are sent to the host together, a moment later', async () => {
  const b = await load();
  b.store('de-settings-unit', '1');
  b.store('de-hidden-status', '1');
  b.store('de-settings-unit', '2');
  assert.equal(b.sent.length, 0, 'batched, not one message per write');
  b.flush();
  assert.deepEqual(b.sent, [{ cmd: 'save_editor_settings', args: { settings: { 'de-settings-unit': '2', 'de-hidden-status': '1' } } }]);
});

test('a setting put back to its default is sent as removed', async () => {
  const b = await load();
  b.store('sse-settings-r1c1', null);
  b.flush();
  assert.deepEqual(b.sent[0].args.settings, { 'sse-settings-r1c1': null });
});

test('nothing else the page stores is sent — only editor-setting keys from localStorage', async () => {
  const b = await load();
  b.store('asc_document_cache', 'the text');
  b.store('guest-username', 'someone');
  b.store('ui-theme-id', 'theme-dark');
  b.store(null, null); // localStorage.clear()
  b.store('de-settings-unit', '1', b.session); // sessionStorage
  b.flush();
  assert.deepEqual(b.sent, []);
});
