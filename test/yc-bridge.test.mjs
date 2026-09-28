import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import path from 'node:path';
import vm from 'node:vm';

// Loads yc-bridge.js against a fake page: a host window (the parent) and one editor frame
// whose location is the web-apps editor, as the real page nests it.
async function loadBridge({ withApi = true } = {}) {
  const src = await readFile(path.resolve(import.meta.dirname, '..', 'bridge', 'yc-bridge.js'), 'utf8');
  const calls = [];
  const api = { asc_Save: (x) => calls.push(['asc_Save', x]) };
  const editorWin = { location: 'office://t/web-apps/apps/documenteditor/main/index.html', Asc: withApi ? { editor: api } : undefined };
  editorWin.document = { defaultView: editorWin, querySelectorAll: () => [] };
  const parent = {};
  const listeners = [];
  const win = {
    parent,
    AscDesktopEditor: { LocalFileSave: (...a) => calls.push(['LocalFileSave', ...a]) },
    addEventListener: (t, cb) => { if (t === 'message') listeners.push(cb); },
    document: { documentElement: {}, querySelectorAll: (q) => (q === 'iframe' ? [{ contentWindow: editorWin, addEventListener() {} }] : []) },
  };
  const ctx = vm.createContext({
    window: win, document: win.document, setInterval: () => 0, setTimeout: () => 0, clearTimeout() {},
    requestAnimationFrame: () => 0, WeakSet, JSON, MutationObserver: class { observe() {} },
  });
  vm.runInContext(src, ctx);
  const post = (data) => listeners.forEach((cb) => cb({ source: parent, data }));
  return { calls, post };
}

test("a host save request runs the editor's own save, not LocalFileSave", async () => {
  const { calls, post } = await loadBridge();
  post({ type: 'yc:office-save' });
  assert.deepEqual(calls, [['asc_Save', false]]);
});

test('a host save request before the editor API exists falls back to LocalFileSave', async () => {
  const { calls, post } = await loadBridge({ withApi: false });
  post({ type: 'yc:office-save' });
  assert.equal(calls[0][0], 'LocalFileSave');
});
