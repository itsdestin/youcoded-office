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

// A window as the editor pages have it: its own BeforeUnloadEvent and listener registry.
function fakeEditorWindow() {
  class BeforeUnloadEvent { constructor() { this.defaultPrevented = false; this._rv = ''; } preventDefault() { this.defaultPrevented = true; } get returnValue() { return this._rv; } set returnValue(v) { this._rv = v; } }
  const listeners = {};
  // The prototype path an editor script could call directly on the window (v0.1.3).
  class EventTarget { addEventListener(type, cb) { (this.listeners[type] = this.listeners[type] || []).push(cb); } }
  // body.onbeforeunload sets the window's handler natively; modelled as a plain field here.
  class HTMLBodyElement {}
  HTMLBodyElement.prototype.onbeforeunload = null;
  const w = {
    BeforeUnloadEvent,
    EventTarget,
    HTMLBodyElement,
    onbeforeunload: null,
    addEventListener(type, cb) { (listeners[type] = listeners[type] || []).push(cb); },
    listeners,
    location: 'office://t/web-apps/apps/documenteditor/main/index.html',
  };
  w.document = { defaultView: w, querySelectorAll: () => [], querySelector: () => null };
  return w;
}

test('no editor frame can veto the window unload: the host saves and asks the person itself', async () => {
  const src = await readFile(path.resolve(import.meta.dirname, '..', 'bridge', 'yc-bridge.js'), 'utf8');
  const editor = fakeEditorWindow();
  // A listener the editor registered before the guard reached its frame.
  editor.addEventListener('beforeunload', (e) => { e.preventDefault(); e.returnValue = 'unsaved'; });
  const ticks = [];
  const top = fakeEditorWindow();
  top.parent = {};
  top.AscDesktopEditor = { LocalFileSave() {} };
  top.document = { documentElement: {}, querySelectorAll: (q) => (q === 'iframe' ? [{ contentWindow: editor, addEventListener() {} }] : []) };
  const ctx = vm.createContext({
    window: top, document: top.document, setInterval: (fn) => { ticks.push(fn); return 0; }, setTimeout: () => 0, clearTimeout() {},
    requestAnimationFrame: () => 0, WeakSet, JSON, Object, String, MutationObserver: class { observe() {} },
  });
  vm.runInContext(src, ctx);
  ticks.forEach((fn) => fn()); // the editor frame appears: the next tick guards it
  for (const w of [top, editor]) {
    w.onbeforeunload = () => 'leave?';
    assert.equal(w.onbeforeunload, null, 'onbeforeunload is ignored');
    const before = (w.listeners.beforeunload || []).length;
    w.addEventListener('beforeunload', () => {});
    assert.equal((w.listeners.beforeunload || []).length, before, 'no new beforeunload listener');
    w.EventTarget.prototype.addEventListener.call(w, 'beforeunload', () => {});
    assert.equal((w.listeners.beforeunload || []).length, before, 'nor through EventTarget.prototype');
    const body = new w.HTMLBodyElement();
    body.onbeforeunload = () => 'leave?';
    assert.equal(body.onbeforeunload, null, 'body.onbeforeunload is ignored');
    const ev = new w.BeforeUnloadEvent();
    (w.listeners.beforeunload || []).forEach((cb) => cb(ev));
    assert.equal(ev.defaultPrevented, false, 'an earlier listener cannot veto');
    assert.equal(ev.returnValue, '', 'nor through returnValue');
  }
  // Other events are untouched.
  editor.addEventListener('keydown', () => {});
  assert.equal(editor.listeners.keydown.length, 1);
  const other = { listeners: {} };
  editor.EventTarget.prototype.addEventListener.call(other, 'beforeunload', () => {});
  assert.equal(other.listeners.beforeunload.length, 1, 'only the window itself is guarded');
});
