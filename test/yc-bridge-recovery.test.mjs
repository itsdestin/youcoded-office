import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import path from 'node:path';
import vm from 'node:vm';

// v0.1.24 (finish plan Task 8): YouCoded's host keeps every batch of edits in a recovery journal.
// Opening a document asks it first and replays any edits its file never got (bridge.js's own
// recovery pipeline); and the editor sends its edits every second, not only after a pause.
// BRIDGE overrides the file under test.
const BRIDGE = process.env.BRIDGE || path.resolve(import.meta.dirname, '..', 'bridge', 'yc-bridge.js');

async function load() {
  const src = await readFile(BRIDGE, 'utf8');
  const intervals = [];
  const logged = [];
  const api = { autoSaveGapFast: 2000, intervalWaitAutoSave: 1000 };
  const editorDoc = { head: { appendChild() {} }, getElementById: () => null, querySelector: () => null, querySelectorAll: () => [], createElement: () => ({ setAttribute() {} }), addEventListener() {} };
  const editorWin = { location: 'office://t/web-apps/apps/documenteditor/main/index.html', document: editorDoc, Asc: { editor: api }, localStorage: { getItem: () => null, setItem() {} } };
  editorDoc.defaultView = editorWin;
  let raf = null;
  const win = {
    parent: { postMessage() {} }, location: { origin: 'office://t' }, addEventListener() {},
    _eoLog: (m) => logged.push(m),
    document: { documentElement: {}, querySelectorAll: (q) => (q === 'iframe' ? [{ contentWindow: editorWin, addEventListener() {} }] : []) },
  };
  const ctx = vm.createContext({
    window: win, document: win.document, location: win.location, localStorage: { getItem: () => null, setItem() {} },
    setInterval: (fn) => { intervals.push(fn); return 0; }, setTimeout: () => 0, clearTimeout() {},
    requestAnimationFrame: (fn) => { raf = fn; return 0; }, getComputedStyle: () => ({ visibility: 'visible' }),
    WeakSet, JSON, Object, Promise, MutationObserver: class { observe() {} },
  });
  vm.runInContext(src, ctx);
  const tick = () => { intervals.forEach((fn) => { try { fn(); } catch { /* other passes need more of a page */ } }); if (raf) { const f = raf; raf = null; f(); } };
  return { win, api, tick, logged };
}

/** A host that answers the editor's commands from `answers` (a function's result, or a value). */
function host(answers) {
  const calls = [];
  const invoke = (cmd, args) => {
    calls.push([cmd, args]);
    const a = answers[cmd];
    try { return Promise.resolve(typeof a === 'function' ? a(args) : a); } catch (e) { return Promise.reject(e); }
  };
  return { invoke, calls };
}

test('a document with edits its file never got opens from the journal, with the edits to replay', async () => {
  const { win } = await load();
  const h = host({
    recovery_candidates: [{ id: 'k1', name: 'plan.docx' }],
    recovery_load: { id: 'k1', data: 'QkFTRQ==', changes: ['c1', 'c2'] },
    open_file: () => { throw new Error('must not open the file'); },
  });
  const pending = await win.__ycOpenFile(h.invoke, 'plan.docx');
  assert.deepEqual(JSON.parse(JSON.stringify(pending)), { data: 'QkFTRQ==', path: 'plan.docx', name: 'plan.docx', recovery: { id: 'k1', changes: ['c1', 'c2'] } });
  assert.deepEqual(h.calls.map((c) => c[0]), ['recovery_candidates', 'recovery_load']);
  assert.deepEqual(JSON.parse(JSON.stringify(h.calls[1][1])), { id: 'k1' });
});

test('with nothing to recover the file opens as before', async () => {
  const { win } = await load();
  const h = host({ recovery_candidates: [], open_file: 'RklMRQ==' });
  const pending = await win.__ycOpenFile(h.invoke, 'dir\\plan.docx');
  assert.deepEqual(JSON.parse(JSON.stringify(pending)), { data: 'RklMRQ==', path: 'dir\\plan.docx', name: 'plan.docx' });
  assert.deepEqual(h.calls.map((c) => c[0]), ['recovery_candidates', 'open_file']);
});

test('a journal that cannot be loaded falls back to the file, and says so in the log', async () => {
  const { win, logged } = await load();
  const h = host({ recovery_candidates: [{ id: 'k1' }], recovery_load: () => { throw new Error('gone'); }, open_file: 'RklMRQ==' });
  const pending = await win.__ycOpenFile(h.invoke, 'plan.docx');
  assert.equal(pending.data, 'RklMRQ==');
  assert.equal(pending.recovery, undefined);
  assert.ok(logged.some((l) => /recovering unsaved changes failed: gone/.test(l)));
});

test('a file that cannot be opened is reported to the host, not left on "Opening…"', async () => {
  const { win } = await load();
  const posted = [];
  win.parent.postMessage = (m) => posted.push(m);
  const h = host({ recovery_candidates: [], open_file: () => { throw new Error('x2t failed'); } });
  await assert.rejects(win.__ycOpenFile(h.invoke, 'plan.docx'), /x2t failed/);
  assert.deepEqual(JSON.parse(JSON.stringify(posted)), [{ type: 'yc:office-failed' }]);
});

test('the editor sends its edits every second, even while the person keeps typing', async () => {
  const { api, tick } = await load();
  tick();
  assert.equal(api.intervalWaitAutoSave, 0);
  assert.equal(api.autoSaveGapFast, 1000);
});

// v0.1.25 (Task 8 fix round 1): just before its window closes, the host asks for the newest edits.
test('asked before a close, the editor sends its edits at once, then answers', async () => {
  const src = await readFile(BRIDGE, 'utf8');
  const order = [];
  const listeners = [];
  const api = { lastSaveTime: new Date(), _autoSave() { order.push(['autosave', this.lastSaveTime.getTime()]); } };
  const editorDoc = { head: { appendChild() {} }, getElementById: () => null, querySelector: () => null, querySelectorAll: () => [], createElement: () => ({ setAttribute() {} }), addEventListener() {} };
  const editorWin = { location: 'office://t/web-apps/apps/documenteditor/main/index.html', document: editorDoc, Asc: { editor: api }, localStorage: { getItem: () => null, setItem() {} } };
  editorDoc.defaultView = editorWin;
  const parent = { postMessage: (m) => order.push(['answer', m.type]) };
  const win = {
    parent, location: { origin: 'office://t' }, addEventListener: (t, cb) => { if (t === 'message') listeners.push(cb); },
    document: { documentElement: {}, querySelectorAll: (q) => (q === 'iframe' ? [{ contentWindow: editorWin, addEventListener() {} }] : []) },
  };
  const ctx = vm.createContext({ window: win, document: win.document, location: win.location, localStorage: { getItem: () => null, setItem() {} },
    setInterval: () => 0, setTimeout: () => 0, clearTimeout() {}, requestAnimationFrame: () => 0, getComputedStyle: () => ({}), WeakSet, JSON, Object, Promise, Date, MutationObserver: class { observe() {} } });
  vm.runInContext(src, ctx);
  listeners.forEach((cb) => cb({ source: parent, data: { type: 'yc:office-journal' } }));
  assert.deepEqual(order, [['autosave', 0], ['answer', 'yc:office-journaled']]);
});
