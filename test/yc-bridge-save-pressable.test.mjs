import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import path from 'node:path';
import vm from 'node:vm';

// v0.1.42: Save stays pressable. The host autosaves, and the editor greyed Save out whenever
// nothing was waiting (Destin: "the regular save button is greyed out, so how would i just save a
// doc?"). The bridge turns on the editor's own `forcesave` option, which keeps the button enabled,
// and a plain Save with nothing waiting tells the host instead of doing nothing. BRIDGE overrides
// the file under test.
const BRIDGE = process.env.BRIDGE || path.resolve(import.meta.dirname, '..', 'bridge', 'yc-bridge.js');

async function load({ modified = false } = {}) {
  const src = await readFile(BRIDGE, 'utf8');
  const listeners = [];
  const intervals = [];
  const store = new Map();
  const ls = { getItem: (k) => (store.has(k) ? store.get(k) : null), setItem: (k, v) => store.set(k, String(v)) };
  const saved = [];
  // Like sdkjs: asc_Save lives on the API's prototype, and the full build defines it again later.
  const proto = {
    isDocumentModified() { return this.modified; },
    asc_Save(...args) { saved.push(['first build', ...args]); },
  };
  const api = Object.create(proto);
  api.modified = modified;
  const fullBuildArrives = () => { proto.asc_Save = function (...args) { saved.push(args); }; };
  const main = { appOptions: { isEdit: true, forcesave: false }, buttonChecks: [], disableSaveButton(v) { this.buttonChecks.push([v, this.appOptions.forcesave]); } };
  const editorDoc = { head: { appendChild() {} }, getElementById: () => null, querySelector: () => null, querySelectorAll: () => [], createElement: () => ({ setAttribute() {} }), addEventListener() {} };
  const editorWin = {
    location: 'office://t/web-apps/apps/documenteditor/main/index.html',
    localStorage: ls, document: editorDoc, dispatchEvent() {}, Event: function () {},
    Asc: { editor: api }, editor: api,
    DE: { getController: (n) => (n === 'Main' ? main : null) },
  };
  editorDoc.defaultView = editorWin;
  const toHost = [];
  const parent = { postMessage: (m) => toHost.push(m) };
  const win = {
    parent, location: { origin: 'office://t' }, localStorage: ls,
    addEventListener: (t, cb) => { if (t === 'message') listeners.push(cb); },
    document: { documentElement: {}, querySelectorAll: (q) => (q === 'iframe' ? [{ contentWindow: editorWin, addEventListener() {} }] : []) },
  };
  let raf = null;
  const ctx = vm.createContext({
    window: win, document: win.document, location: win.location, localStorage: ls,
    setInterval: (fn) => { intervals.push(fn); return 0; }, setTimeout: () => 0, clearTimeout() {},
    requestAnimationFrame: (fn) => { raf = fn; return 0; }, getComputedStyle: () => ({ visibility: 'visible' }),
    WeakSet, JSON, Object, MutationObserver: class { observe() {} },
  });
  vm.runInContext(src, ctx);
  const post = (data) => listeners.forEach((cb) => cb({ source: parent, data }));
  const tick = () => { intervals.forEach((fn) => fn()); if (raf) { const f = raf; raf = null; f(); } };
  // The bridge walks the frames on its first theme; that is when it meets the editor.
  post({ type: 'yc:office-theme', theme: { tokens: { canvas: '#000000' } } });
  tick();
  fullBuildArrives();
  return { api, main, saved, toHost, post, tick, proto };
}

test('Save stays enabled: the editor is told to keep it pressable, and re-checks it at once', async () => {
  const { main } = await load();
  assert.equal(main.appOptions.forcesave, true);
  assert.deepEqual(main.buttonChecks[0], [false, true]);
});

test('a plain Save with nothing waiting tells the host instead of doing nothing', async () => {
  const { api, saved, toHost } = await load({ modified: false });
  api.asc_Save();
  assert.deepEqual(saved, []);
  assert.ok(toHost.some((m) => m && m.type === 'yc:office-save-pressed'));
});

test('a Save with changes waiting saves as before, and the host hears nothing extra', async () => {
  const { api, saved, toHost } = await load({ modified: true });
  api.asc_Save();
  assert.deepEqual(saved, [[]]);
  assert.ok(!toHost.some((m) => m && m.type === 'yc:office-save-pressed'));
});

test('a Save As or PDF export (asc_Save with arguments) always passes through, even unmodified', async () => {
  const { api, saved, toHost } = await load({ modified: false });
  api.asc_Save(false, true, undefined, { pdf: true });
  assert.equal(saved.length, 1);
  assert.equal(saved[0][1], true);
  assert.ok(!toHost.some((m) => m && m.type === 'yc:office-save-pressed'));
});

test("the host's own autosave never asks for a save of an unmodified document", async () => {
  const { saved, toHost, post } = await load({ modified: false });
  post({ type: 'yc:office-save' });
  assert.ok(!toHost.some((m) => m && m.type === 'yc:office-save-pressed'));
  assert.deepEqual(saved, []);
});

test("the editor's save is the one current at each call — its full build's, not one kept from before", async () => {
  // load() lets the bridge meet the editor first, then the full build redefines asc_Save.
  const { api, saved, post } = await load({ modified: true });
  post({ type: 'yc:office-save' }); // the host's autosave
  api.asc_Save();                   // the person's Save
  assert.deepEqual(saved, [[false], []]);
});
