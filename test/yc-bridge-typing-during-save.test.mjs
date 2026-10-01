import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import path from 'node:path';
import vm from 'node:vm';

// v0.1.20: typing during a save is no longer thrown away. sdkjs's desktop save
// (DesktopOfflineAppDocumentStartSave) holds a "block interaction" long action from the moment
// it takes the document's bytes until the host has translated and written the file, and
// web-apps turns the keyboard off for that whole time. Measured in the YouCoded dev window
// (2026-09-30): keys typed while a 20 MB workbook saved (~5 s) were lost — 4 whole cells. The
// bytes are taken synchronously at the start, so the bridge lifts the block right after that
// and swallows the editor's own matching "end" later. BRIDGE overrides the file under test.
const BRIDGE = process.env.BRIDGE || path.resolve(import.meta.dirname, '..', 'bridge', 'yc-bridge.js');

const BLOCK = 1, OTHER_TYPE = 0, SAVE = 7, OTHER_ID = 3;

async function load() {
  const src = await readFile(BRIDGE, 'utf8');
  const listeners = [];
  const intervals = [];
  const store = new Map();
  const ls = { getItem: (k) => (store.has(k) ? store.get(k) : null), setItem: (k, v) => store.set(k, String(v)) };
  // A minimal sdkjs: a long-action counter with its start/end calls, and the desktop save pair.
  const log = [];
  const api = {
    counter: 0,
    modified: true,
    isDocumentModified() { return this.modified; },
    sync_StartAction(type, id) { log.push(['start', type, id]); if (type === BLOCK) this.counter++; },
    sync_EndAction(type, id) { log.push(['end', type, id]); if (type === BLOCK) this.counter = Math.max(0, this.counter - 1); },
    asc_Save() { editorWin.DesktopOfflineAppDocumentStartSave(false); },
  };
  const editorDoc = { head: { appendChild() {} }, getElementById: () => null, querySelector: () => null, querySelectorAll: () => [], createElement: () => ({ setAttribute() {} }), addEventListener() {} };
  const editorWin = {
    location: 'office://t/web-apps/apps/spreadsheeteditor/main/index.html',
    localStorage: ls, document: editorDoc, dispatchEvent() {}, Event: function () {},
    Asc: { editor: api, c_oAscAsyncActionType: { BlockInteraction: BLOCK, Information: OTHER_TYPE }, c_oAscAsyncAction: { Save: SAVE, Open: OTHER_ID } },
    editor: api,
    saves: [],
    // sdkjs: start the block, then hand the bytes to the bridge's LocalFileSave (which ends the
    // save later with DesktopOfflineAppDocumentEndSave).
    DesktopOfflineAppDocumentStartSave(isSaveAs) { api.sync_StartAction(BLOCK, SAVE); editorWin.saves.push(isSaveAs === true ? 'saveas' : 'save'); },
    DesktopOfflineAppDocumentEndSave() { api.sync_EndAction(BLOCK, SAVE); },
  };
  editorDoc.defaultView = editorWin;
  const parent = { postMessage() {} };
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
  return { api, editorWin, post, tick, log };
}

test('an autosave lets typing through as soon as the editor has taken the bytes', async () => {
  const { api, editorWin, post } = await load();
  post({ type: 'yc:office-save' });
  assert.deepEqual(editorWin.saves, ['save']);
  // The bytes are taken; the host is still translating. The keyboard must not be off now.
  assert.equal(api.counter, 0);
  // The editor's own end of that save arrives later: it is the block already lifted, so it
  // must not end anything else.
  api.sync_StartAction(BLOCK, OTHER_ID); // something else blocks meanwhile (an image upload)
  editorWin.DesktopOfflineAppDocumentEndSave(0);
  assert.equal(api.counter, 1, 'the save\'s late end must not end the other block');
  api.sync_EndAction(BLOCK, OTHER_ID);
  assert.equal(api.counter, 0);
});

test('the editor\'s own save (Ctrl+S) is lifted the same way once the bridge has seen the frame', async () => {
  const { api, editorWin, tick, post } = await load();
  post({ type: 'yc:office-mode', slim: false }); // any message that makes the bridge walk the frames
  tick();
  editorWin.DesktopOfflineAppDocumentStartSave(false);
  assert.equal(api.counter, 0);
  editorWin.DesktopOfflineAppDocumentEndSave(0);
  assert.equal(api.counter, 0);
  // A second save behaves the same (the swallow is per save, not once).
  editorWin.DesktopOfflineAppDocumentStartSave(false);
  assert.equal(api.counter, 0);
  editorWin.DesktopOfflineAppDocumentEndSave(0);
  assert.equal(api.counter, 0);
});

test('a Save As keeps its block (its dialog is up), and its end still ends it', async () => {
  const { api, editorWin, tick, post } = await load();
  post({ type: 'yc:office-mode', slim: false }); // any message that makes the bridge walk the frames
  tick();
  editorWin.DesktopOfflineAppDocumentStartSave(true);
  assert.equal(api.counter, 1);
  editorWin.DesktopOfflineAppDocumentEndSave(0);
  assert.equal(api.counter, 0);
});

test('a save that never reaches its end does not swallow the next save\'s block', async () => {
  const { api, editorWin, tick, post } = await load();
  post({ type: 'yc:office-mode', slim: false }); // any message that makes the bridge walk the frames
  tick();
  editorWin.DesktopOfflineAppDocumentStartSave(false); // lifted; its end never comes (gave up)
  editorWin.DesktopOfflineAppDocumentStartSave(true); // a Save As next: its block must hold
  assert.equal(api.counter, 1);
  editorWin.DesktopOfflineAppDocumentEndSave(0);
  assert.equal(api.counter, 0);
});

test('an editor that defines its save start again after the bridge saw it is still lifted', async () => {
  const { api, editorWin, tick, post } = await load();
  post({ type: 'yc:office-mode', slim: false });
  tick();
  // The spreadsheet editor sets DesktopOfflineAppDocumentStartSave again once it has loaded
  // (measured in the dev window): the bridge's next pass wraps the new one.
  editorWin.DesktopOfflineAppDocumentStartSave = function (isSaveAs) { api.sync_StartAction(BLOCK, SAVE); editorWin.saves.push(isSaveAs === true ? 'saveas' : 'save'); };
  post({ type: 'yc:office-mode', slim: false });
  tick();
  editorWin.DesktopOfflineAppDocumentStartSave(false);
  assert.equal(api.counter, 0);
  editorWin.DesktopOfflineAppDocumentEndSave(0);
  assert.equal(api.counter, 0);
});
