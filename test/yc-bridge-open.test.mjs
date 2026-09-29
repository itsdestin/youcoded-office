import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import path from 'node:path';
import vm from 'node:vm';

// v0.1.14 (Task 2 fix round 1): Ctrl+O and the editor's own Open do nothing (files open from the
// Office start screen; the editor's Open reloaded the document and dropped unsaved edits), and
// Word's TXT encoding dialog — a choice x2t ignores — is answered OK without being shown.
async function load(opts = {}) {
  const src = await readFile(path.resolve(import.meta.dirname, '..', 'bridge', 'yc-bridge.js'), 'utf8');
  const intervals = [];
  const opened = [];
  const keyListeners = [];
  const clicked = [];
  // One TXT dialog (encoding only) and one CSV dialog (encoding and delimiter).
  const dialog = (ids) => {
    const w = { style: {}, querySelector: (q) => {
      if (q === '[result="ok"]') return { click: () => clicked.push(ids.join('+')) };
      const m = /^#([\w-]+)$/.exec(q);
      return m && ids.includes(m[1]) ? {} : null;
    } };
    return w;
  };
  const observers = [];
  const txt = dialog(['id-codepages-combo']);
  const csv = dialog(['id-codepages-combo', 'id-delimiters-combo']);
  let shown = true;
  const editorDoc = { body: {}, querySelectorAll: (q) => (q === '.asc-window.open-dlg' || q === '.asc-window' ? [txt, csv] : []), addEventListener() {} };
  const editorWin = {
    location: 'office://t/web-apps/apps/documenteditor/main/index.html', document: editorDoc,
    addEventListener: (t, cb, cap) => { if (t === 'keydown') keyListeners.push({ win: 'editor', cb, cap }); },
    getComputedStyle: () => ({ display: shown ? 'block' : 'none' }),
    MutationObserver: class { constructor(cb) { this.cb = cb; this.targets = []; observers.push(this); } observe(target, opts) { this.targets.push(target); this.opts = opts; } },
  };
  editorDoc.defaultView = editorWin;
  const win = {
    parent: { postMessage() {} },
    AscDesktopEditor: { LocalFileOpen: () => opened.push('open') },
    addEventListener: (t, cb, cap) => { if (t === 'keydown') keyListeners.push({ win: 'host', cb, cap }); },
    document: { documentElement: {}, head: { appendChild() {} }, createElement: () => ({ setAttribute() {} }), querySelectorAll: (q) => (q === 'iframe' ? [{ contentWindow: editorWin, addEventListener() {} }] : []), addEventListener() {}, getElementById: () => null },
  };
  const ctx = vm.createContext({
    window: win, document: win.document, location: { origin: 'office://t' }, localStorage: { getItem: () => null, setItem() {} },
    setInterval: (fn) => { intervals.push(fn); return intervals.length; }, setTimeout: () => 0, clearTimeout() {},
    requestAnimationFrame: () => 0, WeakSet, JSON, MutationObserver: class { observe() {} }, getComputedStyle: () => ({}),
  });
  vm.runInContext(src, ctx);
  if (opts.hiddenAtFirst) shown = false;
  intervals.forEach((fn) => { try { fn(); } catch { /* other passes need more of a page */ } });
  return { win, opened, keyListeners, clicked, txt, csv, observers, show: () => { shown = true; } };
}

test('Ctrl+O is swallowed in every editor window before the editor sees it', async () => {
  const { keyListeners } = await load();
  assert.deepEqual(keyListeners.map((l) => [l.win, l.cap]).sort(), [['editor', true], ['host', true]]);
  let prevented = 0, stopped = 0;
  const ev = (key, extra = {}) => ({ key, ctrlKey: true, shiftKey: false, altKey: false, preventDefault: () => prevented++, stopImmediatePropagation: () => stopped++, ...extra });
  keyListeners[0].cb(ev('o'));
  keyListeners[0].cb(ev('O'));
  assert.equal(prevented, 2); assert.equal(stopped, 2);
  keyListeners[0].cb(ev('s')); // Ctrl+S still reaches the editor
  keyListeners[0].cb(ev('O', { shiftKey: true }));
  assert.equal(prevented, 2);
});

test("the editor's own Open does nothing", async () => {
  const { win, opened } = await load();
  await win.AscDesktopEditor.LocalFileOpen();
  assert.deepEqual(opened, []);
});

test('Word\'s TXT encoding dialog is answered OK and never shown; the CSV one stays', async () => {
  const { clicked, txt, csv } = await load();
  assert.deepEqual(clicked, ['id-codepages-combo']);
  assert.equal(txt.style.visibility, 'hidden');
  assert.equal(csv.style.visibility, undefined);
});

// Fix round 2: answered the moment it is SHOWN, never before — OK on a dialog not yet shown would
// let its show() put it and its mask up for good. Fix round 3: the observers watch only the body's
// children (a window added) and each window's own style/class (shown), not the whole page.
test("the TXT dialog is answered by an observer only once it is shown", async () => {
  const { clicked, observers, show, txt } = await load({ hiddenAtFirst: true });
  assert.equal(observers.length, 2);
  const [shownObs, bodyObs] = observers;
  assert.deepEqual(JSON.parse(JSON.stringify(bodyObs.opts)), { childList: true });
  assert.deepEqual(JSON.parse(JSON.stringify(shownObs.opts)), { attributes: true, attributeFilter: ['style', 'class'] });
  assert.deepEqual(clicked, []);
  assert.equal(txt.style.visibility, 'hidden');
  // A window added to the body is watched for being shown.
  const win2 = { classList: { contains: (c) => c === 'asc-window' } };
  bodyObs.cb([{ addedNodes: [win2, { classList: { contains: () => false } }] }]);
  assert.equal(shownObs.targets.includes(win2), true);
  show();
  shownObs.cb([]);
  assert.deepEqual(clicked, ['id-codepages-combo']);
  shownObs.cb([]);
  assert.deepEqual(clicked, ['id-codepages-combo'], 'answered once');
});
