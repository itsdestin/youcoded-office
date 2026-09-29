import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import path from 'node:path';
import vm from 'node:vm';

// v0.1.4: the bridge keeps the editor quiet and offline (no "New feature" tips, no external-links
// warning and no link re-fetch), themes the icons, hides rulers in slim mode, and loads the
// theme's web font only from the editor's own origin. BRIDGE overrides the file under test, so
// the same tests can be run against an older bridge to see them fail.
const BRIDGE = process.env.BRIDGE || path.resolve(import.meta.dirname, '..', 'bridge', 'yc-bridge.js');

function storage() {
  const m = new Map();
  return { getItem: (k) => (m.has(k) ? m.get(k) : null), setItem: (k, v) => m.set(k, String(v)), map: m };
}

// A host page (the bridge's own window) with one same-origin editor frame whose code has loaded.
async function load({ editorCode = true } = {}) {
  const src = await readFile(BRIDGE, 'utf8');
  const ls = storage();
  const listeners = [];
  const intervals = [];
  const baseApi = function () {};
  baseApi.prototype.onNeedUpdateExternalReferenceOnOpen = function () { this.sendEvent('asc_onNeedUpdateExternalReferenceOnOpen'); };
  const WorkbookView = function () {};
  let timers = 0;
  WorkbookView.prototype.initExternalReferenceUpdateTimer = function () { timers++; };
  const added = [];
  const TooltipManager = { addTips(arr) { added.push(arr); } };
  const rulerCalls = [];
  const head = { children: [], appendChild(el) { this.children.push(el); } };
  const editorDoc = {
    head,
    getElementById: (id) => head.children.find((c) => c.id === id) || null,
    querySelector: (q) => {
      const m = /^link\[data-yc-font="(.*)"\]$/.exec(q);
      return m ? head.children.find((c) => c['data-yc-font'] === m[1]) || null : null;
    },
    querySelectorAll: () => [],
    createElement: (tag) => ({ tag, setAttribute(k, v) { this[k] = v; } }),
    addEventListener() {},
  };
  const editorWin = {
    location: 'office://t/web-apps/apps/documenteditor/main/index.html',
    localStorage: ls,
    document: editorDoc,
    dispatchEvent() {},
    Event: function () {},
    Asc: { editor: { asc_SetViewRulers: (v) => rulerCalls.push(v) } },
  };
  editorDoc.defaultView = editorWin;
  if (editorCode) {
    editorWin.AscCommon = { baseEditorsApi: baseApi };
    editorWin.AscCommonExcel = { WorkbookView };
    editorWin.Common = { UI: { TooltipManager } };
  }
  const parent = { postMessage() {} };
  const win = {
    parent,
    location: { origin: 'office://t' },
    localStorage: ls,
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
  return { ls, post, tick, editorWin, baseApi, WorkbookView, added, head, rulerCalls, timers: () => timers };
}

const theme = (extra = {}) => ({
  tokens: { panel: '#111111', fg: '#eeeeee', inset: '#222222', edge: '#333333', accent: '#ff0000', canvas: '#000000', 'font-sans': 'Nunito, sans-serif' },
  dark: true, wallpaper: false, panelsOpacity: 1, panelsBlur: 0, fontLinks: [], ...extra,
});

test('the "New feature" tips are marked seen before the editor starts', async () => {
  const { ls } = await load({ editorCode: false });
  for (const name of ['help-tip-comment-filter', 'help-tip-chart-elements', 'de-help-tip-signature', 'sse-help-tip-solver', 'sse-help-tip-cellFormat', 'pe-help-tip-master-tab']) {
    assert.equal(ls.getItem(name), '1', name);
  }
});

test('a "New feature" tip added later is marked seen instead of shown; other tips pass through', async () => {
  const { tick, editorWin, added, ls } = await load();
  tick();
  editorWin.Common.UI.TooltipManager.addTips({
    fresh: { name: 'de-help-tip-brand-new', isNewFeature: true },
    warning: { text: 'connection lost' },
  });
  assert.deepEqual(Object.keys(added[0]), ['warning']);
  assert.equal(ls.getItem('de-help-tip-brand-new'), '1');
});

test('opening a file with external links neither asks nor re-fetches them', async () => {
  const { tick, baseApi, WorkbookView, timers } = await load();
  tick();
  const sent = [];
  const api = Object.create(baseApi.prototype);
  api.sendEvent = (e) => sent.push(e);
  api.onNeedUpdateExternalReferenceOnOpen();
  assert.deepEqual(sent, []);
  Object.create(WorkbookView.prototype).initExternalReferenceUpdateTimer();
  assert.equal(timers(), 0);
});

test('toolbar icons follow the theme\'s text colour, not the editor\'s near-black default', async () => {
  const { post, tick, head } = await load();
  post({ type: 'yc:office-theme', theme: theme() });
  tick();
  const css = head.children.find((c) => c.id === 'yc-office-theme').textContent;
  assert.match(css, /--icon-gray-primary:#eeeeee !important/);
});

test('the slim editor starts without rulers; the full editor keeps them', async () => {
  const slim = await load();
  slim.post({ type: 'yc:office-mode', slim: true });
  assert.equal(slim.ls.getItem('de-hidden-rulers'), '1');
  assert.equal(slim.ls.getItem('pe-hidden-rulers'), '1');
  assert.deepEqual(slim.rulerCalls, [false]);
  const full = await load();
  full.post({ type: 'yc:office-mode', slim: false });
  assert.equal(full.ls.getItem('de-hidden-rulers'), '0');
});

test('the theme font loads only through the editor\'s own origin, never from Google directly', async () => {
  const { post, tick, head } = await load();
  const own = 'office://t/yc-fonts/css?u=https%3A%2F%2Ffonts.googleapis.com%2Fcss2%3Ffamily%3DNunito';
  post({ type: 'yc:office-theme', theme: theme({ fontLinks: ['https://fonts.googleapis.com/css2?family=Nunito', 'office://other/yc-fonts/css?u=x', own] }) });
  tick();
  const links = head.children.filter((c) => c.tag === 'link').map((l) => l.href);
  assert.deepEqual(links, [own]);
});

// v0.1.7, the polish pass (Destin, 2026-09-28): the File tab shows only what works inside
// YouCoded, scrollbars wear the app's thumb colours, gallery tiles stay square.
const cssOf = (head) => head.children.find((c) => c.id === 'yc-office-theme').textContent;
const hiddenIds = (css) => {
  const m = /([^{}]*)\{ display: none !important; \}/g;
  const ids = new Set();
  for (const r of css.matchAll(m)) for (const id of r[1].matchAll(/#file-menu-panel #([\w-]+)/g)) ids.add(id[1]);
  return ids;
};

test('the File tab hides what cannot work in YouCoded and keeps what does', async () => {
  const { post, tick, head } = await load();
  post({ type: 'yc:office-theme', theme: theme() });
  tick();
  const hidden = hiddenIds(cssOf(head));
  // The start screen opens and creates; the tab closes; the host refuses the save dialog,
  // save_file_as, print_document and remove_note_separator; YouCoded has its own Versions;
  // the rest needs a document server or the internet.
  for (const id of ['fm-btn-local-open', 'fm-btn-recent', 'fm-btn-create', 'fm-btn-exit', 'fm-btn-download',
    'fm-btn-save-desktop', 'fm-btn-save-copy', 'fm-btn-export-pdf', 'fm-btn-print', 'fm-btn-print-with-preview',
    'fm-btn-eo-note-separator', 'fm-btn-history', 'fm-btn-rights', 'fm-btn-help', 'fm-btn-suggest']) {
    assert.ok(hidden.has(id), `${id} is hidden`);
  }
  for (const id of ['fm-btn-return', 'fm-btn-save', 'fm-btn-info', 'fm-btn-settings']) {
    assert.ok(!hidden.has(id), `${id} stays`);
  }
});

test('panel and menu scrollbars wear the app\'s thumb and hover colours', async () => {
  const { post, tick, head } = await load();
  post({ type: 'yc:office-theme', theme: theme({ tokens: { ...theme().tokens, 'scrollbar-thumb': '#123456', 'scrollbar-hover': '#654321' } }) });
  tick();
  const css = cssOf(head);
  assert.match(css, /::-webkit-scrollbar-thumb \{ background: #123456; border-radius: 4px; \}/);
  assert.match(css, /::-webkit-scrollbar-thumb:hover \{ background: #654321; \}/);
  assert.match(css, /\.ps-scrollbar-x\.always-visible-x \{ background: #123456 !important; border: 0 !important; border-radius: 3px !important; \}/);
  assert.match(css, /--canvas-scroll-thumb:#123456 !important/);
});

test('a host without scrollbar tokens still gets themed scrollbars (the edge colour)', async () => {
  const { post, tick, head } = await load();
  post({ type: 'yc:office-theme', theme: theme() });
  tick();
  assert.match(cssOf(head), /::-webkit-scrollbar-thumb \{ background: #333333;/);
});

test('style-gallery tiles are square; only the gallery frame is rounded', async () => {
  const { post, tick, head } = await load();
  post({ type: 'yc:office-theme', theme: theme({ tokens: { ...theme().tokens, 'radius-sm': '16px', 'radius-md': '26px' } }) });
  tick();
  const css = cssOf(head);
  assert.match(css, /\.combo-dataview \.view \.item, \.combo-dataview \.view \.item canvas[^{]*\{ border-radius: 0 !important; \}/);
  assert.match(css, /\.combo-dataview \.view \{ border-radius: 16px 0 0 16px !important; \}/);
  // Other gallery items and swatches: a few pixels at most, whatever the theme's small radius.
  assert.match(css, /--border-radius-dataview-item:min\(16px, 3px\) !important/);
});

test('the canvases get the theme\'s scrollbar and header colours whenever the theme changes', async () => {
  const { post, tick, editorWin } = await load();
  const skins = [];
  editorWin.Asc.editor.asc_setSkin = (s) => skins.push(s);
  post({ type: 'yc:office-theme', theme: theme({ tokens: { ...theme().tokens, 'scrollbar-thumb': '#123456' } }) });
  tick();
  tick(); // nothing changed: not handed over again
  assert.equal(skins.length, 1);
  assert.equal(skins[0]['canvas-scroll-thumb'], '#123456');
  assert.equal(skins[0]['canvas-cell-title-background'], '#111111');
  assert.equal(skins[0].name, undefined, 'the editor keeps its own light/dark theme');
  post({ type: 'yc:office-theme', theme: theme({ tokens: { ...theme().tokens, 'scrollbar-thumb': '#abcdef' } }) });
  tick();
  assert.equal(skins.length, 2);
  assert.equal(skins[1]['canvas-scroll-thumb'], '#abcdef');
});
