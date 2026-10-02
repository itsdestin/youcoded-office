import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import path from 'node:path';
import vm from 'node:vm';

// v0.1.31 (perf investigation 2026-10-01: "the editor feels chuggy to scroll or run my cursor
// through; doesn't update smoothly with theme/resize changes"). Measured causes, each pinned here:
// - :has() rules in the theme sheet made every style recalculation in the editor ~23x slower;
//   their elements now carry classes set from script.
// - every theme change made the editor re-lay out and redraw (a synthetic resize), even when only
//   colours changed.
// - the 150 ms pass that finds the editor's frames never stopped.
// BRIDGE overrides the file under test, so these can be run against an older bridge to see them fail.
const BRIDGE = process.env.BRIDGE || path.resolve(import.meta.dirname, '..', 'bridge', 'yc-bridge.js');

// Every MutationObserver the bridge makes, so a test can play the editor's DOM changes to it.
function observers() {
  const list = [];
  class MO {
    constructor(cb) { this.cb = cb; this.targets = []; list.push(this); }
    observe(target, opts) { this.targets.push({ target, opts }); }
  }
  const fire = (target, records) => list.filter((o) => o.targets.some((t) => t.target === target)).forEach((o) => o.cb(records));
  return { MO, list, fire };
}

function el(tag, { id, cls = [], style = '' } = {}) {
  const classes = new Set(cls);
  const listeners = {};
  return {
    nodeType: 1, tagName: tag.toUpperCase(), id, _style: style, style: {}, firstElementChild: null,
    classList: {
      contains: (c) => classes.has(c), add: (c) => classes.add(c), remove: (c) => classes.delete(c),
      toggle: (c, on) => (on ? classes.add(c) : classes.delete(c)),
    },
    classes,
    getAttribute(k) { return k === 'style' ? this._style : null; },
    matches: () => false, querySelector: () => null, getElementsByTagName: () => [],
    addEventListener(t, cb) { (listeners[t] = listeners[t] || []).push(cb); },
    emit(t) { (listeners[t] || []).forEach((cb) => cb()); },
  };
}

// A host page (the bridge's own window) with one same-origin editor frame. `matches` answers the
// editor document's querySelectorAll by selector, as the real page would.
async function load({ matches = {}, drawn = false } = {}) {
  const src = await readFile(BRIDGE, 'utf8');
  const { MO, list, fire } = observers();
  const head = { children: [], appendChild(c) { this.children.push(c); } };
  const body = el('body');
  const statusbar = el('div', { id: 'statusbar' });
  const queries = [];
  const all = () => Object.values(matches).flat().concat([body, statusbar]);
  const editorDoc = {
    head, body, documentElement: el('html'),
    getElementById: (id) => (id === 'statusbar' ? statusbar : head.children.find((c) => c.id === id) || null),
    getElementsByClassName: (c) => all().filter((e) => e.classList.contains(c)),
    querySelector: (q) => (drawn && q === '#editor_sdk, #ws-canvas-outer, #id_main_view' ? {} : null),
    querySelectorAll: (q) => { queries.push(q); return matches[q] || []; },
    createElement: (tag) => ({ tag, setAttribute(k, v) { this[k] = v; } }),
    addEventListener() {},
  };
  let resizes = 0;
  const editorWin = {
    location: 'office://t/web-apps/apps/documenteditor/main/index.html',
    document: editorDoc, MutationObserver: MO,
    dispatchEvent() { resizes++; }, Event: function () {}, addEventListener() {},
  };
  editorDoc.defaultView = editorWin;
  const frameEl = el('iframe');
  frameEl.contentWindow = editorWin;
  let frames = [frameEl];
  const listeners = [];
  const parent = { postMessage() {} };
  const topDoc = { documentElement: el('html'), body: el('body'), querySelectorAll: (q) => (q === 'iframe' ? frames : []), getElementById: () => null, querySelector: () => null };
  const win = { parent, location: { origin: 'office://t' }, addEventListener: (t, cb) => { if (t === 'message') listeners.push(cb); }, document: topDoc, MutationObserver: MO };
  const timers = { set: 0, cleared: 0, fns: [] };
  let raf = null;
  const ctx = vm.createContext({
    window: win, document: topDoc, location: win.location,
    // Only the 150 ms pass is counted (the slim toolbar's state poll runs every 250 ms).
    setInterval: (fn, ms) => { if (ms !== 150) return 0; timers.set++; timers.fns.push(fn); return timers.set; },
    clearInterval: () => { timers.cleared++; },
    setTimeout: () => 0, clearTimeout() {},
    requestAnimationFrame: (fn) => { raf = fn; return 0; }, getComputedStyle: () => ({ display: 'block', visibility: 'visible' }),
    MutationObserver: MO,
  });
  vm.runInContext(src, ctx);
  const post = (data) => listeners.forEach((cb) => cb({ source: parent, data }));
  const flush = () => { if (raf) { const f = raf; raf = null; f(); } };
  // The first pass, as the page's 150 ms timer would run it: the frames are guarded and watched.
  timers.fns[0]();
  return {
    post, flush, timers, editorWin, editorDoc, body, statusbar, queries, frameEl, list, fire,
    resizes: () => resizes, setFrames: (f) => { frames = f; },
  };
}

const theme = (tokens = {}) => ({
  tokens: { panel: '#111111', fg: '#eeeeee', inset: '#222222', edge: '#333333', accent: '#ff0000', canvas: '#000000', 'font-sans': 'Nunito, sans-serif', ...tokens },
  dark: true, wallpaper: false, panelsOpacity: 1, panelsBlur: 0, fontLinks: [],
});
const cssOf = (doc) => doc.head.children.find((c) => c.id === 'yc-office-theme').textContent;

test('the theme sheet has no :has() rule left — each one slowed every style recalculation', async () => {
  for (const wallpaper of [false, true]) {
    const { post, flush, editorDoc } = await load();
    post({ type: 'yc:office-theme', theme: { ...theme(), wallpaper, panelsBlur: 12 } });
    flush();
    assert.doesNotMatch(cssOf(editorDoc), /:has\(/);
  }
});

// Each old rule's own selector, matched once from script, puts the class its new rule keys on.
const OLD = {
  'yc-pe-sdk': '#editor-container > #editor_sdk:has(> #id_main_parent)',
  'yc-pe-ct': '#editor-container:has(> #editor_sdk > #id_main_parent)',
  'yc-hide-group': '.group:has(> #slot-btn-compare)',
  'yc-resolved': '.user-comment-item:has(.btn-resolve.comment-resolved)',
  'yc-print-hide': '#id-print-settings tr:has(+ tr #print-combo-printer)',
};
test('what the old :has() rules matched carries the class their new rules key on', async () => {
  const { queries } = await load();
  for (const [cls, sel] of Object.entries(OLD)) {
    const q = queries.find((x) => x.split(', ').includes(sel) || x.includes(sel));
    assert.ok(q, `${sel} is still how ${cls} is found`);
  }
  const printQ = queries.find((x) => x.includes('#print-combo-printer'));
  for (const part of ['tr:has(#print-combo-printer)', 'tr:has(#print-combo-color-printing)', 'tr:has(> td > #print-combo-sides)',
    'tr:has(> td > .pages #print-txt-copies)', 'tr:has(#print-btn-system-dialog)', '.dropdown-menu li:has(> a .menu__icon.btn-print)']) {
    assert.ok(printQ.includes(part), `print hides ${part}`);
  }
  // Tagged: load again with each query answering one element.
  const els = {};
  const matches = {};
  for (const q of queries) if (q.includes(':has(')) { els[q] = el('div'); matches[q] = [els[q]]; }
  await load({ matches });
  for (const [cls, sel] of Object.entries(OLD)) {
    const q = Object.keys(els).find((x) => x.includes(sel));
    assert.ok(els[q].classList.contains(cls), `${cls} set`);
  }
});

test('a part added after the document is drawn is tagged at once; one that stops matching loses its class', async () => {
  const card = el('div', { cls: ['user-comment-item'] });
  // The card is in the page throughout (getElementsByClassName finds it); only what matches changes.
  const matches = { '(in the page)': [card] };
  const page = await load({ matches });
  const q = page.queries.find((x) => x.includes('.user-comment-item:has('));
  // The thread is resolved: the comment list draws the item again.
  matches[q] = [card];
  card.matches = (sel) => sel.split(', ').includes('.user-comment-item');
  page.fire(page.editorDoc.documentElement, [{ addedNodes: [card] }]);
  assert.ok(card.classList.contains('yc-resolved'));
  // Reopened: drawn again, no longer matching.
  matches[q] = [];
  page.fire(page.editorDoc.documentElement, [{ addedNodes: [card] }]);
  assert.ok(!card.classList.contains('yc-resolved'));
});

test('the status bar turned off sets a class on body and lays the editor out again; back on, both undone', async () => {
  const { statusbar, body, fire, resizes } = await load();
  assert.ok(!body.classList.contains('yc-no-statusbar'));
  const before = resizes();
  statusbar._style = 'display: none;';
  fire(statusbar, [{ attributeName: 'style' }]);
  assert.ok(body.classList.contains('yc-no-statusbar'));
  assert.equal(resizes(), before + 1);
  // Its style changing otherwise (still hidden) costs nothing.
  fire(statusbar, [{ attributeName: 'style' }]);
  assert.equal(resizes(), before + 1);
  statusbar._style = 'height: 25px;';
  fire(statusbar, [{ attributeName: 'style' }]);
  assert.ok(!body.classList.contains('yc-no-statusbar'));
  assert.equal(resizes(), before + 2);
});

test('the TXT encoding dialog and its mask are hidden by class while it is up, and the mask comes back after', async () => {
  const dlg = el('div', { cls: ['asc-window', 'open-dlg'] });
  dlg.querySelector = (q) => (q === '#id-codepages-combo' ? {} : null);
  const csv = el('div', { cls: ['asc-window', 'open-dlg'] });
  csv.querySelector = (q) => (q === '#id-codepages-combo' || q === '#id-delimiters-combo' ? {} : null);
  const matches = {};
  const page = await load({ matches });
  matches['.asc-window.open-dlg'] = [dlg];
  page.fire(page.body, [{ addedNodes: [dlg] }]);
  assert.ok(dlg.classList.contains('yc-txt-dlg'));
  assert.ok(page.body.classList.contains('yc-txt-mask-off'));
  // Answered and removed: other dialogs keep their mask.
  matches['.asc-window.open-dlg'] = [];
  page.fire(page.body, [{ removedNodes: [dlg], addedNodes: [] }]);
  assert.ok(!page.body.classList.contains('yc-txt-mask-off'));
  // The CSV dialog (it has a delimiter) keeps both.
  matches['.asc-window.open-dlg'] = [csv];
  page.fire(page.body, [{ addedNodes: [csv] }]);
  assert.ok(!csv.classList.contains('yc-txt-dlg'));
  assert.ok(!page.body.classList.contains('yc-txt-mask-off'));
});

test('a colour-only theme change does not lay the editor out again; the first theme, slim mode and a new font do', async () => {
  const { post, flush, resizes, editorDoc } = await load();
  post({ type: 'yc:office-theme', theme: theme() });
  flush();
  assert.equal(resizes(), 1, 'the first theme hides the title row');
  post({ type: 'yc:office-theme', theme: theme({ panel: '#202020', accent: '#00ff00' }) });
  flush();
  assert.equal(resizes(), 1, 'colours only');
  assert.match(cssOf(editorDoc), /#202020/, 'the new colours are applied');
  post({ type: 'yc:office-mode', slim: true });
  flush();
  assert.equal(resizes(), 2, 'slim hides the bands');
  post({ type: 'yc:office-theme', theme: theme({ panel: '#202020', 'font-sans': 'Inter, sans-serif' }) });
  flush();
  assert.equal(resizes(), 3, 'a new font changes text widths');
  post({ type: 'yc:office-theme', theme: { ...theme({ panel: '#202020', 'font-sans': 'Inter, sans-serif' }), wallpaper: true } });
  flush();
  assert.equal(resizes(), 4, 'a wallpaper makes the desk see-through: the canvases redraw');
});

test('the sheet is built once per theme, not on every pass', async () => {
  const { post, flush, editorDoc } = await load();
  post({ type: 'yc:office-theme', theme: theme() });
  flush();
  const style = editorDoc.head.children.find((c) => c.id === 'yc-office-theme');
  let writes = 0;
  let text = style.textContent;
  Object.defineProperty(style, 'textContent', { get: () => text, set: (v) => { writes++; text = v; } });
  post({ type: 'yc:office-theme', theme: theme() });
  flush();
  assert.equal(writes, 0, 'the same theme again writes nothing');
});

test('the 150 ms pass stops once the document is drawn, and a frame that loads again starts it', async () => {
  const page = await load({ drawn: true });
  assert.equal(page.timers.set, 1);
  assert.equal(page.timers.cleared, 1, 'stopped once drawn');
  page.frameEl.emit('load');
  assert.equal(page.timers.set, 2, 'the frame loaded again: the pass looks for it');
});

test('a frame the editor adds after the pass stopped is guarded at once', async () => {
  const page = await load({ drawn: true });
  const late = { location: 'about:blank', document: { querySelectorAll: () => [] }, addEventListener() {} };
  const lateEl = el('iframe');
  lateEl.contentWindow = late;
  page.setFrames([page.frameEl]);
  page.editorDoc.querySelectorAll = (q) => (q === 'iframe' ? [lateEl] : []);
  page.fire(page.editorDoc.documentElement, [{ addedNodes: [lateEl] }]);
  assert.equal(late.__ycUnloadGuard, true, 'its unload veto is guarded');
  assert.throws(() => new late.RTCPeerConnection(), /peer connections are turned off/);
});

test('parts added many times in one frame (a toolbar re-laid out while the window is dragged) are matched once, before that frame paints', async () => {
  const page = await load();
  const frames = [];
  page.editorWin.requestAnimationFrame = (fn) => frames.push(fn);
  const group = el('div');
  group.matches = (sel) => sel.split(', ').includes('#slot-btn-compare');
  const before = page.queries.length;
  for (let i = 0; i < 10; i++) page.fire(page.editorDoc.documentElement, [{ addedNodes: [group] }]);
  assert.equal(page.queries.length, before, 'nothing matched yet');
  assert.equal(frames.length, 1, 'one frame asked for');
  frames[0]();
  const run = page.queries.slice(before).filter((q) => q.includes(':has('));
  assert.equal(run.length, 1, 'only the entry whose part was added');
  assert.ok(run[0].includes('#slot-btn-compare'));
});

// v0.1.38 (Destin: "not all elements update properly/quickly when I switch themes" — Word's
// rulers and tab-stop box kept the dark theme's colours after a switch to a light one).
test('a theme switch repaints every canvas colour once: the sheet before a light/dark flip, the full skin on a dark-to-dark switch', async () => {
  const page = await load();
  const log = [];
  let current = 'theme-dark';
  page.editorWin.Common = { UI: { Themes: { currentThemeId: () => current, setTheme: (id) => { log.push(['flip', id, cssOf(page.editorDoc).includes('#fafafa')]); current = id; } } } };
  page.editorWin.Asc = { editor: { asc_setSkin: (s) => log.push(['skin', s]) } };
  const dark = { ...theme(), dark: true };
  page.post({ type: 'yc:office-theme', theme: dark });
  page.flush();
  log.length = 0;
  // Dark to light: the flip finds the NEW sheet in place, and no second repaint follows it.
  page.post({ type: 'yc:office-theme', theme: { ...theme({ panel: '#fafafa', fg: '#111111', canvas: '#eeeeee' }), dark: false } });
  page.flush();
  assert.deepEqual(log.map((e) => e.slice(0, 1)), [['flip']]);
  assert.deepEqual(log[0], ['flip', 'theme-light', true], 'the light sheet is written before the flip');
  // Dark to dark (no flip): one skin call carrying the rulers' colours too, from the theme.
  current = 'theme-dark';
  page.post({ type: 'yc:office-theme', theme: dark });
  page.flush();
  log.length = 0;
  page.post({ type: 'yc:office-theme', theme: { ...theme({ panel: '#202020', inset: '#303030', edge: '#404040' }), dark: true } });
  page.flush();
  assert.equal(log.length, 1);
  const [kind, skin] = log[0];
  assert.equal(kind, 'skin');
  assert.equal(skin['canvas-ruler-background'], 'rgba(32,32,32,1)');
  assert.equal(skin['canvas-ruler-margins-background'], '#303030');
  assert.equal(skin['canvas-ruler-border'], '#404040');
  assert.equal(skin['background-toolbar'], 'rgba(32,32,32,1)', 'the slide list');
  assert.equal(skin['canvas-cell-title-background'], '#202020', 'sheet headers');
  assert.ok(!('font-family-base' in skin) && !('border-sidemenu' in skin), 'colours only');
});

// v0.1.39 (Destin: "scrollbars still aren't updating consistently"): a still scrollbar is only
// redrawn by the editor when it moves, so a theme change redraws them all (yc-early.js).
test('a theme change redraws every canvas scrollbar, now and once more in the next frame; an unchanged theme does not', async () => {
  const page = await load();
  let redraws = 0;
  const frames = [];
  page.editorWin.__ycRedrawScrolls = () => { redraws++; };
  page.editorWin.requestAnimationFrame = (fn) => frames.push(fn);
  page.post({ type: 'yc:office-theme', theme: theme() });
  page.flush();
  assert.equal(redraws, 1);
  frames.splice(0).forEach((f) => f());
  assert.equal(redraws, 2);
  page.post({ type: 'yc:office-theme', theme: theme() });
  page.flush();
  assert.equal(redraws, 2, 'the same theme again redraws nothing');
  page.post({ type: 'yc:office-theme', theme: theme({ 'scrollbar-thumb': '#aa3355' }) });
  page.flush();
  assert.equal(redraws, 3);
});
