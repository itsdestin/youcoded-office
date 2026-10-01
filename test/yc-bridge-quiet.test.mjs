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
  // The start screen opens and creates; the tab closes; the host refuses remove_note_separator;
  // "Save copy" needs a document server; YouCoded has its own Versions; the rest needs a document
  // server or the internet.
  for (const id of ['fm-btn-local-open', 'fm-btn-recent', 'fm-btn-create', 'fm-btn-exit',
    'fm-btn-save-copy',
    'fm-btn-eo-note-separator', 'fm-btn-history', 'fm-btn-rights', 'fm-btn-help', 'fm-btn-suggest']) {
    assert.ok(hidden.has(id), `${id} is hidden`);
  }
  // v0.1.12: Save As, Download as (Export) and Export to PDF work now (the host answers
  // dialog.save and save_file_as), so they are back. v0.1.18: so is Print (print_document).
  for (const id of ['fm-btn-return', 'fm-btn-save', 'fm-btn-info', 'fm-btn-settings',
    'fm-btn-download', 'fm-btn-save-desktop', 'fm-btn-export-pdf', 'fm-btn-print', 'fm-btn-print-with-preview']) {
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

// The panels' blur over a wallpaper must not move the menus that open from them. A
// backdrop-filter makes its element the frame of every position:fixed menu inside it, so the
// right panel's menus (the slide background's "Select picture" → From file among them) opened
// ~1400px to the right, off screen (measured in the dev window). The blur sits on a layer
// behind each panel instead.
test('the panel blur never sits on a panel itself, only on a layer behind it', async () => {
  const { post, tick, head } = await load();
  post({ type: 'yc:office-theme', theme: theme({ wallpaper: true, panelsBlur: 12 }) });
  tick();
  const css = head.children.find((c) => c.id === 'yc-office-theme').textContent;
  const rules = css.split('}').filter((r) => r.includes('backdrop-filter'));
  assert.ok(rules.length > 0, 'the blur is still there');
  for (const r of rules) {
    const selectors = r.split('{')[0].split(',').map((s) => s.trim());
    for (const s of selectors) assert.match(s, /::before$/, `${s} carries the blur itself`);
  }
  assert.match(css, /#right-menu::before/);
});

// v0.1.15 (fix round 2): features that open a non-picture file are hidden with their separators,
// and the hidden TXT dialog's mask never takes a click.
test('file features the host cannot serve are hidden, with the TXT dialog\'s mask', async () => {
  const { post, tick, head } = await load();
  post({ type: 'yc:office-theme', theme: theme() });
  tick();
  const css = cssOf(head);
  for (const sel of ['#slot-btn-text-from-file', '#slot-btn-mailrecepients', '#id-right-menu-mail-merge', '#slot-btn-insaudio', '#slot-btn-insvideo',
    '.group:has(> #slot-btn-compare)', '.group:has(> #slot-btn-compare) + .separator', '.group:has(> #slot-btn-data-from-text)',
    '#external-links-btn-change', '#external-links-btn-open', '#external-links-btn-update', '#chart-button-update-data', '#id-dlg-hyperlink-url .select-button']) {
    assert.ok(css.includes(sel), `${sel} is hidden`);
  }
  // Fix round 3: the External links group itself stays, with Break links (which works).
  assert.ok(!css.includes('#slot-btn-data-external-links'), 'External links stays');
  assert.ok(!css.includes('#external-links-btn-delete'), 'Break links stays');
  // Fix round 4: a chart's linked source name stays, as text that cannot be clicked.
  assert.match(css, /#chart-open-external-link \{ pointer-events: none !important; cursor: default !important; color: inherit !important; text-decoration: none !important; border-bottom: none !important; \}/);
  assert.match(css, /\.modals-mask \{ visibility: hidden !important; pointer-events: none !important; \}/);
});

// v0.1.18 (Print): the host prints through the operating system's dialog. The editor's print panel
// keeps its preview, page setup and page choice, and loses what that dialog owns.
test('the print panel loses the rows the system print dialog owns, and "selection" printing', async () => {
  const { post, tick, head } = await load();
  post({ type: 'yc:office-theme', theme: theme() });
  tick();
  const css = cssOf(head);
  const rule = /([^{}]*#print-combo-printer[^{}]*)\{ display: none !important; \}/.exec(css);
  assert.ok(rule, 'one rule hides them');
  const sel = rule[1];
  for (const part of ['tr:has(#print-combo-printer)', 'tr:has(+ tr #print-combo-printer)', 'tr:has(#print-combo-color-printing)',
    'tr:has(> td > #print-combo-sides)', '#print-txt-copies', 'tr:has(#print-btn-system-dialog)',
    '#print-combo-range li[data-value="2"]', '.dropdown-menu li:has(> a .menu__icon.btn-print)', '#slot-btn-dt-print-quick']) {
    assert.ok(sel.includes(part), `hides ${part}`);
  }
  // What stays: the range, pages, page setup, and the Print and Print to PDF buttons.
  for (const id of ['#print-combo-range,', '#print-txt-pages', '#print-combo-pages', '#print-combo-orient', '#print-combo-margins', '#print-btn-print', '#print-btn-print-pdf']) {
    assert.ok(!sel.split(',').some((s) => s.trim() === id.replace(',', '') || s.trim().endsWith(' ' + id.replace(',', ''))), `${id} stays`);
  }
});

test('the print panel gets one stand-in printer, once, so its Print button works', async () => {
  const { tick, editorWin } = await load();
  const calls = [];
  editorWin.DE = { getController: (n) => (n === 'Print' ? { setPrintersInfo: (cur, list) => calls.push([cur, list]) } : null) };
  tick();
  tick();
  assert.equal(calls.length, 1);
  assert.equal(calls[0][0], 'YouCoded');
  assert.equal(calls[0][1][0].name, 'YouCoded');
});

test('Print sends a document\'s own options, and a workbook\'s panel choices once', async () => {
  const { tick } = await load();
  tick();
  const src = await readFile(BRIDGE, 'utf8');
  const win = { AscDesktopEditor: { _currentDocType: 'word' } };
  vm.runInContext(src, vm.createContext({ window: Object.assign(win, { parent: {}, location: {}, addEventListener() {}, document: { querySelectorAll: () => [] } }), document: { querySelectorAll: () => [] }, setInterval: () => 0, setTimeout: () => 0, clearTimeout() {}, requestAnimationFrame: () => 0, WeakSet, JSON, Object, MutationObserver: class { observe() {} } }));
  assert.equal(win.__ycPrintJson({}, '{"nativeOptions":{"pages":"2"}}'), '{"nativeOptions":{"pages":"2"}}');
  assert.equal(win.__ycPrintJson({}, undefined), '');
  win.AscDesktopEditor._currentDocType = 'cell';
  const ad = { asc_getPrintType: () => 0, asc_getStartPageIndex: () => 1, asc_getEndPageIndex: () => null, asc_getActiveSheetsArray: () => [0], asc_getIgnorePrintArea: () => true };
  const ew = { AscDesktopEditor_PrintOptions: { advancedOptions: ad } };
  assert.deepEqual(JSON.parse(win.__ycPrintJson(ew, '{"nativeOptions":{}}')), { adjustOptions: { printType: 0, startPageIndex: 1, endPageIndex: null, activeSheetsArray: [0] }, spreadsheetLayout: { ignorePrintArea: true } });
  // Used up: Ctrl+P (no panel) prints the whole workbook, never the last panel's choices.
  assert.equal(ew.AscDesktopEditor_PrintOptions, null);
  assert.deepEqual(JSON.parse(win.__ycPrintJson(ew, '')), { adjustOptions: { printType: 1 }, spreadsheetLayout: { ignorePrintArea: false } });
});

// v0.1.21 (finish plan Task 6): Office's comments panel wears YouCoded's comment cards — the same
// surfaces, border, radius and neutral avatar as the app's own (desktop components/comments/).
test('comment cards in the editor look like the app\'s: inset card, edge border, neutral avatar, no quote', async () => {
  const { post, tick, head } = await load();
  post({ type: 'yc:office-theme', theme: theme({ tokens: { ...theme().tokens, inset: '#0a0b0c', 'edge-dim': '#202122', 'radius-lg': '14px' } }) });
  tick();
  const css = cssOf(head);
  assert.match(css, /\.user-comment-item \{ [^}]*background-color: #0a0b0c !important; border: 1px solid #202122 !important; border-radius: 14px !important;/);
  // The editor colours each author's initial; the app's avatar is neutral (accent is for state).
  assert.match(css, /\.user-comment-item \.user-info \.color \{ [^}]*background-color: #0a0b0c !important;/);
  assert.match(css, /\.user-comment-item \.user-quote \{ display: none !important; \}/);
  // Resolved: the circle check filled in the accent.
  assert.match(css, /\.btn-resolve\.comment-resolved \{ [^}]*background-color: #[0-9a-f]+ !important; \}/i);
});

// v0.1.23 (finish plan Task 7, Destin: "the side bars and such should gracefully connect to the
// header element to frame the doc/content"; a full-width line under the header; the File tab in
// YouCoded's card style). The ribbon, strips and status bar are one panel frame with no lines in
// it; the desk is the one rounded, outlined hole; the File tab's page is the same hole.
const ruleFor = (css, selectorPart) => css.split('}').filter((r) => r.split('{')[0].includes(selectorPart));

test('the ribbon, side strips and status bar are one frame: painted once, no lines inside it', async () => {
  for (const wallpaper of [false, true]) {
    const { post, tick, head } = await load();
    post({ type: 'yc:office-theme', theme: theme({ wallpaper, tokens: { ...theme().tokens, 'radius-lg': '14px' } }) });
    tick();
    const css = cssOf(head);
    // The frame's bands paint the panel...
    const bands = ruleFor(css, '#toolbar, #left-menu, #right-menu, #statusbar');
    assert.ok(bands.some((r) => /background: (#111111|rgba\(17,17,17,1\))/.test(r)), 'the bands paint the panel');
    // ...and nothing inside them paints it again or draws an edge.
    for (const part of ['#toolbar .toolbar', '#left-menu .tool-menu-btns', '#right-menu .tool-menu-btns', '#statusbar .statusbar', '#toolbar .extra']) {
      const r = ruleFor(css, part).find((x) => x.includes('background: transparent'));
      assert.ok(r && /box-shadow: none/.test(r), `${part} is see-through with no outline`);
    }
    assert.ok(ruleFor(css, '#toolbar .box-controls::before').some((r) => /box-shadow: none/.test(r)), 'the tools row loses its rounded under-line');
    assert.ok(ruleFor(css, '#left-menu .tool-menu-btns').some((r) => /border: 0/.test(r)), 'the strips lose their own edge');
    // The old per-band hairlines are gone.
    assert.doesNotMatch(css, /inset -1px 0 0 #333333/);
    assert.doesNotMatch(css, /#statusbar, \.statusbar \{ box-shadow: inset 0 1px 0/);
    // No inner layer is painted a second time on a wallpaper (the second shade under the ribbon).
    if (wallpaper) assert.doesNotMatch(css, /#toolbar \.toolbar, #statusbar, \.statusbar, #left-menu/);
  }
});

test('the document desk is one hole: the large radius, one edge line, corners in the frame colour', async () => {
  const { post, tick, head } = await load();
  post({ type: 'yc:office-theme', theme: theme({ tokens: { ...theme().tokens, 'radius-lg': '14px' } }) });
  tick();
  const css = cssOf(head);
  const hole = ruleFor(css, '#id_main_parent::after').find((r) => r.includes('border-radius'));
  assert.ok(hole, 'the hole layer exists');
  for (const s of ['#editor-container > #editor_sdk::after', '.layout-ct.vbox > #editor_sdk::after', '#id_main_parent::after']) assert.ok(hole.includes(s), `${s} is a hole`);
  assert.match(hole, /border-radius: 14px/);
  assert.match(hole, /box-shadow: 0 0 0 14px rgba\(17,17,17,1\), inset 0 0 0 1px #333333/);
  assert.match(hole, /pointer-events: none/);
  // The presentation's outer #editor_sdk holds the slide list (frame) and is not cut itself.
  assert.ok(ruleFor(css, '#editor_sdk:has(> #id_main_parent)::after').some((r) => /content: none/.test(r)));
});

test('the File tab: list in the frame, page as the hole, content on YouCoded cards, no accent bar', async () => {
  const { post, tick, head } = await load();
  post({ type: 'yc:office-theme', theme: theme({ tokens: { ...theme().tokens, 'radius-lg': '14px', 'radius-md': '9px' } }) });
  tick();
  const css = cssOf(head);
  assert.ok(ruleFor(css, '#file-menu-panel .panel-menu').some((r) => /border-right: 0/.test(r)), 'no line beside the list');
  const page = ruleFor(css, '#file-menu-panel .panel-context > .content-box').find((r) => r.includes('border-radius'));
  assert.match(page, /background-color: #000000/);
  assert.match(page, /border-radius: 14px/);
  assert.match(page, /inset 0 0 0 1px #333333/);
  const card = ruleFor(css, 'table.main').find((r) => r.includes('border-radius'));
  assert.match(card, /background-color: #111111/);
  assert.match(card, /border: 1px solid #333333/);
  assert.match(card, /border-radius: 14px/);
  assert.ok(ruleFor(css, '.btn-doc-format').some((r) => /border-radius: 14px/.test(r)), 'export formats are cards');
  const open = ruleFor(css, 'li.fm-btn.active').find((r) => r.includes('background-color'));
  assert.match(open, /box-shadow: none/);
  assert.doesNotMatch(css, /inset 3px 0 0 #ff0000/);
});

// Fix round 1 (Destin, 2026-10-01): "the bottom of the inner/outer containers touch each other";
// "stray straight lines that poke past rounded corners".
test('with the status bar off the hole keeps its bottom gap, and no straight line runs past its corners', async () => {
  const { post, tick, head } = await load();
  post({ type: 'yc:office-theme', theme: theme({ tokens: { ...theme().tokens, 'radius-lg': '14px' } }) });
  tick();
  const css = cssOf(head);
  // The editor is laid out 8px shorter when the status bar is hidden, and the frame fills the gap.
  assert.ok(ruleFor(css, 'body:has(#statusbar[style*="display: none"]) #viewport').some((r) => /height: calc\(100% - 8px\)/.test(r)));
  assert.ok(ruleFor(css, 'body:has(#statusbar[style*="display: none"])::after').some((r) => /height: 8px/.test(r) && /background: rgba\(17,17,17,1\)/.test(r)));
  // The holes draw no border of their own (the slide area's left border ran past both corners).
  assert.ok(ruleFor(css, '#id_main_parent').some((r) => /border: 0/.test(r)));
  // The notes divider is inset by the radius at both ends, and gone when the notes are collapsed.
  const divider = ruleFor(css, '#id_bottom_pannels_container').find((r) => r.includes('background-size'));
  assert.match(divider, /border-top-color: transparent/);
  assert.match(divider, /background-size: calc\(100% - 2 \* 14px\) 1px/);
  assert.ok(ruleFor(css, '#id_bottom_pannels_container[style*="height: 4px"]').some((r) => /background-image: none/.test(r)));
});
