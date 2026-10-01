import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import path from 'node:path';
import vm from 'node:vm';

// yc-early.js runs first in the editor page; sdkjs then publishes its classes the way its
// bundles do (`a.AscCommon = a.AscCommon || {}; a.AscCommon.baseEditorsApi = e`). The patch must
// already be in place at that moment — before anything could open a file — with no timer or walk.
async function editorPage() {
  const src = await readFile(path.resolve(import.meta.dirname, '..', 'bridge', 'yc-early.js'), 'utf8');
  const window = {};
  const ctx = vm.createContext({ window, Object });
  vm.runInContext(src, ctx);
  return window;
}

test('the external-links question is silenced the moment sdkjs publishes its editor API', async () => {
  const a = await editorPage();
  function Api() {}
  Api.prototype.onNeedUpdateExternalReferenceOnOpen = function () { return 'asked'; };
  a.AscCommon = a.AscCommon || {};
  a.AscCommon.baseEditorsApi = Api;
  a.AscCommon = a.AscCommon || {}; // later bundles repeat the namespace line
  // A subclass made right away (spreadsheet_api) and an instance, both before any tick.
  const Cell = function () {};
  Cell.prototype = Object.create(Api.prototype);
  assert.equal(new Cell().onNeedUpdateExternalReferenceOnOpen(), undefined);
  assert.equal(a.AscCommon.baseEditorsApi, Api);
});

test('the workbook never schedules its automatic link refresh', async () => {
  const a = await editorPage();
  let scheduled = 0;
  function WorkbookView() {}
  WorkbookView.prototype.initExternalReferenceUpdateTimer = function () { scheduled++; };
  a.AscCommonExcel = a.AscCommonExcel || {};
  a.AscCommonExcel.WorkbookView = WorkbookView;
  new a.AscCommonExcel.WorkbookView().initExternalReferenceUpdateTimer();
  assert.equal(scheduled, 0);
});

test('a namespace replaced by a new object is patched too', async () => {
  const a = await editorPage();
  function Api() {}
  Api.prototype.onNeedUpdateExternalReferenceOnOpen = function () { return 'asked'; };
  a.AscCommon = { baseEditorsApi: Api };
  assert.equal(new Api().onNeedUpdateExternalReferenceOnOpen(), undefined);
});

// v0.1.6: CSP does not cover WebRTC, so the editor page must not be able to open a peer
// connection — and a script must not be able to put the real one back.
test('an editor page cannot open a peer-to-peer connection, or restore one', async () => {
  const a = await editorPage();
  for (const name of ['RTCPeerConnection', 'webkitRTCPeerConnection']) {
    assert.throws(() => new a[name]({ iceServers: [] }), /peer connections are turned off/);
    // A strict-mode page gets a TypeError here; a sloppy one is silently ignored. Either way:
    try { a[name] = function Real() {}; } catch { /* read-only */ }
    assert.throws(() => new a[name](), /peer connections are turned off/, `${name} stays blocked after reassignment`);
    assert.equal(Reflect.deleteProperty(a, name), false);
  }
});

// v0.1.7: sdkjs draws the document and sheet scrollbars on canvases. They lose their arrows and
// get a slim rounded thumb in the theme's real colours (sdkjs's own drawing used only the red
// channel, so a green thumb came out grey).
function fakeCanvas() {
  const ops = [];
  const ctx = new Proxy({}, {
    get: (_, k) => (k === 'ops' ? ops : (...a) => ops.push([k, ...a])),
    set: (_, k, v) => { ops.push([k, v]); return true; },
  });
  return ctx;
}

test('canvas scrollbars have no arrow buttons', async () => {
  const a = await editorPage();
  a.AscCommon = a.AscCommon || {};
  a.AscCommon.ScrollSettings = function () { this.showArrows = true; this.cornerRadius = 0; };
  const s = new a.AscCommon.ScrollSettings();
  assert.equal(s.showArrows, false);
  assert.equal(s.cornerRadius, 0, 'everything else is sdkjs\'s own');
  assert.ok(s instanceof a.AscCommon.ScrollSettings);
});

test('a canvas scrollbar thumb is drawn slim and rounded in the theme colour, then its hover colour', async () => {
  const a = await editorPage();
  a.AscCommon = a.AscCommon || {};
  function ScrollObject() {}
  ScrollObject.prototype._drawScroll = function () { throw new Error('sdkjs drawing'); };
  a.AscCommon.ScrollObject = ScrollObject;
  const so = new ScrollObject();
  so.context = fakeCanvas();
  so.settings = { isVerticalScroll: true, scrollerColor: '#2f7d55', scrollerHoverColor: '#24613f', scrollerActiveColor: '#24613f' };
  so.scroller = { x: 1, y: 40, w: 12, h: 80 };
  so.canvasW = 14; so.canvasH = 600; so.maxScrollY = 500;
  so._drawScroll(0x2f, 0x2f, 0x2f);
  const fills = so.context.ops.filter((o) => o[0] === 'fillStyle').map((o) => o[1]);
  assert.deepEqual(fills, ['rgb(47,125,85)']);
  assert.ok(so.context.ops.some((o) => o[0] === 'arcTo'), 'rounded');
  const moveTo = so.context.ops.find((o) => o[0] === 'moveTo');
  assert.equal(moveTo[1], 4 + 3, 'a 6px thumb centred in the 12px scroller');
  assert.ok(!so.context.ops.some((o) => o[0] === 'stroke'), 'no outline');
  so.context.ops.length = 0;
  so._drawScroll(0x24, 0x24, 0x24);
  assert.deepEqual(so.context.ops.filter((o) => o[0] === 'fillStyle').map((o) => o[1]), ['rgb(36,97,63)']);
  assert.equal(so.scrollColor, 0x24, 'sdkjs\'s own hover bookkeeping still runs');
});

// v0.1.19: the person's editor settings are written into this page's (empty, one-time) storage
// before any editor code reads it — from the host's answer on the page's own origin.
function settingsPage({ status = 200, body = '{}', stored = {} } = {}) {
  const m = new Map(Object.entries(stored));
  const asked = [];
  const window = {
    location: { origin: 'office://tok' },
    localStorage: { getItem: (k) => (m.has(k) ? m.get(k) : null), setItem: (k, v) => m.set(k, String(v)) },
    XMLHttpRequest: function () {
      this.open = (method, url, async) => asked.push({ method, url, async });
      this.send = () => { if (status === 'throw') throw new Error('offline'); this.status = status; this.responseText = body; };
    },
  };
  return { window, m, asked };
}
async function runEarly(window) {
  const src = await readFile(path.resolve(import.meta.dirname, '..', 'bridge', 'yc-early.js'), 'utf8');
  vm.runInContext(src, vm.createContext({ window, Object, JSON }));
}

test('the remembered settings are in storage before the editor starts, asked synchronously of its own origin', async () => {
  const p = settingsPage({ body: JSON.stringify({ 'de-settings-unit': '1', 'sse-settings-r1c1': '1', 'de-settings-zoom': 100 }) });
  await runEarly(p.window);
  assert.deepEqual(p.asked, [{ method: 'GET', url: 'office://tok/yc-settings.json', async: false }]);
  assert.equal(p.m.get('de-settings-unit'), '1');
  assert.equal(p.m.get('sse-settings-r1c1'), '1');
  assert.equal(p.m.has('de-settings-zoom'), false, 'only text values are written');
});

test('a reloaded page keeps the settings it already has', async () => {
  const p = settingsPage({ body: JSON.stringify({ 'de-settings-unit': '1' }), stored: { 'de-settings-unit': '2' } });
  await runEarly(p.window);
  assert.equal(p.m.get('de-settings-unit'), '2');
});

test('no answer from the host leaves the editor on its own defaults, and the page still loads', async () => {
  for (const variant of [{ status: 404, body: 'not found' }, { status: 200, body: '{bad' }, { status: 'throw' }]) {
    const p = settingsPage(variant);
    await runEarly(p.window);
    assert.equal(p.m.size, 0);
  }
});

// Fix round 2 (Destin, 2026-10-01): the scrollbars "overlap the rounded corners of their
// containers", and "elongate/glitch when scrolling a spreadsheet".
test('a canvas thumb travels a track shortened by the corner radius at both ends', async () => {
  const a = await editorPage();
  a.AscCommon = a.AscCommon || {};
  function ScrollObject() {}
  a.AscCommon.ScrollObject = ScrollObject;
  a.__ycScrollInset = 12;
  const draw = (y) => {
    const so = new ScrollObject();
    so.context = fakeCanvas();
    so.settings = { isVerticalScroll: true, scrollerColor: '#2f7d55' };
    so.scroller = { x: 1, y, w: 12, h: 80 };
    so.canvasW = 14; so.canvasH = 600; so.maxScrollY = 500;
    so._drawScroll(0x2f, 0x2f, 0x2f);
    const ys = so.context.ops.filter((o) => o[0] === 'moveTo' || o[0] === 'lineTo' || o[0] === 'arcTo').flatMap((o) => (o[0] === 'arcTo' ? [o[2], o[4]] : [o[2]]));
    return [Math.min(...ys), Math.max(...ys)];
  };
  assert.deepEqual(draw(0), [12, 12 + 77], 'at the top: 12px clear of the corner, the thumb scaled to the shorter track');
  const [, bottom] = draw(520);
  assert.equal(bottom, 600 - 12, 'at the bottom: 12px clear of the corner');
});

test('a sheet\'s view counts at least 1000 rows and 52 columns, so its thumbs keep their size', async () => {
  const a = await editorPage();
  function WorksheetView() {}
  WorksheetView.prototype._initRowsCount = function () { const b = this.nRowsCount; this.nRowsCount = this.used + 1; return b !== this.nRowsCount; };
  WorksheetView.prototype._initColsCount = function () { const b = this.nColsCount; this.setColsCount(this.usedCols + 1); return b !== this.nColsCount; };
  WorksheetView.prototype.setColsCount = function (n) { this.nColsCount = n; };
  a.AscCommonExcel = a.AscCommonExcel || {};
  a.AscCommonExcel.WorksheetView = WorksheetView;
  const small = new a.AscCommonExcel.WorksheetView();
  Object.assign(small, { used: 8, usedCols: 5, nRowsCount: 0, nColsCount: 0 });
  assert.equal(small._initRowsCount(), true);
  small._initColsCount();
  assert.equal(small.nRowsCount, 1000);
  assert.equal(small.nColsCount, 52);
  const big = new a.AscCommonExcel.WorksheetView();
  Object.assign(big, { used: 5000, usedCols: 80, nRowsCount: 0, nColsCount: 0 });
  big._initRowsCount(); big._initColsCount();
  assert.equal(big.nRowsCount, 5001, 'a bigger sheet keeps its own size');
  assert.equal(big.nColsCount, 81);
});
