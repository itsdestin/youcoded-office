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
