import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import path from 'node:path';
import vm from 'node:vm';

// A picture dragged onto the document. sdkjs's desktop drop handler asks the host for the dropped
// files' paths (AscDesktopEditor.GetDropFiles, which Euro-Office's bridge.js does not define — the
// drop threw and the editor showed "An error occurred"). yc-bridge.js answers instead: it sends
// the dropped picture's BYTES to the document's own origin (upload/), and hands sdkjs the bare
// media name, which sdkjs then resolves through LocalFileGetImageUrl like a dialog pick.
async function loadBridge({ frameDesktop = false, status = 200, answer = { 'media/picture-1.png': 'office://t/asc/docmedia/media/picture-1.png' } } = {}) {
  const src = await readFile(path.resolve(import.meta.dirname, '..', 'bridge', 'yc-bridge.js'), 'utf8');
  const sent = [];
  class XMLHttpRequest {
    open(method, url, async) { this.req = { method, url, async, headers: {} }; }
    setRequestHeader(k, v) { this.req.headers[k] = v; }
    send(body) { this.req.body = body; sent.push(this.req); this.status = status; this.responseText = JSON.stringify(typeof answer === 'function' ? answer(sent.length) : answer); }
  }
  const dropListeners = [];
  const editorWin = {
    location: 'office://t/web-apps/apps/documenteditor/main/index.html',
    addEventListener: (t, cb, capture) => { if (t === 'drop') dropListeners.push({ cb, capture: capture, win: editorWin }); },
    AscDesktopEditor: frameDesktop ? {} : undefined,
    added: [],
    AscCommon: { g_oDocumentUrls: { getImageUrl: (n) => 'office://t/asc/docmedia/media/' + n } },
  };
  editorWin.Asc = { editor: { _addImageUrl: (urls) => editorWin.added.push(urls) } };
  editorWin.document = { defaultView: editorWin, querySelectorAll: () => [], querySelector: () => null };
  const ticks = [];
  const later = [];
  const win = {
    parent: {},
    location: { origin: 'office://t' },
    AscDesktopEditor: {},
    addEventListener: () => {},
    document: { documentElement: {}, querySelectorAll: (q) => (q === 'iframe' ? [{ contentWindow: editorWin, addEventListener() {} }] : []) },
  };
  const ctx = vm.createContext({
    window: win, document: win.document, location: win.location, XMLHttpRequest,
    setInterval: (cb) => { ticks.push(cb); return 0; }, setTimeout: (cb) => { later.push(cb); return 0; }, clearTimeout() {},
    requestAnimationFrame: () => 0, WeakSet, JSON, Object, MutationObserver: class { observe() {} },
  });
  vm.runInContext(src, ctx);
  ticks.forEach((t) => t()); // the timers, among them the walk that reaches the editor frames
  const drop = (files) => dropListeners.forEach((l) => l.cb({ dataTransfer: { files } }));
  return { desktop: win.AscDesktopEditor, frameDesktop: editorWin.AscDesktopEditor, dropListeners, drop, sent, editorWin, runLater: () => later.splice(0).forEach((f) => f()) };
}

const file = (name, type) => ({ name, type, size: 3 });

test("the editor frame's own AscDesktopEditor answers drops too", async () => {
  const b = await loadBridge({ frameDesktop: true });
  b.drop([file('cat.png', 'image/png')]);
  assert.deepEqual([...b.frameDesktop.GetDropFiles()], ['picture-1.png']);
});

test('the editor frames are watched for drops before sdkjs sees them', async () => {
  const { dropListeners } = await loadBridge();
  assert.equal(dropListeners.length, 1);
  assert.equal(dropListeners[0].capture, true);
});

test('a dropped picture goes to the document as bytes and comes back as its media name', async () => {
  const { desktop, drop, sent } = await loadBridge();
  const pic = file('cat.png', 'image/png');
  drop([file('notes.txt', 'text/plain'), pic]);
  assert.deepEqual([...desktop.GetDropFiles()], ['picture-1.png']);
  assert.equal(sent.length, 1);
  assert.equal(sent[0].method, 'POST');
  assert.equal(sent[0].url, 'office://t/upload/drop');
  assert.equal(sent[0].async, false);
  assert.equal(sent[0].headers['Content-Type'], 'image/png');
  assert.equal(sent[0].body, pic);
  assert.equal(desktop.IsImageFile('picture-1.png'), true);
  // one drop, one answer: asking again (a later drop of text) finds nothing
  assert.deepEqual([...desktop.GetDropFiles()], []);
});

test('a drop with no picture, or a refused one, gives sdkjs nothing so it pastes text instead', async () => {
  const none = await loadBridge();
  none.drop([file('notes.txt', 'text/plain')]);
  assert.deepEqual([...none.desktop.GetDropFiles()], []);
  assert.equal(none.sent.length, 0);
  const refused = await loadBridge({ status: 404 });
  refused.drop([file('huge.png', 'image/png')]);
  assert.deepEqual([...refused.desktop.GetDropFiles()], []);
  assert.equal(refused.desktop.IsImageFile('notes.txt'), false);
});

test('every dropped picture goes into the document, not only the first', async () => {
  const b = await loadBridge({ answer: (n) => ({ ['media/picture-' + n + '.png']: 'office://t/asc/docmedia/media/picture-' + n + '.png' }) });
  b.drop([file('a.png', 'image/png'), file('notes.txt', 'text/plain'), file('b.jpg', 'image/jpeg'), file('c.gif', 'image/gif')]);
  // sdkjs inserts the one name it is given (its loop stops at the first picture) ...
  assert.deepEqual([...b.desktop.GetDropFiles()], ['picture-1.png']);
  assert.equal(b.sent.length, 3);
  // ... and the others are inserted right after, through the same editor, in drop order
  b.runLater();
  assert.deepEqual(b.editorWin.added.map((u) => [...u]), [['office://t/asc/docmedia/media/picture-2.png', 'office://t/asc/docmedia/media/picture-3.png']]);
});
