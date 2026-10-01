import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import path from 'node:path';
import vm from 'node:vm';

// v0.1.21 (finish plan Task 6): the host's comment changes for an open document are made by the
// editor itself, so they show at once and the editor's autosave writes them (a write to the file
// while the editor has it open is erased by that autosave). These stand-ins mirror the sdkjs calls
// measured in the YouCoded dev window (see bridge/yc-comments.js's header).
const SRC = path.resolve(import.meta.dirname, '..', 'bridge', 'yc-comments.js');

function wordApi() {
  const comments = []; // {Id, Data}
  let n = 0;
  const log = [];
  const ld = {
    selection: 'caret-at-start',
    GetSelectionState() { return this.selection; },
    SetSelectionState(s) { log.push(['restore', s]); this.selection = s; },
    UpdateSelection() {},
    ClearSearch() { log.push(['clear']); },
    text: 'The budget grows. The budget review is in May.',
    Search(ss) {
      const els = {};
      let i = 0, at = -1;
      while ((at = this.text.indexOf(ss.text, at + 1)) >= 0) els[i++] = at;
      this.found = els;
      return { Elements: els };
    },
    SelectSearchElement(id) { this.selection = `search-${id}`; log.push(['select', id]); },
  };
  const api = {
    callbacks: {},
    asc_registerCallback(name, fn) { (this.callbacks[name] ||= []).push(fn); },
    fire(name, id) { (this.callbacks[name] || []).forEach((f) => f(id)); },
    private_GetLogicDocument() { return ld; },
    pluginMethod_GetAllComments() { return comments.map((c) => ({ Id: c.Id, Data: JSON.parse(JSON.stringify(c.Data)) })); },
    pluginMethod_AddComment(data) {
      const at = ld.found[ld.selection.replace('search-', '')];
      const Id = `c${++n}`;
      comments.push({ Id, Data: { ...data, QuoteText: at === undefined ? '' : ld.text.slice(at, at + 6), Solved: !!data.Solved, Replies: data.Replies || [] } });
      log.push(['add', Id, ld.selection]);
      this.fire('asc_onAddComment', Id);
      return Id;
    },
    pluginMethod_ChangeComment(id, data) { const c = comments.find((x) => x.Id === id); c.Data = { ...data, Solved: !!data.Solved }; log.push(['change', id]); },
    pluginMethod_RemoveComments(ids) { ids.forEach((id) => comments.splice(comments.findIndex((x) => x.Id === id), 1)); log.push(['remove', ...ids]); },
  };
  const win = { Asc: { editor: api }, AscCommon: { CSearchSettings: class { put_Text(t) { this.text = t; } put_MatchCase() {} } } };
  return { api, win, comments, log, ld };
}

class CellComment {
  constructor() { this.sText = ''; this.sUserName = ''; this.sTime = ''; this.bSolved = false; this.bDocument = true; this.aReplies = []; this.nId = null; }
  ReadFromSimpleObject(o) { if (o.Text) this.sText = o.Text; if (o.UserName) this.sUserName = o.UserName; if (o.Time) this.sTime = o.Time; if (o.Solved) this.bSolved = o.Solved; (o.Replies || []).forEach((r) => { const c = new CellComment(); c.ReadFromSimpleObject(r); this.aReplies.push(c); }); }
  asc_putDocumentFlag(v) { this.bDocument = v; }
  asc_putCol(c) { this.nCol = c; } asc_putRow(r) { this.nRow = r; }
  asc_putText(t) { this.sText = t; } asc_putUserName(u) { this.sUserName = u; } asc_putTime(t) { this.sTime = t; } asc_putSolved(s) { this.bSolved = s; }
  setId() { this.nId = `sheet${this.wsId}_${Math.random().toString(16).slice(2)}`; this.sGuid = `{G-${this.nId}}`; }
  asc_getId() { return this.nId; }
  asc_addReply(r) { this.aReplies.push(r); }
  clone() { const c = Object.assign(new CellComment(), this); c.aReplies = this.aReplies.slice(); return c; }
}

function cellApi() {
  const log = [];
  const sheets = ['Budget', 'Notes'].map((name, i) => ({ name, id: String(7 + i), aComments: [], getName() { return this.name; }, getId() { return this.id; } }));
  const views = sheets.map((model) => ({
    cellCommentator: {
      model,
      getComment(col, row) { return model.aComments.find((c) => c.nCol === col && c.nRow === row) || null; },
      isLockedComment(d, cb) { cb(true); },
      // r false: through history (the editor counts it as a change and autosaves it).
      _addComment(d, r) { log.push(['_addComment', model.name, d.bDocument, r]); model.aComments.push(d); },
      changeComment(id, data) { const i = model.aComments.findIndex((c) => c.nId === id); model.aComments[i] = data; log.push(['change', model.name, id]); },
    },
  }));
  const api = {
    editMode: false,
    wbModel: { aWorksheets: sheets, getActive: () => 0 },
    wb: { getWorksheet: (i) => views[i] },
    asc_getCellEditMode() { return this.editMode; },
    asc_registerCallback() {},
    asc_removeComment(id) { sheets.forEach((s) => { const i = s.aComments.findIndex((c) => c.nId === id); if (i >= 0) s.aComments.splice(i, 1); }); log.push(['remove', id]); },
  };
  const win = { Asc: { editor: api, asc_CCommentData: CellComment } };
  return { api, win, sheets, log };
}

async function load(editor, { drawn = true } = {}) {
  const src = await readFile(SRC, 'utf8');
  const listeners = [];
  const posted = [];
  const timers = [];
  const parent = { postMessage: (m) => posted.push(m) };
  const win = {
    parent, __ycDrawn: drawn, __ycEditorWindow: () => editor.win,
    addEventListener: (t, cb) => { if (t === 'message') listeners.push(cb); },
  };
  const ctx = vm.createContext({
    window: win, setInterval: () => 0, setTimeout: (fn) => { timers.push(fn); return 1; }, JSON, Object, Math, String, parseInt, isNaN,
  });
  vm.runInContext(src, ctx);
  let seq = 0;
  const send = (op, source = parent) => {
    const id = `q${++seq}`;
    listeners.forEach((cb) => cb({ source, data: { type: 'yc:office-comments', id, op } }));
    const r = posted.find((m) => m.type === 'yc:office-comments-result' && m.id === id);
    return r ? JSON.parse(JSON.stringify(r.result)) : undefined;
  };
  const flush = () => { while (timers.length) timers.shift()(); };
  return { send, posted, flush, win };
}

test('Word: the quoted text is commented by the assistant and the person\'s selection is put back', async () => {
  const w = wordApi();
  const { send } = await load(w);
  const r = send({ kind: 'add', key: 'k1', quote: 'budget', occurrence: 1, text: 'Check this', author: 'Assistant' });
  assert.deepEqual(r, { ok: true, id: 'c1' });
  // The second "budget" was selected for the add, then the person's caret came back.
  assert.deepEqual(w.log.find((e) => e[0] === 'add'), ['add', 'c1', 'search-1']);
  assert.equal(w.ld.selection, 'caret-at-start');
  assert.equal(w.comments[0].Data.UserName, 'Assistant');
  assert.equal(w.comments[0].Data.QuoteText, 'budget');
});

test('Word: a quote that is not in the document is refused and nothing is added', async () => {
  const w = wordApi();
  const { send } = await load(w);
  assert.deepEqual(send({ kind: 'add', quote: 'nowhere', text: 'x' }), { ok: false, error: 'quote-not-found' });
  assert.equal(w.comments.length, 0);
  assert.equal(w.ld.selection, 'caret-at-start');
});

test('Word: reply, resolve, reopen, edit and delete change the editor\'s own comment', async () => {
  const w = wordApi();
  const { send } = await load(w);
  send({ kind: 'add', quote: 'budget', text: 'Check', author: 'Assistant' });
  assert.deepEqual(send({ kind: 'reply', id: 'c1', text: 'Done', author: 'You' }), { ok: true, index: 0 });
  assert.deepEqual(send({ kind: 'resolve', id: 'c1' }), { ok: true });
  assert.equal(w.comments[0].Data.Solved, true);
  send({ kind: 'reopen', id: 'c1' });
  assert.equal(w.comments[0].Data.Solved, false);
  send({ kind: 'edit-reply', id: 'c1', index: 0, text: 'Done now' });
  send({ kind: 'edit', id: 'c1', text: 'Check again' });
  const listed = send({ kind: 'list' });
  assert.equal(listed.comments[0].text, 'Check again');
  assert.deepEqual(listed.comments[0].replies.map((r) => [r.author, r.text]), [['You', 'Done now']]);
  assert.equal(listed.comments[0].quote, 'budget');
  assert.deepEqual(send({ kind: 'edit-reply', id: 'c1', index: 5, text: 'x' }), { ok: false, error: 'reply-not-found' });
  assert.deepEqual(send({ kind: 'delete-reply', id: 'c1', index: 0 }), { ok: true });
  assert.deepEqual(send({ kind: 'delete', id: 'c1' }), { ok: true });
  assert.deepEqual(send({ kind: 'resolve', id: 'c1' }), { ok: false, error: 'comment-not-found' });
});

test('an op the host sends again (its answer came too late) runs once', async () => {
  const w = wordApi();
  const { send } = await load(w);
  const a = send({ kind: 'add', key: 'same', quote: 'budget', text: 'Once' });
  const b = send({ kind: 'add', key: 'same', quote: 'budget', text: 'Once' });
  assert.deepEqual(a, b);
  assert.equal(w.comments.length, 1);
});

test('an editor that is not drawn yet says so, so the host keeps the op and asks again', async () => {
  const w = wordApi();
  const { send } = await load(w, { drawn: false });
  assert.deepEqual(send({ kind: 'add', quote: 'budget', text: 'x' }), { ok: false, error: 'editor-not-ready' });
  assert.equal(w.comments.length, 0);
});

test('only the host may send ops', async () => {
  const w = wordApi();
  const { send } = await load(w);
  assert.equal(send({ kind: 'add', quote: 'budget', text: 'x' }, { postMessage() {} }), undefined);
  assert.equal(w.comments.length, 0);
});

test('a comment added in the editor (by anyone) tells the host, once per burst', async () => {
  const w = wordApi();
  const { send, posted, flush } = await load(w);
  send({ kind: 'add', quote: 'budget', text: 'x' }); // registers the editor's events
  w.api.fire('asc_onChangeCommentData', 'c1');
  w.api.fire('asc_onRemoveComment', 'c1');
  flush();
  assert.equal(posted.filter((m) => m.type === 'yc:office-comments-changed').length, 1);
});

test('Excel: a cell comment on a named sheet is a cell comment (not the workbook\'s) and counts as a change', async () => {
  const c = cellApi();
  const { send } = await load(c);
  const r = send({ kind: 'add', key: 'x1', sheet: 'Notes', cell: 'B2', text: 'Rent rises', author: 'Assistant' });
  assert.equal(r.ok, true);
  // The GUID the saved file will use as the thread's id.
  assert.equal(r.guid, c.sheets[1].aComments[0].sGuid);
  assert.deepEqual(c.log[0], ['_addComment', 'Notes', false, false]);
  const added = c.sheets[1].aComments[0];
  assert.equal(added.nCol, 1);
  assert.equal(added.nRow, 1);
  assert.equal(added.sUserName, 'Assistant');
  assert.match(added.sTime, /^\d+$/);
  assert.deepEqual(send({ kind: 'add', sheet: 'Notes', cell: 'B2', text: 'again' }), { ok: false, error: 'cell-has-comment' });
  assert.deepEqual(send({ kind: 'add', sheet: 'Nope', cell: 'B2', text: 'x' }), { ok: false, error: 'sheet-not-found' });
  const listed = send({ kind: 'list' });
  assert.deepEqual(listed.comments.map((x) => [x.sheet, x.cell, x.text, x.guid]), [['Notes', 'B2', 'Rent rises', r.guid]]);
});

test('Excel: replies and resolve change the comment through its own sheet; nothing while a cell is being typed in', async () => {
  const c = cellApi();
  const { send } = await load(c);
  const { id } = send({ kind: 'add', sheet: 'Notes', cell: 'C3', text: 'Why?', author: 'You' });
  assert.deepEqual(send({ kind: 'reply', id, text: 'Because', author: 'Assistant' }), { ok: true, index: 0 });
  assert.deepEqual(c.log.at(-1), ['change', 'Notes', id]);
  send({ kind: 'resolve', id });
  assert.equal(c.sheets[1].aComments[0].bSolved, true);
  c.api.editMode = true;
  assert.deepEqual(send({ kind: 'reply', id, text: 'later' }), { ok: false, error: 'editor-busy' });
  assert.equal(c.sheets[1].aComments[0].aReplies.length, 1);
});

test('Excel: a move takes the whole thread to the new cell', async () => {
  const c = cellApi();
  const { send } = await load(c);
  const { id } = send({ kind: 'add', sheet: 'Budget', cell: 'A1', text: 'Here', author: 'Assistant' });
  send({ kind: 'reply', id, text: 'ok', author: 'You' });
  const r = send({ kind: 'move', id, sheet: 'Notes', cell: 'D4' });
  assert.equal(r.ok, true);
  assert.equal(c.sheets[0].aComments.length, 0);
  const moved = c.sheets[1].aComments[0];
  assert.equal(moved.sText, 'Here');
  assert.equal(moved.aReplies.length, 1);
  assert.equal(moved.nCol, 3);
});

test('cells convert both ways', async () => {
  const { win } = await load(wordApi());
  const { cellToRC, rcToCell } = win.__ycComments;
  assert.deepEqual(JSON.parse(JSON.stringify(cellToRC('AB12'))), { col: 27, row: 11 });
  assert.equal(rcToCell(27, 11), 'AB12');
  assert.equal(cellToRC('12'), null);
});
