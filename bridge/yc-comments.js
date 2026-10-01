/* yc-comments.js — the editor's comments, reached by YouCoded's host (v0.1.21, finish plan Task 6).
 *
 * WHY this exists: YouCoded's assistant (and the app's own reading view) writes comments into
 * Word and Excel files. While Office has a file open, a write to the file itself is erased by the
 * editor's next autosave — the editor saves what IT holds. So while a document is open, the host
 * sends each comment change here instead: the editor makes it (it shows at once, and autosave
 * writes it), and reads come from here too, so they include what is not saved yet.
 *
 * Messages (only from window.parent, the host):
 *   host → editor  {type:'yc:office-comments', id, op}
 *   editor → host  {type:'yc:office-comments-result', id, result}   result: {ok:true, ...} | {ok:false, error}
 *                  {type:'yc:office-comments-changed'}               a comment was added, changed or
 *                                                                    removed here (by anyone), so the
 *                                                                    host's reading views refresh
 * op.kind: list | add | reply | resolve | reopen | edit | edit-reply | delete | delete-reply | move.
 * Comments are named by the editor's own ids; replies by their place in the thread (0-based).
 * Each op carries `key` (the host's request id): an op seen before answers with its first result
 * instead of running twice — the host retries an op whose answer it did not get in time.
 *
 * Measured in the YouCoded dev window (2026-09-30) against this bundle's sdkjs:
 *  - Word: asc_addComment comments on the current selection, so the quoted text is found
 *    (LogicDocument.Search), selected, commented, and the person's own selection put back.
 *    pluginMethod_GetAllComments/_ChangeComment/_RemoveComments read, replace and remove.
 *  - Excel: a comment is added on a given sheet and cell through that sheet's commentator. Its
 *    data must carry documentFlag false (a new asc_CCommentData defaults to a WORKBOOK comment,
 *    which the save drops), and it must go through _addComment's history path, or the editor
 *    never counts it as a change and autosave never writes it.
 *  - Both editors fire asc_onAddComment / asc_onChangeCommentData / asc_onRemoveComment.
 */
(function () {
  var done = {};      // op key -> its result (see `key` above)
  var doneKeys = [];  // oldest first, so the memory stays small
  var DONE_MAX = 200;

  function editorWindow() {
    // yc-bridge.js finds the editor frame (it is rebuilt on open); reuse its walk.
    try { return window.__ycEditorWindow ? window.__ycEditorWindow() : null; } catch (e) { return null; }
  }
  function apiOf(w) {
    try { return w && ((w.Asc && w.Asc.editor) || w.editor) || null; } catch (e) { return null; }
  }
  function isCell(api) { return !!(api && api.wbModel && api.wb); }

  // ── Changes the host must hear about ──
  // WHY debounced: one assistant op can fire several events (a move is a remove and an add), and
  // the host re-reads every comment on each push; one push per burst is enough.
  var changedTimer = 0;
  function changed() {
    if (changedTimer) return;
    changedTimer = setTimeout(function () { changedTimer = 0; window.parent.postMessage({ type: 'yc:office-comments-changed' }, '*'); }, 300);
  }
  var EVENTS = ['asc_onAddComment', 'asc_onAddComments', 'asc_onChangeCommentData', 'asc_onRemoveComment'];
  // ── The reply and edit boxes answer keys as the app's do (v0.1.34) ──
  // WHY: the panel is styled as YouCoded's comment boxes (yc-bridge.js commentsCss), which send on
  // Enter (Shift+Enter: a new line) and close on Escape — and the editor's own Close button is
  // gone from the reply box, so Escape is how it closes. Only inside a reply or edit box.
  function keys(w) {
    var d;
    try { d = w && w.document; } catch (e) { return; }
    if (!d || d.__ycCommentKeys) return;
    d.__ycCommentKeys = true;
    d.addEventListener('keydown', function (e) {
      var ta = e.target;
      if (!ta || ta.tagName !== 'TEXTAREA' || typeof ta.closest !== 'function') return;
      var box = ta.closest('.reply-ct, .inner-edit-ct');
      if (!box) return;
      var hit = null;
      if (e.key === 'Escape') hit = box.querySelector('.btn-close, .btn-inner-close');
      else if (e.key === 'Enter' && !e.shiftKey && !e.ctrlKey && !e.metaKey && !e.altKey && !e.isComposing) {
        hit = box.querySelector('.btn-reply, .btn-inner-edit');
        if (hit && (hit.disabled || /\bdisabled\b/.test(hit.className))) { e.preventDefault(); return; }
      }
      if (!hit) return;
      e.preventDefault();
      e.stopPropagation();
      hit.click();
    }, true);
  }
  function listen() {
    keys(editorWindow());
    var api = apiOf(editorWindow());
    if (!api || api.__ycComments || typeof api.asc_registerCallback !== 'function') return;
    api.__ycComments = true;
    EVENTS.forEach(function (n) { try { api.asc_registerCallback(n, changed); } catch (e) { /* not this editor */ } });
  }

  // ── Cells ──
  function cellToRC(a1) {
    var m = /^([A-Z]{1,3})([1-9][0-9]{0,6})$/.exec(String(a1 || '').toUpperCase());
    if (!m) return null;
    var col = 0;
    for (var i = 0; i < m[1].length; i++) col = col * 26 + (m[1].charCodeAt(i) - 64);
    return { col: col - 1, row: parseInt(m[2], 10) - 1 };
  }
  function rcToCell(col, row) {
    var s = '', n = col + 1;
    while (n > 0) { var r = (n - 1) % 26; s = String.fromCharCode(65 + r) + s; n = Math.floor((n - 1) / 26); }
    return s + (row + 1);
  }
  function sheetIndex(api, name) {
    var all = api.wbModel.aWorksheets;
    if (!name) return all.length === 1 ? 0 : api.wbModel.getActive();
    for (var i = 0; i < all.length; i++) if (all[i].getName() === name) return i;
    return -1;
  }
  function sheetIndexById(api, wsId) {
    var all = api.wbModel.aWorksheets;
    for (var i = 0; i < all.length; i++) if (all[i].getId() === wsId) return i;
    return -1;
  }

  // ── Reading: one shape for both editors ──
  // {id, text, author, time, solved, replies:[{text, author, time}], quote? (Word), sheet?, cell?, guid? (Excel)}
  function num(t) { var n = parseInt(t, 10); return isNaN(n) ? 0 : n; }
  // The host gets real times: an editor's shown time is local time written as UTC (localNow below).
  function real(t) { var n = num(t); return n ? n + new Date(n).getTimezoneOffset() * 60000 : 0; }
  function fromSimple(id, d) {
    return {
      id: id, text: d.Text || '', author: d.UserName || '', time: real(d.Time), solved: !!d.Solved,
      quote: typeof d.QuoteText === 'string' ? d.QuoteText : '',
      replies: (d.Replies || []).map(function (r) { return { text: r.Text || '', author: r.UserName || '', time: real(r.Time) }; }),
    };
  }
  function list(api) {
    if (isCell(api)) {
      var out = [];
      api.wbModel.aWorksheets.forEach(function (ws) {
        ws.aComments.forEach(function (c) {
          if (c.bDocument) return; // a workbook-level comment has no cell; the file keeps none
          out.push({
            id: c.asc_getId(), text: c.sText || '', author: c.sUserName || '', time: num(c.sOOTime) || real(c.sTime), solved: !!c.bSolved,
            sheet: ws.getName(), cell: rcToCell(c.nCol, c.nRow), guid: c.sGuid || '',
            replies: (c.aReplies || []).map(function (r) { return { text: r.sText || '', author: r.sUserName || '', time: num(r.sOOTime) || real(r.sTime) }; }),
          });
        });
      });
      return out;
    }
    return (api.pluginMethod_GetAllComments() || []).map(function (c) { return fromSimple(c.Id, c.Data || {}); });
  }

  function findCellComment(api, id) {
    var all = api.wbModel.aWorksheets;
    for (var i = 0; i < all.length; i++) {
      for (var j = 0; j < all[i].aComments.length; j++) {
        if (all[i].aComments[j].asc_getId() === id) return { index: i, comment: all[i].aComments[j] };
      }
    }
    return null;
  }
  // WHY through the comment's own sheet: asc_removeComment only reaches the sheet in front (and the
  // workbook's comments), so a comment on another sheet stayed while "deleted" was answered.
  function removeCell(api, found) {
    api.wb.getWorksheet(found.index).cellCommentator.removeComment(found.comment.asc_getId());
  }
  function findWord(api, id) {
    var all = api.pluginMethod_GetAllComments() || [];
    for (var i = 0; i < all.length; i++) if (all[i].Id === id) return all[i].Data || {};
    return null;
  }

  // ── Writing ──
  function now() { return String(Date.now()); }
  // WHY two clocks: both editors keep a comment's shown time as the person's LOCAL time written
  // as if it were UTC (web-apps' utcDateToString: now - timezone offset), and the real UTC time
  // apart (asc_putOnlyOfficeTime), which the spreadsheet's save writes as the comment's date. A
  // reply stamped with the real time showed hours off (measured: 1:49 AM beside a 6:49 PM comment
  // in UTC-7), and a workbook comment with no UTC time was saved with no date at all.
  function localNow() { return String(Date.now() - new Date().getTimezoneOffset() * 60000); }

  // Word: select the quoted text, comment on it, then give the person their selection back.
  function addWord(api, op, data) {
    var ld = api.private_GetLogicDocument && api.private_GetLogicDocument();
    if (!ld || typeof ld.Search !== 'function') return { ok: false, error: 'editor-not-ready' };
    var quote = String(op.quote || '');
    if (!quote) return { ok: false, error: 'quote-not-found' };
    var state = ld.GetSelectionState();
    var id = null;
    try {
      var ss = new (editorWindow().AscCommon.CSearchSettings)();
      ss.put_Text(quote);
      ss.put_MatchCase(true);
      var eng = ld.Search(ss);
      var ids = Object.keys(eng.Elements || {});
      if (!ids.length) return { ok: false, error: 'quote-not-found' };
      // The quote's place among its repeats (selector.occurrence); a later edit can leave fewer.
      var pick = ids[Math.min(Math.max(0, op.occurrence | 0), ids.length - 1)];
      ld.SelectSearchElement(+pick);
      id = api.pluginMethod_AddComment(data);
    } finally {
      try { if (typeof ld.ClearSearch === 'function') ld.ClearSearch(); } catch (e) { /* nothing highlighted */ }
      try { ld.SetSelectionState(state); ld.UpdateSelection(); } catch (e) { /* the editor redraws its own selection */ }
    }
    return id ? { ok: true, id: id } : { ok: false, error: 'apply-failed' };
  }

  // Excel: a comment on a given sheet and cell. See the header for documentFlag and _addComment.
  function addCell(api, op, fields) {
    var idx = sheetIndex(api, op.sheet);
    if (idx < 0) return { ok: false, error: 'sheet-not-found' };
    var rc = cellToRC(op.cell);
    if (!rc) return { ok: false, error: 'invalid-cell' };
    var cc = api.wb.getWorksheet(idx).cellCommentator;
    if (cc.getComment(rc.col, rc.row, false, true)) return { ok: false, error: 'cell-has-comment' };
    var w = editorWindow();
    var d = new w.Asc.asc_CCommentData();
    d.ReadFromSimpleObject(fields);
    d.asc_putDocumentFlag(false);
    if (typeof d.asc_putOnlyOfficeTime === 'function') d.asc_putOnlyOfficeTime(now());
    d.wsId = cc.model.getId();
    d.setId();
    d.asc_putCol(rc.col);
    d.asc_putRow(rc.row);
    var added = false;
    cc.isLockedComment(d, function (ok) { if (ok !== false) { cc._addComment(d, false); added = true; } });
    // The GUID is the comment's id in the saved file too (threadedComment id): the host builds the
    // same id the file will give it.
    return added ? { ok: true, id: d.asc_getId(), guid: d.sGuid || '' } : { ok: false, error: 'apply-failed' };
  }

  // Excel: change a comment through its own sheet's commentator (asc_changeComment only reaches
  // the sheet in front). `edit` changes a copy of the comment; replies are rebuilt from the copy.
  function changeCell(api, id, edit) {
    var f = findCellComment(api, id);
    if (!f) return { ok: false, error: 'comment-not-found' };
    var copy = f.comment.clone();
    var r = edit(copy);
    if (r && r.ok === false) return r;
    api.wb.getWorksheet(f.index).cellCommentator.changeComment(id, copy);
    return r || { ok: true };
  }
  function newCellReply(api, text, author) {
    var w = editorWindow();
    var d = new w.Asc.asc_CCommentData();
    d.asc_putText(text);
    d.asc_putUserName(author);
    d.asc_putTime(localNow());
    if (typeof d.asc_putOnlyOfficeTime === 'function') d.asc_putOnlyOfficeTime(now());
    return d;
  }

  // Word: replace a comment's whole data (the plugin method's own way to change one).
  function changeWord(api, id, edit) {
    var d = findWord(api, id);
    if (!d) return { ok: false, error: 'comment-not-found' };
    var copy = JSON.parse(JSON.stringify(d));
    copy.Replies = copy.Replies || [];
    var r = edit(copy);
    if (r && r.ok === false) return r;
    api.pluginMethod_ChangeComment(id, copy);
    return r || { ok: true };
  }
  function badIndex(list, i) { return typeof i !== 'number' || i < 0 || i >= list.length; }

  function apply(api, op) {
    var cell = isCell(api);
    var text = typeof op.text === 'string' ? op.text : '';
    var author = typeof op.author === 'string' && op.author ? op.author : 'Assistant';
    switch (op.kind) {
      case 'list': return { ok: true, comments: list(api) };
      case 'add': {
        // WHY no Time for Word: asc_addComment stamps the editor's own (local-time) convention when
        // none is given, the same as a comment the person adds there.
        var fields = { Text: text, UserName: author };
        if (cell) { fields.Time = localNow(); return addCell(api, op, fields); }
        return addWord(api, op, fields);
      }
      case 'reply':
        if (cell) {
          return changeCell(api, op.id, function (c) {
            c.asc_addReply(newCellReply(api, text, author));
            return { ok: true, index: c.aReplies.length - 1 };
          });
        }
        return changeWord(api, op.id, function (c) {
          c.Replies.push({ Text: text, UserName: author, Time: localNow() });
          return { ok: true, index: c.Replies.length - 1 };
        });
      case 'resolve':
      case 'reopen': {
        var solved = op.kind === 'resolve';
        if (cell) return changeCell(api, op.id, function (c) { c.asc_putSolved(solved); });
        return changeWord(api, op.id, function (c) { c.Solved = solved; });
      }
      case 'edit':
        if (cell) return changeCell(api, op.id, function (c) { c.asc_putText(text); });
        return changeWord(api, op.id, function (c) { c.Text = text; });
      case 'edit-reply':
        if (cell) {
          return changeCell(api, op.id, function (c) {
            if (badIndex(c.aReplies, op.index)) return { ok: false, error: 'reply-not-found' };
            c.aReplies[op.index].asc_putText(text);
          });
        }
        return changeWord(api, op.id, function (c) {
          if (badIndex(c.Replies, op.index)) return { ok: false, error: 'reply-not-found' };
          c.Replies[op.index].Text = text;
        });
      case 'delete-reply':
        if (cell) {
          return changeCell(api, op.id, function (c) {
            if (badIndex(c.aReplies, op.index)) return { ok: false, error: 'reply-not-found' };
            c.aReplies.splice(op.index, 1);
          });
        }
        return changeWord(api, op.id, function (c) {
          if (badIndex(c.Replies, op.index)) return { ok: false, error: 'reply-not-found' };
          c.Replies.splice(op.index, 1);
        });
      case 'delete':
        if (cell) {
          var gone = findCellComment(api, op.id);
          if (!gone) return { ok: false, error: 'comment-not-found' };
          removeCell(api, gone);
          return findCellComment(api, op.id) ? { ok: false, error: 'apply-failed' } : { ok: true };
        }
        if (!findWord(api, op.id)) return { ok: false, error: 'comment-not-found' };
        api.pluginMethod_RemoveComments([op.id]);
        return { ok: true };
      case 'move': {
        // WHY add-then-remove: neither editor moves a comment's anchor through its API, so the thread
        // (text, author, replies, resolved) is made again at the new place and only THEN removed from
        // the old one — a move whose add fails (the quote is not there, the cell has a comment)
        // leaves the thread where it was instead of losing it. Its id changes, as a moved comment's
        // id does in the file too.
        if (cell) {
          var f = findCellComment(api, op.id);
          if (!f) return { ok: false, error: 'comment-not-found' };
          var c = f.comment;
          var moved = {
            Text: c.sText, UserName: c.sUserName, Time: c.sTime || localNow(), Solved: c.bSolved,
            Replies: (c.aReplies || []).map(function (r) { return { Text: r.sText, UserName: r.sUserName, Time: r.sTime }; }),
          };
          var added = addCell(api, op, moved);
          if (added.ok) removeCell(api, f);
          return added;
        }
        var d = findWord(api, op.id);
        if (!d) return { ok: false, error: 'comment-not-found' };
        var copy = addWord(api, op, { Text: d.Text, UserName: d.UserName, Time: d.Time, Solved: d.Solved, Replies: d.Replies || [] });
        if (copy.ok) api.pluginMethod_RemoveComments([op.id]);
        return copy;
      }
    }
    return { ok: false, error: 'unknown-op' };
  }
  function answer(id, result) { window.parent.postMessage({ type: 'yc:office-comments-result', id: id, result: result }, '*'); }

  function run(id, op) {
    if (!op || typeof op.kind !== 'string') { answer(id, { ok: false, error: 'unknown-op' }); return; }
    var key = typeof op.key === 'string' ? op.key : null;
    if (key && op.kind !== 'list' && Object.prototype.hasOwnProperty.call(done, key)) { answer(id, done[key]); return; }
    var api = apiOf(editorWindow());
    // Not drawn yet (or rebuilt): the host keeps the op and asks again. Typing in a cell: an Excel
    // comment added now would end the cell's edit, so it waits too (the host retries).
    if (!api || window.__ycDrawn !== true) { answer(id, { ok: false, error: 'editor-not-ready' }); return; }
    if (isCell(api) && op.kind !== 'list' && typeof api.asc_getCellEditMode === 'function' && api.asc_getCellEditMode()) {
      answer(id, { ok: false, error: 'editor-busy' });
      return;
    }
    listen();
    var result;
    try { result = apply(api, op); } catch (e) { result = { ok: false, error: 'apply-failed' }; }
    if (key && op.kind !== 'list' && result && result.ok) {
      done[key] = result;
      doneKeys.push(key);
      if (doneKeys.length > DONE_MAX) delete done[doneKeys.shift()];
    }
    answer(id, result);
  }

  window.addEventListener('message', function (e) {
    if (e.source !== window.parent) return; // only the host frames us
    var d = e.data || {};
    if (d.type === 'yc:office-comments') run(d.id, d.op);
  });
  // The person's own comments (added in the editor's panel) are news to the host too: listen as
  // soon as the editor exists, and again whenever its frame is rebuilt.
  setInterval(listen, 500);

  // Tests reach the pieces through this (the add-on's own test suite).
  window.__ycComments = { cellToRC: cellToRC, rcToCell: rcToCell };
})();
