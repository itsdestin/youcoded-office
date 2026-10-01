/* yc-bridge.js — runs INSIDE the Office editor page (the AGPL add-on side).
 *
 * WHY this exists: the editors live on their own sealed origin, so YouCoded
 * cannot style them directly. The host posts its theme here; this script turns
 * YouCoded's tokens into OnlyOffice/Euro-Office's own CSS variables and applies
 * them to every same-origin editor frame (web-apps nests its UI in iframes that
 * appear after load). It also hides the editor's own title row — YouCoded's
 * tabs own the document name — and, in "slim" mode, all of the editor's chrome
 * so the host can draw its own one-row toolbar.
 *
 * It also keeps the editor quiet and offline in ways its own settings can't: no "New feature"
 * tips, and no "links to external sources" warning (see quietEditor below).
 *
 * Prototype for the design stage (2026-09-28). Messages:
 *   {type:'yc:office-theme', theme:{tokens, dark, wallpaper, panelsOpacity, panelsBlur, fontLinks}}
 *   {type:'yc:office-mode', slim:boolean}
 */
(function () {
  var STYLE_ID = 'yc-office-theme';

  // ── No unload veto from the editor ──
  // WHY: the host (YouCoded) keeps every edit in a recovery journal, starts the save of unsaved
  // changes as a window closes, and itself lists unsaved files before a quit or the last window's
  // close (its unsaved-files prompt). The editor's own "leave page?"
  // veto (a beforeunload handler that fires while the document is modified) cannot be answered
  // inside YouCoded — it silently cancelled the window close, or forced the host to override
  // every veto, which also dropped the host's own guard for unsaved text-file edits. So in every
  // same-origin editor frame: an onbeforeunload handler is ignored, beforeunload listeners are
  // not registered, and any that got in before this guard cannot veto (preventDefault and
  // returnValue do nothing on a BeforeUnloadEvent here). v0.1.3 also closes the two side doors:
  // body.onbeforeunload (it sets the window's handler natively, past the window property, and a
  // handler's returned string vetoes without touching the event), and
  // EventTarget.prototype.addEventListener.call(window, 'beforeunload', …).
  function guardUnload(win) {
    try {
      if (!win || win.__ycUnloadGuard) return;
      win.__ycUnloadGuard = true;
      try { win.onbeforeunload = null; } catch (e) { /* read-only: the definition below wins */ }
      Object.defineProperty(win, 'onbeforeunload', { configurable: true, get: function () { return null; }, set: function () {} });
      var add = win.addEventListener;
      win.addEventListener = function (type, listener, options) {
        if (String(type).toLowerCase() === 'beforeunload') return;
        return add.call(this, type, listener, options);
      };
      ['HTMLBodyElement', 'HTMLFrameSetElement'].forEach(function (name) {
        var proto = win[name] && win[name].prototype;
        if (proto) Object.defineProperty(proto, 'onbeforeunload', { configurable: true, get: function () { return null; }, set: function () {} });
      });
      var ET = win.EventTarget && win.EventTarget.prototype;
      if (ET && ET.addEventListener) {
        var etAdd = ET.addEventListener;
        ET.addEventListener = function (type, listener, options) {
          if (this === win && String(type).toLowerCase() === 'beforeunload') return;
          return etAdd.call(this, type, listener, options);
        };
      }
      var P = win.BeforeUnloadEvent && win.BeforeUnloadEvent.prototype;
      if (P) {
        Object.defineProperty(P, 'returnValue', { configurable: true, get: function () { return ''; }, set: function () {} });
        P.preventDefault = function () {};
      }
    } catch (e) { /* not same-origin, or already sealed: nothing to guard here */ }
  }
  // ── No peer-to-peer connections (v0.1.6) ──
  // WHY: CSP does not cover WebRTC, so an RTCPeerConnection could send a document's text off the
  // machine. yc-early.js removes it in each editor page before sdkjs runs; this covers the host
  // page itself (same origin, so an editor frame could reach parent.RTCPeerConnection) and the
  // blank frames the editor makes later, on the same walk that guards unload.
  function blockPeers(win) {
    var names = ['RTCPeerConnection', 'webkitRTCPeerConnection'];
    for (var i = 0; i < names.length; i++) {
      var stub = function () { throw new Error('YouCoded Office: peer connections are turned off'); };
      try { Object.defineProperty(win, names[i], { value: stub, writable: false, configurable: false, enumerable: false }); }
      catch (e) { /* already sealed, or not same-origin */ }
    }
  }
  // ── Pictures dragged onto the document ──
  // WHY: sdkjs's desktop drop handler asks the host for the dropped files' PATHS
  // (AscDesktopEditor.GetDropFiles, then IsImageFile) — Euro-Office's bridge.js defines neither, so
  // every drop (a picture, or text from another app) threw and the editor showed "An error
  // occurred". A web page never learns a dropped file's path, and YouCoded's host must never read
  // a path the frame names. So each editor frame's drop is noted first (capture phase, before
  // sdkjs's own handler), and GetDropFiles sends the first dropped picture's BYTES to this
  // document's own origin (upload/, which checks type and size and stores it in the document's
  // pictures) and answers its bare media name. sdkjs then resolves that name through
  // LocalFileGetImageUrl, as for a picture chosen in the file dialog. No picture → [] and sdkjs
  // pastes the drop's text instead.
  var dropped = null, dropWin = null;
  function watchDrops(win) {
    try {
      if (!win || win.__ycDropWatch || !win.addEventListener) return;
      win.__ycDropWatch = true;
      win.addEventListener('drop', function (e) { dropped = e.dataTransfer ? e.dataTransfer.files : null; dropWin = win; }, true);
    } catch (e) { /* not same-origin */ }
  }
  function uploadPicture(f) {
    try {
      var x = new XMLHttpRequest();
      x.open('POST', location.origin + '/upload/drop', false); // sdkjs asks synchronously
      x.setRequestHeader('Content-Type', f.type);
      x.send(f);
      if (x.status !== 200) return null;
      var key = Object.keys(JSON.parse(x.responseText))[0] || '';
      return key ? key.replace(/^media\//, '') : null;
    } catch (e) { return null; }
  }
  // WHY every picture, in two steps (fix round 1): sdkjs's drop handler inserts only the first
  // picture GetDropFiles names (its loop stops there). So the first goes back to sdkjs as usual,
  // and the rest are inserted just after, through the same frame's editor — the same
  // _addImageUrl call sdkjs itself makes, with the addresses its own getImageUrl gives.
  // The editor's own log line (bridge.js's _eoLog → js_log → the app's log; main keeps lines
  // that say "failed").
  function logLine(msg) {
    try {
      if (window._eoLog) window._eoLog(msg);
      else if (window.__TAURI__) window.__TAURI__.core.invoke('js_log', { msg: msg });
    } catch (e) { /* nowhere left to say it */ }
  }
  function droppedPictures() {
    var files = dropped || [], win = dropWin;
    dropped = null; dropWin = null;
    var names = [];
    for (var i = 0; i < files.length; i++) {
      var f = files[i];
      if (!f || !/^image\//.test(f.type || '')) continue;
      var name = uploadPicture(f);
      if (name) names.push(name);
    }
    if (names.length > 1 && win) {
      var rest = names.slice(1);
      setTimeout(function () {
        // WHY checked and logged (fix round 2): these are sdkjs internals; if a build renames them,
        // the extra pictures must show up in the app's log as not inserted, not vanish silently.
        var why = null;
        try {
          var api = (win.Asc && win.Asc.editor) || win.editor;
          var urls = win.AscCommon && win.AscCommon.g_oDocumentUrls;
          if (!api || typeof api._addImageUrl !== 'function') why = 'no _addImageUrl';
          else if (!urls || typeof urls.getImageUrl !== 'function') why = 'no getImageUrl';
          else api._addImageUrl(rest.map(function (n) { return urls.getImageUrl(n); }));
        } catch (e) { why = String((e && e.message) || e); }
        if (why) logLine('[YC] drop: inserting ' + rest.length + ' more dropped picture(s) failed: ' + why);
      }, 0);
    }
    return names.slice(0, 1);
  }
  function extendDesktopEditor(win) {
    var d;
    try { d = win.AscDesktopEditor; } catch (e) { return; }
    if (!d || d.GetDropFiles) return;
    d.GetDropFiles = droppedPictures;
    d.IsImageFile = function (name) { return /\.(png|jpe?g|gif|bmp|svg|webp|ico)$/i.test(String(name || '')); };
  }
  // ── The editor's own Open (v0.1.14, Task 2 fix round 1) ──
  // WHY: Ctrl+O (editor-patches.js, twice) and the editor's Open call LocalFileOpen, which asks for
  // a document in a dialog and then reloads the editor — dropping unsaved edits. Files open from
  // YouCoded's Office start screen; the host never opens files through this (its "open-file" event
  // goes straight to open_file). So Ctrl+O is swallowed in every editor window before the editor's
  // own handlers (capture, on the window), and LocalFileOpen does nothing. main refuses a document
  // dialog too (open_dialog answers only picture requests).
  function isCtrlO(e) { return (e.ctrlKey || e.metaKey) && !e.shiftKey && !e.altKey && (e.key === 'o' || e.key === 'O'); }
  function noEditorOpen(win) {
    try {
      if (win && !win.__ycNoOpenKeys && win.addEventListener) {
        win.__ycNoOpenKeys = true;
        win.addEventListener('keydown', function (e) { if (isCtrlO(e)) { e.preventDefault(); e.stopImmediatePropagation(); } }, true);
      }
    } catch (e) { /* not same-origin */ }
    var d;
    try { d = win.AscDesktopEditor; } catch (e) { return; }
    if (!d || d.__ycNoOpen) return;
    d.__ycNoOpen = true;
    d.LocalFileOpen = function () { logLine('[YC] the editor\'s own Open is not used; files open from the Office start screen'); return Promise.resolve(); };
  }
  // ── Recovered changes (v0.1.24, finish plan Task 8) ──
  // WHY: YouCoded's host keeps every batch of edits the editor sends (save_changes) in a recovery
  // journal, so a crash, a kill, or a window closed before the next save loses nothing. When the
  // host opens a document (its "open-file" event), the host is asked first whether this document
  // has edits its file never got; if so the editor opens the journal's starting point and replays
  // them through its own recovery pipeline — exactly what Euro-Office's start-screen "Recover" does
  // (bridge.js _loadEditorBin → _recoveryEnqueue → _recoveryMarkModified). Otherwise, or if the
  // journal cannot be read, the file opens as before. build/patch.mjs points bridge.js's
  // open-file listener here. Answers what bridge.js puts in window._pendingFileData.
  window.__ycOpenFile = function (invoke, filePath) {
    var name = String(filePath).replace(/\\/g, '/').split('/').pop();
    return Promise.resolve()
      .then(function () { return invoke('recovery_candidates'); })
      .then(function (list) {
        var c = list && list[0];
        if (!c) return null;
        return invoke('recovery_load', { id: c.id }).then(function (r) {
          return { data: r.data, path: filePath, name: name, recovery: { id: r.id, changes: r.changes || [] } };
        });
      })
      .catch(function (e) { logLine('[YC] recovering unsaved changes failed: ' + ((e && e.message) || e)); return null; })
      .then(function (rec) {
        return rec || invoke('open_file', { path: filePath }).then(function (b64) { return { data: b64, path: filePath, name: name }; });
      });
  };
  // ── Edits reach the journal every second (v0.1.24, finish plan Task 8) ──
  // WHY: sdkjs sends a batch of edits only after the person pauses for a second, and at most every
  // two (its autosave: intervalWaitAutoSave, autoSaveGapFast), so a crash in the middle of a long
  // stretch of typing lost all of it. Measured in the YouCoded dev window 2026-09-30: with no wait
  // and a one-second gap, a batch arrives each second while typing, a few hundred bytes each. The
  // batch only sends the edits; it never saves the document (that stays the host's autosave).
  function streamEdits(win) {
    try {
      var api = (win.Asc && win.Asc.editor) || win.editor;
      if (!api || api.__ycStream || typeof api.autoSaveGapFast !== 'number') return;
      api.__ycStream = true;
      api.intervalWaitAutoSave = 0;
      api.autoSaveGapFast = 1000;
    } catch (e) { /* not an editor frame */ }
  }
  // ── A TXT's encoding (v0.1.14, Task 2 fix round 1) ──
  // WHY: Word's Export → TXT asks for an encoding, but x2t writes TXT as UTF-8 whatever it is told
  // (measured 2026-09-29: windows-1252, UTF-16 and ISO-8859-1 all came out UTF-8). A choice that is
  // ignored is never shown: the dialog (encoding only — the CSV one also has a delimiter, which x2t
  // honours, and stays) is hidden by CSS and answered OK here, so the export simply goes ahead.
  function acceptTxtOptions(win) {
    var doc;
    try { doc = win.document; } catch (e) { return; }
    if (!doc || !doc.querySelectorAll) return;
    var dlgs = doc.querySelectorAll('.asc-window.open-dlg');
    // WHY classes set here (v0.1.31): the sheet used to hide this dialog and its mask with
    // `body:has(...)` rules, and a body-wide :has() made every style recalculation in the editor
    // about 23x slower (perf investigation 2026-10-01). The observers below call this the moment a
    // window is added, shown or removed — before the next paint — so nothing flashes.
    var encodingUp = false, delimiterUp = false;
    for (var i = 0; i < dlgs.length; i++) {
      var w = dlgs[i];
      var isTxt = !!w.querySelector('#id-codepages-combo'), isCsv = !!w.querySelector('#id-delimiters-combo');
      if (isTxt) encodingUp = true;
      if (isCsv) delimiterUp = true;
      if (!isTxt || isCsv) continue;
      if (w.classList && !w.classList.contains('yc-txt-dlg')) w.classList.add('yc-txt-dlg');
      if (w.__ycAccepted) continue;
      if (w.style) w.style.visibility = 'hidden';
      // WHY only once it is shown (fix round 2): OK on a dialog not yet shown would close it
      // before its show() — which would then put it (and its click-blocking mask) up for good.
      var shown;
      try { shown = (win.getComputedStyle || getComputedStyle)(w).display !== 'none'; } catch (e) { shown = true; }
      var ok = w.querySelector('[result="ok"]');
      if (!ok || !shown) continue;
      w.__ycAccepted = true;
      ok.click();
    }
    // The mask goes while an encoding dialog is in the page and no delimiter one is — the same
    // test the old body rule made; the editor itself removes the window once it is answered.
    var body = doc.body, maskOff = encodingUp && !delimiterUp;
    if (body && body.classList && body.classList.contains('yc-txt-mask-off') !== maskOff) body.classList.toggle('yc-txt-mask-off', maskOff);
  }
  // WHY a MutationObserver per editor document (fix round 2): the 150 ms pass alone left the hidden
  // dialog's mask over the editor for up to that long, eating clicks. The observer answers it the
  // moment it is added or shown, before the next paint. WHY this narrow (fix round 3): the editor
  // restyles its canvas cursor and rulers constantly; only windows being added (the body's
  // children) and a window's own style/class (how it is shown) matter here.
  function watchTxtOptions(win) {
    var doc;
    try { doc = win.document; } catch (e) { return; }
    if (!doc || !doc.body || doc.__ycTxtWatch) return;
    var MO = win.MutationObserver || (typeof MutationObserver !== 'undefined' ? MutationObserver : null);
    if (!MO) return;
    doc.__ycTxtWatch = true;
    var shown = new MO(function () { acceptTxtOptions(win); });
    var existing = doc.querySelectorAll('.asc-window');
    for (var k = 0; k < existing.length; k++) shown.observe(existing[k], { attributes: true, attributeFilter: ['style', 'class'] });
    new MO(function (records) {
      for (var i = 0; i < records.length; i++) {
        var added = records[i].addedNodes || [];
        for (var j = 0; j < added.length; j++) {
          var n = added[j];
          if (n && n.classList && n.classList.contains('asc-window')) shown.observe(n, { attributes: true, attributeFilter: ['style', 'class'] });
        }
      }
      acceptTxtOptions(win);
    }).observe(doc.body, { childList: true });
  }
  function guardAll(win) {
    streamEdits(win);
    extendDesktopEditor(win);
    noEditorOpen(win);
    watchTxtOptions(win);
    acceptTxtOptions(win);
    watchDrops(win);
    blockPeers(win);
    guardUnload(win);
    quietEditor(win);
    watchDom(win);
    var frames;
    try { frames = win.document.querySelectorAll('iframe'); } catch (e) { return; }
    for (var i = 0; i < frames.length; i++) {
      watchFrameLoad(frames[i]);
      try { if (frames[i].contentWindow) guardAll(frames[i].contentWindow); } catch (e) { /* cross-origin */ }
    }
  }

  // ── Classes in place of :has() (v0.1.31) ──
  // WHY: measured 2026-10-01 (perf investigation): three `body:has(...)` rules in the theme sheet
  // made every style recalculation in the editor ~23x slower (6.2 ms against 0.27 ms on a 69-page
  // document), and the editor recalculates style on nearly every caret move, scroll step and
  // keystroke — it was the main reason scrolling, typing and resizing felt chuggy. Chromium has to
  // re-check a stylesheet :has() rule on every DOM change near it; a one-off querySelectorAll with
  // :has() costs nothing afterwards. So the sheet keys on plain classes, and these selectors (the
  // old rules' own, word for word) are matched from script when the parts they look at are added.
  // Which rows and menu items printCss hides: the printer, colour, sides and copies rows (each
  // with the label row above it), the separators, "Print using the system dialog", and "Print
  // selection" in the right-click menus.
  var PRINT_ROWS = '#id-print-settings tr:has(#print-combo-printer), #id-print-settings tr:has(+ tr #print-combo-printer),' +
    ' #id-print-settings tr:has(#print-combo-color-printing), #id-print-settings tr:has(+ tr #print-combo-color-printing),' +
    ' #id-print-settings tr:has(> td > #print-combo-sides), #id-print-settings tr:has(+ tr > td > #print-combo-sides),' +
    ' #id-print-settings tr:has(> td > .pages #print-txt-copies), #id-print-settings tr:has(> td #print-txt-copies):not(:has(#print-txt-pages)),' +
    ' #id-print-settings tr:has(> td > .separator), #id-print-settings tr:has(#print-btn-system-dialog),' +
    ' .dropdown-menu li:has(> a .menu__icon.btn-print)';
  // Each: the class, the old rule's selector, and what — when added to the page — can change what
  // that selector matches (only that entry is matched again then). The comment list redraws an
  // item when it is resolved or reopened, so the item (or its button) is added again.
  var TAGS = [
    // The presentation's #editor_sdk holds the slide list beside the hole (frameCss).
    ['yc-pe-sdk', '#editor-container > #editor_sdk:has(> #id_main_parent)', '#id_main_parent'],
    ['yc-pe-ct', '#editor-container:has(> #editor_sdk > #id_main_parent)', '#id_main_parent'],
    // Compare and Get data, with their separators (polishCss says why).
    ['yc-hide-group', '.group:has(> #slot-btn-compare), .group:has(> #slot-btn-data-from-text),' +
      ' .group:has(> #slot-btn-interface-theme):not(:has(> #slot-btn-dark-document))',
      '#slot-btn-compare, #slot-btn-data-from-text, #slot-btn-interface-theme'],
    // A resolved thread's text is muted (commentsCss).
    ['yc-resolved', '.user-comment-item:has(.btn-resolve.comment-resolved)', '.user-comment-item, .btn-resolve'],
    // A comment or reply being edited shows "Editing" on its row in place of its icons (v0.1.34).
    ['yc-editing', '.user-comment-item:has(> .inner-edit-ct), .reply-item-ct:has(> .inner-edit-ct)', '.user-comment-item, .inner-edit-ct'],
    ['yc-print-hide', PRINT_ROWS, '#id-print-settings, .menu__icon.btn-print'],
  ];
  var TAG_TRIGGER = TAGS.map(function (t) { return t[2]; }).join(', ') + ', #statusbar';
  // due: which TAGS entries to match again (all of them when left out).
  function tagPass(doc, due) {
    for (var i = 0; i < TAGS.length; i++) {
      if (due && !due[i]) continue;
      var cls = TAGS[i][0], want;
      try { want = doc.querySelectorAll(TAGS[i][1]); } catch (e) { continue; /* no :has() in this engine */ }
      var keep = new Set();
      for (var j = 0; j < want.length; j++) {
        keep.add(want[j]);
        // WHY check first: an unchanged class write still counts as a change to the element.
        if (want[j].classList && !want[j].classList.contains(cls)) want[j].classList.add(cls);
      }
      var had = doc.getElementsByClassName ? [].slice.call(doc.getElementsByClassName(cls)) : [];
      for (var k = 0; k < had.length; k++) if (!keep.has(had[k])) had[k].classList.remove(cls);
    }
  }
  // The status bar turned off (View → Status bar): the editor is laid out 8px shorter (frameCss).
  // The editor hides it with an inline display: none, so its style attribute alone is watched.
  function watchStatusbar(win, doc) {
    var bar = doc.getElementById && doc.getElementById('statusbar');
    var MO = win.MutationObserver || (typeof MutationObserver !== 'undefined' ? MutationObserver : null);
    if (!bar || bar.__ycWatch || !doc.body || !doc.body.classList || !MO) return;
    bar.__ycWatch = true;
    var sync = function () {
      var off = String((bar.getAttribute && bar.getAttribute('style')) || '').indexOf('display: none') >= 0;
      if (doc.body.classList.contains('yc-no-statusbar') === off) return;
      doc.body.classList.toggle('yc-no-statusbar', off);
      // The editor lays its bands out in script; the class lands just after it measured them.
      try { win.dispatchEvent(new win.Event('resize')); } catch (e) { /* not ours */ }
    };
    new MO(sync).observe(bar, { attributes: true, attributeFilter: ['style'] });
    sync();
  }
  // One observer per same-origin document, for what appears after the document is drawn (when
  // the 150 ms pass has stopped, below): parts the classes above key on, frames the editor makes
  // later (each is guarded at once, as the pass did), and scripts it loads later (sdkjs defines
  // its save entry point again when its full build arrives — keepTypingDuringSave re-wraps it).
  function watchDom(win) {
    var doc;
    try { doc = win.document; } catch (e) { return; }
    var MO = win.MutationObserver || (typeof MutationObserver !== 'undefined' ? MutationObserver : null);
    if (!doc || !doc.documentElement || doc.__ycDomWatch || !MO) return;
    doc.__ycDomWatch = true;
    tagPass(doc);
    watchStatusbar(win, doc);
    var onScript = function () { keepTypingDuringSave(win); streamEdits(win); quietEditor(win); };
    // WHY matched in the next animation frame, and only the entries whose parts were added
    // (measured in the perf rig): dragging the window narrower makes the editor move whole toolbar
    // groups into its "More" box many times a second, and matching every entry on each move cost
    // more than the :has() rules had. A frame's callback runs before that frame is painted, so
    // nothing newly added is ever shown untagged; a group that only moved keeps its class.
    var due = null;
    var runDue = function () { var d = due; due = null; if (d) { tagPass(doc, d); watchStatusbar(win, doc); quietEditor(win); } };
    new MO(function (records) {
      var frames = false;
      for (var i = 0; i < records.length; i++) {
        var added = records[i].addedNodes || [];
        for (var j = 0; j < added.length; j++) {
          var n = added[j];
          if (!n || n.nodeType !== 1) continue;
          if (n.tagName === 'SCRIPT') { n.addEventListener('load', onScript); continue; }
          var deep = !!n.firstElementChild;
          if (n.tagName === 'IFRAME' || (deep && n.getElementsByTagName('iframe').length)) frames = true;
          if (!(n.matches(TAG_TRIGGER) || (deep && n.querySelector(TAG_TRIGGER)))) continue;
          var first = !due;
          due = due || {};
          for (var k = 0; k < TAGS.length; k++) if (!due[k] && (n.matches(TAGS[k][2]) || (deep && n.querySelector(TAGS[k][2])))) due[k] = true;
          due.statusbar = true;
          if (first) { if (win.requestAnimationFrame) win.requestAnimationFrame(runDue); else runDue(); }
        }
      }
      if (frames) { guardAll(window); kick(); }
    }).observe(doc.documentElement, { childList: true, subtree: true });
  }
  // A frame that loads (again) brings a new document: guard and theme it, and watch it draw.
  var seen = new WeakSet();
  function watchFrameLoad(f) {
    if (!f || seen.has(f) || !f.addEventListener) return;
    seen.add(f);
    f.addEventListener('load', function () { kick(); schedule(); });
  }
  // ── No "New feature" tips ──
  // WHY: Euro-Office pops "New" onboarding tips over the toolbar (TooltipManager, web-apps
  // SynchronizeTip.js); YouCoded explains its own features, and in the slim editor the tips point
  // at buttons that are hidden. A tip is skipped once its name is set in localStorage
  // (TooltipManager._addTips / _getNeedShow read Common.localStorage.getItem(name), which is plain
  // localStorage here). The editor's origin is per document and fresh, so this runs on every
  // load, BEFORE the editor frame starts. Names: every tip `name` in the built web-apps app.js of
  // euro-office-lite v0.17.21-alpha (plus the source tree's, for other builds of the same tag).
  var SEEN_TIPS = [
    'help-tip-comment-filter', 'help-tip-chart-elements', 'help-tip-redact-tab', 'help-tip-mark-for-redaction',
    'help-tip-apply-redaction',
    'de-help-tip-multipage-view-statusbar', 'de-help-tip-multipage-view-toolbar', 'de-help-tip-header-footer-tab',
    'de-help-tip-signature', 'de-help-tip-fill-status',
    'de-form-tip-create', 'de-form-tip-roles', 'de-form-tip-save', 'de-form-tip-settings',
    'de-form-tip-settings-group', 'de-form-tip-settings-key', 'de-form-tip-submit',
    'sse-help-tip-table-tab', 'sse-help-tip-solver', 'sse-help-tip-cellFormat', 'sse-help-tip-rtl-dir',
    'pe-help-tip-master-tab', 'pe-help-tip-gif-payback',
    'pdfe-help-tip-annot-rect', 'pdfe-help-tip-create-link', 'pdfe-help-tip-pdf-charts',
  ];
  function markTipsSeen(win) {
    try {
      var ls = win.localStorage;
      SEEN_TIPS.forEach(function (n) { if (!ls.getItem(n)) ls.setItem(n, '1'); });
    } catch (e) { /* no storage in this frame: the tips' own CSS hide below still applies */ }
  }
  markTipsSeen(window);

  // ── Keep the editor quiet and offline, per frame, as soon as its code exists ──
  function quietEditor(win) {
    try {
      // A tip added later than the seed above (a build with new tip names): a "New feature" tip
      // is marked seen instead of queued.
      var TM = win.Common && win.Common.UI && win.Common.UI.TooltipManager;
      if (TM && TM.addTips && !TM.__ycQuiet) {
        TM.__ycQuiet = true;
        var addTips = TM.addTips;
        TM.addTips = function (arr) {
          var keep = {};
          for (var k in arr) {
            if (!Object.prototype.hasOwnProperty.call(arr, k)) continue;
            var tip = arr[k];
            if (tip && tip.isNewFeature) { try { tip.name && win.localStorage.setItem(tip.name, '1'); } catch (e) { /* no storage */ } continue; }
            keep[k] = tip;
          }
          return addTips.call(this, keep);
        };
      }
    } catch (e) { /* not same-origin or no tips here */ }
    try {
      // "This workbook contains links to one or more external sources that could be unsafe"
      // (web-apps ExternalLinks.js onNeedUpdateExternalReferenceOnOpen). WHY neither of its
      // buttons: "Update"/"Continue" fetches the linked workbooks (reaching outside the
      // document), and "Turn off AutoUpdate"/"Don't update" with auto-update on calls
      // asc_setUpdateLinks(false, true), which writes a history point, so the file CHANGES and
      // autosave rewrites it. The third option touches neither: sdkjs asks for the dialog from
      // baseEditorsApi.prototype.onNeedUpdateExternalReferenceOnOpen (sdkjs common/apiBase.js,
      // called by the word, cell and slide APIs at open), so that call does nothing here, and
      // WorkbookView.prototype.initExternalReferenceUpdateTimer (cell/view/WorkbookView.js,
      // which re-fetches every link 30 s after open when the workbook's own setting says
      // "always") does nothing either. The cached values stored in the file stay as they are.
      // Data > External links still OFFERS an update, but it cannot reach outside (no network
      // under the CSP). v0.1.5: yc-early.js does this first, inside the editor page; this later
      // walk is only a second line for a page that was not patched.
      var base = win.AscCommon && win.AscCommon.baseEditorsApi;
      if (base && base.prototype && !base.prototype.__ycQuiet) {
        base.prototype.__ycQuiet = true;
        base.prototype.onNeedUpdateExternalReferenceOnOpen = function () {};
      }
      var WV = win.AscCommonExcel && win.AscCommonExcel.WorkbookView;
      if (WV && WV.prototype && !WV.prototype.__ycQuiet) {
        WV.prototype.__ycQuiet = true;
        WV.prototype.initExternalReferenceUpdateTimer = function () {};
      }
    } catch (e) { /* not same-origin or not an editor frame */ }
    try {
      // Print (v0.1.18): the print panel's Print button stays greyed out until it has a printer.
      // The host prints through the operating system's dialog, which lists the real printers, so
      // the panel gets one stand-in (its row is hidden — printCss). WHY here and not only through
      // the relay's get_printers: editor-patches.js asks for printers on Windows only.
      if (!win.__ycPrinter) {
        var app = win.DE || win.SSE || win.PE;
        var ctrl = app && app.getController && app.getController('Print');
        if (ctrl && ctrl.setPrintersInfo) {
          win.__ycPrinter = true;
          ctrl.setPrintersInfo('YouCoded', [{ name: 'YouCoded', color_supported: true, duplex_supported: true }]);
        }
      }
    } catch (e) { /* not an editor frame, or its print panel isn't built yet — the next pass tries */ }
  }

  blockPeers(window);
  guardUnload(window);
  var latest = null;   // last theme posted by the host
  var slim = false;

  function rgba(color, alpha) {
    var m = /^#([0-9a-f]{6})/i.exec(String(color).trim());
    if (!m) return color;
    var n = parseInt(m[1], 16);
    return 'rgba(' + ((n >> 16) & 255) + ',' + ((n >> 8) & 255) + ',' + (n & 255) + ',' + alpha + ')';
  }

  // ── The File tab (backstage): only what works inside YouCoded (v0.1.7) ──
  // WHY each is hidden (audited in all three editors, 2026-09-28): Open / Open Recent / Create
  // new — the Office start screen does these; Close / Exit — the document's tab does it; "Save copy"
  // — it needs a document server's save-as address; "Note lines" (remove_note_separator) — a
  // command YouCoded's host refuses; Version history — YouCoded has its own Versions; Access
  // rights, Help, Suggest a feature — they need a document server or the internet. Suggest is also
  // switched off in the editor's own config (build/patch.mjs); this list is the second line, and
  // covers the ones no config reaches. Kept: Back, Save (the same save Ctrl+S and autosave make),
  // Save As, Download as and Export to PDF (v0.1.12: the host answers dialog.save and
  // save_file_as — each writes a separate file, the document stays on its own), Print (v0.1.18:
  // the host answers print_document with the system print dialog — see printCss), Info (Document
  // info) and Advanced settings.
  var HIDDEN_FILE_ITEMS = [
    'fm-btn-local-open', 'fm-btn-recent', 'fm-btn-create', 'fm-btn-exit', 'fm-btn-back',
    'fm-btn-save-copy',
    'fm-btn-eo-note-separator',
    'fm-btn-history', 'fm-btn-rights', 'fm-btn-help', 'fm-btn-suggest', 'fm-btn-rename',
  ];

  // ── One frame around the document (v0.1.23, finish plan Task 7) ──
  // WHY: Destin, 2026-09-29: "the side bars and such should gracefully connect to the header
  // element to frame the doc/content"; "header and sidebars still dont connect properly to frame
  // the edit area"; and a full-width line under the header in a different colour. Before, the
  // ribbon was its own rounded card with a hairline all round, the side strips and status bar each
  // drew their own hairline, and on a glass theme the ribbon's lower band painted the panel twice
  // (a second shade between the ribbon and the strips). Now the ribbon, the side strips, the open
  // side panels, the formula bar and the status bar are ONE panel surface with no lines inside
  // it, and the document's desk is the one hole cut into it: the theme's large radius on all four
  // corners and a single edge hairline round it — the same shape YouCoded's host gives the whole
  // editor (EditorFrame's card), so the two radii match.
  // The hole: Word's #editor_sdk (rulers and page), the sheet's #editor_sdk (headers and grid —
  // the formula bar stays in the frame above it), the presentation's #id_main_parent (slide and
  // notes — the slide list stays in the frame beside it). Its corners are painted by a layer over
  // it whose spread shadow (the frame's own colour) fills only what lies outside the rounded
  // shape; the hole clips the rest. So the canvases underneath are never resized or re-laid out,
  // and on a wallpaper theme the desk stays see-through.
  var HOLES = '#editor-container > #editor_sdk, .layout-ct.vbox > #editor_sdk, #id_main_parent';
  function frameCss(t, o) {
    var I = ' !important';
    var lg = t['radius-lg'] || '12px';
    // Every part of the frame, painted once (a glass panel painted twice reads as a second shade).
    var bands = '#toolbar, #left-menu, #right-menu, #statusbar, #cell-editing-box, #cell-editing-box + .layout-resizer';
    // Their inner layers, which the editor paints too: see-through, no outline, no card corners.
    var inner = '#toolbar .toolbar, #toolbar .box-controls, #toolbar .box-tabs, #toolbar section.tabs, #toolbar .extra, #statusbar .statusbar,' +
      ' #right-menu .right-panel > .content-box, #right-menu .right-panel .content-box,' +
      ' #left-menu .tool-menu-btns, #left-menu .left-panel, #right-menu .tool-menu-btns, #right-menu .right-panel,' +
      ' #left-panel-search, #left-panel-chat, #left-panel-navigation, #left-panel-thumbnails, #id_panel_thumbnails';
    return bands + ' { background: ' + o.panel + I + '; box-shadow: none' + I + '; }' +
      inner + ' { background: transparent' + I + '; box-shadow: none' + I + '; border-radius: 0' + I + '; }' +
      // The strips' own edge (--border-sidemenu) and the tools row's rounded under-line.
      '#left-menu .tool-menu-btns, #right-menu .tool-menu-btns, #left-menu .left-panel, #right-menu .right-panel { border: 0' + I + '; }' +
      '#toolbar .box-controls::before, #toolbar .box-controls::after { box-shadow: none' + I + '; border: 0' + I + '; background: transparent' + I + '; }' +
      // Fix round 2 (Destin, 2026-10-01: "we should highlight/darken the selected home/file/view/etc
      // tab and make them round on fill/hover etc. same for bottom tabs of spreadsheets"): the
      // ribbon's tabs and the sheet tabs work like YouCoded's own document tabs (ui/DocumentTabs):
      // the open one sits on the inset fill in the full text colour, the others are fg-2 and take
      // the inset fill on hover and the edge fill on press, all with the medium radius. The fill is
      // a layer inside each tab, 3px clear of the row's edges, so the tab's own box (which the
      // editor measures) does not change; the editor's square fill, underline and borders go.
      '#toolbar section.tabs li.ribtab, #statusbar_bottom > li.list-item { position: relative' + I + '; isolation: isolate' + I + '; background: transparent' + I + '; box-shadow: none' + I + '; border: 0' + I + '; }' +
      '#toolbar section.tabs li.ribtab::after { display: none' + I + '; }' +
      '#toolbar section.tabs li.ribtab::before, #statusbar_bottom > li.list-item::before { content: ""' + I + '; position: absolute' + I + '; inset: 3px 1px' + I + ';' +
      ' z-index: -1' + I + '; border-radius: ' + (t['radius-md'] || '8px') + I + '; background: transparent' + I + '; pointer-events: none' + I + '; display: block' + I + '; }' +
      '#toolbar section.tabs li.ribtab > a, #statusbar_bottom > li.list-item > span { color: ' + (t['fg-2'] || t.fg) + I + '; }' +
      '#toolbar section.tabs li.ribtab:hover::before, #statusbar_bottom > li.list-item:hover::before { background: ' + t.inset + I + '; }' +
      '#toolbar section.tabs li.ribtab:active::before, #statusbar_bottom > li.list-item:active::before { background: ' + t.edge + I + '; }' +
      '#toolbar section.tabs li.ribtab.active::before, #statusbar_bottom > li.list-item.active::before { background: ' + t.inset + I + '; }' +
      '#toolbar section.tabs li.ribtab.active > a, #statusbar_bottom > li.list-item.active > span { color: ' + t.fg + I + '; font-weight: 500' + I + '; }' +
      // A sheet tab's own box: no borders, no square fill, no accent bar — unless the person gave the
      // sheet a colour (the editor sets it inline on the tab), which stays.
      '#statusbar_bottom > li.list-item > span { border: 0' + I + '; box-shadow: none' + I + '; }' +
      '#statusbar_bottom > li.list-item > span:not([style*="background"]) { background: transparent' + I + '; }' +
      // Fix round 2 ("strange fill/background boundaries", the ribbon's More button in its own box):
      // the overflow "More" box painted the panel a second time (a second shade on a glass theme) and
      // drew a separator down its left edge. It is part of the frame now.
      '#toolbar .more-box { background: transparent' + I + '; box-shadow: none' + I + '; }' +
      '#toolbar .more-box > .separator { display: none' + I + '; }' +
      '#statusbar, #statusbar .statusbar, #toolbar .toolbar { border: 0' + I + '; }' +
      // The presentation's slide list sits in the frame, beside the hole. WHY classes (v0.1.31): these
      // and every other rule here once used :has(); TAGS sets the class from script instead.
      '#editor-container > #editor_sdk.yc-pe-sdk { background: ' + o.panel + I + '; }' +
      '#editor-container > #editor_sdk.yc-pe-sdk { overflow: visible' + I + '; }' +
      '#editor-container.yc-pe-ct { background: transparent' + I + '; }' +
      // Fix round 3 (Destin, 2026-10-01: "still dark sharp-cornered background bleeding out around the
      // left edge of the slides container"): on a glass (wallpaper) theme the presentation's
      // #editor_sdk painted the frame's see-through panel under everything — under the slide list,
      // whose own canvas paints the panel again, and under the hole, whose corners paint it again —
      // so a darker square-cornered band showed around the hole's left edge and corners, and the
      // slide itself sat on a tinted desk; sdkjs also painted the slide list's canvas and the 4px
      // splitter beside the hole in the panel colour made opaque (a darker block than the glass).
      // Now #editor_sdk paints nothing; the slide list's box and the splitter paint the glass panel
      // once each (the list's canvas no longer fills its background — yc-early.js clearThumbsBack);
      // the hole's corners and its 1px seam on the right are the only other paint. The slide stays
      // on the see-through desk, like Word's page.
      (o.wallpaper ? '#editor-container > #editor_sdk.yc-pe-sdk { background: transparent' + I + '; }' +
        '#editor_sdk > #id_panel_thumbnails, #id_panel_thumbnails_split { background: ' + o.panel + I + '; }' +
        '#id_main_parent { box-shadow: 1px 0 0 ' + o.panel + I + '; }' +
        // Fix round 4 (Destin, 2026-10-01, a wallpaper theme: "still spots in the excel viewer that
        // have the weird darker background"): the sheet's grid canvas is opaque (sdkjs draws it
        // without transparency, so its headers can only be a solid colour), but the scrollbar
        // strips beside it were see-through — so the solid header band stopped square just short
        // of the rounded top-right corner, with a lighter column under the curve, and likewise at
        // the bottom-left. The two scrollbar strips and the square between them now carry the
        // headers' solid panel colour, so the grid has one solid border all round and the hole's
        // rounded corners cut it cleanly.
        '#ws-v-scrollbar, #ws-h-scrollbar, #ws-scrollbar-corner { background-color: ' + t.panel + I + '; }' +
        // Fix round 4 (sweep): with a left panel open (comments, search) the sheet's area starts 4px
        // after the panel, and those 4px were painted by nothing — a dark line down the sheet's left
        // edge on a wallpaper theme. The panel's resizer only shows while a panel is open; the
        // sheet's area (the one item there with no id — Word's and PowerPoint's are
        // #editor-container and have no gap) fills the 4px with the frame then, and only then.
        '#viewport-hbox-layout > .layout-resizer.after:not([style*="display: none"]) ~ .layout-item:not([id]) { box-shadow: -4px 0 0 ' + o.panel + I + '; }' : '') +
      // Fix round 1: the slide area's own gaps (the 4px between the slide and its notes) showed
      // the frame colour through it — a band across the hole. They are the desk's colour now.
      '#id_main_parent { background-color: ' + (o.wallpaper ? 'transparent' : (t.canvas || t.panel)) + I + '; }' +
      '#editor-container > #editor_sdk:not(.yc-pe-sdk), .layout-ct.vbox > #editor_sdk, #id_main_parent { position: relative' + I + '; overflow: hidden' + I + '; }' +
      HOLES.split(', ').map(function (s) { return s + '::after'; }).join(', ') +
      ' { content: ""' + I + '; position: absolute' + I + '; inset: 0' + I + '; z-index: 1000' + I + '; pointer-events: none' + I + ';' +
      ' border-radius: ' + lg + I + '; box-shadow: 0 0 0 ' + lg + ' ' + o.panel + ', inset 0 0 0 1px ' + t.edge + I + '; }' +
      // Fix round 1 (Destin, 2026-10-01: "stray straight lines that poke past rounded corners"):
      // a hole draws no border of its own — the slide area's 1px left border ran straight past
      // the hole's rounded corners, top and bottom. The hole's outline is the layer's alone.
      HOLES + ' { border: 0' + I + '; }' +
      // ...and the presentation's notes divider (an inline 1px border-top across the whole slide
      // area) is drawn inset by the radius at each end, so it never meets the rounded sides; with
      // the notes turned off its 4px stub sits on the hole's bottom edge, and draws no line at all.
      '#id_bottom_pannels_container { border-top-color: transparent' + I + '; background-image: linear-gradient(' + t.edge + ', ' + t.edge + ')' + I + ';' +
      ' background-size: calc(100% - 2 * ' + lg + ') 1px' + I + '; background-position: top center' + I + '; background-repeat: no-repeat' + I + '; }' +
      '#id_bottom_pannels_container[style*="height: 4px"] { background-image: none' + I + '; }' +
      // Fix round 1 (Destin: "the bottom of the inner/outer containers touch each other"): with
      // the status bar turned off (View → Status bar, remembered between documents) nothing of the
      // frame was left under the hole, so its bottom edge sat on YouCoded's card edge. The editor
      // is then laid out 8px shorter — the same gap the frame keeps beside the hole — and a strip
      // of the frame colour fills those 8px. (v0.1.31: the class comes from watchStatusbar, not :has().)
      'body.yc-no-statusbar #viewport { height: calc(100% - 8px)' + I + '; bottom: auto' + I + '; }' +
      'body.yc-no-statusbar::after { content: ""' + I + '; position: fixed' + I + '; left: 0' + I + '; right: 0' + I + '; bottom: 0' + I + '; height: 8px' + I + '; background: ' + o.panel + I + '; pointer-events: none' + I + '; }' +
      // The presentation's #editor_sdk holds the hole and must not be cut itself.
      '#editor-container > #editor_sdk.yc-pe-sdk::after { content: none' + I + '; }' +
      // The selected strip item: one step down the depth ladder, the medium radius, like the File
      // tab's open item and YouCoded's own selected rows.
      '.tool-menu-btns .btn-category { border-radius: ' + (t['radius-md'] || '8px') + I + '; }' +
      '.tool-menu-btns .btn-category.active, .tool-menu-btns .btn-category.active:hover { background-color: ' + t.inset + I + '; }';
  }

  // ── The File tab: the same frame, YouCoded's cards inside (v0.1.23, finish plan Task 7) ──
  // WHY: Destin, 2026-09-29 (P-file): "make this match youcoded's card styling. this still doesn't
  // blend well into or properly separate from the file/home/insert/etc buttons, or the
  // focused/selected side pane. should all be rounded and such while also separating from the
  // youcoded theme frame". So the File tab is laid out like the editor itself: its list sits in the
  // panel frame (no line beside it), right under the ribbon's tab row (the stub of the ribbon's
  // tools that showed between them is covered); the page beside it is the canvas hole, with the
  // large radius and an edge hairline, inset from the frame's edges; what the page shows — the
  // document's info, the settings, the export formats — sits on YouCoded's cards (panel, edge
  // border, large radius, 16px padding) with eyebrow section titles (guide G-7). The open item is
  // one step down the depth ladder with the medium radius, like the strip's selected button.
  // Opaque, even over a wallpaper (v0.1.8): the File tab covers the document, and a see-through
  // panel showed the page's text through its list and settings (Meadow Mist, 2026-09-28).
  function fileTabCss(t) {
    var I = ' !important';
    var md = t['radius-md'] || '8px', lg = t['radius-lg'] || '12px';
    var F = '#file-menu-panel';
    var page = F + ' .panel-context > .content-box';
    return F + ' { top: 28px' + I + '; background-color: ' + t.panel + I + '; }' +
      F + ' .panel-menu { background-color: transparent' + I + '; border-right: 0' + I + '; padding: 8px 8px 12px' + I + '; }' +
      F + ' .panel-menu li.fm-btn { height: 32px' + I + '; padding: 0 12px' + I + '; margin-bottom: 2px' + I + '; border-radius: ' + md + I + '; }' +
      F + ' .panel-menu li.fm-btn > a { font-size: 13px' + I + '; color: ' + t.fg + I + '; }' +
      F + ' .panel-menu li.fm-btn:hover:not(.disabled) { background-color: ' + t.inset + I + '; }' +
      F + ' .panel-menu li.fm-btn.active:not(.disabled) { background-color: ' + t.inset + I + '; box-shadow: none' + I + '; }' +
      F + ' .panel-menu li.fm-btn.active:not(.disabled) > a { font-weight: 600' + I + '; }' +
      F + ' #fm-btn-return { margin-bottom: 12px' + I + '; }' +
      F + ' .panel-context { background-color: transparent' + I + '; }' +
      // The page: the canvas hole, clear of the frame's right and bottom edges by the same 8px the
      // list keeps from its left.
      page + ' { background-color: ' + (t.canvas || t.panel) + I + '; border-radius: ' + lg + I + '; box-shadow: inset 0 0 0 1px ' + t.edge + I + ';' +
      ' margin: 8px 8px 8px 0' + I + '; height: calc(100% - 16px)' + I + '; width: auto' + I + '; right: 0' + I + '; }' +
      F + ' .panel-context .header, ' + F + ' .panel-context h1, ' + F + ' .panel-context .title { color: ' + t.fg + I + '; }' +
      // A page's own title (Advanced settings, Export) above its card; Info's title is the first row
      // of its card, so it reads as the card's title (guide §2.2: 16px medium).
      page + ' .flex-settings > .header, ' + page + ' .content-container > .header { font-size: 18px' + I + '; font-weight: 600' + I + '; color: ' + t.fg + I + '; }' +
      page + ' table.main td.header { font-size: 16px' + I + '; font-weight: 500' + I + '; color: ' + t.fg + I + '; padding-top: 8px' + I + '; }' +
      // The cards: Info's table, Advanced settings' table, and the export formats' tiles.
      page + ' table.main, ' + page + ' .flex-settings > table { background-color: ' + t.panel + I + '; border: 1px solid ' + t.edge + I + '; border-radius: ' + lg + I + ';' +
      ' border-collapse: separate' + I + '; padding: 8px 16px 16px' + I + '; max-width: 760px' + I + '; width: auto' + I + '; }' +
      // Section titles inside a card are eyebrows (guide G-7): uppercase, 11px, muted.
      page + ' table.main td.title label, ' + page + ' .flex-settings td.group-name label {' +
      ' text-transform: uppercase' + I + '; font-size: 11px' + I + '; letter-spacing: 0.06em' + I + '; font-weight: 500' + I + '; color: ' + (t['fg-muted'] || t.fg) + I + '; }' +
      // Each export format is a small card around its file icon (the icon filled the old tile).
      page + ' .format-items .btn-doc-format { background-color: ' + t.panel + I + '; border: 1px solid ' + t.edge + I + '; border-radius: ' + lg + I + ';' +
      ' box-sizing: content-box' + I + '; padding: 12px 14px' + I + '; }' +
      page + ' .format-items .format-item { margin: 0 12px 12px 0' + I + '; }' +
      page + ' .format-items .btn-doc-format:hover { background-color: ' + t.inset + I + '; }' +
      page + ' .divider { background-color: transparent' + I + '; border: 0' + I + '; }';
  }

  // The polish pass (v0.1.7, Destin: "weird rounded pills but also square outlines", "all of the
  // scrollbars are unstyled", "hard to separate some of the side panels and menus … from the
  // document area", "some of the options under the file tab are just odd or don't seem to work").
  function polishCss(t, o) {
    var I = ' !important';
    var css = '';
    // A theme that leaves out a radius gets YouCoded's default for it, never "undefined".
    var sm = t['radius-sm'] || '4px', md = t['radius-md'] || '8px';
    // ── 1. Roundness: galleries are tiles in a grid, not pills ──
    // WHY: the style galleries (Home > cell styles, the document and slide style galleries, table
    // templates) are a grid of tiles whose square 1px borders overlap by a pixel. Rounding the
    // tile (and its preview canvas) drew a pill inside each square outline. The tiles stay
    // square; the gallery's frame is the one rounded shape, with the small radius on its outer
    // corners, and the "more" button closes it on the right.
    css += '.combo-dataview .view .item, .combo-dataview .view .item canvas, .combo-dataview .view .item img,' +
      ' .combo-dataview .dropdown-menu .item, .combo-dataview .dropdown-menu .item canvas, .combo-dataview .dropdown-menu .item img { border-radius: 0' + I + '; }' +
      '.combo-dataview .view { border-radius: ' + sm + ' 0 0 ' + sm + I + '; }' +
      '.combo-dataview .button button { border-radius: 0 ' + sm + ' ' + sm + ' 0' + I + '; }' +
      '.rtl .combo-dataview .view { border-radius: 0 ' + sm + ' ' + sm + ' 0' + I + '; }' +
      '.rtl .combo-dataview .button button { border-radius: ' + sm + ' 0 0 ' + sm + I + '; }';
    // WHY: colour swatches (palettes, the font / highlight colour bars under their buttons) are
    // content, not controls — a tiny radius at most, so their outline and fill stay one shape.
    css += '.color-palette .palette-color-item, .theme-colorpalette .color-item, .dataview .item.color, .btn-color .caret-swatch,' +
      ' .color-preview, .color-transparent { border-radius: ' + o.tile + I + '; }';

    // ── 2. Scrollbars: slim, rounded, the app's thumb colours ──
    // WHY: YouCoded's renderer styles its scrollbars (globals.css: 8px, transparent track,
    // --scrollbar-thumb / --scrollbar-hover, rounded thumb); the editor's panels and menus showed
    // the system's grey bars, and its own "perfect scrollbar" drew a square bordered thumb with
    // grip stripes. Both now match the app. The document and sheet canvases draw theirs in
    // script — yc-early.js makes those slim and rounded, and the --canvas-scroll-* colours
    // above give them the same thumb colours.
    css += '::-webkit-scrollbar { width: 8px; height: 8px; }' +
      '::-webkit-scrollbar-track, ::-webkit-scrollbar-corner { background: transparent; }' +
      // Fix round 2 (Destin: scrollbars "overlap the rounded corners of their containers"): every
      // track stops 6px short of its ends, clear of a rounded corner; the menus' own scrollbars
      // (perfect-scrollbar) likewise, their rail clipped so the thumb never reaches the corner.
      '::-webkit-scrollbar-track { margin: 6px; }' +
      '.ps-container > .ps-scrollbar-y-rail { margin-top: 6px' + I + '; max-height: calc(100% - 12px)' + I + '; overflow: hidden' + I + '; border-radius: 3px' + I + '; }' +
      '.ps-container > .ps-scrollbar-x-rail { margin-left: 6px' + I + '; max-width: calc(100% - 12px)' + I + '; overflow: hidden' + I + '; border-radius: 3px' + I + '; }' +
      '::-webkit-scrollbar-thumb { background: ' + o.thumb + '; border-radius: 4px; }' +
      '::-webkit-scrollbar-thumb:hover { background: ' + o.thumbHover + '; }' +
      '::-webkit-scrollbar-button { display: none; }' +
      '.ps-container .ps-scrollbar-y-rail, .ps-container .ps-scrollbar-x-rail, .ps-container:hover .ps-scrollbar-y-rail,' +
      ' .ps-container:hover .ps-scrollbar-x-rail, .ps-container .ps-scrollbar-y-rail.hover, .ps-container .ps-scrollbar-x-rail.hover,' +
      ' .ps-container.ps-in-scrolling .ps-scrollbar-y-rail, .ps-container.ps-in-scrolling .ps-scrollbar-x-rail,' +
      ' .ps-container .ps-scrollbar-y-rail.in-scrolling, .ps-container .ps-scrollbar-x-rail.in-scrolling { background-color: transparent' + I + '; }' +
      '.ps-container .ps-scrollbar-y-rail { width: 8px' + I + '; }' +
      '.ps-container .ps-scrollbar-x-rail { height: 8px' + I + '; }' +
      '.ps-container .ps-scrollbar-y, .ps-container .ps-scrollbar-y.always-visible-y { width: 6px' + I + '; right: 1px' + I + '; }' +
      '.ps-container .ps-scrollbar-x, .ps-container .ps-scrollbar-x.always-visible-x { height: 6px' + I + '; bottom: 1px' + I + '; }' +
      '.ps-container .ps-scrollbar-y, .ps-container .ps-scrollbar-x, .ps-container .ps-scrollbar-y.always-visible-y,' +
      ' .ps-container .ps-scrollbar-x.always-visible-x { background: ' + o.thumb + I + '; border: 0' + I + '; border-radius: 3px' + I + '; }' +
      '.ps-container .ps-scrollbar-y-rail:hover .ps-scrollbar-y, .ps-container .ps-scrollbar-x-rail:hover .ps-scrollbar-x,' +
      ' .ps-container .ps-scrollbar-y-rail.in-scrolling .ps-scrollbar-y, .ps-container .ps-scrollbar-x-rail.in-scrolling .ps-scrollbar-x,' +
      ' .ps-container .ps-scrollbar-y-rail:hover .ps-scrollbar-y.always-visible-y, .ps-container .ps-scrollbar-x-rail:hover .ps-scrollbar-x.always-visible-x { background: ' + o.thumbHover + I + '; }' +
      '.ps-container .ps-scrollbar-y div, .ps-container .ps-scrollbar-x div { display: none' + I + '; }';

    // ── 3. Separation: canvas behind, panel for bands, edges between regions ──
    // WHY: the ribbon, the side strips, the open side panels, the formula bar, the status bar
    // and the document desk were all near one colour with no line between them. YouCoded's
    // layering (guide §2.1, §2.4): the desk is the canvas; bands and side panes are the panel;
    // fields are inset. Since v0.1.23 the bands are one frame with no lines inside it and the desk
    // is the one outlined hole (frameCss says why). Outlines are shadows, not borders, so the
    // editor's own layout (it measures these boxes in script) does not move.
    css += frameCss(t, o);
    // WHY: a menu or dropdown opened over the ribbon was the ribbon's own colour with a faint
    // line; YouCoded's menus (G-21) are the panel with an edge border and a floating-layer shadow,
    // and their rows round their hover fill.
    css += '.dropdown-menu:not(.internal-menu) { border: 1px solid ' + t.edge + I + '; box-shadow: 0 8px 24px rgba(0,0,0,' + o.shadow + ')' + I + '; }' +
      '.dropdown-menu:not(.internal-menu) > li > a { border-radius: min(' + sm + ', 6px)' + I + '; margin: 0 4px' + I + '; }' +
      '.asc-window { border: 1px solid ' + t.edge + I + '; box-shadow: 0 12px 40px rgba(0,0,0,' + o.shadow + ')' + I + '; }' +
      '.asc-window > .header { border-bottom: 1px solid ' + t.edge + I + '; }';

    // ── 4. The File tab (backstage) ──
    // WHY: items that cannot work here are hidden (HIDDEN_FILE_ITEMS says why for each), with the
    // dividers that only separated them. What remains is styled like YouCoded's own side list:
    // the panel surface, rounded rows, hover one step down the depth ladder, the open item on the
    // inset surface; the page beside it is the canvas, so the list and the page read as two areas.
    css += HIDDEN_FILE_ITEMS.map(function (id) { return '#file-menu-panel #' + id; }).join(', ') + ' { display: none' + I + '; }' +
      '#file-menu-panel .panel-menu li.devider, #file-menu-panel .panel-menu li.devider-small { display: none' + I + '; }' +
      // WHY: Advanced settings stays (its choices apply to the open document), but its
      // "Interface theme" row goes — YouCoded sets the editor's theme from the app's own on every
      // pass, so a pick there snapped straight back.
      '#file-menu-panel tr.themes { display: none' + I + '; }' +
      // Fix round 3 (Task 7): "Tab style" and "Use toolbar color as tabs background" change only how
      // the editor draws its own ribbon tabs, and the frame styling (frameCss) draws those itself —
      // so both did nothing visible. With the Interface theme row already gone they were the whole
      // Appearance group: its title and the divider after it go with them (all three editors).
      '#file-menu-panel tr.appearance, #file-menu-panel tr.tab-style, #file-menu-panel tr.tab-background,' +
      ' #file-menu-panel tr.tab-background + tr.divider-group, #fms-cmb-tab-style, #fms-chb-tab-background { display: none' + I + '; }' +
      // v0.1.14: Word's TXT encoding dialog is answered for the person (acceptTxtOptions says why);
      // hidden from its first frame so it never flashes. The CSV one has a delimiter and stays.
      // v0.1.31: acceptTxtOptions sets both classes the moment the dialog is added (before paint).
      '.asc-window.open-dlg.yc-txt-dlg { visibility: hidden' + I + '; }' +
      // ...and its mask never blocks a click while it is up (fix round 2).
      'body.yc-txt-mask-off .modals-mask { visibility: hidden' + I + '; pointer-events: none' + I + '; }' +
      // v0.1.15 (fix round 2): features that open a file of a kind YouCoded's host never hands the
      // editor (main answers only picture dialogs) are hidden, not left to do nothing. Word: Insert
      // → Text from file; Collaboration → Compare and Combine (a second document); Mail merge
      // (recipients from a spreadsheet). Spreadsheet: Data → Get data (TXT/CSV, XML) and External
      // links' Change source, Open source and Update values (fix round 3, checked in the dev window
      // on a workbook linked to another: Change source asks for another workbook — the same dialog
      // chart settings → Edit links opens in Word and PowerPoint; Open source asks a document
      // server to open it and window.open is denied; Update values ends in "Error: updating is
      // failed". Break links works — it removes the link and the file saves — so it and the
      // dialog stay). Presentation: Insert → Audio / Video (the
      // host does not offer media, so the editor never builds them; hidden in case a build does).
      // Every editor: the hyperlink dialog's "Select file" button. Each goes with its separator.
      // v0.1.31 (Destin, 2026-10-01: "remove some of these theme related buttons that won't work
      // anymore"): View → Interface Theme. YouCoded sets the editor's theme from the app's own on
      // every theme change, so a pick there was overruled (the same reason File → Advanced
      // settings' theme row goes, above). Its group goes too where it holds nothing else (sheet,
      // slides); Word's keeps Dark Document, which only darkens the page and works with any dark
      // YouCoded theme (checked in the perf rig: readable, kept across dark themes, locked by the
      // editor itself on light ones).
      '#slot-btn-interface-theme,' +
      ' #slot-btn-text-from-file, #slot-btn-mailrecepients, #id-right-menu-mail-merge, #slot-btn-insaudio, #slot-btn-insvideo,' +
      ' .group.yc-hide-group, .group.yc-hide-group + .separator,' +
      ' #external-links-btn-change, #external-links-btn-open, #external-links-btn-update, #chart-button-update-data,' +
      ' #id-dlg-hyperlink-url .select-button { display: none' + I + '; }' +
      // Fix round 4: Word/PowerPoint chart settings — "Update data" (#chart-button-update-data,
      // hidden above) takes the same failing update path as the dialog's Update values; the linked
      // source's name stays readable, but as plain text: its link opens the source through a
      // document server this host doesn't have.
      '#chart-open-external-link { pointer-events: none' + I + '; cursor: default' + I + '; color: inherit' + I + '; text-decoration: none' + I + '; border-bottom: none' + I + '; }';
    css += fileTabCss(t);
    return css + printCss() + commentsCss(t);
  }

  // ── The comments panel IS YouCoded's comment panel (v0.1.21; matched control by control v0.1.34) ──
  // WHY: Destin asked for Office's comments to look and work exactly like the app's own (desktop
  // renderer components/comments/), and in v0.1.33 they still did not ("notice the reply/close
  // buttons particularly … the checkmark has weird bright spots … the edit/delete buttons aren't the
  // same icons … and they have a blue tint"). Every rule below copies the app's own classes, read
  // from its source and measured in the workbench side by side with the reading view's panel:
  //   pane      CommentsPaneFrame: a rounded box (radius-xl) with an edge border on the panel
  //             surface, inset 8px from the strip; header 33px, padding 6/6/6/12, "Comments" 14px
  //             semibold, an edge line under it; × = CloseButton icon-sm (20px ghost, 12px glyph)
  //   card      CommentCard: inset, 1px edge-dim border, radius-lg, 12px padding, 12px text
  //   avatar    Avatar: 20px circle, inset, edge-dim border, ONE 10px fg-2 initial (the editor
  //             writes two: "PS") — never the editor's per-author colour
  //   name row  name 12px medium fg, then the time 11px fg-muted on the same line; the note
  //             below it in the column beside the avatar, 12px fg-2 (fg-muted once resolved)
  //   actions   CommentActions: Edit and Delete are the app's own pencil and bin (stroke 2, 14px
  //             glyph in an 18px button), fg-faint → fg-2 on hover, shown while the card is
  //             hovered; Resolve is CompleteToggle's circle check (16px, stroke 1.8), fg-faint →
  //             fg-2, filled in the accent with a knocked-out check once resolved. WHY `filter:
  //             none` and no ::before/::after: the editor's dark skin inverts its sprite icons
  //             (filter: invert(1) — that turned the faint grey blue) and draws the resolve tick
  //             as a bordered ::after (the bright spots).
  //   reply box CommentComposer: the field surface (inset, edge-dim border, radius-lg, accent on
  //             focus), 11px text with "Reply…" in fg-muted, and the 16px round accent send arrow
  //             INSIDE the field, dimmed while there is nothing to send. Enter sends, Escape
  //             closes (yc-comments.js) — so the editor's own Close button goes.
  //   editing   InlineEditField: the text in the same field, Cancel (ghost) then Save (primary),
  //             small buttons, in the field's own footer; the row says "Editing ✎" in place of its
  //             icons (CommentRowActions' EditingPill).
  // The quoted text goes, as it did from the app's cards (Destin, round 11). No :has() here (perf,
  // v0.1.31): yc-resolved / yc-editing are set from script (TAGS).
  function svgUrl(svg) { return 'url("data:image/svg+xml,' + encodeURIComponent(svg) + '")'; }
  function strokeIcon(paths, width) {
    return svgUrl('<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="black" stroke-width="' + width + '" stroke-linecap="round" stroke-linejoin="round">' + paths + '</svg>');
  }
  // The app's own glyphs, path for path (CommentActions.tsx EditGlyph / DeleteGlyph, CloseButton,
  // CompleteToggle in SessionCardDetails.tsx, the composer's send arrow).
  var EDIT_PATHS = '<path d="M12 20h9"/><path d="M16.5 3.5a2.12 2.12 0 0 1 3 3L7 19l-4 1 1-4Z"/>';
  var ICON_EDIT = strokeIcon(EDIT_PATHS, 2);
  function coloredIcon(paths, color) {
    return svgUrl('<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="' + color + '" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">' + paths + '</svg>');
  }
  var ICON_DELETE = strokeIcon('<path d="M3 6h18"/><path d="M8 6V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"/><path d="M19 6l-1 14a2 2 0 0 1-2 2H8a2 2 0 0 1-2-2L5 6"/>', 2);
  var ICON_CLOSE = strokeIcon('<path d="M6 18L18 6M6 6l12 12"/>', 2);
  var ICON_MORE = strokeIcon('<circle cx="12" cy="12" r="1"/><circle cx="19" cy="12" r="1"/><circle cx="5" cy="12" r="1"/>', 2);
  var ICON_ADD = strokeIcon('<path d="M21 15a2 2 0 0 1-2 2H7l-4 4V5a2 2 0 0 1 2-2h14a2 2 0 0 1 2 2z"/><path d="M12 7v6M9 10h6"/>', 2);
  var ICON_RESOLVE = strokeIcon('<circle cx="12" cy="12" r="9"/><path d="M8 12.5l2.5 2.5L16 9.5"/>', 1.8);
  function resolvedIcon(fill, knock) {
    return svgUrl('<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="9" fill="' + fill + '" stroke="' + fill + '"/><path d="M8 12.5l2.5 2.5L16 9.5" stroke="' + knock + '"/></svg>');
  }
  function sendIcon(color) {
    return svgUrl('<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="' + color + '" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><path d="M5 12h14M12 5l7 7-7 7"/></svg>');
  }
  function commentsCss(t) {
    var I = ' !important';
    var font = t['font-sans'] ? 'font-family: ' + t['font-sans'] + I + ';' : '';
    var lg = t['radius-lg'] || '12px', xl = t['radius-xl'] || '16px', sm = t['radius-sm'] || '4px';
    var fg2 = t['fg-2'] || t.fg, dimFg = t['fg-dim'] || fg2, muted = t['fg-muted'] || fg2, faint = t['fg-faint'] || muted, dim = t['edge-dim'] || t.edge;
    var card = '.user-comment-item', box = '#comments-box';
    // An icon drawn in the text colour through a mask: no sprite, no filter, no pseudo-elements.
    var masked = function (sel, mask, size, idle, hover) {
      return sel + ' { background: ' + idle + I + '; -webkit-mask: ' + mask + ' center / ' + size + ' ' + size + ' no-repeat' + I + '; mask: ' + mask + ' center / ' + size + ' ' + size + ' no-repeat' + I + '; filter: none' + I + '; box-shadow: none' + I + '; border: 0' + I + '; cursor: pointer' + I + '; }' +
        sel + '::before, ' + sel + '::after { display: none' + I + '; content: none' + I + '; }' +
        (hover ? sel + ':hover { background: ' + hover + I + '; }' : '');
    };
    // The app's small buttons (Button size sm): 11px medium, 4px 10px, radius-lg.
    var smallBtn = 'min-width: 0' + I + '; width: auto' + I + '; height: auto' + I + '; padding: 4px 10px' + I + '; ' + font + ' font-size: 11px' + I + '; font-weight: 500' + I + '; line-height: 16px' + I + '; border-radius: ' + lg + I + '; border: 0' + I + '; box-shadow: none' + I + '; filter: none' + I + ';';
    var primary = smallBtn + ' background-color: ' + t.accent + I + '; color: ' + t['on-accent'] + I + ';';
    var ghost = smallBtn + ' background-color: transparent' + I + '; color: ' + dimFg + I + ';';
    // A button whose own words are not the app's ("OK", "Close"): the app's word drawn instead.
    var relabel = function (sel, word) {
      return sel + ' { font-size: 0' + I + '; } ' + sel + '::after { content: "' + word + '"; font-size: 11px; }';
    };
    var field = 'background-color: ' + t.inset + I + '; border: 1px solid ' + dim + I + '; border-radius: ' + lg + I + '; box-shadow: none' + I + ';';
    var fieldText = font + ' font-size: 11px' + I + '; line-height: 1.375' + I + '; color: ' + t.fg + I + '; background: transparent' + I + '; border: 0' + I + '; outline: none' + I + '; box-shadow: none' + I + '; resize: none' + I + '; padding: 6px 10px' + I + ';';
    var css = '';

    // ── The pane ──
    // WHY no box of its own (v0.1.36, the framing sweep): the panel sits in the editor's frame, like
    // the paragraph and cell settings panels beside it, and takes the frame's surface — see-through
    // on a wallpaper theme. Its own rounded panel-coloured box read as a card inside a card there;
    // only the comment cards are cards, as in the app's reading-view panel.
    css += '#left-panel-comments, ' + box + ', ' + box + ' .messages-ct, ' + box + ' .dataview-ct, ' + box + ' .new-comment-ct { background: transparent' + I + '; }' +
      '#left-panel-comments { padding: 0' + I + '; }' +
      box + ' { border: 0' + I + '; border-radius: 0' + I + '; box-shadow: none' + I + '; }' +
      '#comments-header { display: flex' + I + '; align-items: center' + I + '; gap: 2px' + I + '; height: 33px' + I + '; padding: 6px 6px 6px 12px' + I + '; box-sizing: border-box' + I + '; border-bottom: 1px solid ' + t.edge + I + '; background: transparent' + I + '; }' +
      '#comments-header label { flex: 1' + I + '; order: 0' + I + '; margin: 0' + I + '; ' + font + ' font-size: 14px' + I + '; font-weight: 600' + I + '; line-height: 20px' + I + '; color: ' + t.fg + I + '; }' +
      '#comments-header > div { float: none' + I + '; margin: 0' + I + '; }' +
      '#comments-header #comments-btn-close { order: 9' + I + '; }' +
      // The header's buttons are the app's CloseButton (ghost, icon-sm): 20px, radius-lg, fg-dim,
      // the inset fill and fg on hover; their glyphs take the button's colour (no editor tint).
      '#comments-header .btn { width: 20px' + I + '; height: 20px' + I + '; min-width: 0' + I + '; padding: 0' + I + '; border: 0' + I + '; border-radius: ' + lg + I + '; background: transparent' + I + '; color: ' + dimFg + I + '; box-shadow: none' + I + '; display: inline-flex' + I + '; align-items: center' + I + '; justify-content: center' + I + '; }' +
      '#comments-header .btn:hover, #comments-header .btn.active, #comments-header .btn:active { background-color: ' + t.inset + I + '; color: ' + t.fg + I + '; }' +
      // Each header glyph is drawn like the app's icons — a stroke in the button's own colour — so no
      // editor sprite (the add-comment one is blue) and no editor filter shows: × is CloseButton's,
      // "…" and "add" are the same stroke family (lucide's more-horizontal, message-square-plus).
      '#comments-header .btn svg, #comments-header .btn .caption, #comments-header .btn .inner-box-caret { display: none' + I + '; }' +
      '#comments-header .btn::after { content: ""' + I + '; display: block' + I + '; width: 12px' + I + '; height: 12px' + I + '; background: currentColor' + I + '; filter: none' + I + '; }' +
      '#comments-btn-close .btn::after { -webkit-mask: ' + ICON_CLOSE + ' center / 12px 12px no-repeat' + I + '; mask: ' + ICON_CLOSE + ' center / 12px 12px no-repeat' + I + '; }' +
      '#comments-btn-sort .btn::after { width: 14px' + I + '; height: 14px' + I + '; -webkit-mask: ' + ICON_MORE + ' center / 14px 14px no-repeat' + I + '; mask: ' + ICON_MORE + ' center / 14px 14px no-repeat' + I + '; }' +
      '#comments-btn-add .btn::after { width: 14px' + I + '; height: 14px' + I + '; -webkit-mask: ' + ICON_ADD + ' center / 14px 14px no-repeat' + I + '; mask: ' + ICON_ADD + ' center / 14px 14px no-repeat' + I + '; }' +
      box + ' .dataview-ct .item { padding: 4px 8px' + I + '; background: transparent' + I + '; border: 0' + I + '; box-shadow: none' + I + '; }' +
      box + ' .dataview-ct .item:first-child { padding-top: 8px' + I + '; }';

    // ── The card ──
    css += card + ' { ' + font + ' position: relative' + I + '; background-color: ' + t.inset + I + '; border: 1px solid ' + dim + I + '; border-radius: ' + lg + I + '; padding: 12px' + I + '; font-size: 12px' + I + '; line-height: 16px' + I + '; color: ' + fg2 + I + '; box-shadow: none' + I + '; }' +
      // The thread under the pointer, or the one the person is on, gets the app's ring (CommentsMargin:
      // ring-2 ring-accent/60) — the same cue that lights its highlight in the document.
      box + ' .dataview-ct .item:hover ' + card + ', ' + box + ' .dataview-ct .item.selected ' + card + ' { box-shadow: 0 0 0 2px ' + rgba(t.accent, 0.6) + I + '; }' +
      card + ' .user-info { display: flex' + I + '; align-items: flex-start' + I + '; gap: 8px' + I + '; height: auto' + I + '; margin: 0' + I + '; padding-right: 64px' + I + '; }' +
      card + ' .reply-item-ct .user-info { padding-right: 44px' + I + '; }' +
      card + ' .user-info .color { flex: none' + I + '; display: block' + I + '; width: 20px' + I + '; height: 20px' + I + '; margin: 0' + I + '; box-sizing: border-box' + I + '; line-height: 18px' + I + '; border-radius: 50%' + I + '; background-color: ' + t.inset + I + '; border: 1px solid ' + dim + I + '; color: ' + fg2 + I + '; font-size: 0' + I + '; font-weight: 500' + I + '; text-align: center' + I + '; overflow: hidden' + I + '; }' +
      card + ' .user-info .color::first-letter { font-size: 10px; }' +
      card + ' .user-info-text { display: flex' + I + '; flex-direction: row' + I + '; flex: 1 1 auto' + I + '; width: auto' + I + '; align-items: baseline' + I + '; gap: 6px' + I + '; min-width: 0' + I + '; margin: 0' + I + '; padding: 0' + I + '; line-height: 16px' + I + '; }' +
      card + ' .user-name { ' + font + ' color: ' + t.fg + I + '; font-weight: 500' + I + '; font-size: 12px' + I + '; line-height: 16px' + I + '; overflow: hidden' + I + '; text-overflow: ellipsis' + I + '; white-space: nowrap' + I + '; padding: 0' + I + '; flex: 0 0 auto' + I + '; max-width: 75%' + I + '; }' +
      // The editor's date is long ("9/23/26, 10:00 AM") where the app's is "1d ago": it gives way
      // before the name does.
      card + ' .user-date { ' + font + ' color: ' + muted + I + '; font-size: 11px' + I + '; line-height: 16px' + I + '; white-space: nowrap' + I + '; overflow: hidden' + I + '; text-overflow: ellipsis' + I + '; flex: 0 100 auto' + I + '; min-width: 0' + I + '; padding: 0' + I + '; }' +
      // The note sits in the column beside the avatar, just under the name (the avatar is 20px,
      // the name line 16px: the app's mt-0.5 lands it 2px under the name).
      card + ' .user-message { ' + font + ' color: ' + fg2 + I + '; font-size: 12px' + I + '; line-height: 16px' + I + '; margin: -2px 0 0 28px' + I + '; padding: 0' + I + '; white-space: pre-wrap' + I + '; }' +
      card + '.yc-resolved .user-message { color: ' + muted + I + '; }' +
      card + ' .user-quote, ' + card + ' .reply-arrow { display: none' + I + '; }' +
      // A workbook's comment shows its cell on a muted line under the name, as the app's card
      // does ("C4"; CommentCard's cellRef) — the editor keeps the cell in the quote's place.
      '.yc-sheet ' + card + ' > .user-quote { display: block' + I + '; order: 0' + I + '; margin: -4px 0 0 28px' + I + '; padding: 0' + I + '; border: 0' + I + '; ' + font + ' font-size: 11px' + I + '; line-height: 16px' + I + '; font-style: normal' + I + '; color: ' + muted + I + '; white-space: nowrap' + I + '; overflow: hidden' + I + '; text-overflow: ellipsis' + I + '; }' +
      '.yc-sheet ' + card + ' > .user-quote + .user-message { margin-top: 2px' + I + '; }' +
      card + ' .reply-item-ct { margin: 8px 0 0' + I + '; padding: 0 0 0 4px' + I + '; position: relative' + I + '; }';

    // ── Edit, delete, resolve ──
    css += card + ' .edit-ct { position: absolute' + I + '; top: 12px' + I + '; right: 12px' + I + '; display: flex' + I + '; align-items: center' + I + '; gap: 2px' + I + '; height: 18px' + I + '; margin: 0' + I + '; }' +
      card + ' .btns-reply-ct { position: absolute' + I + '; top: 0' + I + '; right: 0' + I + '; display: flex' + I + '; gap: 2px' + I + '; margin: 0' + I + '; }' +
      card + ' .edit-ct > div, ' + card + ' .btns-reply-ct > div { float: none' + I + '; margin: 0' + I + '; padding: 0' + I + '; border-radius: ' + sm + I + '; }' +
      card + ' .btn-edit-common, ' + card + ' .btn-delete { width: 18px' + I + '; height: 18px' + I + '; opacity: 0' + I + '; }' +
      card + ':hover .edit-ct .btn-edit-common, ' + card + ':hover .edit-ct .btn-delete, ' + card + ' .reply-item-ct:hover .btns-reply-ct > div,' +
      ' ' + card + ':focus-within .edit-ct > div { opacity: 1' + I + '; }' +
      masked(card + ' .btn-edit-common', ICON_EDIT, '14px', faint, fg2) +
      masked(card + ' .btn-delete', ICON_DELETE, '14px', faint, fg2) +
      card + ' .btn-resolve { width: 16px' + I + '; height: 16px' + I + '; margin-left: 2px' + I + '; }' +
      masked(card + ' .btn-resolve:not(.comment-resolved)', ICON_RESOLVE, '16px', faint, fg2) +
      card + ' .btn-resolve.comment-resolved { background: ' + resolvedIcon(t.accent, t.canvas || t.panel) + ' center / 16px 16px no-repeat' + I + '; -webkit-mask: none' + I + '; mask: none' + I + '; filter: none' + I + '; box-shadow: none' + I + '; border: 0' + I + '; cursor: pointer' + I + '; }' +
      card + ' .btn-resolve.comment-resolved::before, ' + card + ' .btn-resolve.comment-resolved::after { display: none' + I + '; content: none' + I + '; }';

    // ── The reply box: "Reply…" waiting, then the composer with its arrow inside the field ──
    css += card + ' .user-reply { display: flex' + I + '; align-items: center' + I + '; justify-content: space-between' + I + '; height: auto' + I + '; min-height: 29px' + I + '; box-sizing: border-box' + I + '; margin: 8px 0 0' + I + '; padding: 0 4px 0 10px' + I + '; ' + field + ' font-size: 0' + I + '; text-decoration: none' + I + '; cursor: text' + I + '; }' +
      card + ' .user-reply::before { content: "Reply…"; ' + font + ' font-size: 11px; color: ' + muted + '; }' +
      card + ' .user-reply::after { content: ""; width: 16px; height: 16px; border-radius: 50%; opacity: 0.5; background: ' + t.accent + ' ' + sendIcon(t['on-accent']) + ' center / 10px 10px no-repeat; }' +
      card + ' .reply-ct, ' + box + ' .new-comment-ct .inner-ct { display: flex' + I + '; align-items: flex-end' + I + '; margin: 8px 0 0' + I + '; padding: 0 4px 0 0' + I + '; ' + field + ' }' +
      card + ' .reply-ct:focus-within, ' + card + ' .inner-edit-ct:focus-within { border-color: ' + t.accent + I + '; }' +
      card + ' .reply-ct textarea { flex: 1' + I + '; min-width: 0' + I + '; width: auto' + I + '; height: 27px' + I + '; min-height: 27px' + I + '; ' + fieldText + ' }' +
      card + ' textarea::placeholder, ' + box + ' textarea::placeholder { color: ' + muted + I + '; }' +
      card + ' .reply-ct .btn-reply { flex: none' + I + '; width: 16px' + I + '; height: 16px' + I + '; min-width: 0' + I + '; padding: 0' + I + '; margin: 0 0 5px 4px' + I + '; border: 0' + I + '; border-radius: 50%' + I + '; font-size: 0' + I + '; box-shadow: none' + I + '; background: ' + t.accent + ' ' + sendIcon(t['on-accent']) + ' center / 10px 10px no-repeat' + I + '; }' +
      card + ' .reply-ct .btn-reply.disabled, ' + card + ' .reply-ct .btn-reply:disabled { opacity: 0.5' + I + '; }' +
      card + ' .reply-ct .btn-close { display: none' + I + '; }';

    // ── Editing a comment or reply: the field, its own footer, and "Editing ✎" on the row ──
    css += card + ' .inner-edit-ct { display: flex' + I + '; flex-wrap: wrap' + I + '; justify-content: flex-end' + I + '; gap: 6px' + I + '; margin: 4px 0 0 28px' + I + '; padding: 0 6px 6px 0' + I + '; ' + field + ' }' +
      card + ' .inner-edit-ct textarea { flex: 1 1 100%' + I + '; width: 100%' + I + '; min-height: 3.2em' + I + '; ' + fieldText + ' }' +
      card + ' .inner-edit-ct .btn-inner-edit { order: 2' + I + '; ' + primary + ' }' +
      card + ' .inner-edit-ct .btn-inner-close { order: 1' + I + '; ' + ghost + ' }' +
      card + ' .inner-edit-ct .btn-inner-close:hover { color: ' + t.fg + I + '; background-color: ' + t.inset + I + '; }' +
      relabel(card + ' .inner-edit-ct .btn-inner-edit', 'Save') + relabel(card + ' .inner-edit-ct .btn-inner-close', 'Cancel') +
      card + '.yc-editing > .edit-ct, .reply-item-ct.yc-editing > .btns-reply-ct { display: none' + I + '; }' +
      card + '.yc-editing > .user-info::after, .reply-item-ct.yc-editing > .user-info::after { content: "Editing"; margin-left: auto; padding-right: 14px; ' + font + ' font-size: 11px; line-height: 16px; color: ' + muted + '; background: ' + coloredIcon(EDIT_PATHS, muted) + ' right center / 10px 10px no-repeat; }' +
      card + '.yc-editing > .user-info, .reply-item-ct.yc-editing > .user-info { padding-right: 0' + I + '; }';

    // ── The new-comment box at the panel's foot: the same field and buttons ──
    css += box + ' .new-comment-ct { background: transparent' + I + '; padding: 0 8px 8px' + I + '; }' +
      box + ' .new-comment-ct textarea { flex: 1' + I + '; ' + fieldText + ' }' +
      box + ' .new-comment-ct .btn.add { ' + primary + ' margin: 6px 0 0 6px' + I + '; }' +
      box + ' .new-comment-ct .btn.cancel { ' + ghost + ' margin: 6px 0 0' + I + '; }';
    return css;
  }

  // ── Print (v0.1.18, finish plan Task 3) ──
  // WHY: YouCoded's host prints through the operating system's own print dialog, which chooses the
  // printer, copies, two-sided and colour itself. The editor's print panel stays for what only it
  // can do — the preview, the page setup (applied to the document itself) and which pages or
  // sheets — and loses the rows the system dialog owns, so nothing on it is silently ignored:
  // printer, colour, copies, sides, and "Print using the system dialog" (the Print button is that
  // now). "Print selection" (the right-click menus) and a workbook's "Selection" range go too: the
  // host prints from the saved document, which has no selection. Quick print ("print on the last
  // printer without asking") goes: every print here asks.
  function printCss() {
    var I = ' !important';
    // The rows are found by PRINT_ROWS and carry .yc-print-hide, set from script (v0.1.31, see TAGS).
    return '#id-print-settings tr.yc-print-hide, #print-combo-range li[data-value="2"],' +
      ' .dropdown-menu li.yc-print-hide, #slot-btn-dt-print-quick, .btn-quick-print { display: none' + I + '; }';
  }

  // What the print panel chose, for the host (bridge.js's Print, patched in build/patch.mjs). A
  // document or presentation sends the editor's own options (a page list in nativeOptions); a
  // workbook's choices live in the editor window's AscDesktopEditor_PrintOptions (sdkjs puts them
  // there, not in the options text), so they are read out in the shape the PDF export sends.
  window.__ycPrintJson = function (ew, optionsJson) {
    try {
      if (window.AscDesktopEditor && window.AscDesktopEditor._currentDocType === 'cell') {
        var po = ew && ew.AscDesktopEditor_PrintOptions;
        // One print's choices only: Ctrl+P prints without the panel, and must not reuse them.
        if (ew) ew.AscDesktopEditor_PrintOptions = null;
        var ad = po && po.advancedOptions;
        // No panel (Ctrl+P): the whole workbook, as every other editor prints the whole document.
        // WHY said outright: without a print type sdkjs prints only the active sheet (measured).
        if (!ad) return JSON.stringify({ adjustOptions: { printType: 1 }, spreadsheetLayout: { ignorePrintArea: false } });
        var get = function (name, field) { return typeof ad[name] === 'function' ? ad[name]() : ad[field]; };
        return JSON.stringify({
          adjustOptions: { printType: get('asc_getPrintType', 'printType'), startPageIndex: get('asc_getStartPageIndex', 'startPageIndex'), endPageIndex: get('asc_getEndPageIndex', 'endPageIndex'), activeSheetsArray: get('asc_getActiveSheetsArray', 'activeSheetsArray') },
          spreadsheetLayout: { ignorePrintArea: !!get('asc_getIgnorePrintArea', 'ignorePrintArea') },
        });
      }
    } catch (e) { return ''; }
    return typeof optionsJson === 'string' ? optionsJson : '';
  };

  function buildCss(th) {
    var t = th.tokens;
    var glass = th.wallpaper ? Math.max(0.35, Math.min(1, th.panelsOpacity || 0.6)) : 1;
    var panel = rgba(t.panel, glass);
    // WHY a fallback: a host older than v0.1.7's app change does not send the scrollbar tokens.
    var thumb = t['scrollbar-thumb'] || t.edge, thumbHover = t['scrollbar-hover'] || t['fg-faint'] || t.edge;
    // Small radii for things that must stay nearly square (tiles, swatches, checkboxes): the
    // theme's small radius, but never more than a few pixels — a 16px "small" radius (Strawberry
    // Kitty) on a 12px swatch drew a circle inside a square outline.
    var tile = 'min(' + (t['radius-sm'] || '4px') + ', 3px)';
    var shadow = th.dark ? '0.45' : '0.16';
    var v = {
      // Header and toolbar bands: the panel surface, never OnlyOffice's per-app colour.
      '--toolbar-header-document': panel, '--toolbar-header-spreadsheet': panel,
      '--toolbar-header-presentation': panel, '--toolbar-header-pdf': panel, '--toolbar-header-visio': panel,
      '--text-toolbar-header': t.fg, '--text-toolbar-header-on-background-document': t.fg,
      '--text-toolbar-header-on-background-spreadsheet': t.fg, '--text-toolbar-header-on-background-presentation': t.fg,
      '--background-toolbar': panel, '--background-toolbar-additional': panel,
      '--background-normal': t.panel, '--background-pane': panel,
      '--background-primary-dialog-button': t.accent, '--background-accent-button': t.accent, '--text-inverse': t['on-accent'],
      '--background-notification-popover': t.inset,
      // Hover one step down the depth ladder, press one more (design guide §2.4).
      '--highlight-button-hover': t.inset, '--highlight-button-pressed': t.edge, '--highlight-button-pressed-hover': t.edge,
      '--highlight-header-button-hover': t.inset, '--highlight-header-button-pressed': t.edge,
      // Fix round 2: the open ribbon tab is shown by its fill, like YouCoded's tabs — no underline.
      '--highlight-toolbar-tab-underline': 'transparent', '--highlight-toolbar-tab-underline-document': 'transparent',
      '--highlight-toolbar-tab-underline-spreadsheet': 'transparent', '--highlight-toolbar-tab-underline-presentation': 'transparent',
      '--text-normal': t.fg, '--text-normal-pressed': t.fg, '--text-secondary': t['fg-dim'], '--text-tertiary': t['fg-muted'],
      '--text-link': t.link || t.accent, '--text-contrast-background': t.fg,
      '--icon-normal': t.fg, '--icon-normal-pressed': t.fg, '--icon-toolbar-header': t.fg,
      // WHY (v0.1.4): this build's toolbar icons are SVG symbols stroked with currentColor, and
      // `svg.icon { color: var(--icon-gray-primary, #383838) }` — no editor theme defines that
      // variable, so every theme drew the icons near-black, invisible on a dark band (Task 6's
      // dev run). The icon's main stroke follows the theme's text colour and its secondary fill
      // the inset surface, like the rest of YouCoded's icons.
      '--icon-gray-primary': t.fg, '--icon-gray-secondary': t.inset,
      '--border-toolbar': t.edge, '--border-divider': t.edge, '--border-regular-control': t.edge,
      // WHY a full border (v0.1.7): the editor reads --border-sidemenu as a whole `border` value
      // (`var(--scaled-one-pixel) solid var(--border-toolbar)`); a bare colour made it invalid,
      // so the side strips had no edge at all against the document.
      '--border-sidemenu': '1px solid ' + t.edge, '--border-toolbar-active-panel-top': panel, '--border-control-focus': t.accent,
      '--background-fill-input': t.inset, '--border-fill-input': t.edge,
      '--border-preview-hover': t['fg-faint'] || t.edge, '--border-preview-select': t.accent,
      '--canvas-background': th.wallpaper ? 'transparent' : t.canvas,
      '--canvas-content-background': '#fff', '--canvas-page-border': t.edge,
      '--canvas-ruler-background': panel, '--canvas-ruler-border': t.edge, '--canvas-ruler-margins-background': t.inset,
      '--canvas-high-contrast': t.fg,
      // Scrollbars drawn in the document and sheet canvases: the app's own thumb and hover
      // colours, no outline (the outline takes the thumb's own colour), no grip stripes.
      '--canvas-scroll-thumb': thumb, '--canvas-scroll-thumb-hover': thumbHover, '--canvas-scroll-thumb-pressed': thumbHover,
      '--canvas-scroll-thumb-border': thumb, '--canvas-scroll-thumb-border-hover': thumbHover, '--canvas-scroll-thumb-border-pressed': thumbHover,
      '--canvas-scroll-thumb-target': thumb, '--canvas-scroll-thumb-target-hover': thumbHover, '--canvas-scroll-thumb-target-pressed': thumbHover,
      '--canvas-scroll-arrow': t['fg-muted'], '--canvas-scroll-arrow-hover': t.fg, '--canvas-scroll-arrow-pressed': t.fg,
      '--canvas-background-tabs': panel,
      // The sheet's row and column headers: a band of the panel surface, one step down on hover,
      // the edge colour where selected — the same ladder as every other control.
      '--canvas-cell-title-background': t.panel, '--canvas-cell-title-background-hover': t.inset,
      '--canvas-cell-title-background-selected': t.edge, '--canvas-cell-title-border': t.edge,
      '--canvas-cell-title-border-hover': t.edge, '--canvas-cell-title-border-selected': t.edge,
      '--canvas-cell-title-text': t['fg-dim'] || t.fg,
      // Roundness by role (guide G-3), v0.1.7: controls follow the theme — toolbar icons and
      // fields the small radius, buttons and menus the medium, the ribbon and dialogs the large.
      // Content tiles (gallery items, swatches) stay almost square: see `tile` above and the
      // gallery rules below.
      '--border-radius-button-normal': t['radius-md'], '--border-radius-button-base': t['radius-sm'],
      '--border-radius-button-toolbar': t['radius-sm'], '--border-radius-button-category': t['radius-md'],
      '--border-radius-toolbar': t['radius-lg'], '--border-radius-form-control': t['radius-sm'],
      '--border-radius-dropdown-menu': t['radius-md'], '--border-radius-dataview-item': tile,
      '--border-radius-window': t['radius-lg'], '--border-radius-checkbox': 'min(' + (t['radius-sm'] || '4px') + ', 4px)',
      '--font-family-base': t['font-sans'], '--font-family-base-custom': t['font-sans'],
    };
    var decl = '';
    for (var k in v) if (v[k]) decl += k + ':' + v[k] + ' !important;';
    var css = ':root, body, body[class] {' + decl + '}' +
      // YouCoded's tabs carry the file name and the save state, so the editor's title row goes.
      '#app-title { display: none !important; }' +
      // Onboarding tips pop over the document; YouCoded explains its own features.
      '.tooltip.new-feature, .synch-tip, .asc-synchronizetip { display: none !important; }' +
      polishCss(t, { panel: panel, thumb: thumb, thumbHover: thumbHover, tile: tile, shadow: shadow, wallpaper: th.wallpaper });
    if (th.wallpaper) {
      css += 'html, body, #viewport, .layout-region, #editor_sdk, #id_main, #ws-canvas-outer, .ws-canvas-area { background-color: transparent !important; }' +
        // v0.1.23: the bands themselves are painted once each by frameCss — nothing more here, so
        // no inner layer doubles the glass.

        // WHY on a layer behind each panel, not the panel (v0.1.9): a backdrop-filter makes its
        // element the frame of every position:fixed menu inside it, so the right panel's menus
        // (the slide background's "Select picture", among others) opened ~1400px to the right,
        // off screen. Each of these four is already positioned (relative/absolute), so the layer
        // fills it exactly; z-index -1 keeps it under the panel's content. The static
        // .right-panel/.statusbar sit inside #right-menu/#statusbar and share their layer.
        (th.panelsBlur ? '#toolbar::before, #statusbar::before, #left-menu::before, #right-menu::before { content: ""; position: absolute; inset: 0; z-index: -1; pointer-events: none; border-radius: inherit; backdrop-filter: blur(' + th.panelsBlur + 'px); }' : '');
    }
    if (slim) {
      css += '#toolbar, #statusbar, .statusbar, #left-menu, #right-menu, .right-panel, .left-panel { display: none !important; }';
    }
    return css;
  }

  // ── Canvas colours reach the drawing code too (v0.1.7) ──
  // WHY: the document and sheet canvases (their scrollbars, the sheet's row and column headers,
  // the page outline) are drawn by sdkjs from colours it copies out of the CSS variables only
  // when the editor starts or switches between its light and dark themes. A switch between two
  // dark YouCoded themes (Midnight → Halftone) changed neither, so the canvases kept the old
  // theme's colours. So each frame's editor is handed exactly these colours whenever they change
  // (asc_setSkin with no name or type keeps the editor's own theme and only replaces them).
  var SKIN_KEYS = [
    'canvas-page-border', 'canvas-scroll-thumb', 'canvas-scroll-thumb-hover', 'canvas-scroll-thumb-pressed',
    'canvas-scroll-thumb-border', 'canvas-scroll-thumb-border-hover', 'canvas-scroll-thumb-border-pressed',
    'canvas-scroll-thumb-target', 'canvas-scroll-thumb-target-hover', 'canvas-scroll-thumb-target-pressed',
    'canvas-scroll-arrow', 'canvas-scroll-arrow-hover', 'canvas-scroll-arrow-pressed',
    'canvas-cell-title-background', 'canvas-cell-title-background-hover', 'canvas-cell-title-background-selected',
    'canvas-cell-title-border', 'canvas-cell-title-border-hover', 'canvas-cell-title-border-selected', 'canvas-cell-title-text',
  ];
  function skinFor(th) {
    var t = th.tokens, thumb = t['scrollbar-thumb'] || t.edge, hover = t['scrollbar-hover'] || t['fg-faint'] || t.edge;
    var skin = {
      'canvas-page-border': t.edge,
      'canvas-scroll-thumb': thumb, 'canvas-scroll-thumb-hover': hover, 'canvas-scroll-thumb-pressed': hover,
      'canvas-scroll-thumb-border': thumb, 'canvas-scroll-thumb-border-hover': hover, 'canvas-scroll-thumb-border-pressed': hover,
      'canvas-scroll-thumb-target': thumb, 'canvas-scroll-thumb-target-hover': hover, 'canvas-scroll-thumb-target-pressed': hover,
      'canvas-scroll-arrow': t['fg-muted'], 'canvas-scroll-arrow-hover': t.fg, 'canvas-scroll-arrow-pressed': t.fg,
      'canvas-cell-title-background': t.panel, 'canvas-cell-title-background-hover': t.inset,
      'canvas-cell-title-background-selected': t.edge, 'canvas-cell-title-border': t.edge,
      'canvas-cell-title-border-hover': t.edge, 'canvas-cell-title-border-selected': t.edge,
      'canvas-cell-title-text': t['fg-dim'] || t.fg,
    };
    // The track behind the thumb is the desk; a wallpaper theme's see-through desk has no colour
    // to hand over, so the editor keeps its own there.
    if (!th.wallpaper && t.canvas) skin['canvas-background'] = t.canvas;
    var out = {};
    SKIN_KEYS.concat(['canvas-background']).forEach(function (k) { if (skin[k]) out[k] = skin[k]; });
    return out;
  }
  function pushSkin(win) {
    var api;
    try { api = (win.Asc && win.Asc.editor) || win.editor; } catch (e) { return; }
    if (!api || typeof api.asc_setSkin !== 'function') return;
    // WHY only when the colours differ (v0.1.7 dedupe, now off the cache): each call redraws the
    // canvases. It stays on every colour change, not only light/dark flips — a switch between two
    // dark themes changes these colours too, and the editor would keep the old ones (v0.1.7).
    var th = themeFor();
    if (win.__ycSkin === th.skinKey) return;
    try { api.asc_setSkin(th.skin); win.__ycSkin = th.skinKey; } catch (e) { /* editor still starting: next pass */ }
  }

  function applyTo(win) {
    var doc;
    try { doc = win.document; } catch (e) { return; } // not same-origin: not ours
    if (!doc || !doc.head) return;
    // A workbook's comment names its cell where a document's quotes text (commentsCss, v0.1.34).
    try { if (doc.body && /spreadsheeteditor/.test(String(win.location)) && !doc.body.classList.contains('yc-sheet')) doc.body.classList.add('yc-sheet'); } catch (e) { /* not ours */ }
    // The theme's own web font (Meadow Mist's Nunito), loaded in this frame too. WHY only the
    // editor's own origin's /yc-fonts/css route (v0.1.4): the CSP keeps the editor offline, so a
    // Google link would be blocked; YouCoded's main process fetches Google's font hosts for it and
    // serves the stylesheet and files here, on this document's own office://<token> origin.
    (latest.fontLinks || []).forEach(function (href) {
      if (typeof href !== 'string' || href.indexOf(location.origin + '/yc-fonts/css?') !== 0) return;
      if (doc.querySelector('link[data-yc-font="' + href + '"]')) return;
      var l = doc.createElement('link'); l.rel = 'stylesheet'; l.href = href; l.setAttribute('data-yc-font', href);
      doc.head.appendChild(l);
    });
    // Light or dark editor base, so its icon set matches the band it sits on.
    // WHY every pass, not once: setTheme silently ignores a theme id until the
    // editor has registered its themes, which happens after this frame appears
    // (found 2026-09-28: Halftone kept dark icons on a dark band).
    try {
      var Themes = win.Common && win.Common.UI && win.Common.UI.Themes;
      var want = latest.dark ? 'theme-dark' : 'theme-light';
      if (Themes && Themes.currentThemeId && Themes.currentThemeId() !== want) Themes.setTheme(want);
    } catch (e) { /* no theme API in this frame: the variables still apply */ }
    // Escape with nothing of the editor's own open (no menu, no dialog) goes to
    // the app, like a YouCoded page's Escape — the host closes its top layer.
    if (!doc.__ycEsc) {
      doc.__ycEsc = true;
      doc.addEventListener('keydown', function (ev) {
        if (ev.key !== 'Escape' || ev.defaultPrevented) return;
        var busy = [].some.call(doc.querySelectorAll('.dropdown-menu, .asc-window, .modals-mask, .open > .dropdown-menu'), function (el) {
          return el.offsetParent !== null && getComputedStyle(el).visibility !== 'hidden';
        });
        if (!busy) window.parent.postMessage({ type: 'yc:office-esc' }, '*');
      }, true);
    }
    pushSkin(win);
    // Fix round 2: the canvas scrollbars (yc-early.js drawSlim) keep their thumb clear of the
    // document area's rounded corners; they need the theme's large radius to know how far.
    try { win.__ycScrollInset = (parseFloat(latest.tokens && latest.tokens['radius-lg']) || 12); } catch (e) { /* not ours */ }
    // Fix round 3: on a glass theme the slide list's background canvas is cleared, not filled
    // (yc-early.js clearThumbsBack says why); the frame paints the panel under it once.
    try { win.__ycGlassFrame = !!latest.wallpaper; } catch (e) { /* not ours */ }
    var style = doc.getElementById(STYLE_ID);
    if (!style) { style = doc.createElement('style'); style.id = STYLE_ID; doc.head.appendChild(style); }
    var th = themeFor();
    if (style.__ycCss === th.css && style.textContent === th.css) return;
    style.textContent = th.css;
    style.__ycCss = th.css;
    // OnlyOffice positions its bands in script; a resize makes it lay out again without the rows
    // we just hid. WHY only when what is hidden or the font changed (v0.1.31): the resize makes
    // the editor re-lay out and redraw its whole canvas — measured 2026-10-01, most of a theme
    // switch's 100-300 ms freeze. A colour-only change (one dark theme to another) moves nothing.
    if (win.__ycLayout !== th.layout) {
      win.__ycLayout = th.layout;
      win.dispatchEvent(new win.Event('resize'));
    }
  }
  // The theme sheet and canvas colours, built once per theme and mode (v0.1.31). WHY: the 26 KB
  // sheet was rebuilt for every frame on every pass just to find it unchanged. `layout` is what
  // needs the editor to lay out and redraw: slim mode hides its bands, the font changes their text
  // widths, and a wallpaper turns the desk see-through (measured in the perf rig: without a redraw
  // a sheet kept its old scrollbar colours after a switch to a wallpaper theme).
  var themeCache = null;
  function themeFor() {
    if (!themeCache) {
      var skin = skinFor(latest);
      themeCache = {
        css: buildCss(latest), skin: skin, skinKey: JSON.stringify(skin),
        layout: (slim ? 'slim' : 'full') + '|' + ((latest.tokens && latest.tokens['font-sans']) || '') + '|' + !!latest.wallpaper,
      };
    }
    return themeCache;
  }

  // ── Typing keeps working while a save runs (v0.1.20) ──
  // WHY: sdkjs's desktop save (DesktopOfflineAppDocumentStartSave) starts a "block interaction"
  // long action and only ends it in DesktopOfflineAppDocumentEndSave, which bridge.js calls after
  // YouCoded has translated the whole document and written the file. web-apps turns the keyboard
  // off for that whole time, so whatever the person typed meanwhile was thrown away — measured in
  // the YouCoded dev window 2026-09-30: 2-4 characters lost per save on 5-20 MB documents, and four
  // whole cells on a 20 MB workbook (its save holds the block ~5 s). The document's bytes are taken
  // synchronously inside StartSave (bridge.js LocalFileSave reads asc_nativeGetFile before its first
  // await), so the block has done its job the moment StartSave returns: it is lifted there, and the
  // editor's own end of that save, later, is swallowed so it cannot end some other long action.
  // Typing after the bytes were taken is newer than the save, so the editor stays "modified" and the
  // host's follow-up save writes it. A Save As keeps its block (its dialog is up), as before.
  function keepTypingDuringSave(win) {
    try {
      var start = win.DesktopOfflineAppDocumentStartSave;
      var api = (win.Asc && win.Asc.editor) || win.editor;
      var A = win.Asc;
      if (typeof start !== 'function' || start.__yc || !api || typeof api.sync_EndAction !== 'function' || !A || !A.c_oAscAsyncActionType || !A.c_oAscAsyncAction) return;
      var BLOCK = A.c_oAscAsyncActionType.BlockInteraction, SAVE = A.c_oAscAsyncAction.Save;
      // WHY the state lives on the window, not in this call: the editor can define StartSave again
      // after its API exists (measured in the spreadsheet editor), so a later pass re-wraps
      // StartSave while the API keeps the wrappers of the first pass; both must see one state.
      // lifted: this save's block was lifted early, and its own end is still to come.
      // started: the StartSave under way really started its block.
      var st = win.__ycSaveState || (win.__ycSaveState = { lifted: false, started: false });
      // sdkjs's StartSave/EndSave call the global `editor`; patch it and Asc.editor if they differ.
      var apis = [api];
      if (win.editor && win.editor !== api) apis.push(win.editor);
      apis.forEach(function (a) {
        if (typeof a.sync_EndAction !== 'function' || a.sync_EndAction.__yc) return;
        var end = a.sync_EndAction;
        var startAction = a.sync_StartAction;
        a.sync_StartAction = function (type, id) {
          if (type === BLOCK && id === SAVE) st.started = true;
          return startAction.apply(this, arguments);
        };
        a.sync_EndAction = function (type, id) {
          if (st.lifted && type === BLOCK && id === SAVE) { st.lifted = false; return; }
          return end.apply(this, arguments);
        };
        a.sync_EndAction.__yc = true;
      });
      var patched = function (isSaveAs) {
        // A previous save that never reached its end (it gave up) must not swallow this one's.
        st.lifted = false;
        st.started = false;
        var r = start.apply(this, arguments);
        // Only a block this call really started (StartSave's encryption branch starts none).
        if (isSaveAs !== true && st.started) { api.sync_EndAction(BLOCK, SAVE); st.lifted = true; }
        st.started = false;
        return r;
      };
      patched.__yc = true;
      win.DesktopOfflineAppDocumentStartSave = patched;
    } catch (e) { /* not this editor's API: saves keep the editor's own behaviour */ }
  }

  function walk(win) {
    blockPeers(win);
    guardUnload(win);
    quietEditor(win);
    keepTypingDuringSave(win);
    if (latest) applyTo(win);
    var frames;
    try { frames = win.document.querySelectorAll('iframe'); } catch (e) { return; }
    for (var i = 0; i < frames.length; i++) {
      var f = frames[i];
      watchFrameLoad(f);
      try { if (f.contentWindow) walk(f.contentWindow); } catch (e) { /* cross-origin */ }
    }
  }

  var pending = false;
  function schedule() {
    if (pending) return;
    pending = true;
    requestAnimationFrame(function () { pending = false; walk(window); });
  }

  // ── Slim mode's toolbar: YouCoded draws the buttons, the editor does the work ──
  // The host's one-row bar sends a command name; this presses the editor's own
  // (hidden) toolbar button of the same name, so every file type gets exactly
  // the editor's behaviour. The same ids exist in the document, spreadsheet and
  // presentation editors (probed 2026-09-28), except alignment in slides, which
  // is a menu there — the host leaves those buttons out for presentations.
  var COMMANDS = ['undo', 'redo', 'bold', 'italic', 'underline', 'markers', 'numbering', 'align-left', 'align-center', 'align-right'];
  function editorDoc(win) {
    var frames;
    try { frames = win.document.querySelectorAll('iframe'); } catch (e) { return null; }
    for (var i = 0; i < frames.length; i++) {
      var w;
      try { w = frames[i].contentWindow; if (/web-apps\/apps/.test(String(w.location))) return w.document; } catch (e) { continue; }
      var inner = editorDoc(w);
      if (inner) return inner;
    }
    return null;
  }
  function button(doc, cmd) {
    var el = doc && doc.getElementById('id-toolbar-btn-' + cmd);
    if (!el) return null;
    return el.tagName === 'BUTTON' ? el : el.querySelector('button') || el;
  }
  function editorApi() {
    var doc = editorDoc(window);
    var w = doc && doc.defaultView;
    try { return w && ((w.Asc && w.Asc.editor) || w.editor) || null; } catch (e) { return null; }
  }
  // yc-comments.js (v0.1.21) reaches the editor through the same walk.
  window.__ycEditorWindow = function () { var doc = editorDoc(window); return (doc && doc.defaultView) || null; };
  function save() {
    var api = editorApi();
    // The frame may have been rebuilt since the last walk: make sure this save lets typing through.
    var edoc = editorDoc(window);
    if (edoc && edoc.defaultView) keepTypingDuringSave(edoc.defaultView);
    // WHY the modified check (v0.1.12): a Save As writes a separate file but still clears the
    // editor's "modified" flag, and asc_Save then skips the save — the document's own file would
    // silently miss the edits. The host asks only when it knows edits are unsaved, so an editor
    // that says "nothing changed" goes straight to LocalFileSave, which saves the current content.
    var unchanged = api && typeof api.isDocumentModified === 'function' && !api.isDocumentModified();
    if (api && typeof api.asc_Save === 'function' && !unchanged) { api.asc_Save(false); return; }
    if (window.AscDesktopEditor) window.AscDesktopEditor.LocalFileSave('', '', null, 0, null);
  }
  // ── Edits now, the window is closing (v0.1.25, finish plan Task 8 fix round 1) ──
  // WHY: edits reach the host's recovery journal about once a second (streamEdits); a window closed
  // inside that second would drop the last of them. The host asks just before the close; sdkjs's own
  // autosave step is run at once with its timer cleared (it sends the edits — changes only, never a
  // save of the document), and the answer goes to the host AFTER them, so it knows they arrived.
  function sendEditsNow() {
    try {
      var api = editorApi();
      if (api && typeof api._autoSave === 'function') {
        api.lastSaveTime = new Date(0);
        api._autoSave();
      }
    } catch (e) { logLine('[YC] sending edits before close failed: ' + ((e && e.message) || e)); }
    window.parent.postMessage({ type: 'yc:office-journaled' }, '*');
  }
  function run(cmd) {
    // v0.1.21: open the editor's comments panel (left strip). Only opens — its button toggles, and
    // the host may ask twice. Used by YouCoded's photographs of the restyled panel.
    if (cmd === 'comments') {
      var doc = editorDoc(window);
      var b = doc && doc.getElementById('left-btn-comments');
      if (b && !/\bactive\b/.test(b.className)) b.click();
      return;
    }
    if (COMMANDS.indexOf(cmd) < 0) return;
    var b = button(editorDoc(window), cmd);
    if (b) b.click();
  }
  // Report which commands are on (bold at the caret) and which can run, so the
  // host's buttons light and dim like the editor's own.
  var lastState = '';
  function reportState() {
    var doc = editorDoc(window);
    if (!doc) return;
    var state = {};
    COMMANDS.forEach(function (c) {
      var b = button(doc, c);
      if (!b) return;
      state[c] = { on: /\bactive\b/.test(b.className), enabled: !/\bdisabled\b/.test(b.className) && !b.disabled };
    });
    var json = JSON.stringify(state);
    if (json !== lastState) { lastState = json; window.parent.postMessage({ type: 'yc:office-state', state: state }, '*'); }
  }

  // ── Rulers: none in the slim editor ──
  // WHY: the approved slim design (office-review-3 runs3/after) has no rulers; the editor draws
  // them in its own canvas, so CSS can't hide them. The document and presentation editors read
  // 'de-hidden-rulers' / 'pe-hidden-rulers' from localStorage when they start (web-apps
  // Main.js), so this is set before the editor frame exists (the host posts the mode before it
  // asks for the file). Set both ways, because the origin is this document's own: the full
  // editor keeps its rulers. An editor already running is told directly (asc_SetViewRulers).
  function setRulers() {
    try {
      window.localStorage.setItem('de-hidden-rulers', slim ? '1' : '0');
      window.localStorage.setItem('pe-hidden-rulers', slim ? '1' : '0');
    } catch (e) { /* no storage: the running-editor call below still applies */ }
    var api = editorApi();
    try { if (api && api.asc_SetViewRulers) api.asc_SetViewRulers(!slim); } catch (e) { /* not this editor */ }
  }

  // ── Remember the person's editor settings (v0.1.19) ──
  // WHY: the editor keeps File → Advanced settings and its view toggles in localStorage, which on
  // this document's one-time origin is gone when the document closes. The editor frame shares
  // this page's origin, so each of its localStorage writes arrives here as a 'storage' event
  // (this page's own writes — rulers, tips — do not). Keys that look like an editor setting are
  // collected for a moment and sent to the host, which keeps only its allow-list (desktop
  // editor-settings.ts) and hands them to the next document's yc-early.js. Nothing is sent for
  // anything else the page stores.
  var SETTING = /^(?:de|sse|pe)-|^app-settings-/;
  var changedSettings = null, settingsTimer = 0;
  function sendSettings() {
    settingsTimer = 0;
    var batch = changedSettings; changedSettings = null;
    if (!batch) return;
    try {
      var p = window.__TAURI__ && window.__TAURI__.core.invoke('save_editor_settings', { settings: batch });
      if (p && p.catch) p.catch(function () { /* not remembered: the choice still holds in this document */ });
    } catch (e) { /* no relay */ }
  }
  window.addEventListener('storage', function (e) {
    try {
      if (!e || typeof e.key !== 'string' || !SETTING.test(e.key) || e.storageArea !== window.localStorage) return;
      (changedSettings = changedSettings || {})[e.key] = e.newValue; // null: put back to the default
      if (!settingsTimer) settingsTimer = setTimeout(sendSettings, 300);
    } catch (err) { /* keep the editor running */ }
  });

  window.addEventListener('message', function (e) {
    if (e.source !== window.parent) return; // only the host frames us
    var d = e.data || {};
    if (d.type === 'yc:office-theme' && d.theme && d.theme.tokens) { latest = d.theme; themeCache = null; schedule(); }
    if (d.type === 'yc:office-mode') { slim = !!d.slim; themeCache = null; setRulers(); schedule(); }
    if (d.type === 'yc:office-cmd' && typeof d.cmd === 'string') { run(d.cmd); setTimeout(reportState, 50); }
    // WHY: autosave is the host's decision (3 s after the last change), but the save itself must
    // be the EDITOR's own (asc_Save), the same call Ctrl+S makes. sdkjs records where a save it
    // started ends (History.LastUserSavedIndex, set when DesktopOfflineAppDocumentEndSave answers
    // it); calling LocalFileSave directly skips that bookkeeping, so after every save the editor
    // still reported the document as modified and the host saved again every 3 s, forever
    // (measured in the YouCoded dev window, 2026-09-28). asc_Save still ends in LocalFileSave,
    // which sends the bytes and calls save_file. The direct call stays only as a fallback for a
    // frame whose editor API is not reachable yet.
    if (d.type === 'yc:office-save') save();
    if (d.type === 'yc:office-journal') sendEditsNow();
  });
  setInterval(function () { if (slim) reportState(); }, 250);
  // Editor frames appear late and are rebuilt on open; re-walk when the tree changes.
  new MutationObserver(schedule).observe(document.documentElement, { childList: true, subtree: true });

  // "document:opened" fires when the editor ACCEPTS a file, seconds before it
  // has drawn it. The host keeps its own loading state up until this says the
  // document is really on screen: the editor frame exists and its load mask is gone.
  var announced = false;
  function drawn(win) {
    var frames;
    try { frames = win.document.querySelectorAll('iframe'); } catch (e) { return false; }
    for (var i = 0; i < frames.length; i++) {
      var w, d;
      try { w = frames[i].contentWindow; d = w.document; } catch (e) { continue; }
      if (/web-apps\/apps/.test(String(w.location))) {
        var masked = [].some.call(d.querySelectorAll('.asc-loadmask, .loadmask'), function (m) { return m.offsetParent !== null; });
        if (!masked && d.querySelector('#editor_sdk, #ws-canvas-outer, #id_main_view')) return true;
      }
      if (drawn(w)) return true;
    }
    return false;
  }
  // Slim mode lives in a narrow pane. Fitting the whole PAPER width there shrank the text to
  // about half size (office-review#B-inline, Destin: "make these documents fit/fill the pane
  // better"), so a document fits its TEXT column instead: fit to width, then enlarge by the
  // paper-to-text ratio (1in margins leave ~78% of a Letter/A4 page for text) — the blank
  // margins run off the sides. Measured 2026-09-28: 51% → 65% in a 480px pane. Spreadsheets
  // and slides keep the editor's own fit (they have no paper margins to lose).
  function fitWidth() {
    var doc = editorDoc(window);
    var w = doc && doc.defaultView;
    try {
      var api = w && (w.editor || (w.Asc && w.Asc.editor));
      if (!api || !api.zoomFitToWidth) return;
      api.zoomFitToWidth();
      var z = api.WordControl && api.WordControl.m_nZoomValue;
      if (z && api.zoomCustomMode && api.zoom) { api.zoomCustomMode(); api.zoom(Math.round(z / 0.78)); }
    } catch (e) { /* not every editor has these */ }
  }
  var fitTimer = null;
  window.addEventListener('resize', function () {
    if (!slim) return;
    clearTimeout(fitTimer);
    fitTimer = setTimeout(fitWidth, 200);
  });
  // ── The pass that finds the editor's frames while they load ──
  // WHY it stops once the document is drawn (v0.1.31): measured 2026-10-01, each 150 ms pass cost
  // 2-4 ms of the editor's main thread (15-26 ms of CPU every second per open document, forever),
  // and a pass landing in a busy frame was a hitch while scrolling. After the document is drawn,
  // the observers above (watchDom: frames, scripts and panels added later; watchTxtOptions;
  // watchStatusbar), each frame's load event and the host's theme and mode messages cover what
  // the pass used to catch. A frame that loads again starts it again (watchFrameLoad), so a
  // rebuilt editor is guarded, themed and announced as before.
  var loop = null, looping = false;
  function kick() {
    if (looping) return;
    looping = true;
    loop = setInterval(pass, 150);
  }
  function stopLoop() {
    looping = false;
    try { clearInterval(loop); } catch (e) { /* no timers here */ }
  }
  function pass() {
    guardAll(window); // editor frames appear late; guard each as soon as it exists
    if (latest) walk(window); // cheap when nothing changed; catches late theme registration
    var now = drawn(window);
    // yc-comments.js (v0.1.21) runs the host's comment changes only on a drawn document.
    window.__ycDrawn = now;
    if (now && !announced) {
      announced = true;
      // Slim: fit the text first and only then say "drawn", so the host never shows (or
      // photographs) the page at the editor's own small fit-page zoom. Fitted again a moment
      // later because the editor re-lays out its view once after the first paint.
      var announce = function () { window.parent.postMessage({ type: 'yc:office-loaded' }, '*'); };
      if (slim) { setTimeout(function () { fitWidth(); setTimeout(function () { fitWidth(); announce(); }, 700); }, 300); }
      else announce();
    }
    if (!now) announced = false;
    else stopLoop();
  }
  kick();
})();
