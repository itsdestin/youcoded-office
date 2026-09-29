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
  // WHY: the host (YouCoded) saves every open document before its window closes and asks the
  // person itself when a save failed (Review / Close anyway). The editor's own "leave page?"
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
    for (var i = 0; i < dlgs.length; i++) {
      var w = dlgs[i];
      if (w.__ycAccepted || !w.querySelector('#id-codepages-combo') || w.querySelector('#id-delimiters-combo')) continue;
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
    extendDesktopEditor(win);
    noEditorOpen(win);
    watchTxtOptions(win);
    acceptTxtOptions(win);
    watchDrops(win);
    blockPeers(win);
    guardUnload(win);
    quietEditor(win);
    var frames;
    try { frames = win.document.querySelectorAll('iframe'); } catch (e) { return; }
    for (var i = 0; i < frames.length; i++) {
      try { if (frames[i].contentWindow) guardAll(frames[i].contentWindow); } catch (e) { /* cross-origin */ }
    }
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
  var seen = new WeakSet();

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
    // fields are inset; a 1px edge hairline sits between regions. Hairlines are inset shadows,
    // not borders, so the editor's own layout (it measures these boxes in script) does not move.
    css += '#toolbar > .toolbar, #toolbar .toolbar.toolbar-mask { box-shadow: inset 0 0 0 1px ' + t.edge + I + '; }' +
      '#statusbar, .statusbar { box-shadow: inset 0 1px 0 ' + t.edge + I + '; }' +
      '#left-menu .tool-menu-btns, #left-menu.tool-menu, .tool-menu.left .tool-menu-btns { box-shadow: inset -1px 0 0 ' + t.edge + I + '; }' +
      '#right-menu .tool-menu-btns, .tool-menu.right .tool-menu-btns { box-shadow: inset 1px 0 0 ' + t.edge + I + '; }' +
      '.left-panel, #left-panel-search, #left-panel-comments, #left-panel-chat, #left-panel-navigation, #left-panel-thumbnails { background-color: ' + o.panel + I + '; box-shadow: inset -1px 0 0 ' + t.edge + I + '; }' +
      '.right-panel, #right-menu .right-panel { background-color: ' + o.panel + I + '; box-shadow: inset 1px 0 0 ' + t.edge + I + '; }' +
      '#cell-editing-box { box-shadow: inset 0 -1px 0 ' + t.edge + I + '; }';
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
      '#file-menu-panel tr.themes, #file-menu-panel tr:has(#fms-cmb-theme) { display: none' + I + '; }' +
      // v0.1.14: Word's TXT encoding dialog is answered for the person (acceptTxtOptions says why);
      // hidden from its first frame so it never flashes. The CSV one has a delimiter and stays.
      '.asc-window.open-dlg:has(#id-codepages-combo):not(:has(#id-delimiters-combo)) { visibility: hidden' + I + '; }' +
      // ...and its mask never blocks a click while it is up (fix round 2).
      'body:has(.asc-window.open-dlg #id-codepages-combo):not(:has(.asc-window.open-dlg #id-delimiters-combo)) .modals-mask { visibility: hidden' + I + '; pointer-events: none' + I + '; }' +
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
      '#slot-btn-text-from-file, #slot-btn-mailrecepients, #id-right-menu-mail-merge, #slot-btn-insaudio, #slot-btn-insvideo,' +
      ' .group:has(> #slot-btn-compare), .group:has(> #slot-btn-compare) + .separator,' +
      ' .group:has(> #slot-btn-data-from-text), .group:has(> #slot-btn-data-from-text) + .separator,' +
      ' #external-links-btn-change, #external-links-btn-open, #external-links-btn-update, #chart-button-update-data,' +
      ' #id-dlg-hyperlink-url .select-button { display: none' + I + '; }' +
      // Fix round 4: Word/PowerPoint chart settings — "Update data" (#chart-button-update-data,
      // hidden above) takes the same failing update path as the dialog's Update values; the linked
      // source's name stays readable, but as plain text: its link opens the source through a
      // document server this host doesn't have.
      '#chart-open-external-link { pointer-events: none' + I + '; cursor: default' + I + '; color: inherit' + I + '; text-decoration: none' + I + '; border-bottom: none' + I + '; }' +
      // Opaque, even over a wallpaper: the File tab covers the document, and a see-through
      // panel showed the page's text through its list and settings (Meadow Mist, 2026-09-28).
      '#file-menu-panel .panel-menu { background-color: ' + t.panel + I + '; border-right: 1px solid ' + t.edge + I + '; padding: 12px 8px 16px' + I + '; }' +
      '#file-menu-panel .panel-menu li.fm-btn { height: 32px' + I + '; padding: 0 12px' + I + '; margin-bottom: 2px' + I + '; border-radius: ' + md + I + '; }' +
      '#file-menu-panel .panel-menu li.fm-btn > a { font-size: 13px' + I + '; color: ' + t.fg + I + '; }' +
      '#file-menu-panel .panel-menu li.fm-btn:hover:not(.disabled) { background-color: ' + t.inset + I + '; }' +
      '#file-menu-panel .panel-menu li.fm-btn.active:not(.disabled) { background-color: ' + t.inset + I + '; box-shadow: inset 3px 0 0 ' + t.accent + I + '; }' +
      '#file-menu-panel .panel-menu li.fm-btn.active:not(.disabled) > a { font-weight: 600' + I + '; }' +
      '#file-menu-panel #fm-btn-return { margin-bottom: 12px' + I + '; }' +
      '#file-menu-panel, #file-menu-panel .panel-context { background-color: ' + (t.canvas || t.panel) + I + '; }' +
      '#file-menu-panel .panel-context .header, #file-menu-panel .panel-context h1, #file-menu-panel .panel-context .title { color: ' + t.fg + I + '; }';
    return css + printCss();
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
    return '#id-print-settings tr:has(#print-combo-printer), #id-print-settings tr:has(+ tr #print-combo-printer),' +
      ' #id-print-settings tr:has(#print-combo-color-printing), #id-print-settings tr:has(+ tr #print-combo-color-printing),' +
      ' #id-print-settings tr:has(> td > #print-combo-sides), #id-print-settings tr:has(+ tr > td > #print-combo-sides),' +
      ' #id-print-settings tr:has(> td > .pages #print-txt-copies), #id-print-settings tr:has(> td #print-txt-copies):not(:has(#print-txt-pages)),' +
      ' #id-print-settings tr:has(> td > .separator), #id-print-settings tr:has(#print-btn-system-dialog),' +
      ' #print-combo-range li[data-value="2"],' +
      ' .dropdown-menu li:has(> a .menu__icon.btn-print), #slot-btn-dt-print-quick, .btn-quick-print { display: none' + I + '; }';
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
      '--highlight-toolbar-tab-underline': t.accent, '--highlight-toolbar-tab-underline-document': t.accent,
      '--highlight-toolbar-tab-underline-spreadsheet': t.accent, '--highlight-toolbar-tab-underline-presentation': t.accent,
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
        '#toolbar .toolbar, #statusbar, .statusbar, #left-menu, #right-menu, .right-panel { background: ' + panel + ' !important; }' +
        // WHY on a layer behind each panel, not the panel (v0.1.9): a backdrop-filter makes its
        // element the frame of every position:fixed menu inside it, so the right panel's menus
        // (the slide background's "Select picture", among others) opened ~1400px to the right,
        // off screen. Each of these four is already positioned (relative/absolute), so the layer
        // fills it exactly; z-index -1 keeps it under the panel's content. The static
        // .right-panel/.statusbar sit inside #right-menu/#statusbar and share their layer.
        (th.panelsBlur ? '#toolbar .toolbar::before, #statusbar::before, #left-menu::before, #right-menu::before { content: ""; position: absolute; inset: 0; z-index: -1; pointer-events: none; border-radius: inherit; backdrop-filter: blur(' + th.panelsBlur + 'px); }' : '');
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
    var skin = skinFor(latest), key = JSON.stringify(skin);
    if (win.__ycSkin === key) return;
    try { api.asc_setSkin(skin); win.__ycSkin = key; } catch (e) { /* editor still starting: next pass */ }
  }

  function applyTo(win) {
    var doc;
    try { doc = win.document; } catch (e) { return; } // not same-origin: not ours
    if (!doc || !doc.head) return;
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
    var style = doc.getElementById(STYLE_ID);
    if (!style) { style = doc.createElement('style'); style.id = STYLE_ID; doc.head.appendChild(style); }
    var css = buildCss(latest);
    if (style.textContent === css) return;
    style.textContent = css;
    // OnlyOffice positions its bands in script; a resize makes it lay out again
    // without the rows we just hid.
    win.dispatchEvent(new win.Event('resize'));
  }

  function walk(win) {
    blockPeers(win);
    guardUnload(win);
    quietEditor(win);
    if (latest) applyTo(win);
    var frames;
    try { frames = win.document.querySelectorAll('iframe'); } catch (e) { return; }
    for (var i = 0; i < frames.length; i++) {
      var f = frames[i];
      if (!seen.has(f)) { seen.add(f); f.addEventListener('load', function () { schedule(); }); }
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
  function save() {
    var api = editorApi();
    // WHY the modified check (v0.1.12): a Save As writes a separate file but still clears the
    // editor's "modified" flag, and asc_Save then skips the save — the document's own file would
    // silently miss the edits. The host asks only when it knows edits are unsaved, so an editor
    // that says "nothing changed" goes straight to LocalFileSave, which saves the current content.
    var unchanged = api && typeof api.isDocumentModified === 'function' && !api.isDocumentModified();
    if (api && typeof api.asc_Save === 'function' && !unchanged) { api.asc_Save(false); return; }
    if (window.AscDesktopEditor) window.AscDesktopEditor.LocalFileSave('', '', null, 0, null);
  }
  function run(cmd) {
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

  window.addEventListener('message', function (e) {
    if (e.source !== window.parent) return; // only the host frames us
    var d = e.data || {};
    if (d.type === 'yc:office-theme' && d.theme && d.theme.tokens) { latest = d.theme; schedule(); }
    if (d.type === 'yc:office-mode') { slim = !!d.slim; setRulers(); schedule(); }
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
  setInterval(function () {
    guardAll(window); // editor frames appear late; guard each as soon as it exists
    if (latest) walk(window); // cheap when nothing changed; catches late theme registration
    var now = drawn(window);
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
  }, 150);
})();
