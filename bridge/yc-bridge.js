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
  function guardAll(win) {
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
  }

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

  function buildCss(th) {
    var t = th.tokens;
    var glass = th.wallpaper ? Math.max(0.35, Math.min(1, th.panelsOpacity || 0.6)) : 1;
    var panel = rgba(t.panel, glass);
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
      '--border-sidemenu': t.edge, '--border-toolbar-active-panel-top': panel, '--border-control-focus': t.accent,
      '--background-fill-input': t.inset, '--border-fill-input': t.edge,
      '--canvas-background': th.wallpaper ? 'transparent' : t.canvas,
      '--canvas-content-background': '#fff', '--canvas-page-border': t.edge,
      '--canvas-ruler-background': panel, '--canvas-ruler-border': t.edge, '--canvas-ruler-margins-background': t.inset,
      '--canvas-high-contrast': t.fg, '--canvas-scroll-thumb': t.edge, '--canvas-scroll-thumb-hover': t['fg-faint'],
      '--canvas-scroll-arrow': t['fg-muted'], '--canvas-background-tabs': panel,
      // Roundness: YouCoded's four radii by role (guide G-3).
      '--border-radius-button-normal': t['radius-md'], '--border-radius-button-base': t['radius-md'],
      '--border-radius-button-toolbar': t['radius-md'], '--border-radius-button-category': t['radius-md'],
      '--border-radius-toolbar': t['radius-lg'], '--border-radius-form-control': t['radius-md'],
      '--border-radius-dropdown-menu': t['radius-md'], '--border-radius-dataview-item': t['radius-md'],
      '--border-radius-window': t['radius-lg'], '--border-radius-checkbox': t['radius-sm'],
      '--font-family-base': t['font-sans'], '--font-family-base-custom': t['font-sans'],
    };
    var decl = '';
    for (var k in v) if (v[k]) decl += k + ':' + v[k] + ' !important;';
    var css = ':root, body, body[class] {' + decl + '}' +
      // YouCoded's tabs carry the file name and the save state, so the editor's title row goes.
      '#app-title { display: none !important; }' +
      // Onboarding tips pop over the document; YouCoded explains its own features.
      '.tooltip.new-feature, .synch-tip, .asc-synchronizetip { display: none !important; }';
    if (th.wallpaper) {
      css += 'html, body, #viewport, .layout-region, #editor_sdk, #id_main, #ws-canvas-outer, .ws-canvas-area { background-color: transparent !important; }' +
        '#toolbar .toolbar, #statusbar, .statusbar, #left-menu, #right-menu, .right-panel { background: ' + panel + ' !important;' +
        (th.panelsBlur ? ' backdrop-filter: blur(' + th.panelsBlur + 'px);' : '') + ' }';
    }
    if (slim) {
      css += '#toolbar, #statusbar, .statusbar, #left-menu, #right-menu, .right-panel, .left-panel { display: none !important; }';
    }
    return css;
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
    if (api && typeof api.asc_Save === 'function') { api.asc_Save(false); return; }
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
