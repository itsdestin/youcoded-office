/* yc-early.js — the FIRST script in every Euro-Office editor page (web-apps/apps/<editor>/main/
 * index*.html; build/patch.mjs puts the tag at the top of <head> at bundle time).
 *
 * WHY here and at bundle time (v0.1.5): yc-bridge.js (the host page) could only reach an editor
 * frame on its 150 ms walk, which raced the editor's own start — a fast open could ask about
 * external links before the walk arrived. This script runs in the editor's own window before any
 * sdkjs code, and patches each class the moment sdkjs publishes it, synchronously.
 *
 * What it patches — the "links to external sources" warning (web-apps ExternalLinks.js) and the
 * automatic link refresh. Neither of the warning's buttons is acceptable: "Update"/"Continue"
 * tries to fetch the linked workbooks (reaching outside the document), and "Turn off
 * AutoUpdate"/"Don't update" with auto-update on calls asc_setUpdateLinks(false, true), which
 * writes a history point, so the file CHANGES and autosave rewrites it. So:
 *   - AscCommon.baseEditorsApi.prototype.onNeedUpdateExternalReferenceOnOpen (sdkjs
 *     common/apiBase.js — the one place the word, cell and slide APIs ask for the warning at
 *     open) does nothing;
 *   - AscCommonExcel.WorkbookView.prototype.initExternalReferenceUpdateTimer (sdkjs
 *     cell/view/WorkbookView.js — re-fetches every link 30 s after open, and after each update,
 *     when the workbook says "always") does nothing.
 * The values the file already stores stay as they are. Data > External links still OFFERS an
 * update, but it cannot reach anything outside: the editor has no network (its CSP) and no way to
 * open another file on its own.
 *
 * sdkjs publishes both as `window.AscCommon = window.AscCommon || {}; AscCommon.baseEditorsApi = e`
 * (and the same for AscCommonExcel.WorkbookView), after the prototype is complete — so a setter
 * on each name sees the finished class.
 *
 * It also takes away peer-to-peer connections (v0.1.6). WHY: the page's CSP stops fetches and
 * sockets off the document's own origin, but Chromium does not apply CSP to WebRTC — an
 * RTCPeerConnection could still send a document's text to another machine. The editor works
 * offline and never needs one, so each constructor throws, and cannot be put back.
 */
(function () {
  function blockPeers(win) {
    var names = ['RTCPeerConnection', 'webkitRTCPeerConnection'];
    for (var i = 0; i < names.length; i++) {
      var name = names[i];
      var stub = function () { throw new Error('YouCoded Office: peer connections are turned off'); };
      try { Object.defineProperty(win, name, { value: stub, writable: false, configurable: false, enumerable: false }); }
      catch (e) { /* already sealed */ }
    }
  }
  blockPeers(window);

  // ── The person's editor settings, remembered between documents (v0.1.19) ──
  // WHY here, first, and synchronously: each document opens on its own one-time office://<token>
  // origin, so this page's localStorage starts empty, and web-apps reads its settings (units,
  // spell check, zoom, the view toggles…) from localStorage as it starts. The host keeps the ones
  // the person chose (yc-bridge.js sends each change) and serves them on this origin; they are
  // written in before any editor code runs. A key the page already has is left alone: this page
  // may be a reload of a document whose settings changed since it opened, and those are newer.
  // The host sends only its allow-list of settings keys (desktop editor-settings.ts); this checks
  // only that each is text.
  function seedSettings(win) {
    try {
      var ls = win.localStorage;
      var x = new win.XMLHttpRequest();
      x.open('GET', win.location.origin + '/yc-settings.json', false);
      x.send();
      if (x.status !== 200) return;
      var saved = JSON.parse(x.responseText);
      if (!saved || typeof saved !== 'object') return;
      for (var k in saved) {
        if (!Object.prototype.hasOwnProperty.call(saved, k) || typeof saved[k] !== 'string') continue;
        if (ls.getItem(k) === null) ls.setItem(k, saved[k]);
      }
    } catch (e) { /* no host answer or no storage: the editor starts with its own defaults */ }
  }
  seedSettings(window);

  function onDefine(owner, name, patch) {
    var value = owner[name];
    if (value) { try { patch(value); } catch (e) { /* keep the editor running */ } }
    try {
      Object.defineProperty(owner, name, {
        configurable: true,
        enumerable: true,
        get: function () { return value; },
        set: function (v) {
          if (v && v !== value) { try { patch(v); } catch (e) { /* keep the editor running */ } }
          value = v;
        },
      });
    } catch (e) { /* not definable: yc-bridge's later walk still patches */ }
  }
  function quietApi(Api) {
    if (Api && Api.prototype) Api.prototype.onNeedUpdateExternalReferenceOnOpen = function () {};
  }
  function quietWorkbookView(View) {
    if (View && View.prototype) View.prototype.initExternalReferenceUpdateTimer = function () {};
  }
  // ── A sheet's scrollbars keep their size (fix round 2) ──
  // WHY (Destin, 2026-10-01: the scrollbars "elongate/glitch when scrolling a spreadsheet"): a
  // sheet's scroll range is only its used rows and columns plus what is on screen, and sdkjs adds
  // rows as the view reaches the end. So scrolling a small sheet ran the thumb to the bottom of its
  // track, then the range grew and the thumb jumped back up and changed size — every few wheel
  // turns (measured in the workbench: 433px → 327px and back up the track). Like a fresh Google
  // sheet, a sheet's view now counts at least 1000 rows and 52 columns (A–AZ), so the thumb is
  // small and keeps its size until well past anything a small sheet reaches; a bigger sheet keeps
  // its own size. Only the view's count changes — the workbook, its used range and what is saved
  // or printed do not.
  var MIN_ROWS = 1000, MIN_COLS = 52;
  function steadySheetScroll(View) {
    var P = View && View.prototype;
    if (!P || P.__ycSteady || typeof P._initRowsCount !== 'function' || typeof P._initColsCount !== 'function') return;
    var rows = P._initRowsCount, cols = P._initColsCount;
    P._initRowsCount = function () {
      var before = this.nRowsCount, changed = rows.apply(this, arguments);
      if (this.nRowsCount < MIN_ROWS) this.nRowsCount = MIN_ROWS;
      return changed || before !== this.nRowsCount;
    };
    P._initColsCount = function () {
      var before = this.nColsCount, changed = cols.apply(this, arguments);
      if (this.nColsCount < MIN_COLS && typeof this.setColsCount === 'function') this.setColsCount(MIN_COLS);
      return changed || before !== this.nColsCount;
    };
    P.__ycSteady = true;
  }
  // ── Slim, rounded scrollbars in the document and sheet canvases (v0.1.7) ──
  // WHY (Destin, 2026-09-28: "all of the scrollbars are unstyled"): sdkjs draws these scrollbars
  // itself, on canvases (common/scroll.js ScrollObject), so CSS cannot reach them. Its own drawing
  // is a square, outlined thumb with grip stripes and arrow buttons at both ends — and it paints
  // the thumb in grey only, from the RED channel of the theme colour, so a green thumb came out
  // dark grey. YouCoded's scrollbars are a slim rounded thumb in the theme's --scrollbar-thumb /
  // --scrollbar-hover colours with no arrows (renderer globals.css). So:
  //   - every ScrollSettings is made without arrows (the thumb then has the whole track; the
  //     wheel, a drag and a click on the track still scroll), and
  //   - ScrollObject.prototype._drawScroll draws a 6px rounded thumb in the real colours, over a
  //     clear track (the desk shows through). sdkjs animates hover by passing a grey level
  //     between the default and hover colours' red channels; that is mapped back onto the two
  //     real colours, so the fade still works.
  function hexRgb(c) {
    var m = /^#?([0-9a-f]{6})/i.exec(String(c || '').trim());
    if (!m) return null;
    var n = parseInt(m[1], 16);
    return [(n >> 16) & 255, (n >> 8) & 255, n & 255];
  }
  function thumbColour(settings, level) {
    var base = hexRgb(settings.scrollerColor), hover = hexRgb(settings.scrollerHoverColor), active = hexRgb(settings.scrollerActiveColor);
    if (!base) return settings.scrollerColor || '#888';
    hover = hover || base; active = active || hover;
    var t = 0;
    if (active[0] !== hover[0] && Math.round(level) === active[0]) return 'rgb(' + active.join(',') + ')';
    if (hover[0] !== base[0]) t = Math.max(0, Math.min(1, (level - base[0]) / (hover[0] - base[0])));
    else if (Math.round(level) !== base[0]) t = 1;
    return 'rgb(' + [0, 1, 2].map(function (i) { return Math.round(base[i] + (hover[i] - base[i]) * t); }).join(',') + ')';
  }
  function slimSettings(Orig) {
    if (typeof Orig !== 'function' || Orig.__ycSlim) return Orig;
    var Slim = function () { Orig.apply(this, arguments); this.showArrows = false; };
    Slim.prototype = Orig.prototype;
    for (var k in Orig) if (Object.prototype.hasOwnProperty.call(Orig, k)) Slim[k] = Orig[k];
    Slim.__ycSlim = true;
    return Slim;
  }
  function drawSlim(fillLevel, targetLevel, strokeLevel) {
    var s = this.settings, ctx = this.context, sc = this.scroller;
    this.scrollColor = fillLevel; this.targetColor = targetLevel; this.strokeColor = strokeLevel;
    if (!ctx || !s || !sc) return;
    var br = window.AscCommon && window.AscCommon.AscBrowser;
    var dpr = (br && br.retinaPixelRatio) || 1;
    var th = Math.max(2, Math.round(6 * dpr));
    ctx.clearRect(0, 0, this.canvasW, this.canvasH);
    var x, y, w, h;
    if (s.isVerticalScroll && this.maxScrollY != 0) {
      x = Math.round(sc.x + (sc.w - th) / 2); w = th;
      y = Math.max(0, Math.round(sc.y)); h = Math.min(this.canvasH - y, Math.round(sc.h));
    } else if (s.isHorizontalScroll && this.maxScrollX != 0) {
      y = Math.round(sc.y + (sc.h - th) / 2); h = th;
      x = Math.max(0, Math.round(sc.x)); w = Math.min(this.canvasW - x, Math.round(sc.w));
    } else return;
    // Fix round 2 (Destin, 2026-10-01: the scrollbars "overlap the rounded corners of their
    // containers"): the canvases sit along the document area's rounded edges, so the thumb's whole
    // travel is squeezed into the track minus the corner radius at each end (the host page's bridge
    // hands that radius over as __ycScrollInset). It is proportional, so the thumb still moves with
    // the content and reaches both ends of its shorter track.
    var inset = Math.max(0, Number(window.__ycScrollInset) || 0) * dpr;
    if (s.isVerticalScroll) {
      var LH = this.canvasH, iy = Math.min(inset, LH / 4), ky = (LH - 2 * iy) / LH;
      y = Math.round(iy + y * ky); h = Math.max(th, Math.round(h * ky));
    } else {
      var LW = this.canvasW, ix = Math.min(inset, LW / 4), kx = (LW - 2 * ix) / LW;
      x = Math.round(ix + x * kx); w = Math.max(th, Math.round(w * kx));
    }
    if (w <= 0 || h <= 0) return;
    var r = th / 2;
    ctx.beginPath();
    ctx.moveTo(x + r, y);
    ctx.lineTo(x + w - r, y); ctx.arcTo(x + w, y, x + w, y + r, r);
    ctx.lineTo(x + w, y + h - r); ctx.arcTo(x + w, y + h, x + w - r, y + h, r);
    ctx.lineTo(x + r, y + h); ctx.arcTo(x, y + h, x, y + h - r, r);
    ctx.lineTo(x, y + r); ctx.arcTo(x, y, x + r, y, r);
    ctx.closePath();
    ctx.fillStyle = thumbColour(s, fillLevel);
    ctx.fill();
  }
  function slimScroll(ScrollObject) {
    if (ScrollObject && ScrollObject.prototype) ScrollObject.prototype._drawScroll = drawSlim;
  }
  // Like onDefine, but the value itself is replaced (a constructor wrapped) rather than patched.
  function onDefineWrap(owner, name, wrap) {
    var value = owner[name] ? wrap(owner[name]) : owner[name];
    try {
      Object.defineProperty(owner, name, {
        configurable: true,
        enumerable: true,
        get: function () { return value; },
        set: function (v) { try { value = wrap(v); } catch (e) { value = v; } },
      });
    } catch (e) { /* not definable: the editor keeps its own scrollbars */ }
  }

  onDefine(window, 'AscCommon', function (ns) {
    onDefine(ns, 'baseEditorsApi', quietApi);
    onDefineWrap(ns, 'ScrollSettings', slimSettings);
    onDefine(ns, 'ScrollObject', slimScroll);
  });
  onDefine(window, 'AscCommonExcel', function (ns) { onDefine(ns, 'WorkbookView', quietWorkbookView); onDefine(ns, 'WorksheetView', steadySheetScroll); });
})();
