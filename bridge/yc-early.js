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
  onDefine(window, 'AscCommon', function (ns) { onDefine(ns, 'baseEditorsApi', quietApi); });
  onDefine(window, 'AscCommonExcel', function (ns) { onDefine(ns, 'WorkbookView', quietWorkbookView); });
})();
