/* The add-on's __TAURI__ relay (design §3a): euro-office-lite's bridge.js calls Tauri's
 * JS API; inside YouCoded the editor page has no host API at all, so every call becomes a
 * postMessage to the framing app page, which forwards it to main and posts the answer back.
 * Loaded before bridge.js. */
(function () {
  var seq = 0, pending = {}, listeners = {}, announced = false;
  window.addEventListener('message', function (e) {
    if (e.source !== window.parent) return;
    var d = e.data || {};
    if (d.yc === 'rpc-result' && pending[d.id]) {
      var p = pending[d.id]; delete pending[d.id];
      if (d.error) p.reject(new Error(d.error)); else p.resolve(d.result);
    }
    if (d.yc === 'event') (listeners[d.name] || []).forEach(function (cb) { cb({ event: d.name, payload: d.payload }); });
  });
  function invoke(cmd, args) {
    // WHY answered here (v0.1.18, Print): the editor's print panel waits for a printer list before
    // its Print button works. The host prints through the operating system's dialog, which lists
    // the real printers, so the panel gets one stand-in entry (its row is hidden — yc-bridge.js
    // printCss) and nothing is asked of the host.
    if (cmd === 'plugin:printer|get_printers') return Promise.resolve(JSON.stringify([{ name: 'YouCoded', is_default: true }]));
    return new Promise(function (resolve, reject) {
      var id = ++seq; pending[id] = { resolve: resolve, reject: reject };
      window.parent.postMessage({ yc: 'rpc', id: id, cmd: cmd, args: args || {} }, '*');
    });
  }
  window.__TAURI__ = {
    core: { invoke: invoke },
    // WHY a ready signal: the host must not send "open-file" before bridge.js listens for it.
    // The spike guessed with a 1.5 s timer; the first listen() for it is the exact moment.
    event: { listen: function (name, cb) {
      (listeners[name] = listeners[name] || []).push(cb);
      if (name === 'open-file' && !announced) { announced = true; window.parent.postMessage({ yc: 'ready' }, '*'); }
      return Promise.resolve(function () {});
    } },
    dialog: {
      confirm: function () { return Promise.resolve(true); },
      message: function () { return Promise.resolve(); },
      // WHY (Insert → Picture → From file): bridge.js's OpenFilenameDialog calls this, and without
      // it the call threw and nothing happened. The host shows its own system dialog and answers
      // Tauri's shape (a list when multiple, else one entry, null when cancelled) — with opaque
      // handles, never folders; the editor passes one on to copy-to-media, which only the host
      // can resolve. Only the two fields the dialog needs go across.
      open: function (o) {
        o = o || {};
        return invoke('open_dialog', { multiple: !!o.multiple, filters: Array.isArray(o.filters) ? o.filters : [] });
      },
      // WHY (v0.1.12, Save As / Download as / Export to PDF): bridge.js's LocalFileSave asks this
      // where the file goes, then hands the answer to save_file_as. The host shows its own save
      // dialog and answers a handle ending in the chosen name — never a folder — or null when
      // cancelled. Only the filters go across; a defaultPath from the frame is not a folder the
      // host would ever start in.
      save: function (o) {
        o = o || {};
        return invoke('save_dialog', { filters: Array.isArray(o.filters) ? o.filters : [] });
      },
    },
    window: { getCurrentWindow: function () { return { setTitle: function () { return Promise.resolve(); }, close: function () { return Promise.resolve(); }, onCloseRequested: function () { return Promise.resolve(function () {}); } }; } },
  };
})();
