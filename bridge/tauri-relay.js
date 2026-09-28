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
    dialog: { confirm: function () { return Promise.resolve(true); }, message: function () { return Promise.resolve(); } },
    window: { getCurrentWindow: function () { return { setTitle: function () { return Promise.resolve(); }, close: function () { return Promise.resolve(); }, onCloseRequested: function () { return Promise.resolve(function () {}); } }; } },
  };
})();
