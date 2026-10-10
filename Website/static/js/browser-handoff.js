/* Memory-only file transfer between tool tabs. No file bytes enter browser storage. */
(function () {
  'use strict';
  var MAX_BYTES = 25 * 1024 * 1024;
  var TIMEOUT = 120000;
  var channels = [];

  function close(channel) {
    channel.close();
    var at = channels.indexOf(channel);
    if (at >= 0) channels.splice(at, 1);
  }
  function channelFor(token) {
    if (!window.BroadcastChannel) throw new Error('This browser cannot carry files between tools. Download the file and open it in the next tool.');
    var channel = new BroadcastChannel('officeimo-handoff-' + token);
    channels.push(channel);
    return channel;
  }
  function valid(file) {
    return file && typeof file.name === 'string' && file.name.length > 0 && file.name.length <= 512 &&
      file.buffer instanceof ArrayBuffer && file.buffer.byteLength <= MAX_BYTES;
  }

  function send(url, getFile) {
    var destination = new URL(url, window.location.href);
    if (destination.origin !== window.location.origin) return Promise.reject(new Error('Files can only be carried to another OfficeIMO tool.'));
    var random = new Uint32Array(4);
    window.crypto.getRandomValues(random);
    var token = Array.prototype.map.call(random, function (n) { return n.toString(16).padStart(8, '0'); }).join('');
    var channel;
    try { channel = channelFor(token); } catch (error) { return Promise.reject(error); }
    destination.hash = 'handoff=' + token;
    // Open synchronously from the click so loading the bytes cannot consume popup permission.
    var tab = window.open(destination.href, '_blank');
    if (!tab) { close(channel); return Promise.reject(new Error('Allow the new tool tab to open, then try again.')); }
    tab.opener = null;
    return new Promise(function (resolve, reject) {
      var file = null, receiverReady = false, sent = false, settled = false;
      var timer = window.setTimeout(function () { finish(new Error('The next tool did not receive the file. Try again.')); }, TIMEOUT);
      function finish(error) {
        if (settled) return;
        settled = true;
        window.clearTimeout(timer);
        file = null;
        close(channel);
        if (error) reject(error); else resolve();
      }
      channel.cancel = function () { finish(new Error('The file transfer was cancelled when the tab closed.')); };
      function transfer() {
        if (!file || !receiverReady || sent) return;
        sent = true;
        try { channel.postMessage({ type: 'file', file: file }); }
        catch (error) { finish(error); }
      }
      channel.onmessage = function (event) {
        if (event.data && event.data.type === 'ready') { receiverReady = true; transfer(); }
        else if (event.data && event.data.type === 'received') finish();
      };
      Promise.resolve().then(getFile).then(function (value) {
        if (settled) return;
        if (!valid(value)) throw new Error('The file exceeds the browser limit or could not be read.');
        file = value;
        transfer();
      }).catch(function (error) {
        if (settled) return;
        try { channel.postMessage({ type: 'error', message: 'The previous tool could not read the file. Choose it here instead.' }); }
        catch (ignored) { /* The sender still settles its own failure below. */ }
        finish(error);
      });
    });
  }

  function receive() {
    var token = new URLSearchParams(window.location.hash.slice(1)).get('handoff');
    if (!token || !/^[0-9a-f]{32}$/.test(token)) return Promise.resolve(null);
    window.history.replaceState(null, '', window.location.pathname + window.location.search);
    return new Promise(function (resolve, reject) {
      var channel;
      try { channel = channelFor(token); } catch (error) { reject(error); return; }
      var timer, interval, settled = false;
      function finish(file, error) {
        if (settled) return;
        settled = true;
        window.clearTimeout(timer);
        window.clearInterval(interval);
        close(channel);
        if (error) reject(error); else resolve(file);
      }
      channel.cancel = function () { finish(null, new Error('The file transfer was cancelled when the tab closed.')); };
      channel.onmessage = function (event) {
        var message = event.data;
        if (message && message.type === 'file') {
          if (!valid(message.file)) { finish(null, new Error('The transferred file is invalid or too large.')); return; }
          channel.postMessage({ type: 'received' });
          finish(message.file);
        } else if (message && message.type === 'error') finish(null, new Error(message.message));
      };
      timer = window.setTimeout(function () { finish(null, new Error('The file transfer expired. Choose the file again.')); }, TIMEOUT);
      interval = window.setInterval(function () { channel.postMessage({ type: 'ready' }); }, 500);
      channel.postMessage({ type: 'ready' });
    });
  }

  // Retire the old dedicated handoff store. It held file bytes, and had no other product data.
  if (window.indexedDB) {
    try { window.indexedDB.deleteDatabase('officeimo-browser-tools'); } catch (error) { /* Storage can be disabled. */ }
  }
  window.addEventListener('pagehide', function () { channels.slice().forEach(function (channel) { if (channel.cancel) channel.cancel(); else close(channel); }); });
  window.OfficeIMOBrowserHandoff = { send: send, receive: receive };
})();
