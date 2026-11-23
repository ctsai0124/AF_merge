/* AF → aftool 一鍵匯入：aftool 頁面端橋接。把背景服務轉來的 ZIP 以 window.postMessage
 * 交給頁面上的 static/af_import.js（只送往本頁同源），並把結果回報給背景服務。 */
(function () {
  'use strict';
  chrome.runtime.onMessage.addListener(function (m, sender, sendResponse) {
    if (!m || m.type !== 'af-zip' || typeof m.b64 !== 'string') return;
    var bin = atob(m.b64), bytes = new Uint8Array(bin.length);
    for (var i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
    var blob = new Blob([bytes], { type: 'application/zip' });
    var id = Math.random().toString(36).slice(2) + Date.now().toString(36);
    var done = false;
    function finish(r) { if (done) return; done = true; window.removeEventListener('message', onMsg); sendResponse(r); }
    function onMsg(e) {
      if (e.source !== window || e.origin !== location.origin || !e.data || e.data.source !== 'aftool-page') return;
      if (e.data.id === id && e.data.type === 'af-ext-result') finish({ ok: !!e.data.ok, filename: e.data.filename || '', error: e.data.error || '' });
    }
    window.addEventListener('message', onMsg);
    window.postMessage({ source: 'aftool-extension', type: 'af-zip', id: id, name: String(m.name || 'AF校對清冊.zip').slice(0, 200), blob: blob }, location.origin);
    setTimeout(function () { finish({ ok: false, error: 'aftool 頁面沒有回應（可能是舊版頁面，請重新整理 aftool 後再按一次）。' }); }, 20000);
    return true;
  });
})();
