/* AF → aftool 一鍵匯入：AF 系統頁面（fund.dgpa.gov.tw）上的按鈕與流程 */
(function () {
  'use strict';
  if (window.__afAftoolExt) return;
  window.__afAftoolExt = true;
  var isTop = window === window.top;
  var running = false, claimed = false, toastEl = null, hideTimer = null;

  function toast(text, kind) {
    if (!isTop) return;
    if (!toastEl) {
      toastEl = document.createElement('div');
      toastEl.id = 'af-aftool-toast';
      toastEl.setAttribute('role', 'status');
      toastEl.style.cssText = 'position:fixed;right:18px;bottom:74px;z-index:2147483647;max-width:360px;padding:10px 14px;border-radius:8px;font:13px/1.5 "Microsoft JhengHei",sans-serif;box-shadow:0 4px 16px rgba(0,0,0,.25);';
      document.documentElement.appendChild(toastEl);
    }
    var c = kind === 'err' ? ['#fdeeee', '#7a0016', '#b00020'] : kind === 'ok' ? ['#eaf5ec', '#1d4d2c', '#2a6e3f'] : ['#eef4fb', '#1a3a5c', '#1a3a5c'];
    toastEl.style.background = c[0]; toastEl.style.color = c[1]; toastEl.style.border = '1px solid ' + c[2];
    toastEl.style.display = 'block';
    toastEl.textContent = text;
    clearTimeout(hideTimer);
    if (kind === 'ok') hideTimer = setTimeout(function () { toastEl.style.display = 'none'; }, 8000);
  }

  function blobToB64(blob) {
    return new Promise(function (resolve, reject) {
      var r = new FileReader();
      r.onload = function () { resolve(String(r.result).split(',')[1] || ''); };
      r.onerror = function () { reject(new Error('讀取報表失敗')); };
      r.readAsDataURL(blob);
    });
  }

  async function start() {
    if (running) { toast('一鍵匯入進行中，請稍候…'); return; }
    running = true; claimed = false;
    try {
      var r = await AFCore.run({
        status: function (t) { toast(t); },
        claimed: function () { return claimed; },
        beforePopup: function () { return chrome.runtime.sendMessage({ type: 'expect-popup' }); }
      });
      if (r.claimed) { toast('已在「報表列印」視窗繼續處理…', 'ok'); return; }
      toast('已取得報表，傳送到 aftool…');
      var b64 = await blobToB64(r.blob);
      var resp = await chrome.runtime.sendMessage({ type: 'deliver', name: r.name, b64: b64 });
      if (resp && resp.ok) toast('✅ 已匯入 aftool：' + (resp.filename || '校對清冊') + '。多校時切換「兼辦機關」後再按一次。', 'ok');
      else toast('已取得報表，但匯入 aftool 失敗：' + ((resp && resp.error) || '未知原因'), 'err');
    } catch (e) {
      toast('一鍵匯入失敗：' + (e && e.message || e) + '（可改用手動「下載報表」，再把 ZIP 拖進 aftool）', 'err');
    } finally {
      running = false;
    }
  }

  function addButton() {
    if (document.getElementById('af-aftool-btn')) return;
    var b = document.createElement('button');
    b.id = 'af-aftool-btn';
    b.type = 'button';
    b.textContent = 'AF → aftool 一鍵匯入';
    b.title = '自動選「校對清冊(A)」、EXCEL、填公版密碼，取得報表並送進 aftool';
    b.style.cssText = 'position:fixed;right:18px;bottom:18px;z-index:2147483647;padding:10px 18px;border:0;border-radius:24px;background:#1a3a5c;color:#fff;font:bold 15px "Microsoft JhengHei",sans-serif;box-shadow:0 4px 14px rgba(0,0,0,.3);cursor:pointer;';
    b.addEventListener('click', function (e) { e.preventDefault(); e.stopPropagation(); start(); });
    document.documentElement.appendChild(b);
  }

  chrome.runtime.onMessage.addListener(function (m, sender, sendResponse) {
    if (!m) return;
    if (m.type === 'run' && isTop) { start(); sendResponse({ ok: true }); }
    else if (m.type === 'claimed') { claimed = true; }
  });

  if (isTop) {
    addButton();
    // 主視窗剛按了「報表列印」而開出這個視窗時，由這裡接手
    if (AFCore.hasDownloadButton()) {
      chrome.runtime.sendMessage({ type: 'popup-ready' }).then(function (resp) {
        if (resp && resp.run) start();
      }, function () {});
    }
  }
})();
