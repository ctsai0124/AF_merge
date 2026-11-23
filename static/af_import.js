/* v1.3.20 AF 校對清冊 ZIP 匯入（aftool 頁面端）
 * 1. AF 上傳框接受 .zip：送 /af/unzip 在伺服器記憶體中解開（預設公版密碼），
 *    取回 ZIP 內原始檔名的 Excel，再交給既有 setFile('af', …)。
 * 2. 公版密碼解不開時，請使用者輸入密碼（可選擇記在本機 localStorage，不存伺服器）。
 * 3. 接收 fund.dgpa.gov.tw 書籤（af_oneclick.js）以 postMessage 傳來的 ZIP。
 */
(function () {
  'use strict';
  var FUND_ORIGIN = 'https://fund.dgpa.gov.tw';
  var PW_KEY = 'aftool.afZipPassword';
  var AFTOOL = location.origin;
  var origSetFile = window.setFile;
  if (typeof origSetFile !== 'function') return;

  function $(id) { return document.getElementById(id); }
  function isZip(f) {
    var n = (f && f.name || '').toLowerCase();
    return /\.zip$/.test(n) || f.type === 'application/zip' || f.type === 'application/x-zip-compressed';
  }
  function getSavedPw() { try { return localStorage.getItem(PW_KEY) || ''; } catch (e) { return ''; } }
  function setSavedPw(v) { try { v ? localStorage.setItem(PW_KEY, v) : localStorage.removeItem(PW_KEY); } catch (e) {} }

  function msgBox() {
    var el = $('afi-msg');
    if (!el) {
      el = document.createElement('div');
      el.id = 'afi-msg'; el.className = 'afi-msg'; el.setAttribute('role', 'status');
      var dz = $('drop-af'); dz.parentNode.insertBefore(el, dz.nextSibling);
    }
    return el;
  }
  function showMsg(kind, text) { var el = msgBox(); el.className = 'afi-msg ' + kind; el.textContent = text; }
  function banner(kind, text) {
    var b = $('afi-banner');
    if (!b) {
      b = document.createElement('div'); b.id = 'afi-banner'; b.className = 'afi-banner';
      b.setAttribute('role', 'status');
      var main = document.querySelector('main.container');
      main.insertBefore(b, main.firstChild);
    }
    b.className = 'afi-banner' + (kind === 'err' ? ' err' : '');
    b.style.display = 'block'; b.textContent = text;
  }
  function hidePw() { var p = $('afi-pw'); if (p) p.remove(); }

  async function unzipAndSet(file, password, fromPrompt) {
    hidePw();
    showMsg('info', '正在解開 ZIP…');
    var fd = new FormData();
    fd.append('file', file, file.name || 'AF.zip');
    var pw = password || getSavedPw();
    if (pw) fd.append('password', pw);
    var res, data;
    try {
      res = await fetch('/af/unzip', { method: 'POST', body: fd, credentials: 'same-origin' });
    } catch (e) {
      showMsg('err', '無法連線到 aftool，請稍後再試。');
      return false;
    }
    if (res.ok) {
      var blob = await res.blob();
      var name = decodeURIComponent(res.headers.get('X-AF-Filename') || 'AF校對清冊.xls');
      var used = res.headers.get('X-AF-Password');
      if (fromPrompt && fromPrompt.remember && used === 'custom') setSavedPw(pw);
      var xls = new File([blob], name, { type: 'application/vnd.ms-excel' });
      origSetFile('af', xls);
      var n = $('name-af'); if (n) n.textContent = '✔ ' + name + '（由 ZIP 解開）';
      showMsg('ok', '已解開 ZIP：' + name + (used === 'public' ? '（公版密碼）' : used === 'custom' ? '（您輸入的密碼）' : ''));
      return true;
    }
    try { data = await res.json(); } catch (e) { data = {}; }
    if (data.code === 'password_required' || data.code === 'bad_password') {
      if (!password && getSavedPw() && data.code === 'bad_password') setSavedPw('');
      askPassword(file, data.error || '請輸入密碼');
      return false;
    }
    showMsg('err', data.error || 'ZIP 解開失敗，請重新從 AF 系統下載。');
    return false;
  }

  function askPassword(file, text) {
    hidePw();
    var tpl = $('afi-pw-tpl'); if (!tpl) return;
    var node = tpl.content.firstElementChild.cloneNode(true);
    var dz = $('drop-af'); dz.parentNode.insertBefore(node, dz.nextSibling);
    $('afi-pw-text').textContent = text;
    showMsg('err', '');
    var input = $('afi-pw-input');
    input.focus();
    var go = function () {
      var v = input.value;
      if (!v) { input.focus(); return; }
      var remember = $('afi-pw-remember').checked;
      unzipAndSet(file, v, { remember: remember });
    };
    $('afi-pw-ok').onclick = go;
    input.onkeydown = function (e) { if (e.key === 'Enter') { e.preventDefault(); go(); } };
    $('afi-pw-cancel').onclick = function () { hidePw(); showMsg('info', '已取消。'); };
  }

  // 包裝既有 setFile：AF 收到 .zip 時改走解壓流程，其餘維持原樣
  window.setFile = function (key, f) {
    if (key === 'af' && f && isZip(f)) { unzipAndSet(f); return; }
    return origSetFile.apply(this, arguments);
  };

  // ── 書籤 ──
  var BOOKMARKLET = "javascript:(function(){var A='" + AFTOOL + "';" +
    "if(location.hostname!=='fund.dgpa.gov.tw'){alert('請先登入 eCPA，打開 AF 系統（fund.dgpa.gov.tw）的「資料校對及報表列印」頁面，再按這個書籤。');return;}" +
    "var w=null;try{w=window.open('','aftool_af_import');var b=!w||w.closed||w.location.href==='about:blank';}catch(e){b=false;}if(b){w=window.open(A+'/?import=af','aftool_af_import');}if(!w){alert('瀏覽器擋住了 aftool 視窗，請允許 fund.dgpa.gov.tw 開啟彈出式視窗後再按一次。');return;}" +
    "window.__aftoolWin=w;var s=document.createElement('script');s.src=A+'/static/af_oneclick.js?t='+Date.now();" +
    "s.onerror=function(){alert('載入 aftool 一鍵匯入程式失敗，請確認網路連線。');};(document.body||document.documentElement).appendChild(s);})();";

  function setupUi() {
    var inp = $('inp-af');
    if (inp) {
      inp.setAttribute('accept', '.zip,.xlsx,.xls');
      inp.setAttribute('aria-label', '選擇本月 AF 資料（AF 系統下載的 .zip，或 .xlsx／.xls）');
    }
    var hint = document.querySelector('#drop-af .drop-hint');
    if (hint) hint.textContent = '或點擊選擇檔案　.zip / .xlsx / .xls';
    var desc = $('desc-af');
    if (desc) {
      var em = desc.querySelectorAll('.em');
      if (em.length) em[em.length - 1].textContent = '.zip（AF 下載檔，免解壓）　.xlsx　.xls';
      var tpl = $('afi-guide-tpl');
      if (tpl && !$('afi-guide')) desc.parentNode.insertBefore(tpl.content.firstElementChild.cloneNode(true), desc.nextSibling);
    }
    var bm = $('afi-bookmarklet');
    if (bm) {
      bm.setAttribute('href', BOOKMARKLET);
      bm.addEventListener('click', function (e) {
        e.preventDefault();
        alert('請把這個按鈕「拖曳」到瀏覽器的書籤列，之後在 AF 系統頁面按書籤使用。');
      });
    }
    var cp = $('afi-copy-bm');
    if (cp) cp.addEventListener('click', function () {
      var done = function () { cp.textContent = '已複製'; setTimeout(function () { cp.textContent = '複製書籤網址'; }, 2000); };
      if (navigator.clipboard) navigator.clipboard.writeText(BOOKMARKLET).then(done, function () { window.prompt('請複製以下內容：', BOOKMARKLET); });
      else window.prompt('請複製以下內容：', BOOKMARKLET);
    });
    if (/[?&]import=af\b/.test(location.search)) {
      banner('', '等待 AF 系統傳送校對清冊…（請不要關閉 AF 系統視窗）');
    }
  }

  // ── 接收書籤送來的 ZIP ──
  window.addEventListener('message', function (e) {
    if (e.origin !== FUND_ORIGIN || !e.data || typeof e.data !== 'object') return;
    var d = e.data;
    if (d.type === 'af-hello') {
      e.source.postMessage({ type: 'aftool-ready' }, FUND_ORIGIN);
      return;
    }
    if (d.type === 'af-status' && typeof d.message === 'string') {
      banner('', 'AF 系統：' + d.message.slice(0, 300));
    } else if (d.type === 'af-error' && typeof d.message === 'string') {
      banner('err', '一鍵匯入失敗：' + d.message.slice(0, 500) + '（也可以改用手動下載後拖進下方框）');
    } else if (d.type === 'af-zip' && d.blob instanceof Blob) {
      if (d.blob.size > 16 * 1024 * 1024) { banner('err', '收到的檔案過大，已略過。'); return; }
      var name = (typeof d.name === 'string' && d.name) ? d.name.slice(0, 200) : 'AF校對清冊.zip';
      if (!/\.zip$/i.test(name)) name += '.zip';
      banner('', '已從 AF 系統收到檔案，正在解開…');
      e.source.postMessage({ type: 'aftool-received' }, FUND_ORIGIN);
      unzipAndSet(new File([d.blob], name, { type: 'application/zip' })).then(function (ok) {
        if (ok) {
          banner('', '✅ 已從 AF 系統匯入校對清冊：' + (files.af ? files.af.name : '') + '。接著上傳固定清冊或薪資清冊後即可執行。');
          var dz = $('drop-af'); if (dz && dz.scrollIntoView) dz.scrollIntoView({ block: 'center' });
          try { window.focus(); } catch (err) {}
        } else {
          banner('err', '已收到 AF 檔案，但無法自動解開，請看下方 AF 上傳框的提示。');
        }
      });
    }
  });

  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', setupUi);
  else setupUi();
})();
