/* aftool「AF → aftool」書籤一鍵匯入（v1.3.20，於 fund.dgpa.gov.tw 已登入的頁面執行）
 *
 * 流程：找到「報表列印」表單（目前頁、同源 frame／iframe，或按「報表列印」後彈出的視窗）
 *   → 報表種類選「校對清冊(A)」、報表格式 EXCEL、人員類別「全部」、身分證號密碼填公版密碼
 *   → 以 fetch 送出與「下載報表」按鈕相同的 ASP.NET postback，直接拿到 ZIP（不存檔）
 *   → 以 postMessage 把 ZIP 交給 aftool 視窗（只送往 aftool 來源）。
 * 只使用瀏覽器現有的登入狀態；不讀取、不傳送任何 cookie 或帳密，也不送出其他表單。
 * 元素一律以畫面文字定位（「下載報表」「校對清冊(A)」「EXCEL」「密碼」），以適應頁面改版。
 */
(function () {
  'use strict';
  var AFTOOL = 'https://aftool.leotsai.me';
  var PUBLIC_PASSWORD = '!@Abcdefg1234567';
  var VERSION = '1.3.20';
  if (window.__aftoolOneClickRunning) { alert('一鍵匯入進行中，請稍候…'); return; }
  window.__aftoolOneClickRunning = true;

  var target = window.__aftoolWin || null;
  var ready = false, queue = [], received = false;

  window.addEventListener('message', function (e) {
    if (e.origin !== AFTOOL || !e.data) return;
    if (e.data.type === 'aftool-ready') { ready = true; target = e.source; flush(); }
    if (e.data.type === 'aftool-received') received = true;
  });
  function flush() {
    if (!ready || !target) return;
    while (queue.length) { try { target.postMessage(queue.shift(), AFTOOL); } catch (err) {} }
  }
  function post(m) { queue.push(m); flush(); }
  function status(text) { post({ type: 'af-status', message: text }); }
  function fail(text) {
    post({ type: 'af-error', message: text });
    window.__aftoolOneClickRunning = false;
    alert('AF → aftool 一鍵匯入失敗：\n' + text + '\n\n可以改用手動「下載報表」，再把 ZIP 拖進 aftool。');
  }
  var helloTimer = setInterval(function () {
    if (ready || !target) return;
    try { target.postMessage({ type: 'af-hello', v: VERSION }, AFTOOL); } catch (err) {}
  }, 300);
  setTimeout(function () { clearInterval(helloTimer); }, 60000);

  function sleep(ms) { return new Promise(function (r) { setTimeout(r, ms); }); }
  function norm(s) {
    return String(s || '').replace(/[\s\u3000]/g, '').replace(/（/g, '(').replace(/）/g, ')').toUpperCase();
  }
  function textOf(el) {
    if (!el) return '';
    var t = el.tagName === 'INPUT' ? (el.value || el.alt || el.title) : (el.textContent || el.title || el.alt || '');
    return norm(t);
  }
  function rowText(el) {
    var tr = el.closest && el.closest('tr');
    var t = tr ? tr.textContent : (el.parentNode ? el.parentNode.textContent : '');
    if (el.id && el.ownerDocument) {
      var lab = el.ownerDocument.querySelector('label[for="' + el.id.replace(/"/g, '') + '"]');
      if (lab) t += lab.textContent;
    }
    return norm(t);
  }
  var CLICKABLE = 'input[type=submit],input[type=button],input[type=image],button,a,[onclick]';
  function findClickable(doc, label, exact) {
    var want = norm(label), best = null;
    var els = doc.querySelectorAll(CLICKABLE);
    for (var i = 0; i < els.length; i++) {
      var t = textOf(els[i]);
      if (!t) continue;
      if (t === want) return els[i];
      if (!exact && t.indexOf(want) >= 0 && t.length < want.length + 8 && !best) best = els[i];
    }
    return best;
  }

  // 目前視窗、同源 frame／iframe、上層視窗、被我們捕捉到的彈出視窗
  // 彈出視窗清單掛在 window 上，第二次按書籤（多校切換）時仍找得到上次的「報表列印」視窗
  var popups = window.__aftoolPopups = window.__aftoolPopups || [];
  function collectDocs() {
    var out = [], seen = [];
    function walk(win) {
      try {
        if (!win || win.closed || seen.indexOf(win) >= 0) return;
        seen.push(win);
        var d = win.document;
        if (!d || !d.location || d.location.href === 'about:blank') return;
        out.push(d);
        var fs = d.querySelectorAll('iframe,frame');
        for (var i = 0; i < fs.length; i++) walk(fs[i].contentWindow);
      } catch (err) { /* 跨網域 frame 略過 */ }
    }
    walk(window);
    try { walk(window.top); } catch (err) {}
    popups.forEach(walk);
    try { if (window.opener) walk(window.opener); } catch (err) {}
    return out;
  }
  function findReportDoc() {
    var docs = collectDocs();
    for (var i = 0; i < docs.length; i++) {
      var btn = findClickable(docs[i], '下載報表');
      if (btn) return { doc: docs[i], btn: btn };
    }
    return null;
  }
  function findInDocs(label, exact) {
    var docs = collectDocs();
    for (var i = 0; i < docs.length; i++) {
      var el = findClickable(docs[i], label, exact);
      if (el) return el;
    }
    return null;
  }
  function capturePopups(win) {
    try {
      if (win.__aftoolOpenPatched) return;
      var orig = win.open;
      var list = window.__aftoolPopups;
      win.open = function () { var p = orig.apply(this, arguments); if (p && list.indexOf(p) < 0) list.push(p); return p; };
      win.__aftoolOpenPatched = true;
    } catch (err) {}
  }
  async function clickAndWait(el, check, timeoutMs) {
    capturePopups(el.ownerDocument.defaultView);
    el.click();
    var end = Date.now() + timeoutMs;
    while (Date.now() < end) {
      await sleep(500);
      var r = check();
      if (r) return r;
    }
    return null;
  }

  async function locateReportForm() {
    var r = findReportDoc();
    if (r) return r;
    var printBtn = findInDocs('報表列印', true);
    if (!printBtn) {
      var menu = findInDocs('資料校對及報表列印');
      if (!menu) {
        var top = findInDocs('用人費用資料維護');
        if (top) {
          status('展開「用人費用資料維護」選單…');
          await clickAndWait(top, function () { return findInDocs('資料校對及報表列印'); }, 5000);
          menu = findInDocs('資料校對及報表列印');
        }
      }
      if (menu) {
        status('切換到「資料校對及報表列印」頁面…（若畫面換頁後沒有繼續，請再按一次書籤）');
        printBtn = await clickAndWait(menu, function () { return findReportDoc() || findInDocs('報表列印', true); }, 20000);
        if (printBtn && printBtn.doc) return printBtn;
      }
    }
    if (!printBtn) return null;
    status('開啟「報表列印」視窗…');
    return await clickAndWait(printBtn, findReportDoc, 20000);
  }

  // ── 表單操作（同時適用於畫面上的表單與 fetch 取回後解析的表單）──
  function selectOption(sel, pred) {
    for (var i = 0; i < sel.options.length; i++) {
      if (pred(norm(sel.options[i].text), sel.options[i])) {
        var changed = sel.selectedIndex !== i;
        sel.selectedIndex = i;
        for (var j = 0; j < sel.options.length; j++) {
          if (j === i) sel.options[j].setAttribute('selected', 'selected'); else sel.options[j].removeAttribute('selected');
        }
        return changed;
      }
    }
    return null;
  }
  function isPostBack(el) { return /__doPostBack|WebForm_DoPostBack/.test(el.getAttribute('onchange') || el.getAttribute('onclick') || ''); }
  function fire(el, type) {
    try { el.dispatchEvent(new (el.ownerDocument.defaultView || window).Event(type, { bubbles: true })); } catch (err) {}
  }

  function applyChoices(form, live) {
    var report = null, autopost = [];
    var sels = form.querySelectorAll('select');
    for (var i = 0; i < sels.length; i++) {
      var s = sels[i], ch = null, row = rowText(s);
      if (!report) {
        ch = selectOption(s, function (t) { return t.indexOf('校對清冊(A)') >= 0; });
        if (ch !== null) report = s;
      }
      if (ch === null && /EXCEL/.test(norm(s.textContent)) && /格式|EXCEL/.test(row)) {
        ch = selectOption(s, function (t) { return t.indexOf('EXCEL') >= 0; });
      }
      if (ch === null && row.indexOf('人員類別') >= 0) {
        ch = selectOption(s, function (t) { return t === '全部' || t.indexOf('全部') === 0; });
      }
      if (ch) {
        if (isPostBack(s)) autopost.push(s);
        else if (live) fire(s, 'change');
      }
    }
    // 報表格式若是選項按鈕
    var radios = form.querySelectorAll('input[type=radio]');
    for (var k = 0; k < radios.length; k++) {
      var rd = radios[k];
      var lab = rd.id ? form.ownerDocument.querySelector('label[for="' + rd.id + '"]') : null;
      var lt = norm((lab && lab.textContent) || (rd.nextSibling && rd.nextSibling.textContent) || rd.value);
      if (lt.indexOf('EXCEL') >= 0 || norm(rd.value) === 'EXCEL') {
        if (!rd.checked) { rd.checked = true; rd.setAttribute('checked', 'checked'); if (isPostBack(rd)) autopost.push(rd); }
      }
    }
    // 密碼：表單內所有密碼欄（含確認密碼）
    var pws = form.querySelectorAll('input[type=password]');
    if (!pws.length) {
      var texts = form.querySelectorAll('input[type=text],input:not([type])');
      pws = Array.prototype.filter.call(texts, function (t) { return rowText(t).indexOf('密碼') >= 0; });
    }
    Array.prototype.forEach.call(pws, function (p) {
      p.disabled = false; p.value = PUBLIC_PASSWORD; p.setAttribute('value', PUBLIC_PASSWORD);
      if (live) { fire(p, 'input'); fire(p, 'change'); }
    });
    return { report: report, autopost: autopost, pwCount: pws.length };
  }

  function serialize(form, extra) {
    var p = new URLSearchParams();
    var els = form.elements;
    for (var i = 0; i < els.length; i++) {
      var el = els[i], tag = el.tagName, type = (el.type || '').toLowerCase();
      if (!el.name || el.disabled) continue;
      if (tag === 'INPUT' && /^(submit|button|image|reset|file)$/.test(type)) continue;
      if (tag === 'BUTTON') continue;
      if ((type === 'checkbox' || type === 'radio') && !el.checked) continue;
      if (tag === 'SELECT') {
        for (var j = 0; j < el.options.length; j++) if (el.options[j].selected) p.append(el.name, el.options[j].value);
        continue;
      }
      p.append(el.name, el.value);
    }
    Object.keys(extra || {}).forEach(function (k) { p.set(k, extra[k]); });
    return p;
  }
  function postbackTarget(el) {
    var code = (el.getAttribute('href') || '') + ' ' + (el.getAttribute('onclick') || '') + ' ' + (el.getAttribute('onchange') || '');
    var m = code.match(/__doPostBack\(\s*['"]([^'"]+)['"]\s*,\s*['"]([^'"]*)['"]/) ||
            code.match(/WebForm_PostBackOptions\(\s*["']([^"']+)["']\s*,\s*["']([^"']*)["']/);
    return m ? { target: m[1], arg: m[2] } : null;
  }
  function popupUrl(el, base) {
    var code = (el.getAttribute('onclick') || '') + ' ' + (el.getAttribute('href') || '');
    var m = code.match(/(?:window\.open|location(?:\.href)?\s*=|location\.assign)\s*\(?\s*['"]([^'"]+\.(?:aspx|ashx|zip)[^'"]*)['"]/i);
    return m ? new URL(m[1], base).href : null;
  }
  function formAction(form, doc) {
    return new URL(form.getAttribute('action') || doc.location.href, doc.location.href).href;
  }

  function isZipBytes(buf) { var b = new Uint8Array(buf, 0, Math.min(4, buf.byteLength)); return b[0] === 0x50 && b[1] === 0x4b && b[2] === 3 && b[3] === 4; }
  function fileNameFrom(res) {
    var cd = res.headers.get('content-disposition') || '';
    var m = cd.match(/filename\*\s*=\s*UTF-8''([^;]+)/i) || cd.match(/filename\s*=\s*"?([^";]+)"?/i);
    if (!m) return 'AF校對清冊.zip';
    try { return decodeURIComponent(m[1]); } catch (err) { return m[1]; }
  }
  function decodeHtml(buf, res) {
    var ct = res.headers.get('content-type') || '';
    var enc = (ct.match(/charset=([\w-]+)/i) || [])[1] || 'utf-8';
    try { return new TextDecoder(enc).decode(buf); } catch (err) { return new TextDecoder('utf-8').decode(buf); }
  }
  function findFollowUrl(html, base) {
    var pats = [/window\.open\(\s*['"]([^'"]+)['"]/i, /location(?:\.href)?\s*=\s*['"]([^'"]+)['"]/i,
                /<iframe[^>]+src\s*=\s*['"]([^'"]+\.(?:ashx|aspx|zip)[^'"]*)['"]/i,
                /href\s*=\s*['"]([^'"]+\.(?:zip|ashx)[^'"]*)['"]/i];
    for (var i = 0; i < pats.length; i++) {
      var m = html.match(pats[i]);
      if (m && !/ecpa\.dgpa\.gov\.tw\/?$/.test(m[1]) && !/^javascript:/i.test(m[1])) {
        try { var u = new URL(m[1].replace(/&amp;/g, '&'), base); if (u.origin === location.origin) return u.href; } catch (err) {}
      }
    }
    return null;
  }
  function alertText(html) {
    var m = html.match(/alert\(\s*['"]([^'"]{2,200})['"]/);
    return m ? m[1].replace(/\\n/g, ' ') : '';
  }

  async function fetchZip(url, opts, depth) {
    var res = await fetch(url, Object.assign({ credentials: 'include', redirect: 'follow' }, opts || {}));
    var buf = await res.arrayBuffer();
    if (isZipBytes(buf)) return { blob: new Blob([buf], { type: 'application/zip' }), name: fileNameFrom(res) };
    var html = decodeHtml(buf, res);
    var msg = alertText(html);
    if (msg && /錯誤|失敗|請|無|逾時|閒置|登入|密碼/.test(msg)) return { error: 'AF 系統訊息：「' + msg + '」', html: html };
    if ((depth || 0) < 2) {
      var next = findFollowUrl(html, res.url || url);
      if (next) return fetchZip(next, { method: 'GET' }, (depth || 0) + 1);
    }
    return { html: html, res: res, error: msg ? 'AF 系統訊息：「' + msg + '」' : '' };
  }

  async function run() {
    if (!target) {
      fail('瀏覽器擋住了 aftool 視窗，請允許 fund.dgpa.gov.tw 開啟彈出式視窗後再按一次書籤。');
      return;
    }
    status('尋找「報表列印」…');
    var found = await locateReportForm();
    if (!found) {
      fail('找不到「報表列印」或「下載報表」按鈕。請先進入「用人費用資料維護 → 資料校對及報表列印」頁面再按書籤。');
      return;
    }
    var doc = found.doc, btn = found.btn;
    var form = btn.form || btn.closest('form') || doc.forms[0];
    if (!form) { fail('「報表列印」頁面沒有可送出的表單。'); return; }
    var choice = applyChoices(form, true);
    if (!choice.report) { fail('報表種類找不到「校對清冊(A)」選項。'); return; }
    if (!choice.pwCount) { fail('找不到「身分證號密碼」欄位。'); return; }
    status('已選「校對清冊(A)」、EXCEL，並填入公版密碼，正在取得報表…');

    // 有 AutoPostBack 的下拉選單時，先在背景送一次變更，再用新回傳的表單下載
    var curForm = form, curBtn = btn, base = formAction(form, doc);
    for (var round = 0; round < 3 && choice.autopost.length; round++) {
      var ap = choice.autopost[0];
      var t = postbackTarget(ap) || { target: ap.name, arg: '' };
      var r0 = await fetch(base, { method: 'POST', credentials: 'include', body: serialize(curForm, { __EVENTTARGET: t.target, __EVENTARGUMENT: t.arg }) });
      var html0 = await r0.text();
      var pd = new DOMParser().parseFromString(html0, 'text/html');
      var nb = findClickable(pd, '下載報表');
      if (!nb) { fail('切換報表選項後頁面結構改變，找不到「下載報表」。'); return; }
      curBtn = nb; curForm = nb.form || nb.closest('form') || pd.forms[0];
      base = new URL(curForm.getAttribute('action') || base, r0.url || base).href;
      choice = applyChoices(curForm, false);
    }

    var result;
    var direct = popupUrl(curBtn, base);
    var type = (curBtn.type || '').toLowerCase();
    if (direct && !postbackTarget(curBtn) && curBtn.tagName !== 'A' && type !== 'submit') {
      result = await fetchZip(direct, { method: 'GET' });
    } else {
      var extra = {};
      var pb = postbackTarget(curBtn);
      if (curBtn.tagName === 'INPUT' && type === 'image' && curBtn.name) {
        extra[curBtn.name + '.x'] = '1'; extra[curBtn.name + '.y'] = '1';
      } else if ((curBtn.tagName === 'INPUT' || curBtn.tagName === 'BUTTON') && curBtn.name) {
        extra[curBtn.name] = curBtn.value || '';
      }
      if (pb && !(curBtn.name && pb.target === curBtn.name && type === 'submit')) {
        extra.__EVENTTARGET = pb.target; extra.__EVENTARGUMENT = pb.arg;
      } else if (!pb) {
        extra.__EVENTTARGET = extra.__EVENTTARGET || ''; extra.__EVENTARGUMENT = '';
      }
      result = await fetchZip(base, { method: 'POST', body: serialize(curForm, extra) });
    }
    if (!result.blob) {
      fail((result.error || 'AF 系統回應的不是 ZIP 檔。') + '（如果持續發生，請把報表列印視窗的「另存網頁」HTML 提供給 aftool 維護者調整）');
      return;
    }
    status('已取得報表，傳送到 aftool…');
    post({ type: 'af-zip', blob: result.blob, name: result.name });
    window.__aftoolOneClickRunning = false;
    for (var w = 0; w < 40 && !received; w++) await sleep(250);
    if (!received) { alert('已取得 AF 報表，但 aftool 視窗沒有回應。請確認 aftool 分頁已開啟，再按一次書籤。'); return; }
    try { target.focus(); } catch (err) {}
  }

  run().catch(function (err) { fail('執行時發生錯誤：' + (err && err.message || err)); });
})();
