/* v1.3.25：點一下複製（AF 公版密碼等）。
 * 任何帶 data-copy 的元素：滑鼠點擊或鍵盤 Enter／空白鍵 → 複製 data-copy 的值，
 * 旁邊浮出「已複製」約 1.5 秒（絕對定位，不推動版面）。用事件委派，所以由
 * <template> 複製出來的說明區塊也適用；不依賴其他腳本。
 */
(function () {
  'use strict';
  function fallbackCopy(text) {
    var ta = document.createElement('textarea');
    ta.value = text;
    ta.setAttribute('readonly', '');
    ta.style.cssText = 'position:fixed;top:0;left:-9999px;opacity:0';
    document.body.appendChild(ta);
    var sel = document.getSelection(), saved = sel && sel.rangeCount ? sel.getRangeAt(0) : null;
    ta.select(); ta.setSelectionRange(0, text.length);
    var ok = false;
    try { ok = document.execCommand('copy'); } catch (e) { ok = false; }
    document.body.removeChild(ta);
    if (saved && sel) { sel.removeAllRanges(); sel.addRange(saved); }
    return ok;
  }
  function copyText(text) {
    if (navigator.clipboard && window.isSecureContext) {
      return navigator.clipboard.writeText(text).then(function () { return true; },
        function () { return fallbackCopy(text); });
    }
    return Promise.resolve(fallbackCopy(text));
  }
  function feedback(el, ok) {
    var tip = el.querySelector('.afi-copied');
    if (!tip) {
      tip = document.createElement('span');
      tip.className = 'afi-copied'; tip.setAttribute('role', 'status'); tip.setAttribute('aria-live', 'polite');
      el.appendChild(tip);
    }
    tip.textContent = ok ? '已複製' : '請手動選取複製';
    tip.classList.toggle('err', !ok);
    tip.classList.add('show');
    clearTimeout(el._afiCopyTimer);
    el._afiCopyTimer = setTimeout(function () { tip.classList.remove('show'); }, 1500);
  }
  function activate(e, el) {
    e.preventDefault();
    e.stopPropagation();          // 不要觸發外層 <details>／<summary> 展開收合
    copyText(el.getAttribute('data-copy') || '').then(function (ok) { feedback(el, ok); });
  }
  document.addEventListener('click', function (e) {
    var el = e.target.closest && e.target.closest('[data-copy]');
    if (el) activate(e, el);
  }, true);
  document.addEventListener('keydown', function (e) {
    if (e.key !== 'Enter' && e.key !== ' ' && e.key !== 'Spacebar') return;
    var el = e.target.closest && e.target.closest('[data-copy]');
    if (el) activate(e, el);
  }, true);
})();
