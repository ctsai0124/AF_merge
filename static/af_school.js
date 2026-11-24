/*
 * v1.3.23：上傳 AF 時自動帶入「該校」的固定清冊，並防止 AF 與固定清冊不同校。
 *
 * 1. AF 一進來（xlsx／xls、ZIP 解壓、擴充功能、範例資料皆同）就呼叫 /af/identify，
 *    以 AF 內容或檔名的機關代碼判定學校。
 * 2. 已登入人事資料倉庫、且這所學校在帳號的學校清單中：自動切換「選擇學校」，
 *    (a) 該校有已存的固定清冊檔案 → 套用；(b) 沒有 → 套用該校人員的序號清冊。
 *    (c) 帳號沒有這所學校／沒登入 → 明確提示這所學校尚無固定清冊。
 * 3. 兩份檔案都在時，再比對人員重疊（只回傳人數），AF 與固定清冊的學校不同或
 *    人員幾乎不重疊時，顯示醒目警告；按「執行排序／產製互核結果表」前需再確認。
 * 使用者手動上傳的固定清冊、或自己在下拉選單改選的學校，一律尊重，不會被自動蓋掉。
 * 人事資料倉庫相關函式（acct*）不存在時（未啟用該功能）只做第 3 點。
 */
(function () {
  'use strict';
  if (typeof files === 'undefined' || typeof window.checkReady !== 'function') return;

  var state = {
    af: null,            // /af/identify 結果
    afFile: null, rosterFile: null,
    roster: { source: '', schoolId: '' },   // source: saved | warehouse | manual | sample
    overlap: null, pending: 0, note: null, ack: null, busy: false,
  };
  var tag = null;                 // 下一次固定清冊變動的來源標記
  var userPicked = false;         // 使用者自己操作「選擇學校」下拉選單

  function $(id) { return document.getElementById(id); }
  function cred() { return typeof acctCred !== 'undefined' ? acctCred : null; }
  function savedId() { return typeof acctSavedRosterSchoolId !== 'undefined' ? acctSavedRosterSchoolId : null; }
  function selectedId() { return typeof acctSelectedSchoolId !== 'undefined' ? acctSelectedSchoolId : null; }
  function isSample() { return typeof isSampleDataActive !== 'undefined' && isSampleDataActive; }
  function hasAcct() { return typeof window.acctCheckSavedRoster === 'function' && !!$('acct-school-select'); }
  function options() {
    var sel = $('acct-school-select');
    if (!sel) return [];
    return Array.prototype.map.call(sel.options, function (o) {
      return { schoolId: o.value, name: (o.textContent || '').replace(/（[^（）]*）\s*$/, '').trim() };
    });
  }
  function optionName(id) {
    var o = options().filter(function (x) { return x.schoolId === id; })[0];
    return o ? o.name : '';
  }
  function afLabel() {
    var a = state.af || {};
    return a.school_name || a.code || 'AF 的學校';
  }

  // ── 畫面：固定清冊區塊頂端的提示框 ──
  var box = document.createElement('div');
  box.id = 'af-school-box';
  box.setAttribute('role', 'status');
  box.style.cssText = 'display:none;margin:0 0 10px;padding:10px 12px;border-radius:6px;font-size:13px;line-height:1.6';
  var desc = $('desc-roster');
  if (desc && desc.parentNode) desc.parentNode.insertBefore(box, desc);

  var STYLES = {
    info: 'background:#eef4fb;border:1px solid #9fb8d6;color:#1a3a5c',
    ok: 'background:#eafaf0;border:1px solid #2a6e3f;color:#145a30',
    warn: 'background:#fdeeee;border:2px solid #b00020;color:#7a0016;font-weight:600',
  };

  function mismatch() {
    if (!state.af || !files.af || !files.roster || isSample()) return null;
    var target = state.af.account_school_id;
    var rid = state.roster.schoolId;
    var ov = state.overlap;
    if (rid && target && rid !== target) {
      return { roster: optionName(rid) || rid, why: 'school' };
    }
    if (rid && !target && state.af.code && !(ov && ov.same_school === true)) {
      return { roster: optionName(rid) || rid, why: 'school' };
    }
    if (ov && ov.same_school === false) {
      return { roster: rid ? (optionName(rid) || rid) : '', why: 'people', ov: ov };
    }
    return null;
  }

  function mismatchText(m) {
    var t = 'AF 是「' + afLabel() + '」，但目前的固定清冊'
      + (m.roster ? '是「' + m.roster + '」' : '與 AF 的人員幾乎不重疊')
      + (m.ov ? '（AF ' + m.ov.af_total + ' 人中只有 ' + m.ov.matched + ' 人在固定清冊內）' : '')
      + '，兩份很可能不是同一所學校。';
    if (cred() && hasAcct() && state.af && !state.af.account_school_id) {
      t += '您的人事資料倉庫帳號沒有「' + afLabel() + '」，這所學校尚未登錄固定清冊。';
    }
    return t;
  }

  function render() {
    var m = mismatch();
    var html = '', kind = '';
    if (m) {
      kind = 'warn';
      html = '⚠ ' + esc(mismatchText(m)) + '請改用「' + esc(afLabel()) + '」的固定清冊；'
        + '若確定無誤，按執行時會再請您確認一次。';
      if (cred() && state.af.account_school_id && state.roster.schoolId !== state.af.account_school_id) {
        html += ' <button type="button" id="af-school-switch" style="margin-left:6px;font-size:12px;padding:3px 8px">改用「'
          + esc(optionName(state.af.account_school_id) || afLabel()) + '」的固定清冊</button>';
      }
    } else if (state.note) {
      kind = state.note.kind;
      html = esc(state.note.text);
    }
    if (!html) { box.style.display = 'none'; box.innerHTML = ''; return; }
    box.innerHTML = html;
    box.setAttribute('data-kind', kind);
    box.style.cssText = 'display:block;margin:0 0 10px;padding:10px 12px;border-radius:6px;font-size:13px;line-height:1.6;' + STYLES[kind];
    var btn = $('af-school-switch');
    if (btn) btn.onclick = function () { switchTo(state.af.account_school_id, true); };
  }
  function esc(s) {
    return String(s == null ? '' : s).replace(/[&<>"']/g, function (c) {
      return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c];
    });
  }
  function note(kind, text) { state.note = text ? { kind: kind, text: text } : null; render(); }

  // ── 後端判定 ──
  function identify(withRoster) {
    var af = files.af;
    if (!af) return Promise.resolve(null);
    var fd = new FormData();
    fd.append('af', af, af.name);
    fd.append('schools', JSON.stringify(options()));
    if (withRoster && files.roster) fd.append('roster', files.roster, files.roster.name);
    state.pending++;
    return fetch('/af/identify', { method: 'POST', body: fd })
      .then(function (r) { return r.ok ? r.json() : null; })
      .catch(function () { return null; })
      .then(function (d) { state.pending--; return d; });
  }

  function refreshOverlap() {
    if (!files.af || !files.roster || isSample()) { state.overlap = null; render(); return Promise.resolve(); }
    var af = files.af, ro = files.roster;
    return identify(true).then(function (d) {
      if (files.af !== af || files.roster !== ro) return;   // 期間又換了檔案
      if (d) {
        state.overlap = d.overlap || null;
        if (!state.af) state.af = d;
        else state.af.account_school_id = d.account_school_id || state.af.account_school_id;
      }
      render();
    });
  }

  // ── 自動切換：(a) 已存檔案 → (b) 序號清冊 → (c) 提示 ──
  function switchTo(target, manualButton) {
    if (!target || !hasAcct()) return Promise.resolve(false);
    var sel = $('acct-school-select');
    var name = optionName(target) || afLabel();
    sel.value = target;
    state.busy = true;
    note('info', '正在帶入「' + name + '」的固定清冊…');
    return Promise.resolve(origCheck(target)).then(function () {
      if (state.roster.schoolId === target && files.roster) return 'saved';
      return Promise.resolve(window.acctApplyAllStaffRoster()).then(function () {
        return state.roster.schoolId === target && files.roster ? 'warehouse' : '';
      });
    }).catch(function () { return ''; }).then(function (how) {
      state.busy = false;
      if (how === 'saved') note('ok', '✅ 已依 AF 學校自動切換為「' + name + '」已存的固定清冊檔案。如需使用別份，直接重新上傳即可覆蓋。');
      else if (how === 'warehouse') note('ok', '✅ 「' + name + '」沒有已存的固定清冊檔案，已改用人事資料倉庫中該校人員的序號清冊。如需使用別份，直接重新上傳即可覆蓋。');
      else note('warn', '⚠ 「' + name + '」尚未登錄固定清冊（人事資料倉庫沒有存檔，也無法產生序號清冊），請手動上傳這所學校的固定清冊。');
      return refreshOverlap().then(function () { return !!how; });
    });
  }

  function autoSwitch() {
    if (!state.af || isSample()) return Promise.resolve();
    var target = state.af.account_school_id;
    var name = afLabel();
    var rosterIsAuto = !files.roster || state.roster.source === 'saved' || state.roster.source === 'warehouse';
    if (cred() && hasAcct() && options().length) {
      if (!target) {
        note('warn', '⚠ AF 是「' + name + '」，但您的人事資料倉庫帳號沒有這所學校，這所學校還沒有登錄固定清冊。請手動上傳「' + name + '」的固定清冊。');
        return refreshOverlap();
      }
      if (files.roster && state.roster.schoolId === target) {
        note('ok', '✅ 固定清冊與 AF 同為「' + (optionName(target) || name) + '」。');
        return refreshOverlap();
      }
      if (!rosterIsAuto) { note(null); return refreshOverlap(); }   // 使用者自己上傳的：不覆蓋，只做防呆
      return switchTo(target);
    }
    if (!files.roster) {
      note('info', 'AF 學校：「' + name + '」。請上傳這所學校的固定清冊；若已在人事資料倉庫登錄，登入後會自動帶入該校的固定清冊。');
    } else {
      note(null);
    }
    return refreshOverlap();
  }

  // ── 監看檔案變動：所有上傳路徑最後都會呼叫 checkReady() ──
  var origReady = window.checkReady;
  window.checkReady = function () {
    var r = origReady.apply(this, arguments);
    var afChanged = files.af !== state.afFile;
    var rosterChanged = files.roster !== state.rosterFile;
    state.afFile = files.af; state.rosterFile = files.roster;
    if (rosterChanged) {
      state.roster = files.roster
        ? (isSample() ? { source: 'sample', schoolId: '' } : (tag || { source: 'manual', schoolId: '' }))
        : { source: '', schoolId: '' };
      tag = null;
      state.overlap = null;
      state.ack = null;
    }
    if (afChanged) {
      state.af = null; state.overlap = null; state.ack = null; note(null);
      if (files.af && !isSample()) {
        var af = files.af;
        identify(!!files.roster).then(function (d) {
          if (files.af !== af) return;
          state.af = d;
          state.overlap = d && d.overlap || null;
          if (!d) { render(); return; }
          return autoSwitch();
        });
      }
    } else if (rosterChanged && !state.busy) {
      if (state.roster.source === 'manual' && state.note && state.note.kind !== 'warn') note(null);
      refreshOverlap();
    }
    return r;
  };

  // ── 標記人事資料倉庫帶入的固定清冊屬於哪一校 ──
  var origCheck = window.acctCheckSavedRoster;
  if (typeof window.acctApplySavedRoster === 'function') {
    var origSaved = window.acctApplySavedRoster;
    window.acctApplySavedRoster = function () {
      tag = { source: 'saved', schoolId: savedId() || '' };
      var p = origSaved.apply(this, arguments);
      return Promise.resolve(p).then(function (v) { tag = null; return v; }, function (e) { tag = null; throw e; });
    };
  }
  if (typeof window.acctApplyAllStaffRoster === 'function') {
    var origAll = window.acctApplyAllStaffRoster;
    window.acctApplyAllStaffRoster = function () {
      var sel = $('acct-school-select');
      tag = { source: 'warehouse', schoolId: (sel && sel.value) || selectedId() || '' };
      var p = origAll.apply(this, arguments);
      return Promise.resolve(p).then(function (v) { tag = null; return v; }, function (e) { tag = null; throw e; });
    };
  }
  if (typeof origCheck === 'function') {
    var sel = $('acct-school-select');
    if (sel) sel.addEventListener('change', function () { userPicked = true; }, true);
    // 登入（手動或自動偵測）後，原本會替清單第一間學校自動套用存檔；
    // 已知 AF 學校時改套用 AF 那一校，帳號沒有這所學校時不套用別校的檔案。
    window.acctCheckSavedRoster = function (schoolId) {
      if (userPicked) {                 // 使用者自己改選學校：手動優先
        userPicked = false;
        return origCheck.apply(this, arguments);
      }
      if (!files.af || isSample() || state.busy) return origCheck.apply(this, arguments);
      var af = files.af;
      return identify(!!files.roster).then(function (d) {
        if (files.af !== af) return origCheck(schoolId);
        if (d) { state.af = d; state.overlap = d.overlap || null; }
        var target = d && d.account_school_id;
        if (!target) {
          note('warn', '⚠ AF 是「' + afLabel() + '」，但您的人事資料倉庫帳號沒有這所學校，所以沒有自動套用其他學校的固定清冊。請手動上傳「' + afLabel() + '」的固定清冊。');
          return;
        }
        return switchTo(target);
      });
    };
  }

  // ── 不同校防呆：執行排序／產製互核結果表前再確認 ──
  function fileKey() { return [files.af && files.af.name, files.af && files.af.size, files.af && files.af.lastModified,
    files.roster && files.roster.name, files.roster && files.roster.size, files.roster && files.roster.lastModified].join('|'); }
  document.addEventListener('click', function (e) {
    var t = e.target && e.target.closest ? e.target.closest('#btn-go, #btn-cross') : null;
    if (!t || t.disabled) return;
    if (state.pending > 0 || state.busy) {
      e.preventDefault(); e.stopImmediatePropagation();
      alert('正在核對 AF 與固定清冊是否為同一所學校，請稍候一兩秒再按一次。');
      return;
    }
    var m = mismatch();
    if (!m || state.ack === fileKey()) return;
    e.preventDefault(); e.stopImmediatePropagation();
    var msg = mismatchText(m) + '\n\n確定仍要用這份固定清冊繼續嗎？';
    if (window.confirm(msg)) {
      state.ack = fileKey();
      t.click();
    }
  }, true);

  window.aftoolSchoolRoster = { state: state, render: render };   // 供測試／除錯
})();
