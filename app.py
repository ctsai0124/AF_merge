<!DOCTYPE html>
<html lang="zh-TW">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>AF 用人費用欄位排序系統</title>
<style>
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:'Microsoft JhengHei','微軟正黑體',Arial,sans-serif;background:#fff;color:#111;font-size:14px;min-height:100vh;display:flex;flex-direction:column}
/* ── 頂部 ── */
.header{border-bottom:3px solid #1a3a5c;padding:10px 32px;display:flex;align-items:center;justify-content:space-between;flex-shrink:0}
.header-left{display:flex;align-items:center;gap:14px}
.header-logo{width:38px;height:38px;border:2px solid #1a3a5c;display:flex;align-items:center;justify-content:center;font-size:12px;font-weight:700;color:#1a3a5c;letter-spacing:-.5px;flex-shrink:0}
.header-title{font-size:17px;font-weight:700;color:#1a3a5c;letter-spacing:.5px}
.header-sub{font-size:11px;color:#555;margin-top:2px}
.header-right{font-size:11px;color:#777;text-align:right;line-height:1.8}
/* ── 主體 ── */
.container{max-width:940px;margin:0 auto;padding:24px 24px 32px;width:100%;flex:1}
/* ── 區塊 ── */
.section{border:1px solid #b0b8c4;margin-bottom:18px}
.section-head{background:#1a3a5c;color:#fff;padding:7px 16px;font-size:13px;font-weight:700;letter-spacing:.8px;display:flex;align-items:center;gap:10px;user-select:none}
.step-badge{background:#fff;color:#1a3a5c;font-size:11px;font-weight:700;border-radius:2px;min-width:20px;height:20px;display:flex;align-items:center;justify-content:center;padding:0 4px;flex-shrink:0}
.section-body{padding:18px 20px}
/* ── 上傳區 ── */
.upload-row{display:grid;grid-template-columns:1fr 1fr;gap:18px}
.upload-block{border:1px solid #b0b8c4;padding:0}
.upload-block-head{background:#e8ecf2;padding:7px 12px;font-size:12px;font-weight:700;color:#1a3a5c;border-bottom:1px solid #b0b8c4;display:flex;align-items:center;justify-content:space-between}
.upload-block-body{padding:14px 14px 12px}
.upload-desc{font-size:11px;color:#555;margin-bottom:10px;line-height:1.7}
.upload-desc .em{color:#1a3a5c;font-weight:700}
.file-row{display:flex;align-items:center;gap:8px}
.btn-file{border:1px solid #777;background:#fff;color:#333;font-size:12px;padding:5px 14px;cursor:pointer;font-family:inherit;white-space:nowrap;transition:.1s}
.btn-file:hover{background:#eef0f4}
.file-name{font-size:12px;color:#444;flex:1;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.file-name.ok{color:#2a6e3f;font-weight:700}
.file-name.none{color:#999}
.upload-input{display:none}
.template-link{display:inline-flex;align-items:center;gap:4px;font-size:11px;color:#1a3a5c;text-decoration:none;margin-top:8px;border:1px solid #1a3a5c;padding:3px 10px}
.template-link:hover{background:#eef3f9}
/* ── 提示框 ── */
.hint-box{background:#fffbe6;border:1px solid #d4aa00;padding:8px 12px;font-size:12px;color:#5a3e00;margin-top:14px;line-height:1.7}
.hint-box .hi{font-weight:700}
/* ── 操作列 ── */
.action-bar{display:flex;align-items:center;gap:14px;flex-wrap:wrap}
.btn-go{background:#1a3a5c;color:#fff;border:none;padding:8px 32px;font-size:13px;font-weight:700;cursor:pointer;font-family:inherit;letter-spacing:.5px;transition:.15s}
.btn-go:hover{background:#122840}
.btn-go:disabled{background:#aaa;cursor:default}
.btn-dl{background:#2a6e3f;color:#fff;border:none;padding:8px 22px;font-size:13px;font-weight:700;cursor:pointer;font-family:inherit;letter-spacing:.3px;transition:.15s}
.btn-dl:hover{background:#1e5230}
.btn-reset{background:#fff;color:#444;border:1px solid #999;padding:8px 18px;font-size:12px;cursor:pointer;font-family:inherit;transition:.1s}
.btn-reset:hover{background:#f4f4f4}
/* ── 訊息 ── */
.msg{margin-top:12px}
.ok-msg{background:#eef7ee;border:1px solid #5a9e6f;padding:8px 12px;font-size:12px;color:#1a4d2a}
.warn-msg{background:#fffbe6;border:1px solid #d4aa00;padding:8px 12px;font-size:12px;color:#5a3e00;margin-bottom:4px}
.err-msg{background:#fff2f2;border:1px solid #cc5555;padding:8px 12px;font-size:12px;color:#7a1a1a}
/* ── loader ── */
.loader{display:none;align-items:center;gap:8px;font-size:12px;color:#555}
.spin{width:14px;height:14px;border:2px solid #ccc;border-top-color:#1a3a5c;border-radius:50%;animation:sp .7s linear infinite;flex-shrink:0}
@keyframes sp{to{transform:rotate(360deg)}}
/* ── 結果 ── */
#result-section{display:none}
.result-bar{display:flex;align-items:center;gap:14px;margin-bottom:12px;flex-wrap:wrap}
.r-stat{font-size:12px;color:#333}
.r-stat strong{color:#1a3a5c;font-size:14px}
.tbl-wrap{overflow-x:auto;border:1px solid #b0b8c4}
table{border-collapse:collapse;font-size:12px;white-space:nowrap;width:max-content;min-width:100%}
thead th{background:#1a3a5c;color:#fff;padding:8px 12px;text-align:center;font-weight:700;border-right:1px solid #2a5080;position:sticky;top:0}
thead th:last-child{border-right:none}
tbody tr:nth-child(odd){background:#fff}
tbody tr:nth-child(even){background:#f0f4f9}
tbody tr:hover{background:#dde8f4}
tbody td{padding:6px 12px;border:1px solid #d5dae2;text-align:center;color:#222}
tbody td.name-col{text-align:left;font-weight:600}
.tbl-note{font-size:11px;color:#888;margin-top:6px}
/* ── 底部 ── */
.footer{border-top:1px solid #ccc;padding:8px 24px;font-size:11px;color:#888;text-align:center;flex-shrink:0}
@media(max-width:640px){
  .upload-row{grid-template-columns:1fr}
  .header{padding:10px 16px}
  .container{padding:16px}
  .header-right{display:none}
}
</style>
</head>
<body>

<div class="header">
  <div class="header-left">
    <div class="header-logo">AF</div>
    <div>
      <div class="header-title">AF 用人費用欄位排序系統</div>
      <div class="header-sub">依薪資清冊序號自動重新排列 AF 資料欄位</div>
    </div>
  </div>
  <div class="header-right">系統版本：v1.0<br>114 年度適用</div>
</div>

<div class="container">

  <!-- 步驟一 -->
  <div class="section">
    <div class="section-head">
      <span class="step-badge">1</span>上傳檔案
    </div>
    <div class="section-body">
      <div class="upload-row">

        <!-- 固定清冊檔 -->
        <div class="upload-block">
          <div class="upload-block-head">
            固定清冊檔（人員排列順序）
            <a class="template-link" href="/download-template" download>⬇ 下載範例格式</a>
          </div>
          <div class="upload-block-body">
            <div class="upload-desc">
              包含 <span class="em">input</span> 工作表，欄位：序號、姓名。<br>
              請依薪資清冊正確順序填入，每月人員異動時更新。<br>
              支援格式：<span class="em">.xlsx　.xls</span>
            </div>
            <div class="file-row">
              <input class="upload-input" type="file" id="inp-roster" accept=".xlsx,.xls" onchange="onFile('roster',this)">
              <button class="btn-file" onclick="document.getElementById('inp-roster').click()">選擇檔案</button>
              <span class="file-name none" id="name-roster">尚未選擇</span>
            </div>
          </div>
        </div>

        <!-- AF 資料檔 -->
        <div class="upload-block">
          <div class="upload-block-head">AF 資料檔（本月）</div>
          <div class="upload-block-body">
            <div class="upload-desc">
              由 AF 用人費用管理系統產製，包含 <span class="em">output</span> 工作表。<br>
              每月上傳當月產製之檔案。<br>
              支援格式：<span class="em">.xlsx　.xls</span>
            </div>
            <div class="file-row">
              <input class="upload-input" type="file" id="inp-af" accept=".xlsx,.xls" onchange="onFile('af',this)">
              <button class="btn-file" onclick="document.getElementById('inp-af').click()">選擇檔案</button>
              <span class="file-name none" id="name-af">尚未選擇</span>
            </div>
          </div>
        </div>

      </div>
      <div class="hint-box">
        <span class="hi">※ 注意：</span>固定清冊檔的「姓名」欄位須與 AF 資料檔完全一致（含全名、不可有空格），系統以姓名進行比對。若格式不確定，請先下載右上角範例格式參考。
      </div>
    </div>
  </div>

  <!-- 步驟二 -->
  <div class="section">
    <div class="section-head"><span class="step-badge">2</span>執行排序</div>
    <div class="section-body">
      <div class="action-bar">
        <button class="btn-go" id="btn-go" onclick="doProcess()" disabled>執行排序</button>
        <div class="loader" id="loader"><div class="spin"></div>資料比對處理中，請稍候⋯</div>
      </div>
      <div class="msg" id="msg"></div>
    </div>
  </div>

  <!-- 步驟三（隱藏直到有結果） -->
  <div class="section" id="result-section">
    <div class="section-head"><span class="step-badge">3</span>結果預覽與下載</div>
    <div class="section-body">
      <div class="result-bar">
        <div class="r-stat">比對完成，共 <strong id="stat-total">0</strong> 筆人員</div>
        <button class="btn-dl" onclick="location.href='/download-result'">⬇ 下載排序結果（Excel）</button>
        <button class="btn-reset" onclick="resetAll()">重新上傳</button>
      </div>
      <div class="tbl-wrap">
        <table>
          <thead><tr id="tbl-head"></tr></thead>
          <tbody id="tbl-body"></tbody>
        </table>
      </div>
      <div class="tbl-note" id="tbl-note"></div>
    </div>
  </div>

</div>

<div class="footer">本系統僅供內部作業使用，請勿將資料外傳。</div>

<script>
const files = {roster: null, af: null};

function onFile(key, input) {
  const f = input.files[0];
  if (!f) return;
  files[key] = f;
  const el = document.getElementById('name-' + key);
  el.textContent = '✔ ' + f.name;
  el.className = 'file-name ok';
  checkReady();
}

function checkReady() {
  document.getElementById('btn-go').disabled = !(files.roster && files.af);
}

async function doProcess() {
  const fd = new FormData();
  fd.append('roster', files.roster);
  fd.append('af', files.af);

  document.getElementById('btn-go').disabled = true;
  document.getElementById('loader').style.display = 'flex';
  document.getElementById('msg').innerHTML = '';
  document.getElementById('result-section').style.display = 'none';

  try {
    const res = await fetch('/process', {method: 'POST', body: fd});
    const data = await res.json();

    if (!res.ok || !data.success) {
      setMsg('err', data.error || '發生未知錯誤，請確認檔案格式');
    } else {
      let msgs = '';
      if (data.warnings && data.warnings.length) {
        data.warnings.forEach(w => msgs += `<div class="warn-msg">⚠ ${w}</div>`);
      } else {
        msgs = '<div class="ok-msg">✔ 排序完成，所有人員均成功對應。</div>';
      }
      document.getElementById('msg').innerHTML = msgs;
      renderTable(data.columns, data.preview, data.total);
    }
  } catch (e) {
    setMsg('err', '連線失敗：' + e.message);
  } finally {
    document.getElementById('btn-go').disabled = false;
    document.getElementById('loader').style.display = 'none';
  }
}

function setMsg(type, text) {
  document.getElementById('msg').innerHTML = `<div class="${type}-msg">${type==='err'?'❌':type==='warn'?'⚠':'✔'} ${text}</div>`;
}

function renderTable(cols, rows, total) {
  const head = document.getElementById('tbl-head');
  const body = document.getElementById('tbl-body');
  head.innerHTML = cols.map(c => `<th>${c}</th>`).join('');
  body.innerHTML = rows.map(row =>
    '<tr>' + cols.map(c => {
      const cls = c === '姓名' ? ' class="name-col"' : '';
      return `<td${cls}>${row[c] ?? ''}</td>`;
    }).join('') + '</tr>'
  ).join('');
  document.getElementById('stat-total').textContent = total;
  const note = total > 20 ? `※ 預覽顯示前 20 筆，下載 Excel 含全部 ${total} 筆資料。` : '※ 已顯示全部資料。';
  document.getElementById('tbl-note').textContent = note;
  document.getElementById('result-section').style.display = 'block';
}

function resetAll() {
  files.roster = null; files.af = null;
  ['roster','af'].forEach(k => {
    document.getElementById('inp-' + k).value = '';
    const el = document.getElementById('name-' + k);
    el.textContent = '尚未選擇'; el.className = 'file-name none';
  });
  document.getElementById('msg').innerHTML = '';
  document.getElementById('result-section').style.display = 'none';
  checkReady();
}
</script>
</body>
</html>
