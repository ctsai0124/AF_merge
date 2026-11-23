/* AF → aftool 一鍵匯入：背景服務（只在 fund.dgpa.gov.tw 與 aftool.leotsai.me 之間轉送） */
'use strict';
const AFTOOL = 'https://aftool.leotsai.me';
const FUND = 'https://fund.dgpa.gov.tw/';
let expectPopup = null;   // {tabId, frameId, at}：主視窗剛按了「報表列印」，等彈出視窗接手

const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

chrome.action.onClicked.addListener(async (tab) => {
  if (tab && tab.url && tab.url.startsWith(FUND)) {
    try { await chrome.tabs.sendMessage(tab.id, { type: 'run' }, { frameId: 0 }); return; } catch (e) { /* 頁面還沒載入完成 */ }
  }
  await chrome.tabs.create({ url: AFTOOL + '/#afi-guide' });
});

async function findOrOpenAftool() {
  const tabs = await chrome.tabs.query({ url: AFTOOL + '/*' });
  let tab = tabs.find((t) => t.active) || tabs[0];
  if (tab) {
    await chrome.tabs.update(tab.id, { active: true });
    try { await chrome.windows.update(tab.windowId, { focused: true }); } catch (e) {}
    return tab;
  }
  return chrome.tabs.create({ url: AFTOOL + '/?import=af' });
}

async function deliver(msg) {
  const tab = await findOrOpenAftool();
  for (let i = 0; i < 60; i++) {
    try {
      const resp = await chrome.tabs.sendMessage(tab.id, { type: 'af-zip', name: msg.name, b64: msg.b64 }, { frameId: 0 });
      if (resp) return resp;
    } catch (e) { /* aftool 頁面還在載入 */ }
    await sleep(500);
  }
  return { ok: false, error: 'aftool 分頁沒有回應，請重新整理 aftool 後再按一次。' };
}

chrome.runtime.onMessage.addListener((msg, sender, sendResponse) => {
  if (!sender.tab || !sender.url) return;
  const fromFund = sender.url.startsWith(FUND);
  if (!fromFund) return;
  if (msg.type === 'expect-popup') {
    expectPopup = { tabId: sender.tab.id, frameId: sender.frameId, at: Date.now() };
    sendResponse({ ok: true });
  } else if (msg.type === 'popup-ready') {
    const e = expectPopup;
    const run = !!(e && Date.now() - e.at < 30000 && e.tabId !== sender.tab.id);
    if (run) {
      expectPopup = null;
      chrome.tabs.sendMessage(e.tabId, { type: 'claimed' }, { frameId: e.frameId }).catch(() => {});
    }
    sendResponse({ run });
  } else if (msg.type === 'deliver' && typeof msg.b64 === 'string') {
    if (msg.b64.length > 23 * 1024 * 1024) { sendResponse({ ok: false, error: '檔案過大' }); return; }
    deliver(msg).then(sendResponse, (err) => sendResponse({ ok: false, error: String(err && err.message || err) }));
    return true;
  }
});
