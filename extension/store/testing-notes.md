# Notes for certification / 審核說明（貼到 Partner Center 的 "Notes for certification"）

English:
This extension is a helper for school personnel staff in Taiwan. It only runs on
https://fund.dgpa.gov.tw/* (the Taiwan government "AF" personnel-cost system) and
https://aftool.leotsai.me/* (our own web tool).

- The AF system requires a government employee account (login via https://ecpa.dgpa.gov.tw
  with a government credential), so reviewers cannot log in. The system is also only open
  07:00–19:00 Taiwan time.
- What it does: when the user clicks the floating "AF → aftool 一鍵匯入" button (or the toolbar
  icon) on an AF page, the content script finds the report-print form by its visible labels,
  selects report type "校對清冊(A)", format EXCEL, staff type "全部", fills the shared public
  download password used by our tool, submits the same ASP.NET postback with fetch() (same
  origin, the user's existing session), receives the encrypted ZIP report, and hands it to the
  aftool tab (opened or reused via chrome.tabs). The aftool page decrypts it for sorting.
- No remote code: all logic is packaged in the extension. No permissions other than the two
  host permissions. No cookies, storage, downloads, history or <all_urls>. No analytics.
- How to test without an AF account: open https://aftool.leotsai.me/ to see the receiving
  side and the help text. A mock of the AF report page (fake data) and an automated test are in
  our repository (extension/tests/); we can provide a screen recording on request.
- Privacy policy: https://aftool.leotsai.me/privacy-extension

中文：
本擴充功能供台灣學校人事人員使用，只在 fund.dgpa.gov.tw（AF 用人費用管理資訊系統）與
aftool.leotsai.me（本工具網站）運作。AF 系統需以公務帳號經 eCPA 登入，審核人員無法登入；
可用 aftool 網站查看接收端，或由我們提供模擬頁測試與操作錄影。擴充功能不載入遠端程式碼，
除兩個網站存取權外不要求任何權限，不讀取 cookie、不儲存資料、不傳送給第三方。
