# AF → aftool 一鍵匯入（Microsoft Edge 擴充功能，Manifest V3）

- `edge/`：擴充功能原始碼（manifest.json 所在層＝「載入解壓縮」要選的資料夾）。
  - `af_core.js`：一鍵流程核心（由 `static/af_oneclick.js` 移植；改這裡時兩邊都要同步）。
  - `fund_content.js`：AF 系統頁面右下角按鈕、提示、彈出「報表列印」視窗接手。
  - `background.js`：工具列圖示、開啟／重用 aftool 分頁、轉送 ZIP。
  - `aftool_bridge.js`：aftool 頁面端，以 window.postMessage 交給 `static/af_import.js`。
- `tools/build_zip.py`：打包成 `static/extension/AF-aftool-extension-<版本>.zip`（網站下載與上架用）。
  改了 edge/ 一定要重跑，`test_extension.py` 會檢查 ZIP 與原始碼一致。
- `tools/make_icons.py`、`tools/make_screenshots.py`：圖示與商店截圖（需 Pillow）。
- `tests/`：模擬 AF 頁（`mock_fund.py`、假資料 `fake_af_public.zip`）與 Playwright 端對端測試
  `e2e_extension.py`（`CHROMIUM=<Chrome for Testing 執行檔> python e2e_extension.py`；
  需先在 5061 啟動 staging，或設 `PROD=1` 對正式站）。
- `store/`：商店文字、審核說明、送審步驟、logo 與截圖。

開發人員模式安裝：`edge://extensions` → 開啟「開發人員模式」→「載入解壓縮」→ 選 `edge/` 資料夾。
