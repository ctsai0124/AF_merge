# Leo 送審步驟（Microsoft Edge Add-ons，設為「不公開 Hidden」）

標示「（未確認）」的地方是我無法從官方文件確定的細節，請以實際畫面為準。
官方說明：https://learn.microsoft.com/microsoft-edge/extensions/publish/publish-extension

## 0. 準備好的檔案（都在 repo 的 extension/store/ 與 static/extension/）
- 上傳用套件：`static/extension/AF-aftool-extension-1.0.0.zip`（manifest.json 在 ZIP 根目錄）
- 商店 logo：`extension/store/logo-300x300.png`（官方：1:1，建議 300×300，至少 128×128）
- 截圖：`extension/store/screenshot-*-1280x800.png`（官方：640×480 或 1280×800，選填，最多 6 張）
- 商店文字：`extension/store/listing-zh-TW.md`
- 給審核者的說明：`extension/store/testing-notes.md`
- 隱私權聲明網址：https://aftool.leotsai.me/privacy-extension

## 1. 註冊 Partner Center 的 Microsoft Edge 開發者帳號（只需一次）
1. 用自己的 Microsoft 帳號登入 Partner Center 的 Edge 註冊頁：https://partner.microsoft.com/dashboard/microsoftedge/ （註冊流程中選「個人 Individual」即可；Edge 開發者註冊免費）。
2. 依畫面填寫聯絡資料、同意開發者協議。（是否需要額外身分驗證：未確認）

## 2. 建立擴充功能並上傳
1. Partner Center → Edge → Overview →「Create new extension」。
2. 把 `AF-aftool-extension-1.0.0.zip` 拖進「Upload package」，等驗證通過 →「Continue」。

## 3. Availability（可見度）
1. Visibility 選「**Hidden**」：不會出現在商店搜尋與瀏覽，只有拿到商店連結的人可安裝。
2. Markets：可維持預設（全部市場）或只勾 Taiwan。
3. Save & Continue。

## 4. Properties
- Category：Productivity（生產力）。
- 隱私權相關問題：本擴充功能會處理使用者在 AF 系統的報表（含個資），建議回答「會存取個人資料」並填隱私權聲明網址 https://aftool.leotsai.me/privacy-extension 。（題目確切文字：未確認）
- Website：https://aftool.leotsai.me/ ；Support contact：填您要公開的 email。

## 5. Store listings
- 語言：新增「中文（台灣）／Chinese (Traditional)」。
- Description：貼 `listing-zh-TW.md` 的「詳細說明」。
- Extension logo：`logo-300x300.png`；Screenshots：三張 1280×800。

## 6. Submit
1. 在「Notes for certification」貼 `testing-notes.md` 的內容（說明需公務帳號、無法登入、可看 aftool 接收端）。
2. 按「Publish」送審。官方文件說審核最多約 7 個工作天。
3. 通過後，在 Extension overview 頁面取得商店連結（listing URL），分享給同仁安裝；告訴我連結，我會把 aftool 首頁的「審核中」改成商店安裝連結。

## 之後更新版本
修改 extension/edge 後把 manifest.json 的 version 加一，執行 `python3 extension/tools/build_zip.py` 產生新 ZIP，在 Partner Center 該擴充功能的 Packages 上傳新 ZIP 再送審；已安裝的同仁會自動更新。

## 可能遇到的問題
- 公務電腦若由機關以群組原則管理，可能禁止安裝擴充功能或開發人員模式（ExtensionInstallBlocklist 等政策），需洽機關資訊人員放行此擴充功能。
- 審核者無法登入 AF 系統；若被要求示範，可提供模擬頁測試（extension/tests/）或操作錄影。
