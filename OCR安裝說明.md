# 掃描檔自動辨識（選配）

## 不裝也能用

遇到掃描檔時，系統會列出 AF 的人員清單，讓你對照紙本手動填入四個項目的金額。
填入「應發金額」後會自動檢查加總，填錯會即時標紅。

## 想讓系統預先填好數字

需要額外安裝兩個 Python 套件和 tesseract 中文包。

### 1. requirements.txt 取消註解

```
opencv-python-headless
pytesseract
```

### 2. 安裝系統套件

Zeabur 若使用 Dockerfile，加入：

```dockerfile
RUN apt-get update && apt-get install -y \
    tesseract-ocr tesseract-ocr-chi-tra \
    && rm -rf /var/lib/apt/lists/*
```

若使用 Nixpacks（Zeabur 預設），在專案根目錄建立 `nixpacks.toml`：

```toml
[phases.setup]
aptPkgs = ["tesseract-ocr", "tesseract-ocr-chi-tra"]
```

### 3. 重新部署

程式會自動偵測。裝好後上傳掃描檔，表格會預先填入辨識到的數字；
沒裝就是空白表格，前端完全不用改。

## 準確度說明

- 金額欄位限定只能是 0-9，辨識率高
- 但掃描品質不佳時仍可能出錯，務必對照紙本確認
- 加總檢查是最後一道防線：各項相加 ≠ 應發金額就會標紅
