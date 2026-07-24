# Mac 端薪資清冊 OCR — 修正與第二階段

## 一、先修正解析規則

第一階段結果 15/17（88%），兩筆失敗原因已確認，請依下列修改
`~/payroll-ocr/parse_tokens.py`。

### 失敗案例分析

| 姓名 | 現象 | 原因 |
|---|---|---|
| 陳玟莼 | 四項金額正確，應發金額變成 1369 | OCR 讀出 `93940_`，尾端底線導致數字判斷失敗，抓到下一欄的公保自付 |
| 江宇薇 | 整排欄位位移 | 俸點是 `245(無)`，含括號中文不被視為數字，但程式固定「跳過第一個數字當俸點」 |

### 修改內容

**修改 1：`is_num` 容許尾端雜訊**

```python
def is_num(s):
    # 容許尾端出現 OCR 雜訊（_ . , | / - 等）
    return bool(re.fullmatch(r'[\d,]+[\s._,\-/|]*', s))
```

**修改 2：俸點改用數值判斷，不再固定跳過第一個**

原本：

```python
    idx = 1                       # 跳過俸點
    pay = nums[idx]; idx += 1
```

改為：

```python
    # 俸點為 3 位數以內（約 90~800），薪俸則是 4 位數以上。
    # 依數值大小濾掉開頭的俸點，避免俸點格式異常（如「245(無)」）造成整排位移。
    while nums and nums[0] < 1000:
        nums.pop(0)
    if not nums:
        return None
    idx = 0
    pay = nums[idx]; idx += 1
```

### 驗證

```bash
cd ~/payroll-ocr
python3 parse_tokens.py tokens.json
```

**預期結果：17 筆全部加總相符**。若仍有失敗，請把該列的原始 token 貼出：

```bash
python3 -c "
import json
t=json.load(open('tokens.json'))
for x in t:
    if 0.30 < x['y'] < 0.34:   # 依失敗列的位置調整範圍
        print(round(x['x'],3), repr(x['text']), round(x['conf'],2))
"
```

達到 17/17 再進行第二階段。

---

## 二、第二階段：接回伺服器

### 架構

Mac 主動輪詢伺服器，**不需要固定 IP、不開放任何連接埠**。

輪詢間隔由伺服器動態決定，不是固定值：

| 情況 | 間隔 | 說明 |
|---|---|---|
| 有人正在使用網站（10 分鐘內） | 3 秒 | 反應快 |
| 有待辨識工作 | 3 秒 | 立即處理 |
| 閒置 | 60 秒 | 省資源 |
| 伺服器連不上 | 60 秒 | 避免瘋狂重試 |

固定 5 秒的話一天會發出 17,280 次請求；動態調整後約 1,500 次，減少約
90%，且有人使用時反應更快。**Mac 端不需要自行判斷，照伺服器回傳的
`next_poll` 值等待即可。**

```
使用者上傳掃描檔
      ↓
伺服器建立工作，狀態「等待辨識」
      ↓
Mac 每 5 秒詢問一次 → 領到工作 → 下載 PDF
      ↓
Vision 辨識 → 解析 → 回傳結果
      ↓
伺服器接續比對 → 前端顯示結果
```

### 需要的設定值

向系統管理者取得下列兩項，寫進 `~/payroll-ocr/config.json`：

```json
{
  "server": "https://你的網址",
  "key": "與伺服器環境變數 OCR_KEY 相同的密鑰"
}
```

**此檔含密鑰，請設定權限：`chmod 600 config.json`**

### 3. 建立 `~/payroll-ocr/worker.py`

```python
#!/usr/bin/env python3
"""薪資清冊 OCR 工作程式：輪詢伺服器、辨識、回傳"""
import json, os, subprocess, sys, tempfile, time, base64
import urllib.request, urllib.error

HERE = os.path.dirname(os.path.abspath(__file__))
CFG = json.load(open(os.path.join(HERE, 'config.json'), encoding='utf-8'))
SERVER, KEY = CFG['server'].rstrip('/'), CFG['key']
POLL_MIN, POLL_MAX = 3, 60      # 實際間隔由伺服器決定，此為安全範圍

sys.path.insert(0, HERE)
from parse_tokens import group_rows, parse_row      # 沿用第一階段的解析


def req(path, data=None, timeout=30):
    url = SERVER + path
    body = json.dumps(data).encode() if data is not None else None
    r = urllib.request.Request(url, data=body, method='POST' if body else 'GET')
    r.add_header('X-OCR-KEY', KEY)
    if body:
        r.add_header('Content-Type', 'application/json')
    try:
        with urllib.request.urlopen(r, timeout=timeout) as resp:
            if resp.status == 204:
                return None
            return json.loads(resp.read() or b'null')
    except urllib.error.HTTPError as e:
        print(f'  HTTP {e.code}: {e.read()[:200]}', flush=True)
        return None
    except Exception as e:
        print(f'  連線失敗：{e}', flush=True)
        return None


def ocr(pdf_bytes):
    """呼叫 Vision 辨識，回傳解析後的人員清單"""
    with tempfile.NamedTemporaryFile(suffix='.pdf', delete=False) as f:
        f.write(pdf_bytes)
        path = f.name
    try:
        out = subprocess.run(
            ['swift', os.path.join(HERE, 'ocr_extract.swift'), path],
            capture_output=True, timeout=180)
        if out.returncode != 0:
            raise RuntimeError(out.stderr.decode()[:300])
        tokens = json.loads(out.stdout)
        people = [p for p in (parse_row(r) for r in group_rows(tokens)) if p]
        return people
    finally:
        os.unlink(path)


def main():
    print(f'OCR 工作程式啟動｜伺服器 {SERVER}', flush=True)
    wait = POLL_MAX
    while True:
        resp = req('/ocr/claim')

        # 連線失敗：維持最長間隔，避免伺服器離線時瘋狂重試
        if resp is None:
            time.sleep(POLL_MAX)
            continue

        # 伺服器會依「是否有人正在使用網站」回傳建議間隔：
        #   有人操作中 → 3 秒（反應快）
        #   閒置       → 60 秒（省資源）
        wait = max(POLL_MIN, min(POLL_MAX, int(resp.get('next_poll', POLL_MAX))))

        if not resp.get('job_id'):
            time.sleep(wait)
            continue

        jid = resp['job_id']
        job = resp
        print(f'領到工作 {jid}', flush=True)
        try:
            people = ocr(base64.b64decode(job['pdf_b64']))
            ok = sum(1 for p in people if p['加總相符'])
            print(f'  解析 {len(people)} 人，加總相符 {ok}', flush=True)
            req('/ocr/result', {'job_id': jid, 'people': people})
        except Exception as e:
            print(f'  辨識失敗：{e}', flush=True)
            req('/ocr/result', {'job_id': jid, 'error': str(e)[:300]})


if __name__ == '__main__':
    main()
```

測試（先手動跑，確認能連上）：

```bash
cd ~/payroll-ocr
python3 worker.py
```

畫面應顯示「OCR 工作程式啟動」並持續運作。此時到網站上傳一份掃描檔，
終端機應出現「領到工作」。

### 4. 設定開機自動啟動

建立 `~/Library/LaunchAgents/com.payroll.ocr.plist`：

```xml
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN"
  "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
    <key>Label</key>
    <string>com.payroll.ocr</string>
    <key>ProgramArguments</key>
    <array>
        <string>/usr/bin/python3</string>
        <string>/Users/你的帳號/payroll-ocr/worker.py</string>
    </array>
    <key>WorkingDirectory</key>
    <string>/Users/你的帳號/payroll-ocr</string>
    <key>RunAtLoad</key>
    <true/>
    <key>KeepAlive</key>
    <true/>
    <key>StandardOutPath</key>
    <string>/Users/你的帳號/payroll-ocr/worker.log</string>
    <key>StandardErrorPath</key>
    <string>/Users/你的帳號/payroll-ocr/worker.err</string>
</dict>
</plist>
```

**請把 `你的帳號` 換成實際使用者名稱**（用 `whoami` 查詢）。

載入與確認：

```bash
launchctl load ~/Library/LaunchAgents/com.payroll.ocr.plist
launchctl list | grep payroll          # 應出現該項目
tail -f ~/payroll-ocr/worker.log       # 確認有啟動訊息
```

停用：`launchctl unload ~/Library/LaunchAgents/com.payroll.ocr.plist`

### 5. 電源設定

Mac 睡眠時程式不會執行。請設定不休眠：

```bash
sudo pmset -a sleep 0 disablesleep 1
```

或到「系統設定 → 節能」關閉自動休眠。螢幕可以關，硬碟不要休眠。

---

## 注意事項

1. **加總不符的列必須照實回報**，不可自行修補數字。伺服器會把這些列標示為
   需人工確認，這是防止錯誤資料混入稽核結果的最後一道防線。

2. **不可用金額反推姓名**。姓名辨識失敗就標記該列，不要猜。用金額配對會讓
   真正的金額差異被掩蓋，那正是這套工具要抓的東西。

3. **密鑰不要寫進程式碼或提交到版本控制**，只放在 `config.json`。

4. 薪資清冊含個人資料。暫存檔用完即刪（程式已處理），`worker.log` 不會記錄
   金額內容，但仍請勿外傳。

---

## 三、補充：回傳欄位信心值（提升人工確認效率）

伺服器收到辨識結果後，會把**加總不符**或**姓名對不上**的列列出來讓使用者
核對。若能一併提供辨識信心，介面可以精準標示是哪一格可疑。

`parse_tokens.py` 的 `parse_row()` 已回傳 `最低信心`（整列最低值）。若要更
精準，可改為記錄**每個欄位各自的信心**：

```python
    # 在取用 nums 時，同步記錄該 token 的 conf
    num_tokens = [t for t in row[name_i + 1:] if is_num(t['text'])]
    # ...依序取值時，一併取 num_tokens[k]['conf']
    return {
        ...,
        '欄位信心': {'薪俸': c1, '主管加給': c2, '專業加給': c3, '導師特教': c4},
        '最低信心': min(...),
    }
```

伺服器對 `欄位信心` 不存在時會自動略過，因此**這是選配**，不做也不影響運作。

---

## 四、使用者端會看到什麼

辨識完成後，畫面分兩部分：

- **自動通過的列**：加總相符且姓名對得上 → 直接進入比對，使用者不需處理
- **需確認的列**：帶著 OCR 讀到的數值顯示，可疑欄位加紅框，下方列出原因

使用者對照紙本修正數值，四項相加等於應發金額時該列即轉綠。姓名欄有下拉
提示，來源是 AF 名單。真的無法確認的列可清空姓名，該列就不納入比對。

這代表：**Mac 端不需要為了讓數字好看而做任何補正**。讀不準就照實回報，
使用者會處理。反而是自行猜測補值會造成錯誤資料混入稽核結果。
