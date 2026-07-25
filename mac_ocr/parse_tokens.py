#!/usr/bin/env python3
"""
把 ocr_extract.swift 的 token JSON 還原成人員資料。

支援兩種薪資清冊版面：
  橫式：一人一列（編號 職稱 姓名 俸點 薪俸 …）
  直式：一人一欄（左側為項目名稱，上方為序號／身分證字號／姓名／職稱）

若清冊含身分證字號，會一併讀出並以檢查碼驗證，供伺服器端精確配對。
"""
import json, sys, re

Y_TOL = 0.006          # 同一列的 y 座標容差
X_TOL = 0.035          # 直式版面中，token 歸屬欄位的 x 容差

MGR_WORDS = ('校長', '主任', '組長')
TITLE_WORDS = ('校長', '主任', '組長', '教師', '幹事', '校護', '工友', '技工',
               '駕駛', '護理師', '代理教師', '編號', '職稱', '導師', '科任',
               '教保員', '幼師', '教練', '資源班', '暫代', '長代', '借調',
               '輔導教師', '學士')
# 僅列出「不可能出現在職稱中」的字詞，避免誤跳過真正的人員。
# 例：「總務主任」「人事主任」「會計主任」「出納組長」都是實際職稱，不可過濾。
SKIP_WORDS = ('小計', '合計', '總計', '備註', '製表', '用印', '機關首長',
              '承辦人', '核章', '稽核人員')

# 直式版面的項目名稱 → 內部欄位
FIELD_ALIAS = {
    '本俸': '薪俸', '薪俸': '薪俸', '月俸額': '薪俸',
    '專業加給': '專業加給', '學術研究': '專業加給', '學術研究費': '專業加給',
    '職務加給': '主管加給', '主管加給': '主管加給',
    '導師加給': '導師費', '導師費': '導師費',
    '特教加給': '特教加給',
    '地域加給': '地域加給',
    '教保費': '其他加給',            # 契約教保員的加給，計入應發
    '行政工作獎金': '其他加給',
    '工作津貼': '其他加給',
    '導師加給扣款': '扣款',          # 自應發金額中扣除
    '應發數合計': '應發金額', '應發金額': '應發金額', '小計': '應發金額',
}

# 直式清冊若誤走橫式解析，這些項目名稱很容易被當成「姓名」。
# 僅用於候選版面評分，不會直接刪除任何資料。
NON_NAME_LABELS = {
    '公保費', '健保費', '勞保費', '退撫基金', '所得稅', '互助金',
    '儲蓄存款', '實發數', '實發金額', '實扣數', '扣款合計',
    '補發數', '應發數', '應發合計', '應發數合計',
    *FIELD_ALIAS.keys(),
}

# ── 身分證字號驗證 ──────────────────────────────────────
_ID_MAP = {'A':10,'B':11,'C':12,'D':13,'E':14,'F':15,'G':16,'H':17,'I':34,
           'J':18,'K':19,'L':20,'M':21,'N':22,'O':35,'P':23,'Q':24,'R':25,
           'S':26,'T':27,'U':28,'V':29,'W':32,'X':30,'Y':31,'Z':33}


def valid_id(s):
    """台灣身分證字號檢查碼驗證，可用來偵測 OCR 讀錯"""
    s = (s or '').strip().upper()
    if len(s) != 10 or s[0] not in _ID_MAP or not s[1:].isdigit():
        return False
    n = _ID_MAP[s[0]]
    tot = n // 10 + (n % 10) * 9
    for i, d in enumerate(s[1:9]):
        tot += int(d) * (8 - i)
    tot += int(s[9])
    return tot % 10 == 0


def looks_like_id(s):
    return bool(re.fullmatch(r'[A-Za-z][0-9]{9}', (s or '').strip()))


# ── 共用工具 ────────────────────────────────────────────

# 文字辨識可能把千分位逗號讀成全形逗號、頓號或其他相近字元，
# 也可能夾帶空白，比對前一律正規化。
_SEP_CHARS = str.maketrans({
    '，': ',', '、': ',', '‚': ',', '·': ',', '・': ',',
    '．': ',', '。': ',', '｡': ',', '.': ',',      # 千分位常被讀成句點
    '　': '', ' ': '', '\u00a0': '', '\u2009': '', '\u202f': '',
})


def norm_text(s):
    return (s or '').translate(_SEP_CHARS).strip()


def is_num(s):
    # 掃描 OCR 常在金額前後黏到框線、括號或一個近似直線的英文字元。
    # 只容許有限的邊緣雜訊，而且逗號必須符合千分位格式。
    # 例如「4,0001」通常是右側框線被讀成 1；不能直接去掉逗號變成 40001。
    match = re.fullmatch(
        r'[~～LlI|<>＜＞]*(-?[\d,]+?)[._,\-/|()（）\[\]［］{}]*',
        norm_text(s))
    if not match:
        return False
    core = match.group(1).lstrip('-')
    return core.isdigit() or bool(re.fullmatch(r'\d{1,3}(?:,\d{3})+', core))


def to_int(s):
    s = norm_text(s)
    neg = s.startswith('-')
    s = re.sub(r'[^\d]', '', s)
    if not s:
        return 0
    return -int(s) if neg else int(s)


def is_name(s):
    s = (s or '').strip()
    if len(s) < 2 or len(s) > 4:
        return False
    if not all('\u4e00' <= c <= '\u9fff' for c in s):
        return False
    return not any(w in s for w in TITLE_WORDS)


def repair_trailing_gridline_digit(row):
    """
    修復金額右框線被 Vision 黏成尾端「1」的極窄案例。

    只在下列證據同時成立時修復，不使用加總或應發金額反推：
    1. 原字串是千分位後多一個 1（例：4,0001）；
    2. OCR 信心不高於 0.35；
    3. 同一列至少另有兩格完全相同的合法基準值（例：4,000）；
    4. 異常 token 的字框寬度與基準值中位數相差不超過 12%。
    """
    for token in row:
        raw = norm_text(token.get('text', ''))
        match = re.fullmatch(r'(\d{1,3},\d{3})1', raw)
        if not match or float(token.get('conf', 1.0)) > 0.35:
            continue

        base = match.group(1)
        peers = [
            other for other in row
            if other is not token and norm_text(other.get('text', '')) == base
            and float(other.get('w', 0)) > 0
        ]
        if len(peers) < 2 or float(token.get('w', 0)) <= 0:
            continue

        widths = sorted(float(other['w']) for other in peers)
        median_width = widths[len(widths) // 2]
        width_ratio = float(token['w']) / median_width
        if 0.88 <= width_ratio <= 1.12:
            token['text'] = base


def merge_number_fragments(row):
    """
    合併被拆開的數字。

    文字辨識可能把「55,690」讀成「55,」與「690」兩段。麻煩的是排序後
    這兩段之間有時會夾著鄰欄的數字（因為切分位置與欄位中心不一致），
    因此不能只比較相鄰的兩個 token。

    做法：找出所有以逗號結尾的碎片，為它配對右側最接近的三位數字，
    再依距離由近而遠逐一配成對。以逗號結尾是極強的訊號——單獨的
    金額不會以逗號收尾。
    """
    repair_trailing_gridline_digit(row)

    if len(row) < 2:
        return row

    NUMISH = re.compile(r'^-?[\d,]+$')
    # 尾段可能夾帶多餘標點（例：「690.」），比對時允許結尾有分隔字元
    THREE = re.compile(r'^\d{3}[,\-/|()（）\[\]［］]*$')
    MAX_GAP = 0.06                      # 仍遠小於典型欄距

    body = [dict(t) for t in row[1:]]
    heads = [i for i, t in enumerate(body)
             if NUMISH.match(norm_text(t['text']).rstrip(',') + ',')
             and norm_text(t['text']).endswith(',')]
    tails = [i for i, t in enumerate(body)
             if THREE.match(norm_text(t['text']))]

    pairs = []
    for i in heads:
        for j in tails:
            if j == i:
                continue
            d = body[j]['x'] - body[i]['x']
            if 0 < d <= MAX_GAP:
                pairs.append((d, i, j))
    pairs.sort()

    used_h, used_t, merged = set(), set(), {}
    for d, i, j in pairs:
        if i in used_h or j in used_t:
            continue
        used_h.add(i); used_t.add(j)
        merged[i] = j

    out = [dict(row[0])]
    skip = set(merged.values())
    for i, t in enumerate(body):
        if i in skip:
            continue
        if i in merged:
            j = merged[i]
            a, b = body[i], body[j]
            t = dict(a)
            t['text'] = (norm_text(a['text']) + norm_text(b['text'])).rstrip(',')
            t['w'] = (b['x'] + b.get('w', 0) / 2) - (a['x'] - a.get('w', 0) / 2)
            t['x'] = (a['x'] + b['x']) / 2
            t['conf'] = min(a['conf'], b['conf'])
        out.append(t)

    # 合併後可能改變左右順序，重新排序
    return [out[0]] + sorted(out[1:], key=lambda t: t['x'])


def group_rows(tokens):
    tokens = sorted(tokens, key=lambda t: (t['page'], t['y']))
    rows, cur, last = [], [], None
    for t in tokens:
        if cur and (t['page'] != cur[0]['page'] or abs(t['y'] - last) > Y_TOL):
            rows.append(sorted(cur, key=lambda a: a['x']))
            cur = []
        cur.append(t)
        last = t['y']
    if cur:
        rows.append(sorted(cur, key=lambda a: a['x']))
    return [merge_number_fragments(r) for r in rows]


def row_text(row):
    return ''.join(t['text'] for t in row)


# ── 版面判斷 ────────────────────────────────────────────

def _layout_evidence(rows):
    """回傳直式版面的表頭列數與金額項目列數。"""
    header_labels = {'姓名', '身分證字號', '序號', '職稱', '支薪俸級'}
    heads = fields = name_rows = 0
    # 不再只看前 25 列。不同學校的頁首長度差異很大，關鍵列可能在後面。
    for r in rows:
        if not r:
            continue
        first = (r[0]['text'] or '').strip().rstrip('：:')
        if first in header_labels:
            heads += 1
            if first == '姓名' and len(r) > 1:
                name_rows += 1
        if first in FIELD_ALIAS and len(r) > 1:
            fields += 1
    return heads, fields, name_rows


def detect_layout(rows):
    """
    直式版面的特徵：某一列以「姓名」開頭，後面接多個人名；
    且另有一列以「本俸／薪俸」等項目名稱開頭。
    """
    heads, fields, name_rows = _layout_evidence(rows)
    vertical = (
        (name_rows >= 1 and fields >= 1)
        or (heads >= 2 and fields >= 2)
        or fields >= 4
    )
    return 'vertical' if vertical else 'horizontal'


# ── 直式解析 ────────────────────────────────────────────

def _vertical_block(rows):
    """
    解析單一區塊（通常是一頁）的直式資料。

    直式清冊中，姓名／職稱靠左對齊，金額靠右對齊，兩者中心點差距可能
    大於欄距的一半，因此不能用「最接近的中心點」去配對，否則金額會被
    歸到隔壁的人。

    正確做法：
      1. 以「每個人都會有值」的金額列（應發數合計或本俸）建立金額欄位座標
      2. 姓名列與金額欄位皆由左至右排序，依序號對應
      3. 其餘金額列再以最接近的金額欄位座標歸位（此時對齊方式一致）
    """
    def first_row(label_pred):
        return next((r for r in rows
                     if label_pred((r[0]['text'] or '').strip()) and len(r) > 1), None)

    def money_tokens(row, minimum=1000):
        return [t for t in row[1:]
                if is_num(t.get('text', '')) and abs(to_int(t.get('text', ''))) >= minimum]

    # 先找本俸列。標準版面使用 FIELD_ALIAS；掃描品質差時，改找表頭後
    # 第一列「至少三個四位數以上金額」，避免依賴容易誤讀的「本俸」二字。
    base_idx = next(
        (i for i, r in enumerate(rows)
         if FIELD_ALIAS.get((r[0]['text'] or '').strip()) == '薪俸'
         and len(money_tokens(r)) >= 2),
        None)
    if base_idx is None:
        base_idx = next(
            (i for i, r in enumerate(rows[:24])
             if len(money_tokens(r, 10000)) >= 3),
            None)
    if base_idx is None:
        return []
    base_row = rows[base_idx]

    name_row = first_row(lambda s: s == '姓名')
    names = []

    def add_name(t):
        nm = (t['text'] or '').strip()
        if not nm or any(w in nm for w in SKIP_WORDS):
            return
        if len(nm) > 5 or t.get('x', 0) < 0.08:
            return
        if not is_name(nm) or _suspicious_name(nm):
            return
        names.append({'姓名': nm, 'x': t['x'], '_conf': [t['conf']]})

    if name_row:
        for t in name_row[1:]:
            add_name(t)
        name_y = name_row[0]['y']
    else:
        # 鼓山版面每頁表頭常被讀成「上.名／生名／E名」，甚至同一姓名列
        # 被拆成相鄰兩列。先找表頭區中最早一列至少兩個合理人名，再合併
        # 上下 0.012 內的姓名 token。排除 x<0.08 可避開誤讀的列標題。
        candidates = []
        for i, r in enumerate(rows[:base_idx]):
            vals = [t for t in r if t.get('x', 0) >= 0.08
                    and is_name((t.get('text') or '').strip())
                    and not _suspicious_name(t.get('text'))
                    and not any(w in (t.get('text') or '') for w in SKIP_WORDS)]
            if len(vals) >= 2:
                candidates.append((i, vals))
        if not candidates:
            return []
        ni, _ = candidates[0]
        name_y = rows[ni][0]['y']
        for r in rows[:base_idx]:
            if abs(r[0]['y'] - name_y) <= 0.012:
                for t in r:
                    add_name(t)

    names.sort(key=lambda n: n['x'])
    if not names:
        return []

    # 找一列「所有人都有值」的金額列來定義欄位座標
    ref = None
    for want in ('應發金額', '薪俸'):
        for r in rows:
            lab = (r[0]['text'] or '').strip()
            if (FIELD_ALIAS.get(lab) == want
                    and len(money_tokens(r)) >= len(names)):
                ref = r
                break
        if ref:
            break
    if ref is None:
        ref = base_row

    num_xs = sorted(t['x'] for t in money_tokens(ref))
    if len(num_xs) < len(names):
        names = names[:len(num_xs)]
    # 小計／合計位於最右側，取左起與人數相同的欄位即可
    cols = []
    for i, n in enumerate(names):
        n['nx'] = num_xs[i]
        cols.append(n)

    gaps = sorted(b['nx'] - a['nx'] for a, b in zip(cols, cols[1:])) or [X_TOL * 2]
    tol = max(X_TOL, gaps[len(gaps) // 2] * 0.5)

    def assign_num(row, key):
        for t in row[1:]:
            txt = (t['text'] or '').strip()
            if not txt or not is_num(txt):
                continue
            best, bd = None, 9
            for c in cols:
                d = abs(t['x'] - c['nx'])
                if d < bd:
                    best, bd = c, d
            if best is not None and bd <= tol:
                best[key] = to_int(txt)
                best['_conf'].append(t['conf'])

    def assign_text(row, key):
        """身分證、職稱與姓名同為靠左對齊，依姓名欄座標配對"""
        for t in row[1:]:
            txt = (t['text'] or '').strip()
            if not txt or any(w in txt for w in SKIP_WORDS):
                continue
            best, bd = None, 9
            for c in cols:
                d = abs(t['x'] - c['x'])
                if d < bd:
                    best, bd = c, d
            if best is not None and bd <= tol:
                best[key] = txt
                best['_conf'].append(t['conf'])

    # 表頭的「職稱」也可能被讀成「成稱／E稱」；在姓名列與本俸列之間，
    # 排除俸級數字後，依 x 座標把文字配回各人。
    title_row = first_row(lambda s: s == '職稱')
    if title_row:
        assign_text(title_row, '職稱')
    else:
        for r in rows:
            if not (name_y < r[0]['y'] < base_row[0]['y']):
                continue
            if len(money_tokens(r, 1)) >= 2:
                continue
            for t in r:
                txt = (t.get('text') or '').strip()
                if (not txt or is_num(txt) or t.get('x', 0) < 0.08
                        or any(w in txt for w in SKIP_WORDS)):
                    continue
                best, bd = None, 9
                for c in cols:
                    d = abs(t['x'] - c['x'])
                    if d < bd:
                        best, bd = c, d
                if best is not None and bd <= tol:
                    best['職稱'] = txt
                    best['_conf'].append(t['conf'])

    # 鼓山各頁欄名雖會變形，但資料列順序固定。本俸起算的相對位置可作為
    # 文字辨識失敗時的安全備援；精確讀到 FIELD_ALIAS 時仍以標準對應補強。
    positional_fields = {
        0: '薪俸',
        1: '專業加給',
        2: '主管加給',
        3: '導師費',
        7: '特教加給',
        8: '地域加給',
        9: '應發金額',
    }
    for offset, key in positional_fields.items():
        ri = base_idx + offset
        if ri < len(rows):
            assign_num(rows[ri], key)

    for r in rows:
        label = (r[0]['text'] or '').strip()
        if label == '身分證字號':
            assign_text(r, '身分證')
        elif label in FIELD_ALIAS:
            assign_num(r, FIELD_ALIAS[label])

    out = []
    for c in cols:
        pay = c.get('薪俸', 0)
        if not pay:
            continue
        duty = c.get('導師費', 0) + c.get('特教加給', 0)
        s = (pay + c.get('主管加給', 0) + c.get('專業加給', 0) + duty
             + c.get('地域加給', 0) + c.get('其他加給', 0) - c.get('扣款', 0))
        total = c.get('應發金額', 0)
        pid = (c.get('身分證') or '').upper()
        out.append({
            '姓名': c['姓名'], '職稱': c.get('職稱', ''),
            '身分證': pid if looks_like_id(pid) else '',
            '身分證有效': valid_id(pid),
            '薪俸': pay, '主管加給': c.get('主管加給', 0),
            '專業加給': c.get('專業加給', 0), '導師特教': duty,
            '其他加給': c.get('其他加給', 0), '應發金額': total,
            '加總相符': (total > 0 and s == total),
            '加總差額': (total - s) if total else None,
            '最低信心': round(min(c['_conf']), 3),
        })
    return out


def parse_vertical(rows):
    """
    直式清冊通常一頁一組人員，每頁各有自己的「姓名」列。
    以頁碼分區解析，避免不同頁的欄位互相覆蓋。
    """
    pages = {}
    for r in rows:
        pages.setdefault(r[0]['page'], []).append(r)

    out = []
    for pno in sorted(pages):
        page_rows = pages[pno]
        starts = [i for i, r in enumerate(page_rows)
                  if (r[0]['text'] or '').strip() == '姓名' and len(r) > 1]
        if len(starts) <= 1:
            out.extend(_vertical_block(page_rows))
        else:
            # 同一頁有多組（少見）：以「姓名」列為界再切
            for k, s in enumerate(starts):
                begin = 0 if k == 0 else starts[k - 1] + 1
                end = starts[k + 1] - 2 if k + 1 < len(starts) else len(page_rows)
                out.extend(_vertical_block(page_rows[begin:max(end, s + 1)]))

    seen, uniq = set(), []
    for p in out:
        key = p['身分證'] or (p['姓名'], p['薪俸'])
        if key in seen:
            continue
        seen.add(key)
        uniq.append(p)
    return uniq


# ── 橫式解析 ────────────────────────────────────────────

def parse_horizontal_row(row):
    texts = [t['text'] for t in row]
    line = ' '.join(texts)
    if any(w in line for w in SKIP_WORDS):
        return None

    pid = next((s.strip().upper() for s in texts if looks_like_id(s)), '')

    # 橫式報表常是「職別｜薪點｜姓名｜月俸額」。資源班、暫代、專任教練
    # 等職別本身也像中文姓名，若一律取第一個 2～4 字中文 token，整列就會
    # 從職別開始錯位。優先找薪點右側、且右邊緊接大額月俸的姓名；另保留
    # 「姓名｜薪點｜月俸」版面的向左備援。
    name_candidates = [
        i for i, s in enumerate(texts)
        if is_name(s) and not _suspicious_name(s)
    ]
    point_indexes = [
        i for i, s in enumerate(texts)
        if is_num(s) and 0 < abs(to_int(s)) < 1000
    ]

    def has_pay_after(i):
        return any(is_num(texts[j]) and abs(to_int(texts[j])) >= 10000
                   for j in range(i + 1, min(len(texts), i + 5)))

    name_i = None
    for pi in point_indexes:
        right = [i for i in name_candidates
                 if i > pi and row[i]['x'] - row[pi]['x'] <= 0.11
                 and has_pay_after(i)]
        if right:
            name_i = min(right, key=lambda i: row[i]['x'] - row[pi]['x'])
            break
    if name_i is None:
        for pi in point_indexes:
            left = [i for i in name_candidates
                    if i < pi and row[pi]['x'] - row[i]['x'] <= 0.11
                    and has_pay_after(pi)]
            if left:
                name_i = min(left, key=lambda i: row[pi]['x'] - row[i]['x'])
                break
    if name_i is None:
        name_i = name_candidates[0] if name_candidates else None
    if name_i is None:
        return None
    name = texts[name_i]

    title = ' '.join(s for s in texts[:name_i] if not is_num(s))
    has_mgr = any(w in title for w in MGR_WORDS)

    nums = [to_int(s) for s in texts[name_i + 1:] if is_num(s)]
    while nums and 0 < nums[0] < 1000:      # 濾掉俸點
        nums.pop(0)
    if len(nums) < 2:
        return None

    pay, mgr, prof, duty, other, total = nums[0], 0, 0, 0, 0, 0
    total_inferred = False

    # 薪資區最後有一個「小計」，必定等於月俸＋前方各項加給。
    # 找最早成立的前綴和即可切開後方保險、扣款欄，不依賴職稱猜空白欄。
    total_i = next(
        (i for i in range(2, min(len(nums), 7))
         if nums[i] >= pay and nums[i] == sum(nums[:i])),
        None)
    if total_i is not None:
        total = nums[total_i]
        additions = nums[1:total_i]
        # 此類報表的第一個小額欄是「主管加給／特教職加」共用欄：
        # 主管職才歸主管加給，非主管（例如資源班）則須與後方導師費
        # 合併到 AF 的「導師＋特教」。
        # 專業加給通常是第一個至少 15,000 元的加給，可作為分界。
        prof_i = next((i for i, v in enumerate(additions) if v >= 15000), None)
        if prof_i is not None:
            leading_duty_or_mgr = sum(additions[:prof_i])
            if has_mgr:
                mgr = leading_duty_or_mgr
            else:
                duty = leading_duty_or_mgr
            prof = additions[prof_i]
            duty += sum(additions[prof_i + 1:])
        elif '教保員' in title:
            other = sum(additions)
        elif additions:
            # 沒有專業加給的少數人員，保留在「其他加給」參與驗算，
            # 不硬塞入 AF 的四個比對欄位。
            other = sum(additions)
    else:
        # 舊版面沒有可辨識的小計時，保留原有順序解析作為備援。
        idx = 1
        # 主管／特殊職務加給通常低於 15,000，且位於至少 15,000 的
        # 專業加給之前。即使職稱不是主任（例如資源班），也應照欄序讀取。
        small_before_prof = (
            idx + 1 < len(nums)
            and 0 <= nums[idx] < 15000
            and nums[idx + 1] >= 15000
        )
        if (has_mgr or small_before_prof) and idx < len(nums):
            if has_mgr:
                mgr = nums[idx]
            else:
                duty = nums[idx]
            idx += 1
        prof = nums[idx] if idx < len(nums) else 0
        idx += 1
        if idx < len(nums) and 0 < nums[idx] < 20000:
            duty += nums[idx]; idx += 1
        total_candidate = nums[idx] if idx < len(nums) else 0
        # 小計不可能小於月俸；若下一個數字已落入保險／扣款區（例如 3,107），
        # 不能冒充應發金額。暫填薪資項目合計供人工核對，但標記為推算，
        # 絕不讓它自動通過。
        if total_candidate >= pay:
            total = total_candidate
        else:
            total = pay + mgr + prof + duty + other
            total_inferred = True

    s = pay + mgr + prof + duty + other
    return {
        '姓名': name, '職稱': title.strip(),
        '身分證': pid, '身分證有效': valid_id(pid),
        '薪俸': pay, '主管加給': mgr, '專業加給': prof, '導師特教': duty,
        '其他加給': other,
        '應發金額': total,
        '應發金額推算': total_inferred,
        '加總相符': (total > 0 and s == total and not total_inferred),
        '加總差額': (total - s) if total else None,
        '最低信心': round(min(t['conf'] for t in row), 3),
    }


def parse_row(row):
    """保留舊介面：單列橫式解析"""
    return parse_horizontal_row(row)


# ── 主流程 ──────────────────────────────────────────────

def _parse_as(rows, layout):
    if layout == 'vertical':
        return parse_vertical(rows)
    return [p for p in (parse_horizontal_row(r) for r in rows) if p]


def _suspicious_name(name):
    s = (name or '').strip()
    if s in NON_NAME_LABELS:
        return True
    return any(w in s for w in ('健保', '公保', '勞保', '所得稅', '儲蓄',
                                '實發', '應發', '扣款', '加給'))


def _candidate_score(people):
    """評估解析結果是否像一份人員清冊；採比例，避免列數多者占便宜。"""
    if not people:
        return -100.0
    n = len(people)
    plausible = sum(1 for p in people if not _suspicious_name(p.get('姓名')))
    arith_ok = sum(1 for p in people if p.get('加總相符'))
    valid_ids = sum(1 for p in people if p.get('身分證有效'))
    suspicious = n - plausible
    return (10 * plausible / n
            + 4 * arith_ok / n
            + 3 * valid_ids / n
            - 12 * suspicious / n)


def layout_diagnostics(tokens):
    """回傳不含姓名或金額的雙版面診斷摘要，供 worker 異常時記錄。"""
    rows = group_rows(tokens)
    evidence = _layout_evidence(rows)
    detected = detect_layout(rows)
    result = {'evidence': evidence, 'detected': detected, 'candidates': {}}
    for layout in ('vertical', 'horizontal'):
        people = _parse_as(rows, layout)
        result['candidates'][layout] = {
            'people': len(people),
            'score': round(_candidate_score(people), 2),
            'arith_ok': sum(1 for p in people if p.get('加總相符')),
            'suspicious_names': sum(
                1 for p in people if _suspicious_name(p.get('姓名'))),
        }
    return result


def parse_tokens(tokens, preferred_layout=None):
    rows = group_rows(tokens)
    detected = detect_layout(rows)
    candidates = {
        'vertical': _parse_as(rows, 'vertical'),
        'horizontal': _parse_as(rows, 'horizontal'),
    }
    scores = {layout: _candidate_score(people)
              for layout, people in candidates.items()}

    # 自動結構判定與該校既有成功設定都只作為加分，不會永久強制。
    # 若學校日後換版面，另一候選的資料品質更好時仍會自動切換。
    scores[detected] += 1.0
    if preferred_layout in candidates:
        scores[preferred_layout] += 1.5

    layout = max(scores, key=scores.get)
    return candidates[layout], layout


def diagnose(tokens):
    """
    印出解析器實際看到的結構，用來判斷新版面卡在哪裡。
    不會顯示任何金額或姓名以外的個人資料。
    """
    rows = group_rows(tokens)
    layout = detect_layout(rows)
    print(f'=== 版面判定：{layout} ===\n')

    print('【各列的第一欄文字】（直式版面的項目名稱應出現在此）')
    known, unknown = [], []
    for i, r in enumerate(rows[:60]):
        first = (r[0]['text'] or '').strip()
        n = len(r)
        if first in FIELD_ALIAS:
            known.append(first)
            tag = f'→ 對應「{FIELD_ALIAS[first]}」'
        elif first in ('序號', '姓名', '職稱', '身分證字號', '支薪俸級', '編號'):
            tag = '→ 表頭'
        else:
            if first and n > 1:
                unknown.append(first)
            tag = ''
        print(f'  列{i:2d}（{n:2d} 個 token）  {first[:14]:16} {tag}')

    print('\n【已對應的項目】')
    print('  ' + ('、'.join(dict.fromkeys(known)) if known else '（無）'))

    print('\n【未對應的列標題】← 若薪資項目出現在這裡，請加入 FIELD_ALIAS')
    cand = [u for u in dict.fromkeys(unknown)
            if any(k in u for k in ('俸', '給', '費', '額', '金', '計', '薪',
                                    '貼', '津', '獎', '補', '扣', '數'))]
    print('  可能是薪資項目：' + ('、'.join(cand) if cand else '（無）'))
    other = [u for u in dict.fromkeys(unknown) if u not in cand]
    print('  其他：' + ('、'.join(other[:20]) if other else '（無）'))

    print('\n【身分證字號】')
    ids = [t['text'].strip() for t in tokens if looks_like_id(t['text'])]
    good = sum(1 for i in ids if valid_id(i))
    print(f'  找到 {len(ids)} 組，檢查碼有效 {good} 組'
          if ids else '  此清冊未包含身分證字號')

    print('\n【欄位對齊檢查】')
    if layout == 'vertical':
        pages = {}
        for r in rows:
            pages.setdefault(r[0]['page'], []).append(r)
        pr = pages[sorted(pages)[0]]
        nr = next((r for r in pr
                   if (r[0]['text'] or '').strip() == '姓名' and len(r) > 1), None)
        vr = next((r for r in pr
                   if (r[0]['text'] or '').strip() in FIELD_ALIAS and len(r) > 1), None)
        if nr:
            print('  姓名列 x：' + '  '.join(
                f"{(t['text'] or '')[:4]}@{t['x']:.3f}" for t in nr[1:6]))
        if vr:
            print(f"  {(vr[0]['text'] or '')}列 x：" + '  '.join(
                f"{(t['text'] or '')[:7]}@{t['x']:.3f}" for t in vr[1:6]))
        if nr and vr:
            d = min(abs(a['x'] - b['x']) for a in vr[1:3] for b in nr[1:3])
            print(f'  最小水平偏移：{d:.3f}（需小於容差才會對上）')

    print('\n【解析結果】')
    people, _ = parse_tokens(tokens)
    ok = sum(1 for p in people if p['加總相符'])
    print(f'  解析 {len(people)} 人，加總相符 {ok} 人')
    if people and ok < len(people) * 0.8:
        print('\n  ⚠ 相符率偏低，可能原因：')
        print('    · 版面判定錯誤（上方顯示的版面與實際不符）')
        print('    · 有薪資項目未對應（見上方「未對應的列標題」）')
        print('    · 掃描品質不佳導致數字辨識錯誤')


def inspect_person(tokens, name):
    """列出某位人員在各薪資項目列中，被歸到他名下的原始 token"""
    rows = group_rows(tokens)
    pages = {}
    for r in rows:
        pages.setdefault(r[0]['page'], []).append(r)

    for pno in sorted(pages):
        pr = pages[pno]
        nr = next((r for r in pr
                   if (r[0]['text'] or '').strip() == '姓名' and len(r) > 1), None)
        if not nr:
            continue
        hit = next((t for t in nr[1:] if (t['text'] or '').strip() == name), None)
        if not hit:
            continue

        print(f'=== {name}（第 {pno + 1} 頁，姓名 x={hit["x"]:.3f}）===\n')
        for r in pr:
            lab = (r[0]['text'] or '').strip()
            if lab not in FIELD_ALIAS and lab not in ('身分證字號', '職稱'):
                continue
            near = sorted(r[1:], key=lambda t: abs(t['x'] - hit['x']))[:3]
            desc = '  '.join(f"{(t['text'] or '')!r}@{t['x']:.3f}" for t in near)
            print(f'  {lab:10} 最接近的 3 個 token：{desc}')
        return
    print(f'找不到「{name}」')


def main():
    if '--person' in sys.argv:
        i = sys.argv.index('--person')
        name = sys.argv[i + 1]
        del sys.argv[i:i + 2]
        inspect_person(json.load(open(sys.argv[1], encoding='utf-8')), name)
        return
    if '--diagnose' in sys.argv:
        sys.argv.remove('--diagnose')
        diagnose(json.load(open(sys.argv[1], encoding='utf-8')))
        return
    tokens = json.load(open(sys.argv[1], encoding='utf-8'))
    people, layout = parse_tokens(tokens)

    ok = [p for p in people if p['加總相符']]
    with_id = [p for p in people if p['身分證有效']]

    print(f'版面：{"直式（一人一欄）" if layout == "vertical" else "橫式（一人一列）"}')
    print(f'解析 {len(people)} 人｜加總相符 {len(ok)}｜身分證有效 {len(with_id)}\n')
    print(f'{"姓名":6}{"身分證":12}{"職稱":14}{"薪俸":>8}{"主管":>7}'
          f'{"專業":>8}{"導師特教":>8}{"應發":>9}  檢查')
    for p in people:
        mark = '✔' if p['加總相符'] else f'✗ 差{p["加總差額"]}'
        idm = p['身分證'] if p['身分證有效'] else (p['身分證'] + '?' if p['身分證'] else '—')
        print(f'{p["姓名"]:6}{idm:12}{p["職稱"][:12]:14}{p["薪俸"]:>8}'
              f'{p["主管加給"]:>7}{p["專業加給"]:>8}{p["導師特教"]:>8}'
              f'{p["應發金額"]:>9}  {mark}')

    json.dump(people, open('parsed.json', 'w', encoding='utf-8'),
              ensure_ascii=False, indent=2)
    print('\n已輸出 parsed.json')


if __name__ == '__main__':
    main()
