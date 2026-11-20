#!/usr/bin/env python3
"""
把 ocr_extract.swift 的 token JSON 還原成人員資料。

支援三種薪資清冊版面：
  橫式：一人一列（編號 職稱 姓名 俸點 薪俸 …）
  橫式雙行：一人分成薪資／補助兩行（姓名位於第二行）
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
    # 明正國小等清冊：事病假扣薪印為負數、補薪為正數，兩者都計入應發數合計。
    '事病假扣薪': '其他加給',
    '補薪': '其他加給',
    '補發職務加給': '其他加給',
    '補發導師費': '其他加給',
    '特教加給差額': '其他加給',
    '應發數合計': '應發金額', '應發金額': '應發金額', '小計': '應發金額',
}

# 直式項目名稱中，代表「自應發扣除」的列；不論印成正數或負數都以負值計入。
DEDUCT_LABELS = ('事病假扣薪',)

# 掃描常見的形近誤讀，只用於比對列標題（例：「職務加紿」「地域加紿」）。
_LABEL_FIX = str.maketrans({'紿': '給', '洽': '給', '絵': '給', '新': '薪'})

# 直式總計表（依科室／職別彙總，非個人資料）的表頭標記；整頁跳過。
SUMMARY_PAGE_LABELS = ('員工人數', '科室組別')

# 直式表頭各列的列標題，用來偵測「表頭多列被併成同一列」的歪斜掃描。
VERTICAL_HEADER_LABELS = ('序號', '身分證字號', '姓名', '職稱', '支薪俸級', '支新俸級')

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


# 身分證字號首碼的掃描形近誤讀（例：E 讀成 B、英文 O 讀成數字 0）。
_ID_FIRST_LOOKALIKE = {
    # O／Q／D 三者檢查碼相同、無法以檢查碼區分，故數字 0 只還原為最常見的 O。
    'B': 'E', 'E': 'BF', 'F': 'E', '8': 'B', '0': 'O',
    '1': 'I', 'L': 'I', 'I': 'L', 'P': 'R', 'R': 'P', '5': 'S',
}


def repair_id(s):
    """
    檢查碼不符的身分證字號，只嘗試把「首碼」換成形近字母；恰好只有一個
    替代字母能通過檢查碼時才採用，否則原樣回傳（不猜測數字部分）。
    """
    s = re.sub(r'\s', '', (s or '')).upper()
    if len(s) != 10 or valid_id(s) or not s[1:].isdigit():
        return s
    fixes = [c + s[1:] for c in _ID_FIRST_LOOKALIKE.get(s[0], '')
             if valid_id(c + s[1:])]
    return fixes[0] if len(fixes) == 1 else s


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
    s = (s or '').translate(_SEP_CHARS).strip()
    # 同一個千分位符號偶爾同時被讀成「.，」，正規化後會變成兩個逗號。
    # 收斂連續分隔符，讓「40.，760」仍可安全還原成「40,760」。
    return re.sub(r',+', ',', s)


def is_num(s):
    # 掃描 OCR 常在金額前後黏到框線、括號或一個近似直線的英文字元。
    # 只容許有限的邊緣雜訊，而且逗號必須符合千分位格式。
    # 例如「4,0001」通常是右側框線被讀成 1；不能直接去掉逗號變成 40001。
    match = re.fullmatch(
        r'[~～LlI|<>＜＞•]*(-?[\d,]+?)[._,\-/|()（）\[\]［］{}]*',
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
    THREE = re.compile(r'^\d{3}[,\-/|()（）\[\]［］_]*$')
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


def merge_token_sets(*token_sets):
    """
    合併同一份 PDF 的多次 OCR 結果（例如整頁＋分塊），去除重複辨識的同一格文字。

    兩個 token 位於同頁、基線相差不到 0.006，且水平範圍重疊至少為較窄者的
    一半，即視為同一格；保留信心較高者，同分時保留字數較多者（較不可能被截斷）。
    """
    merged = []
    by_page = {}
    for tokens in token_sets:
        for t in tokens or []:
            if not (t.get('text') or '').strip():
                continue
            cand = dict(t)
            left = cand.get('x', 0) - cand.get('w', 0) / 2
            right = cand.get('x', 0) + cand.get('w', 0) / 2
            dup_index = None
            for i in by_page.get(cand.get('page'), []):
                o = merged[i]
                if abs(o.get('y', 0) - cand.get('y', 0)) > 0.006:
                    continue
                ol = o.get('x', 0) - o.get('w', 0) / 2
                orr = o.get('x', 0) + o.get('w', 0) / 2
                overlap = min(right, orr) - max(left, ol)
                narrow = min(right - left, orr - ol)
                if narrow <= 0 or overlap < narrow * 0.5:
                    continue
                dup_index = i
                break
            if dup_index is None:
                by_page.setdefault(cand.get('page'), []).append(len(merged))
                merged.append(cand)
                continue
            o = merged[dup_index]
            better = (float(cand.get('conf', 0)), len(norm_text(cand['text']))) > \
                     (float(o.get('conf', 0)), len(norm_text(o['text'])))
            if better:
                merged[dup_index] = cand
    return merged


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


def _is_two_line_horizontal(rows):
    """判斷一人資料分布在「薪資／政府補助」上下兩行的橫式清冊。"""
    pages = {}
    for r in rows:
        if r:
            pages.setdefault(r[0]['page'], []).append(r)

    for page_rows in pages.values():
        texts = [row_text(r) for r in page_rows[:8]]
        header = ''.join(texts)
        has_primary = all(label in header for label in
                          ('姓名', '月支薪額', '專業加給'))
        subsidy_labels = sum(label in header for label in
                             ('公提勞退', '補助公保', '補助健保', '補助退撫'))
        if has_primary and subsidy_labels >= 2:
            return True
    return False


def detect_layout(rows):
    """
    直式版面的特徵：某一列以「姓名」開頭，後面接多個人名；
    且另有一列以「本俸／薪俸」等項目名稱開頭。
    """
    if _is_two_line_horizontal(rows) or _is_grid_two_line(rows):
        return 'horizontal'

    heads, fields, name_rows = _layout_evidence(rows)
    vertical = (
        (name_rows >= 1 and fields >= 1)
        or (heads >= 2 and fields >= 2)
        or fields >= 4
    )
    return 'vertical' if vertical else 'horizontal'


# ── 直式解析 ────────────────────────────────────────────

def _row_label(r):
    """
    直式清冊的項目名稱有時被拆成兩個 token（「專業」「加給」、「應發數」
    「合計」）。第一個 token 不是已知項目時，與緊鄰的第二個中文 token
    合併後再查表；已知項目或合併後仍未知時，維持原本只看第一個 token。
    """
    first = (r[0]['text'] or '').strip() if r else ''
    if first not in FIELD_ALIAS:
        # 形近字（紿→給）或黏到框線的尾端數字（「補發導師費18」）
        fixed = re.sub(r'\d+$', '', first.translate(_LABEL_FIX))
        if fixed in FIELD_ALIAS:
            first = fixed
    if first in FIELD_ALIAS or len(r) < 2:
        return first
    second = (r[1]['text'] or '').strip().translate(_LABEL_FIX)
    if (second and len(second) <= 4 and not is_num(second)
            and all('\u4e00' <= ch <= '\u9fff' for ch in second)
            and abs(r[1]['x'] - r[0]['x']) < 0.05
            and first + second in FIELD_ALIAS):
        return first + second
    return first


def _is_summary_page(page_rows):
    """
    直式「總計表」依科室／職別彙總（列標題含「員工人數」「科室組別」），
    欄位是職別而非個人，也沒有身分證字號；不可當成人員解析。
    """
    labels = {(r[0].get('text') or '').strip() for r in page_rows if r}
    if not labels & set(SUMMARY_PAGE_LABELS):
        return False
    return not any(looks_like_id((t.get('text') or '').strip())
                   for r in page_rows for t in r)


def _vertical_header_rows_ok(header_rows, n_people):
    """標準逐列解析的姓名列是否恰好每位人員一個姓名（可直接信任）。"""
    for r in header_rows:
        if r and (r[0].get('text') or '').strip() == '姓名':
            got = [t for t in r[1:] if is_name((t.get('text') or '').strip())
                   and not _suspicious_name(t.get('text'))
                   and not any(w in t['text'] for w in SKIP_WORDS)]
            return len(got) == n_people
    return False


def _vertical_header_merged(header_rows):
    """表頭多列（序號／身分證／姓名／職稱）因掃描歪斜被併成同一列。"""
    for r in header_rows:
        labels = {(t.get('text') or '').strip() for t in r
                  if t.get('x', 1) < 0.16}
        if len(labels & set(VERTICAL_HEADER_LABELS)) >= 2:
            return True
    return False


def _vertical_header_columns(rows, base_idx, num_xs):
    """
    逐欄解析直式表頭：不依賴「同一列」分組，而以金額欄（靠右對齊）切出
    每人的欄位範圍，再於欄內由上而下找身分證字號、姓名、職稱。

    用於職稱折成多行、各欄表頭上下錯位，導致序號／身分證／姓名／職稱
    被併成同一列的掃描清冊（例：「薦任七等功六級（590）」三行職稱）。
    需有「小計／合計」表頭界定人員欄的右界，且每一人員欄都恰好讀到
    一個姓名，否則回傳空串列，沿用原本的逐列解析。
    """
    header = [t for r in rows[:base_idx] for t in r]
    # 直式表頭的列標題一律在最左欄；不在左欄者是橫式表頭，不適用。
    labels = [t for t in header
              if (t.get('text') or '').strip() in VERTICAL_HEADER_LABELS
              and t.get('x', 1) < 0.16]
    if not num_xs or not any(t['text'].strip() == '姓名' for t in labels):
        return []
    totals = [t['x'] for t in header
              if (t.get('text') or '').strip() in ('小計', '合計', '總計')]
    if not totals:
        return []
    right = min(totals)
    person_xs = [x for x in num_xs if x < right]
    if not person_xs:
        return []
    top = min(t['y'] for t in labels) - 0.012
    left = max(t['x'] for t in labels) + 0.02
    name_label_y = next((t['y'] for t in labels
                         if (t.get('text') or '').strip() == '姓名'), None)
    out = []
    for i, nx in enumerate(person_xs):
        lo = person_xs[i - 1] if i else left
        cell = sorted((t for t in header
                       if lo < t['x'] < nx and t['y'] >= top
                       and (t.get('text') or '').strip()
                       not in VERTICAL_HEADER_LABELS),
                      key=lambda t: t['y'])
        ids = [t for t in cell if looks_like_id(repair_id(t.get('text')))]
        if len(ids) > 1:
            return []            # 一個金額欄內有兩人：該列金額漏讀，欄界不可信
        pid = ids[0] if ids else None
        cands = [t for t in cell
                 if is_name((t.get('text') or '').strip())
                 and not _suspicious_name(t.get('text'))
                 and not any(w in t['text'] for w in SKIP_WORDS)]
        if pid is not None:
            after = [t for t in cands if t['y'] > pid['y'] - 0.002]
            name = after[0] if after else None
        elif name_label_y is not None and cands:
            near = [t for t in cands if abs(t['y'] - name_label_y) <= 0.012]
            if len(near) > 1:
                return []
            name = min(cands, key=lambda t: abs(t['y'] - name_label_y))
        else:
            name = None
        if name is None:
            return []
        title = next((t for t in cell
                      if t['y'] > name['y'] + 0.002
                      and not is_num(t['text'])
                      and re.search(r'[\u4e00-\u9fff]', t['text'])
                      and not re.fullmatch(r'[\d（）()\s]+', t['text'])), None)
        entry = {'姓名': name['text'].strip(), 'x': name['x'],
                 '_conf': [name['conf']]}
        if pid is not None:
            entry['身分證'] = pid['text'].strip()
            entry['_conf'].append(pid['conf'])
        if title is not None:
            entry['職稱'] = title['text'].strip()
        out.append(entry)
    return out


def _geometric_columns(names, rows, base_idx):
    """
    以所有主要金額列（本俸／專業加給／應發）的金額 x 座標分群，為每個
    姓名取「姓名右側最近」的金額欄。必須一人一欄、彼此不重複，且所有
    主要金額列中多數人都有讀到金額，否則回傳 None 沿用依序配對。
    """
    xs = []
    used_rows = 0
    for r in rows[base_idx:]:
        if FIELD_ALIAS.get(_row_label(r)) not in ('薪俸', '專業加給', '應發金額'):
            continue
        vals = [t['x'] for t in r[1:] if is_num(t.get('text', ''))
                and abs(to_int(t.get('text', ''))) >= 1000]
        if vals:
            used_rows += 1
            xs.extend(vals)
    if used_rows < 2 or not names:
        return None
    xs.sort()
    clusters = []
    for x in xs:
        if clusters and x - clusters[-1][-1] <= 0.02:
            clusters[-1].append(x)
        else:
            clusters.append([x])
    # 只採至少兩列都出現的欄，避免單一雜訊 token 自成一欄
    centers = [sum(c) / len(c) for c in clusters if len(c) >= 2]
    out = []
    for i, n in enumerate(names):
        right = [c for c in centers if c > n['x'] + 0.005]
        nxt = names[i + 1]['x'] if i + 1 < len(names) else None
        if right and (nxt is None or right[0] <= nxt + 0.04):
            out.append(right[0])
        else:
            out.append(None)     # 此人主要金額全數漏讀：不可借用隔壁欄
    found = [x - n['x'] for x, n in zip(out, names) if x is not None]
    if not found:
        return None
    found.sort()
    offset = found[len(found) // 2]
    out = [x if x is not None else n['x'] + offset for x, n in zip(out, names)]
    if len(set(out)) != len(out):
        return None
    return out


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
                     if label_pred(_row_label(r)) and len(r) > 1), None)

    def money_tokens(row, minimum=1000):
        return [t for t in row[1:]
                if is_num(t.get('text', '')) and abs(to_int(t.get('text', ''))) >= minimum]

    # 先找本俸列。標準版面使用 FIELD_ALIAS；掃描品質差時，改找表頭後
    # 第一列「至少三個四位數以上金額」，避免依賴容易誤讀的「本俸」二字。
    base_idx = next(
        (i for i, r in enumerate(rows)
         if FIELD_ALIAS.get(_row_label(r)) == '薪俸'
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
            lab = _row_label(r)
            if (FIELD_ALIAS.get(lab) == want
                    and len(money_tokens(r)) >= len(names)):
                ref = r
                break
        if ref:
            break
    if ref is None:
        ref = base_row

    num_xs = sorted(t['x'] for t in money_tokens(ref))
    # 逐欄表頭解析用「金額最多」的主要金額列界定欄位，避免應發列有斷字。
    col_ref = max((r for r in rows
                   if FIELD_ALIAS.get(_row_label(r)) in ('薪俸', '應發金額',
                                                         '專業加給')),
                  key=lambda r: len(money_tokens(r)), default=ref)
    col_xs = sorted(t['x'] for t in money_tokens(col_ref))
    colwise = _vertical_header_columns(rows, base_idx, col_xs)
    use_colwise = bool(colwise) and (
        _vertical_header_merged(rows[:base_idx])
        or (len(names) != len(colwise)
            and not _vertical_header_rows_ok(rows[:base_idx], len(colwise))))
    if use_colwise:
        names = colwise
        num_xs = col_xs
    # 只有在依序配對「明顯錯位」（某人的金額欄已越過下一人的姓名）時，
    # 才改用幾何配對；一般情形維持原本依序配對。
    crossed = any(num_xs[i] > names[i + 1]['x'] + 0.01
                  for i in range(min(len(names) - 1, len(num_xs))))
    geo = _geometric_columns(names, rows, base_idx) if crossed else None
    if geo is not None and geo != num_xs[:len(names)]:
        # 參考列缺了某人的金額（斷字、漏讀）時，依序配對會讓後面的人整批
        # 錯位到隔壁或小計欄；改依幾何位置配對（金額靠右、姓名靠左，
        # 每人的金額欄是姓名右側最近的金額欄）。
        num_xs = geo
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
    if use_colwise:
        pass                      # 職稱／身分證已由逐欄表頭解析取得
    elif title_row:
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
    # 文字辨識失敗時的安全備援。若本頁已有至少三種可靠欄名（竹滬等
    # 標準直式清冊），就完全停用位置推測，避免空白列消失後欄序位移。
    positional_fields = {
        0: '薪俸',
        1: '專業加給',
        2: '主管加給',
        3: '導師費',
        7: '特教加給',
        8: '地域加給',
        9: '應發金額',
    }
    exact_fields = {
        FIELD_ALIAS[_row_label(r)]
        for r in rows
        if _row_label(r) in FIELD_ALIAS
    }
    if len(exact_fields) < 3:
        for offset, key in positional_fields.items():
            if key in exact_fields:
                continue
            ri = base_idx + offset
            if ri < len(rows):
                assign_num(rows[ri], key)

    for r in rows:
        label = _row_label(r)
        if label == '身分證字號':
            if not use_colwise:
                assign_text(r, '身分證')
        elif FIELD_ALIAS.get(label) == '其他加給':
            # 教保費、行政工作獎金、事病假扣薪、補薪等可能同時出現；各自保存
            # 再加總，避免同一人的多個「其他」項目互相覆蓋。
            assign_num(r, '_x_' + label)
        elif label in FIELD_ALIAS:
            assign_num(r, FIELD_ALIAS[label])
    for c in cols:
        for lab in DEDUCT_LABELS:
            if '_x_' + lab in c:
                c['_x_' + lab] = -abs(c['_x_' + lab])

    def other(c):
        return c.get('其他加給', 0) + sum(
            v for k, v in c.items() if k.startswith('_x_'))

    def comp_sum(c):
        return (c.get('薪俸', 0) + c.get('主管加給', 0) + c.get('專業加給', 0)
                + c.get('導師費', 0) + c.get('特教加給', 0)
                + c.get('地域加給', 0) + other(c) - c.get('扣款', 0))

    # 找不到「應發數合計」標籤時（標籤列被標題文字擠掉），改以算術證明：
    # 取「多數人該列數值恰等於各項加總」的那一列作為應發金額。不猜位置。
    if not any(c.get('應發金額') for c in cols):
        best_row, best_hits = None, 0
        for r in rows[base_idx + 1:]:
            vals = {}
            for t in r:
                txt = (t['text'] or '').strip()
                if not txt or not is_num(txt):
                    continue
                bc, bd = None, 9
                for c in cols:
                    d = abs(t['x'] - c['nx'])
                    if d < bd:
                        bc, bd = c, d
                if bc is not None and bd <= tol:
                    vals[id(bc)] = to_int(txt)
            hits = sum(1 for c in cols if comp_sum(c) and vals.get(id(c)) == comp_sum(c))
            if hits > best_hits:
                best_row, best_hits = r, hits
        if best_row is not None and best_hits >= max(2, (len(cols) + 1) // 2):
            assign_num(best_row, '應發金額')

    out = []
    for c in cols:
        pay = c.get('薪俸', 0)
        if not pay:
            continue
        duty = c.get('導師費', 0) + c.get('特教加給', 0)
        s = (pay + c.get('主管加給', 0) + c.get('專業加給', 0) + duty
             + c.get('地域加給', 0) + other(c) - c.get('扣款', 0))
        total = c.get('應發金額', 0)
        extra = c.get('地域加給', 0) + other(c) - c.get('扣款', 0)
        if total and extra and s != total and s - extra == total:
            # 地域／其他列誤收了鄰列數字：不含這些項目時恰等於印列應發，
            # 以算術證明剔除，不猜測金額。
            for k in ['地域加給', '其他加給', '扣款'] + [
                    k for k in c if k.startswith('_x_')]:
                c.pop(k, None)
            s = total
        pid = repair_id(c.get('身分證'))
        out.append({
            '姓名': c['姓名'], '職稱': c.get('職稱', ''),
            '身分證': pid if looks_like_id(pid) else '',
            '身分證有效': valid_id(pid),
            '薪俸': pay, '主管加給': c.get('主管加給', 0),
            '專業加給': c.get('專業加給', 0), '導師特教': duty,
            # 地域加給、導師加給扣款等不屬 AF 四欄，但計入應發；併入「其他加給」
            # 讓伺服器端 from_ocr 用同一條等式驗算，避免已相符的列被誤送人工確認。
            '其他加給': (c.get('地域加給', 0) + other(c)
                     - c.get('扣款', 0)),
            '應發金額': total,
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
        if _is_summary_page(page_rows):
            continue
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

    # 直式掃描偶爾只漏掉個別人的導師／特教加給，但印列應發總額仍正確。
    # 僅在差額等於同份文件已明確讀到的導師特教金額時補回，並降低信心；
    # 未出現在文件中的新金額不猜測。
    known_duties = {p.get('導師特教', 0) for p in uniq
                    if p.get('導師特教', 0) > 0}
    for p in uniq:
        if p.get('加總相符') or not p.get('應發金額'):
            continue
        diff = p.get('加總差額')
        if diff in known_duties and 0 < diff <= 20000:
            p['導師特教'] += diff
            p['加總相符'] = True
            p['加總差額'] = 0
            p['薪資欄位推算'] = True
            p['最低信心'] = min(p.get('最低信心', 1.0), 0.3)
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
            # 蔡文版面的特教職加為 2,800；真正主管加給則常見
            # 4,320／5,930／10,010。職稱偶爾會被 OCR 截成「兼主1」，
            # 因此除職稱外也保留金額判斷，避免把主管加給併入特教。
            leading_is_special = (
                leading_duty_or_mgr == 2800 or '資源班' in title)
            if leading_is_special and not has_mgr:
                duty = leading_duty_or_mgr
            else:
                mgr = leading_duty_or_mgr
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
            leading_is_special = nums[idx] == 2800 or '資源班' in title
            if leading_is_special and not has_mgr:
                duty = nums[idx]
            else:
                mgr = nums[idx]
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


def _header_column(page_rows, labels, fallback=None):
    for r in page_rows[:8]:
        for t in r:
            text = (t.get('text') or '').strip()
            if any(label in text for label in labels):
                return t['x']
    return fallback


def _two_line_name(block, name_x, top_y):
    """從第二行姓名欄取人名，並處理姓名與職稱黏成同一 token 的情形。"""
    candidates = []
    for t in block:
        if abs(t.get('x', 0) - name_x) > 0.032:
            continue
        dy = t.get('y', top_y) - top_y
        if not 0.006 <= dy <= 0.024:
            continue

        raw = (t.get('text') or '').strip()
        runs = re.findall(r'[\u4e00-\u9fff]{2,10}', raw)
        for run in runs:
            pieces = [run]
            if len(run) > 4:
                # 常見錯誤是「姓名＋職稱」黏在一起；姓名通常位於字串前端。
                pieces.extend((run[:3], run[:4], run[-3:], run[-4:]))
            for piece in pieces:
                if not is_name(piece) or _suspicious_name(piece):
                    continue
                score = (float(t.get('conf', 0))
                         - abs(t.get('x', 0) - name_x) * 2
                         - abs(dy - 0.014) * 2)
                candidates.append((score, piece, t))

    if not candidates:
        return '', 0.0
    _, name, token = max(candidates, key=lambda item: item[0])
    return name, float(token.get('conf', 0))


def _two_line_title(block, title_x, top_y):
    parts = []
    for t in block:
        text = (t.get('text') or '').strip()
        if (abs(t.get('x', 0) - title_x) <= 0.032
                and abs(t.get('y', top_y) - top_y) <= 0.006
                and text and not is_num(text)):
            parts.append((t['x'], text))
    return ' '.join(text for _, text in sorted(parts))


def _two_line_amount(block, column_x, top_y, y_tol=0.006, x_tol=0.018):
    """只讀第一行薪資欄，刻意排除第二行同 x 座標的政府補助數字。"""
    candidates = []
    for t in block:
        text = t.get('text', '')
        if (abs(t.get('x', 0) - column_x) <= x_tol
                and abs(t.get('y', top_y) - top_y) <= y_tol
                and is_num(text)):
            value = to_int(text)
            candidates.append((abs(t['x'] - column_x), -float(t.get('conf', 0)),
                               value, t))
    if not candidates:
        return None, None
    _, _, value, token = min(candidates)
    return value, token


def _infer_two_line_components(people):
    """用同份清冊已辨識的欄值與應發金額補回少數漏讀欄位。"""
    keys = ('薪俸', '主管加給', '專業加給', '_導師', '_特教')
    observed = {key: {} for key in keys}
    for p in people:
        for key in keys:
            value = p.get(key)
            if value:
                observed[key][value] = observed[key].get(value, 0) + 1

    for p in people:
        total = p.get('應發金額')
        if not total:
            continue
        missing = [key for key in keys if p.get(key) is None]
        if not missing:
            continue

        known_sum = sum(p.get(key) or 0 for key in keys if key not in missing)
        target = total - known_sum
        if target < 0:
            continue

        # 空白欄本來就代表 0；非零候選只取同份文件實際出現過的值。
        choices = []
        for key in missing:
            vals = [0] + sorted(observed[key],
                                key=lambda v: (-observed[key][v], v))
            choices.append(vals)

        solutions = []
        def search(i, current, picked):
            if current > target:
                return
            if i == len(missing):
                if current == target:
                    nonzero = sum(bool(v) for v in picked)
                    popularity = sum(observed[k].get(v, 0)
                                     for k, v in zip(missing, picked) if v)
                    solutions.append((nonzero, -popularity, tuple(picked)))
                return
            for value in choices[i]:
                search(i + 1, current + value, picked + [value])
        search(0, 0, [])

        if not solutions:
            continue
        # 優先最少補值，再選文件中出現頻率最高者；同分才視為不確定。
        solutions.sort()
        best = solutions[0]
        if len(solutions) > 1 and solutions[1][:2] == best[:2]:
            continue
        for key, value in zip(missing, best[2]):
            p[key] = value
        if any(best[2]):
            p['_component_inferred'] = True


def _reconcile_two_line_page_totals(people, page_totals):
    """以每頁印列小計補回漏讀欄，且不跨頁挪用金額。"""
    keys = ('薪俸', '主管加給', '專業加給', '_導師', '_特教')
    observed = {key: {} for key in keys}
    for p in people:
        for key in keys:
            value = p.get(key)
            if value:
                observed[key][value] = observed[key].get(value, 0) + 1

    for pno, expected in page_totals.items():
        page_people = [p for p in people if p.get('_page') == pno]
        for key in keys:
            target = expected.get(key)
            if target is None:
                continue
            missing = [p for p in page_people if p.get(key) is None]
            known = sum(p.get(key) or 0 for p in page_people if p.get(key) is not None)
            residual = target - known
            if not missing or residual < 0:
                continue
            if residual == 0:
                for p in missing:
                    p[key] = 0
                continue
            if len(missing) == 1:
                missing[0][key] = residual
                missing[0]['_component_inferred'] = True
                continue

            evidenced = [p for p in missing if p.get('_column_evidence', {}).get(key)]
            if len(evidenced) == 1:
                # 該格確實有墨跡但數字被黏合；用頁小計的唯一差額修復。
                for p in missing:
                    p[key] = residual if p is evidenced[0] else 0
                evidenced[0]['_component_inferred'] = True
                continue

            # 多人同欄同時漏讀時，只接受「每人差額完全相同且該值已在本檔
            # 其他列出現」的情形；不排列組合猜測每個人應拿哪一個金額。
            if residual % len(missing):
                continue
            shared = residual // len(missing)
            if shared not in observed[key]:
                continue
            for p in missing:
                p[key] = shared
                if shared:
                    p['_component_inferred'] = True


def parse_horizontal_two_line(rows):
    """解析一人分成「薪資／補助」上下兩行的橫式清冊。"""
    pages = {}
    for r in rows:
        if r:
            pages.setdefault(r[0]['page'], []).append(r)

    out = []
    page_totals = {}
    page_gross = {}
    for pno in sorted(pages):
        page_rows = pages[pno]
        salary_x = _header_column(page_rows, ('月支薪額',))
        name_x = _header_column(page_rows, ('姓名',))
        if salary_x is None or name_x is None:
            continue

        prof_x = _header_column(page_rows, ('專業加給',), salary_x + 0.045)
        mgr_x = _header_column(page_rows, ('主管加給',), salary_x + 0.09)
        teacher_x = _header_column(page_rows, ('導師職加', '導師加給'),
                                   salary_x + 0.135)
        special_x = _header_column(page_rows, ('特教職加', '特教加給'),
                                   salary_x + 0.18)
        title_x = _header_column(page_rows, ('職稱',), name_x + 0.03)

        gross_x = None
        for r in page_rows[:8]:
            for t in r:
                text = (t.get('text') or '').strip()
                if 0.40 <= t.get('x', 0) <= 0.47 and '應' in text and '金額' in text:
                    gross_x = t['x']
                    break
            if gross_x is not None:
                break
        if gross_x is None:
            gross_x = special_x + 0.087

        # 每頁底部都印有本頁小計；它是修復個別欄位漏讀最可靠的約束。
        footer_labels = [t for r in page_rows for t in r
                         if t.get('y', 0) > 0.80
                         and any(word in (t.get('text') or '')
                                 for word in ('合計', '小計'))]
        footer_y = min((t['y'] for t in footer_labels), default=None)
        total_base = [t for r in page_rows for t in r
                      if footer_y is not None
                      and abs(t.get('y', 0) - footer_y) <= 0.02
                      and abs(t.get('x', 0) - salary_x) <= 0.018
                      and is_num(t.get('text', ''))]
        if total_base:
            total_base_token = max(total_base, key=lambda t: abs(to_int(t['text'])))
            total_tokens = [t for r in page_rows for t in r]
            totals = {}
            for key, x in (('薪俸', salary_x), ('專業加給', prof_x),
                           ('主管加給', mgr_x), ('_導師', teacher_x),
                           ('_特教', special_x)):
                totals[key], _ = _two_line_amount(
                    total_tokens, x, total_base_token['y'], y_tol=0.006)
            # 小計欄空白在此版面明確代表 0，不是 OCR 漏讀。
            page_totals[pno] = {key: value or 0 for key, value in totals.items()}
            # 應發小計可能落在小計列的上行或下行，收集該欄所有候選數字；
            # 之後必須與本頁應發合計「完全相同」才算佐證。
            page_gross[pno] = {
                abs(to_int(t['text'])) for t in total_tokens
                if abs(t.get('y', 0) - total_base_token['y']) <= 0.012
                and abs(t.get('x', 0) - gross_x) <= 0.021
                and is_num(t.get('text', ''))}

        name_header_y = max(
            t['y'] for r in page_rows[:8] for t in r
            if (t.get('text') or '').strip() == '姓名')
        total_ys = [t['y'] for r in page_rows for t in r
                    if t['y'] > name_header_y + 0.01
                    and any(word in (t.get('text') or '')
                            for word in ('合計', '小計'))]
        body_end_y = min(total_ys) - 0.01 if total_ys else 0.98

        starts = []
        for ri, r in enumerate(page_rows):
            base_candidates = [
                t for t in r
                if abs(t.get('x', 0) - salary_x) <= 0.018
                and is_num(t.get('text', ''))
                and 15000 <= abs(to_int(t.get('text', ''))) <= 150000
                and name_header_y < t.get('y', 0) < body_end_y
            ]
            if base_candidates:
                base = min(base_candidates, key=lambda t: abs(t['x'] - salary_x))
                starts.append((ri, base))

        for order, (start_i, base_token) in enumerate(starts):
            end_i = starts[order + 1][0] if order + 1 < len(starts) else len(page_rows)
            block = [t for r in page_rows[start_i:end_i] for t in r
                     if t.get('y', 0) < body_end_y]
            top_y = base_token['y']

            name, name_conf = _two_line_name(block, name_x, top_y)
            if not name:
                continue
            title = _two_line_title(block, title_x, top_y)

            values = {}
            column_evidence = {}
            value_tokens = []
            for key, x in (('薪俸', salary_x), ('專業加給', prof_x),
                           ('主管加給', mgr_x), ('_導師', teacher_x),
                           ('_特教', special_x)):
                column_evidence[key] = any(
                    abs(t.get('x', 0) - x) <= 0.018
                    and abs(t.get('y', top_y) - top_y) <= 0.006
                    and (t.get('text') or '').strip()
                    for t in block)
                value, source = _two_line_amount(block, x, top_y)
                upper = 150000 if key == '薪俸' else 100000
                if value is not None and not (0 <= value <= upper):
                    value, source = None, None
                values[key] = value
                if source:
                    value_tokens.append(source)
            gross, gross_token = _two_line_amount(
                block, gross_x, top_y, y_tol=0.011, x_tol=0.021)
            if gross is not None and not (15000 <= gross <= 250000):
                gross, gross_token = None, None
            if gross_token:
                value_tokens.append(gross_token)

            pid = next(((t.get('text') or '').strip().upper() for t in block
                        if looks_like_id(t.get('text', ''))), '')
            out.append({
                '姓名': name, '姓名信心': round(name_conf, 3),
                '職稱': title, '身分證': pid,
                '身分證有效': valid_id(pid),
                **values, '應發金額': gross,
                '_component_inferred': False,
                '_column_evidence': column_evidence,
                '_value_conf': [float(t.get('conf', 0)) for t in value_tokens],
                '_page': pno, '_order': order,
            })

    _infer_two_line_components(out)
    _reconcile_two_line_page_totals(out, page_totals)

    clean = []
    for p in out:
        for key in ('薪俸', '主管加給', '專業加給', '_導師', '_特教'):
            if p.get(key) is None:
                p[key] = 0
        duty = p.pop('_導師') + p.pop('_特教')
        computed = (p['薪俸'] + p['主管加給'] + p['專業加給'] + duty)
        total_inferred = not p.get('應發金額') or p['應發金額'] != computed
        if total_inferred and computed > 0:
            p['應發金額'] = computed
        confidence = [p.pop('姓名信心'), *p.pop('_value_conf')]
        p.pop('_column_evidence', None)
        if p.pop('_component_inferred') or total_inferred:
            confidence.append(0.3)

        p.update({
            '導師特教': duty, '其他加給': 0,
            '應發金額推算': total_inferred,
            '加總相符': (p['應發金額'] > 0 and computed == p['應發金額']
                     and not total_inferred),
            '加總差額': p['應發金額'] - computed if p['應發金額'] else None,
            '最低信心': round(min(confidence or [0]), 3),
        })
        p.pop('_order', None)
        clean.append(p)

    # 本頁小計佐證：本頁各欄合計與印列小計逐欄完全相同，且應發合計也等於
    # 印列應發小計時，本頁以項目合計推算的應發金額即獲紙本證明。
    for pno, gross_total in page_gross.items():
        members = [p for p in clean if p.get('_page') == pno]
        if not members or not any(p['應發金額推算'] for p in members):
            continue
        totals = page_totals.get(pno) or {}
        if not totals:
            continue
        fields_ok = (
            sum(p['薪俸'] for p in members) == totals.get('薪俸', 0)
            and sum(p['專業加給'] for p in members) == totals.get('專業加給', 0)
            and sum(p['主管加給'] for p in members) == totals.get('主管加給', 0)
            and sum(p['導師特教'] for p in members)
            == totals.get('_導師', 0) + totals.get('_特教', 0))
        if fields_ok and sum(p['應發金額'] for p in members) in gross_total:
            for p in members:
                if p['應發金額推算']:
                    p['應發金額推算'] = False
                    p['加總相符'] = True
                    p['應發金額頁小計佐證'] = True
    for p in clean:
        p.pop('_page', None)
    return clean


# ── 橫式「每人兩行、上下欄共用表頭」清冊 ───────────────────
# 例：員工編號／員工姓名（第二行職稱）｜帳號（第二行俸點）｜月支俸額（第二行
# 特教津貼）｜職務加給（補發薪津）｜專業加給（外縣市交通）｜學術研究（輔導津貼）
# ｜導師費｜公保費｜應發合計｜各項扣款…｜應扣合計｜實發總額。每頁底部有
# 「本頁小計」「群組累計」各兩行，再加「單位人數累計」頁尾。
# 以表頭結構辨識，不依校名；欄位 x 由表頭文字定位，缺字時用同份文件其他頁。

GRID_HEADER_LABELS = ('員工編號', '員工姓名', '月支俸額', '職務加給', '專業加給',
                      '學術研究', '導師費', '公保費', '應發合計', '應扣合計',
                      '實發總額', '特教津貼', '補發薪津', '俸點')
# 上方欄名 → (第一行欄位, 第二行欄位)
GRID_COLUMNS = (
    ('月支俸額', '薪俸', '_特教'),
    ('職務加給', '主管加給', '_補發'),
    ('專業加給', '專業加給', '_交通'),
    ('學術研究', '_學術', '_輔導'),
    ('導師費', '_導師', None),
    ('公保費', '_公保', None),
)
GRID_PAY_KEYS = ('薪俸', '_特教', '主管加給', '_補發', '專業加給', '_交通',
                 '_學術', '_輔導', '_導師', '_公保')
# 找不到表頭字時，相對「月支俸額」欄的典型距離（A4 橫式實測）。
GRID_DEFAULT_OFFSET = {
    '月支俸額': 0.0, '職務加給': 0.046, '專業加給': 0.088, '學術研究': 0.133,
    '導師費': 0.181, '公保費': 0.226, '應發合計': 0.271,
    '應扣合計': 0.634, '實發總額': 0.678,
}
GRID_NUM_SHIFT = 0.007   # 金額靠右對齊，數字中心通常在表頭中心右側
GRID_TOTAL_WORDS = ('小計', '累計', '合計', '總計')
GRID_ZERO_LOOKALIKES = {'Q', 'O', 'o', 'D', '口', 'Di', 'DI', '〇', 'U'}


def _page_groups(rows):
    pages = {}
    for r in rows:
        if r:
            pages.setdefault(r[0]['page'], []).append(r)
    return pages


def _grid_num(text, zero_lookalike=False):
    """讀表格金額；容許掃描框線造成的前後標點（例：「：54160」「35780|」）。"""
    s = norm_text(text).strip()
    if zero_lookalike and s in GRID_ZERO_LOOKALIKES:
        return 0
    s = re.sub(r'^[^\d\-]+', '', s)
    s = re.sub(r'[^\d]+$', '', s)
    if not s:
        return None
    if not (re.fullmatch(r'-?\d+', s) or re.fullmatch(r'-?\d{1,3}(?:,\d{3})+', s)):
        return None
    return to_int(s)


def _grid_is_name(s):
    """同 is_name，但接受擴充 A 區漢字（Vision 偶爾把罕用字讀成該區字元）。"""
    s = (s or '').strip()
    if not 2 <= len(s) <= 4:
        return False
    if not all('\u4e00' <= c <= '\u9fff' or '\u3400' <= c <= '\u4dbf' for c in s):
        return False
    return not any(w in s for w in TITLE_WORDS) and not _suspicious_name(s)


def _grid_header_hits(page_rows):
    """回傳本頁表頭區辨識到的欄名與 x 座標（多欄名黏成一個 token 時依字元位置估算）。"""
    hits = {}
    for r in page_rows:
        if not r or r[0].get('y', 1) > 0.25:
            continue
        for t in r:
            text = norm_text(t.get('text', '')).replace(' ', '')
            for label in GRID_HEADER_LABELS:
                if label not in text or (label in hits and not hits[label][1]):
                    continue
                if len(text) > len(label) + 1:
                    w = float(t.get('w', 0))
                    pos = text.index(label) + len(label) / 2
                    hits.setdefault(label, (t['x'] - w / 2 + w * pos / len(text), True))
                else:
                    hits[label] = (t['x'], False)
    return hits


def _is_grid_two_line(rows):
    """每人兩行、表頭含員工姓名／月支俸額／應發合計等的 A4 橫式清冊。"""
    for page_rows in _page_groups(rows).values():
        hits = _grid_header_hits(page_rows)
        primary = sum(label in hits for label in ('月支俸額', '應發合計', '員工姓名',
                                                   '員工編號'))
        secondary = sum(label in hits for label in ('特教津貼', '補發薪津', '俸點',
                                                     '學術研究', '職務加給',
                                                     '應扣合計', '實發總額'))
        if primary >= 3 and secondary >= 2:
            return True
    return False


def _grid_columns(pages):
    """以每頁表頭定位欄位；缺漏欄名用其他頁的中位數或典型距離補齊。"""
    per_page, collected = {}, {}
    for pno, page_rows in pages.items():
        hits = _grid_header_hits(page_rows)
        per_page[pno] = hits
        for label, (x, estimated) in hits.items():
            if not estimated:
                collected.setdefault(label, []).append(x)
    if not any(per_page.values()):
        return {}
    doc = {label: sorted(xs)[len(xs) // 2] for label, xs in collected.items()}

    out = {}
    for pno, hits in per_page.items():
        if not hits:
            continue
        cols = {}
        for label in GRID_HEADER_LABELS:
            if label in hits and not hits[label][1]:
                cols[label] = hits[label][0]
            elif label in doc:
                cols[label] = doc[label]
            elif label in hits:
                cols[label] = hits[label][0]
        base = cols.get('月支俸額')
        if base is None:
            for label in ('職務加給', '學術研究', '應發合計'):
                if label in cols:
                    base = cols[label] - GRID_DEFAULT_OFFSET[label]
                    break
        if base is None:
            continue
        for label, off in GRID_DEFAULT_OFFSET.items():
            cols.setdefault(label, base + off)
        out[pno] = cols
    return out


def _grid_line_numbers(tokens, y, cols, y_tol, zero_lookalike=False):
    """把同一行的數字依最近表頭欄歸位，回傳 {欄名: (數值, token)}。"""
    labels = list(GRID_DEFAULT_OFFSET)
    found = {}
    for t in tokens:
        if abs(t.get('y', 0) - y) > y_tol:
            continue
        value = _grid_num(t.get('text', ''), zero_lookalike)
        if value is None:
            continue
        x = t['x'] - GRID_NUM_SHIFT
        label = min(labels, key=lambda lb: abs(cols[lb] - x))
        dist = abs(cols[label] - x)
        if dist > 0.024:
            continue
        if label not in found or dist < found[label][2]:
            found[label] = (value, t, dist)
    return {label: (v, t) for label, (v, t, _) in found.items()}


def _grid_name(tokens, y, cols):
    """取第一行姓名欄的 2～4 字中文姓名（處理與員工編號或職稱黏在一起的情形）。"""
    name_x = cols.get('員工姓名', cols['月支俸額'] - 0.113)
    lo = cols.get('員工編號', name_x - 0.055) - 0.02
    hi = cols.get('俸點', name_x + 0.05) - 0.02
    best = None
    for t in tokens:
        if abs(t.get('y', 0) - y) > 0.008 or not lo <= t.get('x', 0) <= hi:
            continue
        for run in re.findall(r'[\u3400-\u4dbf\u4e00-\u9fff]{2,10}', t.get('text', '')):
            pieces = [run] if len(run) <= 4 else [run[:3], run[:4]]
            for piece in pieces:
                if not _grid_is_name(piece):
                    continue
                score = float(t.get('conf', 0)) - abs(t['x'] - name_x)
                if best is None or score > best[0]:
                    best = (score, piece, float(t.get('conf', 0)))
    return (best[1], best[2]) if best else ('', 0.0)


def _grid_title(tokens, y, cols):
    name_x = cols.get('員工姓名', cols['月支俸額'] - 0.113)
    hi = cols.get('俸點', name_x + 0.05) - 0.02
    parts = []
    for t in tokens:
        dy = t.get('y', 0) - y
        text = norm_text(t.get('text', '')).strip('.。．,:：')
        if 0.008 <= dy <= 0.024 and name_x - 0.05 <= t.get('x', 0) <= hi \
                and text and _grid_num(text) is None:
            parts.append((t['x'], text))
    return ' '.join(text for _, text in sorted(parts))


def _grid_anchor_ys(body, cols):
    """人員第一行的 y：同列需有姓名、員工編號、帳號或月支俸額中至少兩項證據。"""
    name_x = cols.get('員工姓名', cols['月支俸額'] - 0.113)
    id_x = cols.get('員工編號', name_x - 0.055)
    acct_x = cols.get('俸點', name_x + 0.05)
    cands = []
    for t in body:
        x, text = t.get('x', 0), norm_text(t.get('text', ''))
        if abs(x - id_x) <= 0.025 and re.search(r'\d{4,}', text) \
                and len(re.sub(r'\D', '', text)) <= 8:
            cands.append(t['y'])
        elif abs(x - acct_x) <= 0.04 and len(re.sub(r'\D', '', text)) >= 9:
            cands.append(t['y'])
        elif abs(x - name_x) <= 0.035 and _grid_name([t], t['y'], cols)[0]:
            cands.append(t['y'])
        elif abs(x - GRID_NUM_SHIFT - cols['月支俸額']) <= 0.024:
            v = _grid_num(text)
            if v is not None and v >= 5000:
                cands.append(t['y'])
    cands.sort()
    clusters = []
    for y in cands:
        if clusters and y - clusters[-1][-1] <= 0.008:
            clusters[-1].append(y)
        else:
            clusters.append([y])
    ys = [sum(c) / len(c) for c in clusters if len(c) >= 2]
    starts = []
    for y in ys:
        # 第二行（職稱／俸點／特教津貼）距第一行約 0.015，不可當成新的人。
        if not starts or y - starts[-1] > 0.024:
            starts.append(y)
    return starts


def _grid_page_totals(tokens, cols, label_y):
    first = _grid_line_numbers(tokens, label_y, cols, 0.008, zero_lookalike=True)
    second = _grid_line_numbers(
        [t for t in tokens if 0.008 < t['y'] - label_y <= 0.026],
        label_y + 0.018, cols, 0.01, zero_lookalike=True)
    totals = {}
    for label, key1, key2 in GRID_COLUMNS:
        if label in first:
            totals[key1] = first[label][0]
        if key2 and label in second:
            totals[key2] = second[label][0]
    if '應發合計' in first:
        totals['應發金額'] = first['應發合計'][0]
    return totals


def _grid_parse_pages(rows):
    pages = _page_groups(rows)
    columns = _grid_columns(pages)
    people, page_totals = [], {}
    for pno in sorted(pages):
        cols = columns.get(pno)
        if not cols:
            continue
        tokens = [t for r in pages[pno] for t in r]
        header_y = max((t['y'] for t in tokens if t['y'] < 0.25 and any(
            lb in norm_text(t.get('text', '')) for lb in GRID_HEADER_LABELS)),
            default=0.1)
        total_labels = [t for t in tokens if t['y'] > header_y + 0.005
                        and any(w in t.get('text', '') for w in GRID_TOTAL_WORDS)
                        and t['x'] < cols['月支俸額']]
        body_end = min((t['y'] for t in total_labels), default=0.98) - 0.008
        body = [t for t in tokens if header_y + 0.005 < t['y'] < body_end]

        page_label = next((t for t in sorted(total_labels, key=lambda t: t['y'])
                           if '小計' in t.get('text', '')), None)
        if page_label:
            page_totals[pno] = _grid_page_totals(tokens, cols, page_label['y'])

        for order, y in enumerate(_grid_anchor_ys(body, cols)):
            name, name_conf = _grid_name(body, y, cols)
            first = _grid_line_numbers(body, y, cols, 0.008)
            second = _grid_line_numbers(
                [t for t in body if 0.008 < t['y'] - y <= 0.024],
                y + 0.015, cols, 0.009)
            values, tokens_used = {}, {}
            for label, key1, key2 in GRID_COLUMNS:
                for key, line in ((key1, first), (key2, second)):
                    if not key:
                        continue
                    hit = line.get(label)
                    values[key] = hit[0] if hit else None
                    if hit:
                        tokens_used[key] = hit[1]
            # 第二行月支俸額欄只會是特教津貼等小額，超過合理範圍即視為誤讀。
            if values.get('_特教') and values['_特教'] >= 10000:
                values['_特教'] = None
            gross = first.get('應發合計')
            deduct = first.get('應扣合計')
            net = first.get('實發總額')
            people.append({
                '姓名': name, '_name_conf': name_conf,
                '職稱': _grid_title(body, y, cols),
                '_values': values, '_tokens': tokens_used,
                '_gross': gross[0] if gross else None,
                '_gross_tok': gross[1] if gross else None,
                '_check': (deduct[0] + net[0]) if deduct and net else None,
                '_page': pno, '_order': order,
            })
    return people, page_totals


def _grid_sum(values):
    return sum(values.get(k) or 0 for k in GRID_PAY_KEYS)


def _grid_repair(people):
    """
    以印列的應發合計（或應扣合計＋實發總額）修復少數 OCR 誤讀，全部有算術證據：
    1. 金額尾端黏到右框線多一個 1（357801 → 35780），去掉後加總恰好相符；
    2. 單一欄漏讀，且差額正好是同份文件該欄已讀到過的金額。
    """
    observed = {k: {} for k in GRID_PAY_KEYS}
    for p in people:
        for k in GRID_PAY_KEYS:
            v = p['_values'].get(k)
            if v:
                observed[k][v] = observed[k].get(v, 0) + 1

    for p in people:
        values = p['_values']
        targets = [t for t in (p['_gross'], p['_check']) if t]
        if not targets:
            continue
        target = p['_gross'] or p['_check']
        if p['_gross'] and p['_check'] and p['_gross'] != p['_check']:
            # 應發合計讀錯時，以應扣＋實發的組合為準（兩者需有一個等於項目加總）。
            if _grid_sum(values) == p['_check']:
                target = p['_check']
        if _grid_sum(values) == target:
            # 應發合計格讀錯，但應扣＋實發與項目加總一致時，以後者為準。
            p['_gross'] = target if p['_gross'] else p['_gross']
            continue

        fixed = False
        for k in GRID_PAY_KEYS:
            v = values.get(k)
            if v and str(v).endswith('1') and len(str(v)) >= 5:
                trial = dict(values, **{k: v // 10})
                if _grid_sum(trial) == target:
                    values[k] = v // 10
                    p['_repaired'] = True
                    fixed = True
                    break
        if fixed:
            continue

        diff = target - _grid_sum(values)
        if diff <= 0:
            continue
        options = [k for k in GRID_PAY_KEYS
                   if not values.get(k) and diff in observed[k]]
        if len(options) == 1:
            values[options[0]] = diff
            p['_repaired'] = True
        if p['_gross'] is None and target:
            p['_gross'] = target


def grid_page_checks(tokens_or_rows):
    """
    驗證用：回傳每頁「解析人員加總 vs 印列本頁小計」比對結果。
    不含姓名，只含頁碼、人數與各欄差額。
    """
    rows = tokens_or_rows
    if rows and isinstance(rows[0], dict):
        rows = group_rows(rows)
    people, page_totals = _grid_parse_pages(rows)
    _grid_repair(people)
    out = []
    for pno in sorted({p['_page'] for p in people} | set(page_totals)):
        page_people = [p for p in people if p['_page'] == pno]
        expected = page_totals.get(pno, {})
        diffs = {}
        for key, target in expected.items():
            if key == '應發金額':
                got = sum(p['_gross'] or _grid_sum(p['_values']) for p in page_people)
            else:
                got = sum(p['_values'].get(key) or 0 for p in page_people)
            if got != target:
                diffs[key] = target - got
        out.append({'page': pno, 'people': len(page_people),
                    'has_subtotal': bool(expected), 'diffs': diffs})
    return out


def parse_horizontal_grid(rows):
    """解析「每人兩行、上下欄共用表頭」的橫式清冊。"""
    people, _ = _grid_parse_pages(rows)
    _grid_repair(people)

    clean = []
    for p in people:
        v = p['_values']
        g = lambda k: v.get(k) or 0
        confs = [p['_name_conf'] if p['姓名'] else 0.0]
        confs += [float(t.get('conf', 0)) for t in p['_tokens'].values()]
        computed = _grid_sum(v)
        # 應發合計格漏讀時，印列的「應扣合計＋實發總額」即等於應發合計。
        gross = p['_gross'] or p['_check']
        printed = bool(p['_gross_tok']) or bool(p['_check'])
        if p['_gross_tok']:
            confs.append(float(p['_gross_tok'].get('conf', 0)))
        inferred = not gross or not printed
        if p.get('_repaired') or inferred:
            confs.append(0.3)
        total = gross if gross else computed
        clean.append({
            '姓名': p['姓名'], '職稱': p['職稱'],
            '身分證': '', '身分證有效': False,
            '薪俸': g('薪俸'),
            '主管加給': g('主管加給'),
            '專業加給': g('專業加給') + g('_學術'),
            '導師特教': g('_導師') + g('_特教'),
            # 補發薪津、外縣市交通、輔導津貼、公保費不屬 AF 四欄，只參與驗算。
            '其他加給': g('_補發') + g('_交通') + g('_輔導') + g('_公保'),
            '應發金額': total,
            '應發金額推算': inferred,
            '加總相符': bool(total) and not inferred and total == computed,
            '加總差額': (total - computed) if total else None,
            '最低信心': round(min(confs), 3),
        })
    return clean


def parse_row(row):
    """保留舊介面：單列橫式解析"""
    return parse_horizontal_row(row)


# ── 主流程 ──────────────────────────────────────────────

def _parse_as(rows, layout):
    if layout == 'vertical':
        return parse_vertical(rows)
    if _is_two_line_horizontal(rows):
        return parse_horizontal_two_line(rows)
    if _is_grid_two_line(rows):
        return parse_horizontal_grid(rows)
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


_FLIP_FIELD_LABELS = ('本俸', '薪俸', '月支俸額', '月支數額', '本薪')


def _unflip_pages(tokens):
    """
    舊版 OCR 存下的診斷檔裡，有些頁面座標上下顛倒（直式清冊的「姓名」列
    跑到「本俸」列下方）。新版擷取器已依文字方向校正；此處只處理同一欄
    標籤「姓名在本俸正下方」這種不可能的版面，把該頁 y 座標鏡射回來。
    """
    pages = {}
    for t in tokens:
        pages.setdefault(t['page'], []).append(t)
    flip = set()
    for page, toks in pages.items():
        names = [t for t in toks if (t.get('text') or '').strip() == '姓名']
        fields = [t for t in toks
                  if (t.get('text') or '').strip() in _FLIP_FIELD_LABELS]
        if not names or not fields:
            continue
        upright = any(f['y'] > n['y'] + 0.01 for n in names for f in fields
                      if abs(f['x'] - n['x']) < 0.04)
        flipped = any(n['y'] > f['y'] + 0.02 for n in names for f in fields
                      if abs(f['x'] - n['x']) < 0.04)
        if flipped and not upright:
            flip.add(page)
    if not flip:
        return tokens
    return [dict(t, y=1 - t['y']) if t['page'] in flip else t for t in tokens]


def _ok_count(people):
    return sum(1 for p in people if p.get('加總相符'))


def _per_page_layout(rows, people):
    """
    同一檔案混用直式與橫式頁面時（例：教師頁直式、職員頁橫式），整份只用
    一種版面會漏掉另一種頁面。各頁依自身結構判斷版面後分組解析；只有在
    「加總相符」人數確實增加時才採用，避免影響單一版面的檔案。
    """
    pages = {}
    for r in rows:
        if r:
            pages.setdefault(r[0]['page'], []).append(r)
    groups = {}
    for page in sorted(pages):
        page_rows = pages[page]
        heads, fields, _ = _layout_evidence(page_rows)
        if not heads and not fields and not any(
                '姓名' in row_text(r) for r in page_rows):
            # 沒有表頭的接續頁跟著前一頁的版面
            layout = next(reversed(groups), None) if groups else None
            if layout is None:
                continue
        else:
            layout = detect_layout(page_rows)
        groups.setdefault(layout, []).extend(page_rows)
    if len(groups) < 2:
        return None
    combined = []
    for layout, group_rows_ in groups.items():
        combined.extend(_parse_as(group_rows_, layout))
    if _ok_count(combined) > _ok_count(people):
        return combined
    return None


# 只用於「沒有應發金額」的列：名稱含這些表格用語者不可能是人名，
# 是表頭、簽核欄或代扣欄被誤當成人員欄，直接剔除以免增加人工確認。
_JUNK_NAME_WORDS = ('單位', '驗收', '會計室', '首長', '長官', '支薪', '姓名',
                    '專業', '自提', '代扣', '總數', '所得', '應稅', '總計',
                    '小計', '合計', '製表', '承辦', '頁次', '備註')


def _drop_label_rows(people):
    keep = []
    for p in people:
        name = (p.get('姓名') or '').strip()
        if not p.get('應發金額') and (
                _suspicious_name(name)
                or any(w in name for w in _JUNK_NAME_WORDS)):
            continue
        keep.append(p)
    return keep


def ensemble_people(base, tiled):
    """
    以人為單位合併「整頁」與「分塊」兩次辨識結果，只會讓結果變好：
      - 整頁已算術相符的列一律保留原樣，不會被分塊結果覆蓋；
      - 整頁未相符、分塊同名（或字形相近且唯一）已相符者，改用分塊列；
      - 分塊多找到、且整頁沒有相近姓名或同額同字的人，才補入。
    分塊列本身也必須算術相符（且非推算應發），否則一律不採用。
    """
    from difflib import SequenceMatcher

    def sim(a, b):
        return SequenceMatcher(None, a or '', b or '').ratio()

    def good(p):
        return bool(p.get('加總相符')) and not p.get('應發金額推算')

    out = list(base)
    used = set()
    for t in tiled:
        if not good(t) or _suspicious_name(t.get('姓名')):
            continue
        name = (t.get('姓名') or '').strip()
        exact = [i for i, b in enumerate(out) if (b.get('姓名') or '').strip() == name]
        close = exact or [i for i, b in enumerate(out)
                          if sim(name, b.get('姓名')) >= 0.6]
        if close:
            if len(close) == 1 and close[0] not in used and not good(out[close[0]]):
                out[close[0]] = t
                used.add(close[0])
            continue
        if any(b.get('薪俸') == t.get('薪俸') and b.get('應發金額') == t.get('應發金額')
               and set(b.get('姓名') or '') & set(name) for b in out):
            continue
        out.append(t)
        used.add(len(out) - 1)
    return out


def parse_tokens(tokens, preferred_layout=None):
    tokens = _unflip_pages(tokens)
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
    people = candidates[layout]
    mixed = _per_page_layout(rows, people)
    if mixed is not None:
        people = mixed
    return _drop_label_rows(people), layout


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
