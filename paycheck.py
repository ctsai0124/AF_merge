"""薪資清冊 PDF × AF 校對清冊 比對模組"""
import io, os, re, tempfile
from difflib import get_close_matches

import pandas as pd

# ── AF 職務加給表別代碼 ────────────────────────────────
MGR_CODES = {'C1009', 'C1001'}       # 主管加給（教育人員 / 職員）
TEACH_CODES = {'C1014'}              # 導師加給＋特教加給（AF 合併為一筆）

# 比對的四個項目
FIELDS = ['薪俸', '專業加給', '主管加給', '導師特教']
FIELD_LABEL = {'薪俸': '薪俸／本俸', '專業加給': '專業加給／學術研究費',
               '主管加給': '職務加給', '導師特教': '導師加給＋特教加給'}


def _n(v):
    v = str(v or '').replace(',', '').strip()
    try:
        return int(float(v)) if v else 0
    except Exception:
        return 0


# ── 1. AF 讀取 ────────────────────────────────────────

def load_af(file_bytes, filename):
    """
    讀 AF 校對清冊，依身分證字號合併多列，回傳 (records, warnings)
    每筆：身分證、姓名、單位、薪俸、專業加給、主管加給、導師特教
    """
    ext = os.path.splitext(filename)[1].lower()
    engine = 'xlrd' if ext == '.xls' else 'openpyxl'
    df = pd.read_excel(io.BytesIO(file_bytes), sheet_name=0,
                       engine=engine, dtype=str, header=0).fillna('')
    df.columns = [str(c).strip() for c in df.columns]

    need = ['身分證字號', '姓名', '薪俸表別', '支領數額',
            '專業加給表別', '支領數額.1', '職務加給表別', '支領數額.2']
    missing = [c for c in need if c not in df.columns]
    if missing:
        raise KeyError('AF 缺少欄位：' + '、'.join(missing))

    recs, warns, unknown = {}, [], {}
    for _, r in df.iterrows():
        pid = str(r['身分證字號']).strip()
        if not pid:
            continue
        d = recs.setdefault(pid, {
            '身分證': pid, '姓名': str(r['姓名']).strip(), '單位': '',
            '薪俸': 0, '專業加給': 0, '主管加給': 0, '導師特教': 0, '列數': 0,
        })
        d['列數'] += 1
        if str(r.get('單位', '')).strip():
            d['單位'] = str(r['單位']).strip()
        if str(r['薪俸表別']).strip():
            d['薪俸'] = _n(r['支領數額'])
        if str(r['專業加給表別']).strip():
            d['專業加給'] = _n(r['支領數額.1'])

        code = str(r['職務加給表別']).strip().upper()
        if code in MGR_CODES:
            d['主管加給'] += _n(r['支領數額.2'])
        elif code in TEACH_CODES:
            d['導師特教'] += _n(r['支領數額.2'])
        elif code:
            unknown.setdefault(code, []).append(d['姓名'])

    if unknown:
        for c, names in unknown.items():
            warns.append(f'AF 出現未知職務加給代碼 {c}（{"、".join(sorted(set(names))[:5])}），未納入比對')

    # 同名不同身分證 → 提醒
    by_name = {}
    for d in recs.values():
        by_name.setdefault(d['姓名'], []).append(d)
    for name, group in by_name.items():
        if len(group) > 1:
            warns.append(f'AF 有 {len(group)} 位同名「{name}」，將以職稱／單位區辨')

    return list(recs.values()), warns


# ── 2. PDF 解析（自動判斷版面）────────────────────────

FIELD_ALIAS = {
    '本俸': '薪俸', '薪俸': '薪俸',
    '專業加給': '專業加給', '學術研究': '專業加給', '學術研究費': '專業加給',
    '職務加給': '主管加給', '主管加給': '主管加給',
    '導師加給': '_導師', '導師費': '_導師',
    '特教加給': '_特教',
    '地域加給': '地域加給',
    '應發數合計': '應發金額', '應發金額': '應發金額',
    '實發數': '實發金額', '實發金額': '實發金額',
}
SKIP_NAME = ('小計', '合計', '總計', '備註', '出納', '人事', '會計', '製表')


def has_text_layer(file_bytes):
    """判斷 PDF 有無文字層（純 Python，不需系統套件）"""
    import pdfplumber
    tmp = tempfile.NamedTemporaryFile(suffix='.pdf', delete=False)
    tmp.write(file_bytes); tmp.close()
    try:
        with pdfplumber.open(tmp.name) as pdf:
            for page in pdf.pages[:3]:
                if (page.extract_text() or '').strip():
                    return True
        return False
    except Exception:
        return False
    finally:
        os.unlink(tmp.name)


def parse_pdf(file_bytes):
    """自動判斷直式／橫式版面並解析，回傳人員清單"""
    import pdfplumber
    tmp = tempfile.NamedTemporaryFile(suffix='.pdf', delete=False)
    tmp.write(file_bytes); tmp.close()
    try:
        with pdfplumber.open(tmp.name) as pdf:
            tables = []
            for page in pdf.pages:
                tables.extend([t for t in page.extract_tables() if t and t[0]])
            if not tables:
                return [], 'unknown'
            head0 = (tables[0][0][0] or '')
            if '序號' in head0 and '姓名' in head0:
                return _parse_vertical(tables), 'vertical'
            return _parse_horizontal(tables), 'horizontal'
    finally:
        os.unlink(tmp.name)


def _finalize(p):
    p['導師特教'] = p.pop('_導師', 0) + p.pop('_特教', 0)
    for f in FIELDS:
        p.setdefault(f, 0)
    return p


def _parse_vertical(tables):
    """直式：一人一欄，表頭為 序號\\n姓名\\n職稱\\n支薪俸級"""
    people = {}
    for tbl in tables:
        head = [c or '' for c in tbl[0]]
        if '序號' not in head[0]:
            continue
        cols = {}
        for ci, cell in enumerate(head[1:], 1):
            parts = [x.strip() for x in (cell or '').split('\n') if x.strip()]
            if not parts or not re.fullmatch(r'\d+', parts[0]):
                continue
            cols[ci] = {'序號': int(parts[0]),
                        '姓名': parts[1] if len(parts) > 1 else '',
                        '職稱': parts[2] if len(parts) > 2 else ''}
        for ci, info in cols.items():
            people.setdefault(info['序號'], dict(info))
        for row in tbl[1:]:
            if not row or not row[0]:
                continue
            key = FIELD_ALIAS.get(row[0].strip())
            if not key:
                continue
            for ci, info in cols.items():
                if ci < len(row):
                    people[info['序號']][key] = _n(row[ci])
    return [_finalize(people[k]) for k in sorted(people)]


def _parse_horizontal(tables):
    """橫式：一人一列"""
    people = []
    for tbl in tables:
        hidx, hmap = None, {}
        for i, row in enumerate(tbl):
            cells = [(c or '').strip() for c in (row or [])]
            if any('姓名' in c for c in cells):
                hidx = i
                for j, c in enumerate(cells):
                    if '姓名' in c:
                        hmap['姓名'] = j
                    elif '職稱' in c:
                        hmap['職稱'] = j
                    elif c in FIELD_ALIAS:
                        hmap[FIELD_ALIAS[c]] = j
                break
        if hidx is None or '姓名' not in hmap:
            continue
        for row in tbl[hidx + 1:]:
            cells = [(c or '').strip() for c in (row or [])]
            if hmap['姓名'] >= len(cells):
                continue
            name = cells[hmap['姓名']]
            if len(name) < 2 or any(k in name for k in SKIP_NAME):
                continue
            if any(p['姓名'] == name for p in people):
                continue
            p = {'姓名': name,
                 '職稱': cells[hmap['職稱']] if '職稱' in hmap and hmap['職稱'] < len(cells) else ''}
            for key, ci in hmap.items():
                if key in ('姓名', '職稱'):
                    continue
                p[key] = _n(cells[ci]) if ci < len(cells) else 0
            people.append(_finalize(p))
    return people


# ── 3. 職稱 → 單位推導 ────────────────────────────────

TITLE_DEPT = {
    '校長': '校長室',
    '教務主任': '教務處', '教學組長': '教務處', '註冊組長': '教務處',
    '設備組長': '教務處', '資訊組長': '教務處', '教務組長': '教務處',
    '學務主任': '學生事務處', '訓導主任': '學生事務處', '訓育組長': '學生事務處',
    '生教組長': '學生事務處', '生活教育組長': '學生事務處', '體育組長': '學生事務處',
    '衛生組長': '學生事務處', '訓導組長': '學生事務處',
    '總務主任': '總務處', '事務組長': '總務處', '出納組長': '總務處', '文書組長': '總務處',
    '輔導主任': '輔導室', '輔導組長': '輔導室', '資料組長': '輔導室', '特教組長': '輔導室',
    '教導主任': '教導處', '人事主任': '人事室', '人事管理員': '人事室',
    '會計主任': '會計室', '主計': '會計室',
    '校護': '健康中心', '護理師': '健康中心',
    '工友': '總務處', '技工': '總務處', '駕駛': '總務處',
}
DEPT_ALIAS = {'訓導處': '學生事務處', '學務處': '學生事務處',
              '輔導處': '輔導室', '學生輔導室': '輔導室'}


def norm_dept(d):
    d = DEPT_ALIAS.get((d or '').strip(), (d or '').strip())
    return d.replace('處', '').replace('室', '').replace('中心', '')


def title_to_dept(title):
    """由職稱推導單位，推不出來回 None。支援「教師兼教務主任」寫法"""
    t = (title or '').strip()
    if not t:
        return None
    if t in TITLE_DEPT:
        return TITLE_DEPT[t]
    for k, v in sorted(TITLE_DEPT.items(), key=lambda x: -len(x[0])):
        if k in t:
            return v
    return None


def _similar_enough(a, b):
    """
    僅接受「字數相同且只差一個字」的情況，視為文字辨識誤差。
    例：吳盂謙 / 吳孟謙 → 可接受
        楊棨翔 / 林怡慈 → 不接受
    """
    if len(a) != len(b) or not a:
        return False
    return sum(1 for x, y in zip(a, b) if x != y) == 1


# ── 4. 配對 ───────────────────────────────────────────

def match(pdf_people, af_records):
    """
    回傳 (pairs, pdf_only, af_only, ambiguous)
    pairs: [(pdf, af, 依據)]
    ambiguous: [(候選 pdf 列, 候選 af 筆)] 需人工指定
    """
    af_by_name = {}
    for a in af_records:
        af_by_name.setdefault(a['姓名'], []).append(a)

    pairs, pdf_only, ambiguous = [], [], []
    used = set()

    def take(p, a, why):
        used.add(a['身分證'])
        pairs.append((p, a, why))

    # 先處理姓名唯一者
    pending = []
    for p in pdf_people:
        cands = [a for a in af_by_name.get(p['姓名'], []) if a['身分證'] not in used]
        if len(cands) == 1:
            take(p, cands[0], '姓名')
        elif len(cands) > 1:
            pending.append((p, cands))
        else:
            pending.append((p, None))

    # 同名多筆 → 逐層區辨
    groups = {}
    for p, c in pending:
        if c:
            groups.setdefault(p['姓名'], {'pdf': [], 'af': c})['pdf'].append(p)

    for name, g in groups.items():
        left_p = list(g['pdf'])
        left_a = [a for a in g['af'] if a['身分證'] not in used]

        # 第1層：單位（由職稱推導）
        for p in list(left_p):
            dept = title_to_dept(p.get('職稱'))
            if not dept:
                continue
            c = [a for a in left_a if a.get('單位') and norm_dept(a['單位']) == norm_dept(dept)]
            if len(c) == 1:
                take(p, c[0], f'單位（{dept}）')
                left_p.remove(p); left_a.remove(c[0])

        # 第2層：主管職推斷
        for p in list(left_p):
            want = p.get('主管加給', 0) > 0
            c = [a for a in left_a if (a.get('主管加給', 0) > 0) == want]
            if len(c) == 1:
                take(p, c[0], '主管職')
                left_p.remove(p); left_a.remove(c[0])

        # 第3層：金額指紋（僅唯一吻合）
        for p in list(left_p):
            fp = tuple(p.get(f, 0) for f in FIELDS)
            c = [a for a in left_a if tuple(a.get(f, 0) for f in FIELDS) == fp]
            if len(c) == 1:
                take(p, c[0], '金額指紋')
                left_p.remove(p); left_a.remove(c[0])

        # 第4層：只剩一對一 → 順推
        if len(left_p) == 1 and len(left_a) == 1:
            take(left_p[0], left_a[0], '順推')
            left_p, left_a = [], []

        if left_p or left_a:
            ambiguous.append({'姓名': name, 'pdf': left_p, 'af': left_a})

    # 姓名完全找不到 → 僅做「字形相近」的模糊比對
    #
    # ⚠ 此處刻意不使用金額比對來配對不同姓名的人。
    #    人員異動時，接任者常沿用相同職務與薪級，四項金額會完全相同
    #    （例：前任訓導組長離職、新任者薪級相同）。若以金額反推，
    #    會把兩個不同的人誤判為同一人，使真正的人員異動被掩蓋。
    #    姓名不同就是不同人，寧可列為「查無此人」交由使用者判斷。
    unmatched = [p for p, c in pending if not c]
    for p in unmatched:
        pool = [a for a in af_records if a['身分證'] not in used]
        m = get_close_matches(p['姓名'], [a['姓名'] for a in pool], n=1, cutoff=0.75)
        if m and _similar_enough(p['姓名'], m[0]):
            c = [a for a in pool if a['姓名'] == m[0]]
            if len(c) == 1:
                take(p, c[0], f'姓名相近（AF 為「{m[0]}」，請確認是否同一人）')
                continue
        pdf_only.append(p)

    af_only = [a for a in af_records if a['身分證'] not in used]
    return pairs, pdf_only, af_only, ambiguous


# ── 5. 比對 ───────────────────────────────────────────

def compare(pdf_people, af_records):
    pairs, pdf_only, af_only, ambiguous = match(pdf_people, af_records)
    results = []
    for p, a, why in pairs:
        diffs = []
        for f in FIELDS:
            pv, av = p.get(f, 0), a.get(f, 0)
            if pv != av:
                diffs.append({'欄位': FIELD_LABEL[f], '清冊': pv, 'AF': av, '差額': av - pv})
        results.append({
            '姓名': a['姓名'], '職稱': p.get('職稱', ''), '單位': a.get('單位', ''),
            '身分證': a['身分證'], '配對依據': why,
            '狀態': 'diff' if diffs else 'ok', '差異': diffs,
            '需確認': why not in ('姓名',),
            '次要': bool(p.get('_次要')), '次要原因': p.get('_次要原因', ''),
        })
    for p in pdf_only:
        results.append({'姓名': p['姓名'], '職稱': p.get('職稱', ''), '單位': '',
                        '身分證': '', '配對依據': '', '狀態': 'pdf_only', '差異': [],
                        '次要': bool(p.get('_次要')), '次要原因': p.get('_次要原因', '')})
    for a in af_only:
        results.append({'姓名': a['姓名'], '職稱': '', '單位': a.get('單位', ''),
                        '身分證': a['身分證'], '配對依據': '', '狀態': 'af_only',
                        '差異': [], '次要': False, '次要原因': ''})

    order = {'diff': 0, 'pdf_only': 1, 'af_only': 2, 'ok': 3}
    results.sort(key=lambda r: (order[r['狀態']], r['姓名']))

    main = [r for r in results if not r.get('次要')]
    sub = [r for r in results if r.get('次要')]
    summary = {
        '一致': sum(1 for r in main if r['狀態'] == 'ok'),
        '有差異': sum(1 for r in main if r['狀態'] == 'diff'),
        '清冊有AF無': sum(1 for r in main if r['狀態'] == 'pdf_only'),
        'AF有清冊無': sum(1 for r in main if r['狀態'] == 'af_only'),
        '待人工指定': len(ambiguous),
        '其他人員': len(sub),
        '其他人員有差異': sum(1 for r in sub if r['狀態'] in ('diff', 'pdf_only')),
    }
    return {'summary': summary, 'results': results, 'ambiguous': ambiguous}


# ── 5b. 排除規則 ─────────────────────────────────────

# 預設排除的職稱關鍵字（非正式人員，通常不在 AF 校對清冊內）
DEFAULT_EXCLUDE_TITLES = [
    '代理', '代課', '長代',            # 代理教師、長期代理、侍親長代、育嬰長代
    '教保員', '助理教保',              # 幼兒園教保人員
    '約僱', '約聘', '臨時', '工讀',     # 臨時性人力
    '替代役', '實習', '志工',
]


def title_head(title):
    """
    取職稱中「兼」之前的部分作為主要職務。
    「教師兼代理主任」→「教師」（是正式教師，不可因含「代理」被排除）
    「代理教師」→「代理教師」
    """
    t = (title or '').strip()
    return t.split('兼')[0] if '兼' in t else t


def suggest_exclude_titles(people):
    """從人員清單中挑出符合預設規則的職稱，供介面建議勾選"""
    found = set()
    for p in people:
        head = title_head(p.get('職稱', ''))
        for kw in DEFAULT_EXCLUDE_TITLES:
            if kw in head:
                t = (p.get('職稱') or '').strip()
                if t:
                    found.add(t)
    return sorted(found)


def mark_exclusions(people, ex_titles=None, ex_names=None, use_default=True):
    """
    標記不需納入稽核的人員，但**不從資料中移除**——
    全部人員仍會完成比對，只是在介面上預設收合。

    比對只看職稱中「兼」之前的部分，避免「教師兼代理主任」這類
    正式人員被誤判為代理人員。
    """
    ex_titles = [t.strip() for t in (ex_titles or []) if t.strip()]
    ex_names = set(n.strip() for n in (ex_names or []) if n.strip())
    defaults = DEFAULT_EXCLUDE_TITLES if use_default else []

    for p in people:
        title = (p.get('職稱') or '').strip()
        head = title_head(title)
        name = (p.get('姓名') or p.get('建議姓名')
                or p.get('原始姓名') or '').strip()

        hit = next((t for t in ex_titles if t in head), None) \
            or next((t for t in defaults if t in head), None)
        if hit:
            p['_次要'] = True
            p['_次要原因'] = f'職稱含「{hit}」'
        elif name in ex_names:
            p['_次要'] = True
            p['_次要原因'] = '個別指定'
        else:
            p['_次要'] = False
    return people


def title_observations(people, af_names):
    """
    統計本次清冊中，各職稱的人有多少比例出現在 AF。
    回傳 {職稱: {'total': n, 'in_af': m}}
    供跨檔案累積，用來判斷哪些職稱屬於非正式人員。
    """
    obs = {}
    for p in people:
        t = (p.get('職稱') or '').strip()
        if not t:
            continue
        d = obs.setdefault(t, {'total': 0, 'in_af': 0})
        d['total'] += 1
        nm = (p.get('姓名') or p.get('建議姓名') or p.get('原始姓名') or '').strip()
        if nm in af_names:
            d['in_af'] += 1
    return obs


def learned_suggestions(stats, min_people=3, min_files=2):
    """
    從累積統計中挑出「疑似非正式人員」的職稱。
    條件：出現過一定人次、來自多份檔案、且從未在 AF 出現過。
    回傳 (建議排除清單, 疑似誤排除清單)
    """
    suggest, wrong = [], []
    for t, d in (stats or {}).items():
        if t in DEFAULT_EXCLUDE_TITLES or any(k in title_head(t) for k in DEFAULT_EXCLUDE_TITLES):
            # 已由預設規則涵蓋；若竟然出現在 AF，代表預設規則可能誤判
            if d.get('in_af', 0) > 0:
                wrong.append({'職稱': t, '出現在AF': d['in_af'], '總人次': d['total']})
            continue
        if (d.get('in_af', 0) == 0 and d.get('total', 0) >= min_people
                and d.get('files', 0) >= min_files):
            suggest.append({'職稱': t, '總人次': d['total'], '檔案數': d['files']})
    suggest.sort(key=lambda x: -x['總人次'])
    return suggest, wrong


# ── 6. 算術自檢 ───────────────────────────────────────

def arith_check(p):
    """
    檢查 薪俸＋主管加給＋導師特教＋專業加給＋其他 是否等於應發金額。
    回傳 (是否通過, 差額)；沒有應發金額資料時回傳 (None, 0) 代表無法檢查
    """
    total = p.get('應發金額', 0)
    if not total:
        return None, 0
    s = (p.get('薪俸', 0) + p.get('專業加給', 0)
         + p.get('主管加給', 0) + p.get('導師特教', 0)
         + p.get('地域加給', 0) + p.get('其他', 0))
    return s == total, total - s


def annotate_arith(people):
    """為每筆資料加上 _arith_ok / _arith_diff 標記"""
    for p in people:
        ok, diff = arith_check(p)
        p['_arith_ok'] = ok
        p['_arith_diff'] = diff
    return people


# ── 9. 由 Mac 端 OCR 結果轉為可比對資料 ─────────────────

def from_ocr(ocr_people, af_records):
    """
    ocr_people：Mac 端 parse_tokens.py 的輸出
    回傳 (可直接比對的人員, 需使用者確認的列)

    規則：
    - 加總相符且姓名可對應 → 直接比對
    - 其餘 → 保留 OCR 讀到的數值交由使用者核對，並標示可疑欄位
    - 絕不以 AF 的金額回填，否則會掩蓋真正的差異
    """
    known = {a['姓名'] for a in af_records}
    good, need = [], []

    for p in ocr_people or []:
        raw = (p.get('姓名') or '').strip()
        fixed, changed = fix_name(raw, known)
        vals = {
            '薪俸': _n(p.get('薪俸', 0)),
            '專業加給': _n(p.get('專業加給', 0)),
            '主管加給': _n(p.get('主管加給', 0)),
            '導師特教': _n(p.get('導師特教', 0)),
        }
        total = _n(p.get('應發金額', 0))
        s = sum(vals.values())
        name_ok = fixed in known
        sum_ok = bool(total) and s == total
        conf = p.get('最低信心')

        if name_ok and sum_ok:
            good.append({'姓名': fixed, '職稱': re.sub(r'^\d+\s*', '', p.get('職稱', '')).strip(),
                         '應發金額': total, '_ocr_name_fixed': changed, **vals})
            continue

        reasons = []
        if not name_ok:
            reasons.append(f'姓名「{raw or "空白"}」無法對應 AF 名單')
        if not total:
            reasons.append('未讀到應發金額，無法驗算')
        elif not sum_ok:
            reasons.append(f'四項相加 {s:,} 與應發金額 {total:,} 不符（差 {total - s:+,}）')
        if conf is not None and conf < 0.5:
            reasons.append(f'辨識信心偏低（{conf}）')

        need.append({
            '原始姓名': raw, '建議姓名': fixed if name_ok else '',
            '職稱': re.sub(r'^\d+\s*', '', p.get('職稱', '')).strip(),
            '應發金額': total, '原因': reasons,
            '姓名可疑': not name_ok, '金額可疑': not sum_ok,
            **vals,
        })

    return good, need


def compare_with_fixed(good, fixed_rows, af_records):
    """合併「自動通過」與「使用者更正後」的資料再比對"""
    merged = list(good)
    for r in fixed_rows or []:
        name = (r.get('姓名') or '').strip()
        if not name:
            continue
        merged.append({
            '姓名': name, '職稱': r.get('職稱', ''),
            '薪俸': _n(r.get('薪俸', 0)), '專業加給': _n(r.get('專業加給', 0)),
            '主管加給': _n(r.get('主管加給', 0)), '導師特教': _n(r.get('導師特教', 0)),
            '應發金額': _n(r.get('應發金額', 0)), '_user_fixed': True,
        })
    return compare(merged, af_records)


def fix_name(raw, known_names):
    """用 AF 名單校正 OCR 姓名，回傳 (校正後, 是否有改動)"""
    raw = (raw or '').strip()
    if not raw or raw in known_names:
        return raw, False
    m = get_close_matches(raw, list(known_names), n=1, cutoff=0.5)
    return (m[0], True) if m else (raw, False)
