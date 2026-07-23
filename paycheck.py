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

    # 姓名完全找不到 → 模糊比對 + 金額指紋救援
    unmatched = [p for p, c in pending if not c]
    for p in unmatched:
        pool = [a for a in af_records if a['身分證'] not in used]
        m = get_close_matches(p['姓名'], [a['姓名'] for a in pool], n=1, cutoff=0.5)
        if m:
            c = [a for a in pool if a['姓名'] == m[0]]
            if len(c) == 1:
                take(p, c[0], f'姓名模糊（AF 為「{m[0]}」）')
                continue
        fp = tuple(p.get(f, 0) for f in FIELDS)
        c = [a for a in pool if tuple(a.get(f, 0) for f in FIELDS) == fp]
        if len(c) == 1:
            take(p, c[0], f'金額指紋（AF 為「{c[0]["姓名"]}」）')
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
        })
    for p in pdf_only:
        results.append({'姓名': p['姓名'], '職稱': p.get('職稱', ''), '單位': '',
                        '身分證': '', '配對依據': '', '狀態': 'pdf_only', '差異': []})
    for a in af_only:
        results.append({'姓名': a['姓名'], '職稱': '', '單位': a.get('單位', ''),
                        '身分證': a['身分證'], '配對依據': '', '狀態': 'af_only', '差異': []})

    order = {'diff': 0, 'pdf_only': 1, 'af_only': 2, 'ok': 3}
    results.sort(key=lambda r: (order[r['狀態']], r['姓名']))

    summary = {
        '一致': sum(1 for r in results if r['狀態'] == 'ok'),
        '有差異': sum(1 for r in results if r['狀態'] == 'diff'),
        '清冊有AF無': sum(1 for r in results if r['狀態'] == 'pdf_only'),
        'AF有清冊無': sum(1 for r in results if r['狀態'] == 'af_only'),
        '待人工指定': len(ambiguous),
    }
    return {'summary': summary, 'results': results, 'ambiguous': ambiguous}
