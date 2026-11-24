"""v1.3.23：AF 學校 ↔ 固定清冊學校的對應與防呆。

1. match_account_school：AF 的機關代碼／校名 → 人事資料倉庫（account.leotsai.me）
   登入帳號學校清單中的 schoolId。先比 orgCode，再比校名（全名或去掉
   「高雄市／市立／某某區」後的核心名稱），必須唯一才採用。
2. roster_overlap：固定清冊與 AF 的人員重疊程度（只回傳人數與比率，不回傳
   姓名／身分證），用來偵測「AF 是 A 校、固定清冊是 B 校」。
"""
import re

import school_identity as SID

CODE_FULL_RE = SID.CODE_FULL_RE
# AF 人員至少這個比例出現在固定清冊，才視為同一所學校的清冊。
# 每月人事異動通常只有少數人，同校清冊重疊率遠高於此；不同校幾乎為 0。
OVERLAP_MIN_RATIO = 0.5
OVERLAP_MIN_PEOPLE = 3


def _norm_code(v):
    v = str(v or '').strip().upper()
    return v if CODE_FULL_RE.match(v) else ''


def match_account_school(code, full_name, schools):
    """schools：[{schoolId, name, orgCode}]。回傳唯一相符的 schoolId 或 ''。"""
    code = _norm_code(code)
    rows = [s for s in (schools or []) if isinstance(s, dict) and s.get('schoolId')]
    if code:
        hit = {str(s['schoolId']) for s in rows if _norm_code(s.get('orgCode')) == code}
        if len(hit) == 1:
            return hit.pop()
        if len(hit) > 1:
            return ''
    full = SID._norm_name(full_name)
    if not full:
        return ''
    hit = {str(s['schoolId']) for s in rows if SID._norm_name(s.get('name')) == full}
    if len(hit) == 1:
        return hit.pop()
    if hit:
        return ''
    core = SID._core_name(full_name)
    if len(core) < 4:
        return ''
    hit = {str(s['schoolId']) for s in rows if SID._core_name(s.get('name')) == core}
    return hit.pop() if len(hit) == 1 else ''


def school_code_from_name(name, schools_table):
    """倉庫學校名稱 → (機關代碼, 全名)；對不到唯一一校時 ('', '')。"""
    return SID.find_school_in_text([name or ''], schools_table or [])


def _clean(v):
    s = '' if v is None else str(v)
    if s.lower() == 'nan':
        return ''
    return re.sub(r'\s+', '', s)


def people_keys(df, id_cols):
    """DataFrame → (姓名集合, 身分證集合)。"""
    names, ids = set(), set()
    if df is None:
        return names, ids
    cols = [str(c).strip() for c in df.columns]
    df = df.copy()
    df.columns = cols
    id_col = next((c for c in id_cols if c in cols), None)
    for _, r in df.iterrows():
        n = _clean(r.get('姓名')) if '姓名' in cols else ''
        i = _clean(r.get(id_col)).upper() if id_col else ''
        if n:
            names.add(n)
        if i:
            ids.add(i)
    return names, ids


def roster_overlap(roster_df, af_df, id_cols):
    """回傳 {'af_total', 'roster_total', 'matched', 'ratio', 'same_school'}。

    以 AF 每一位人員（依身分證，沒有時依姓名）是否出現在固定清冊（身分證或
    姓名任一相符）計算。same_school 為 None 代表人數太少無法判斷。"""
    r_names, r_ids = people_keys(roster_df, id_cols)
    cols = [str(c).strip() for c in af_df.columns]
    af = af_df.copy()
    af.columns = cols
    id_col = next((c for c in id_cols if c in cols), None)
    seen, matched = set(), 0
    for _, r in af.iterrows():
        n = _clean(r.get('姓名')) if '姓名' in cols else ''
        i = _clean(r.get(id_col)).upper() if id_col else ''
        key = i or n
        if not key or key in seen:
            continue
        seen.add(key)
        if (i and i in r_ids) or (n and n in r_names):
            matched += 1
    total = len(seen)
    ratio = round(matched / total, 3) if total else 0.0
    roster_total = len(r_names | r_ids) and max(len(r_names), len(r_ids))
    if total < OVERLAP_MIN_PEOPLE or not roster_total:
        same = None
    else:
        same = ratio >= OVERLAP_MIN_RATIO
    return {'af_total': total, 'roster_total': roster_total,
            'matched': matched, 'ratio': ratio, 'same_school': same}
