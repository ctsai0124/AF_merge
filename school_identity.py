"""學校身分判定：決定各校記憶（排除設定、版面樣式）要記在哪個「學校鍵」底下。

AF 校對清冊（薪資系統匯出）內容本身沒有機關代碼或學校名稱欄位，只有
「單位」（處室名稱）。為了不被改過的檔名（例如「複本 校對清冊」）搞混，
依可信度由高到低判定：

1. af_code       AF 內容中的機關代碼欄（目前格式沒有，日後若加上會優先採用）
2. filename_code 檔名中的機關代碼（9 位數字＋1 個英文字母，例：薪資系統原檔名）
3. af_name       AF 內容中的學校名稱欄（同上，目前格式沒有）
4. pdf_name      薪資清冊 PDF 標題的學校名稱（文字層或 OCR 標題）
   —— 3、4 皆以 static/school.xlsx 對回機關代碼；對不到唯一一校時才以名稱為鍵
5. fingerprint   AF 人員組成指紋：身分證字號經本機密鑰 HMAC 後與已知學校比對
6. filename      清掉「複本／副本／(1)／- Copy／年月」等雜訊後的檔名

舊版本以檔名切出的鍵（legacy_key）只用於讀取時的向後相容。
指紋只存 HMAC 雜湊（截短 16 碼），不存身分證字號本身。
"""
import hashlib
import hmac
import io
import json
import os
import re
import threading
import time
import unicodedata

CODE_RE = re.compile(r'(?<![0-9A-Za-z])(\d{9}[A-Za-z])(?![0-9A-Za-z])')
CODE_FULL_RE = re.compile(r'^\d{9}[A-Za-z]$')
CODE_COLUMNS = ('機關代碼', '機關代號', '機關編號', '單位代碼', '單位代號',
                '學校代碼', '學校代號', '服務機關代碼')
NAME_COLUMNS = ('機關名稱', '學校名稱', '服務機關', '服務機關名稱', '學校', '機關')
STRONG_SOURCES = ('af_code', 'filename_code', 'af_name', 'pdf_name')
SOURCE_LABELS = {
    'af_code': 'AF 內容機關代碼', 'filename_code': '檔名機關代碼',
    'af_name': 'AF 內容學校名稱', 'pdf_name': '清冊標題學校名稱',
    'fingerprint': 'AF 人員組成比對', 'filename': '檔名',
}

# 檔名雜訊：複本／副本／拷貝／Copy of／- Copy／(1)／（2）
_NOISE_RES = [
    re.compile(r'^(?:copy\s+of\s+)', re.I),
    re.compile(r'[\s_\-－]*(?:的)?(?:複本|副本|拷貝|备份|備份)(?:\s*\d+)?[\s_\-－]*'),
    re.compile(r'[\s_\-－]*copy(?:\s*\d+)?\b', re.I),
    re.compile(r'\s*[\(（]\s*\d{1,3}\s*[\)）]'),
]
_ROC_YM_RE = re.compile(r'(?<!\d)1[0-4]\d(?:0[1-9]|1[0-2])(?!\d)')
_LONG_DIGITS_RE = re.compile(r'(?<![0-9A-Za-z])\d{6,}(?![0-9A-Za-z])')

_SHORT_FORMS = (('國小', '國民小學'), ('國中', '國民中學'))
_PREFIX_RE = re.compile(r'^(?:高雄市立|高雄市|市立)?(?:[\u4e00-\u9fff]{1,3}區)?')


# ── 檔名 ─────────────────────────────────────────────
def legacy_key(af_name):
    """v1.3.18 以前的學校鍵：檔名以 _ 切開後倒數第二段，沒有時用整個檔名。"""
    name = af_name or ''
    parts = os.path.splitext(name)[0].split('_')
    sn1 = parts[-2] if len(parts) >= 2 else ''
    return sn1 or (name or '_default')


def filename_code(af_name):
    """檔名中的機關代碼；優先取舊規則那一段，其次檔名中任一處。"""
    seg = legacy_key(af_name)
    if CODE_FULL_RE.match(seg.strip()):
        return seg.strip().upper()
    m = CODE_RE.findall(os.path.splitext(af_name or '')[0])
    codes = {c.upper() for c in m}
    return codes.pop() if len(codes) == 1 else ''


def clean_filename(af_name):
    """去掉副檔名與「複本、副本、(1)、- Copy、年月、時間戳」等雜訊。"""
    base = unicodedata.normalize('NFKC', os.path.splitext(af_name or '')[0])
    for rx in _NOISE_RES:
        base = rx.sub(' ', base)
    base = _LONG_DIGITS_RE.sub(' ', base)
    base = _ROC_YM_RE.sub(' ', base)
    base = re.sub(r'[\s_\-]+', ' ', base).strip(' _-.')
    return base or '_default'


# ── 學校名稱 ↔ 機關代碼 ───────────────────────────────
def _norm_name(s):
    s = unicodedata.normalize('NFKC', str(s or ''))
    s = re.sub(r'\s+', '', s)
    for a, b in _SHORT_FORMS:
        s = s.replace(a, b)
    return s


def _core_name(s):
    return _PREFIX_RE.sub('', _norm_name(s))


def find_school_in_text(texts, schools):
    """在標題文字中找學校；schools 為 [(機關代碼, 學校全名)]。
    先找全名，再找去掉「高雄市／市立／某某區」後的核心名稱；必須唯一才採用。
    回傳 (代碼, 學校全名) 或 ('', '')。"""
    if isinstance(texts, str):
        texts = [texts]
    joined = [_norm_name(t) for t in (texts or []) if t]
    if not joined or not schools:
        return '', ''
    full = {(c, n) for c, n in schools if n and any(_norm_name(n) in t for t in joined)}
    if full:
        # 「三民國民中學」也包含在「高雄市立三民國民中學」裡：取最長的全名
        best = max(len(_norm_name(n)) for _, n in full)
        full = {(c, n) for c, n in full if len(_norm_name(n)) == best}
        return full.pop() if len(full) == 1 else ('', '')
    core = {}
    for c, n in schools:
        k = _core_name(n)
        if len(k) >= 4:
            core.setdefault(k, set()).add((c, n))
    hits = set()
    for k, rows in core.items():
        if any(k in t for t in joined):
            if len(rows) != 1:
                return '', ''          # 同名學校（例：不同區的中正國小）→ 不猜
            hits |= rows
    if len(hits) == 1:
        return hits.pop()
    return '', ''


def lookup_name(name, schools):
    """AF 內容的學校名稱 → (代碼, 全名)；對不到唯一一校時回傳 ('', 正規化名稱)。"""
    code, full = find_school_in_text([name], schools)
    if code:
        return code, full
    return '', _norm_name(name)


# ── AF 內容 ──────────────────────────────────────────
def af_content(af_bytes, af_name):
    """讀 AF 第一個工作表：回傳 {'code', 'name', 'ids'}（ids 僅供即時計算指紋）。"""
    out = {'code': '', 'name': '', 'ids': []}
    if not af_bytes:
        return out
    try:
        import pandas as pd
        ext = os.path.splitext(af_name or '')[1].lower()
        engine = 'xlrd' if ext == '.xls' else 'openpyxl'
        df = pd.read_excel(io.BytesIO(af_bytes), sheet_name=0, engine=engine,
                           dtype=str, header=0).fillna('')
    except Exception:
        return out
    cols = [str(c).strip() for c in df.columns]
    df.columns = cols

    def most_common(col, ok):
        vals = {}
        for v in df[col]:
            v = str(v).strip()
            if v and ok(v):
                vals[v] = vals.get(v, 0) + 1
        return max(vals, key=vals.get) if vals else ''

    for c in CODE_COLUMNS:
        if c in cols:
            v = most_common(c, lambda s: bool(CODE_FULL_RE.match(s)))
            if v:
                out['code'] = v.upper()
                break
    for c in NAME_COLUMNS:
        if c in cols:
            v = most_common(c, lambda s: len(s) >= 3 and not s.isdigit())
            if v:
                out['name'] = v
                break
    if '身分證字號' in cols:
        out['ids'] = sorted({str(v).strip().upper() for v in df['身分證字號']
                             if str(v).strip()})
    return out


# ── 人員組成指紋 ─────────────────────────────────────
_fp_lock = threading.Lock()
FP_MIN_IDS = 5
FP_MATCH = 0.6
FP_MARGIN = 0.2


def _secret(data_dir, create=True):
    path = os.path.join(data_dir, '.school_fp_secret')
    try:
        with open(path, 'rb') as f:
            key = f.read()
        if len(key) >= 32:
            return key
    except OSError:
        pass
    if not create:
        return None
    key = os.urandom(32)
    fd = os.open(path, os.O_WRONLY | os.O_CREAT | os.O_TRUNC, 0o600)
    with os.fdopen(fd, 'wb') as f:
        f.write(key)
    return key


def hash_ids(ids, data_dir, create=True):
    key = _secret(data_dir, create)
    if key is None:
        return []
    return sorted({hmac.new(key, i.encode(), hashlib.sha256).hexdigest()[:16]
                   for i in ids if i})


def _fp_file(data_dir):
    return os.path.join(data_dir, 'school_fingerprints.json')


def _load_fp(data_dir):
    try:
        with open(_fp_file(data_dir), encoding='utf-8') as f:
            d = json.load(f)
        return d if isinstance(d, dict) else {}
    except Exception:
        return {}


def match_fingerprint(ids, data_dir):
    """回傳 (學校鍵, 重疊率)；不夠明確時 ('', 0)。"""
    if len(set(ids)) < FP_MIN_IDS:
        return '', 0.0
    known_all = _load_fp(data_dir)
    if not known_all:
        return '', 0.0              # 尚無任何指紋：不建立密鑰、不寫檔
    cur = set(hash_ids(ids, data_dir, create=False))
    if not cur:
        return '', 0.0
    scores = []
    for school, rec in known_all.items():
        known = set(rec.get('h') or [])
        if len(known) < FP_MIN_IDS:
            continue
        common = len(cur & known)
        if common < FP_MIN_IDS:
            continue
        scores.append((common / min(len(cur), len(known)), school))
    if not scores:
        return '', 0.0
    scores.sort(reverse=True)
    best, school = scores[0]
    if best < FP_MATCH or (len(scores) > 1 and scores[1][0] > best - FP_MARGIN):
        return '', 0.0
    return school, round(best, 3)


def remember_fingerprint(school, ids, data_dir):
    """以本次 AF 人員組成更新該校指紋（取代舊值，跟上人事異動）。"""
    if not school or len(set(ids)) < FP_MIN_IDS:
        return False
    with _fp_lock:
        d = _load_fp(data_dir)
        d[school] = {'h': hash_ids(ids, data_dir), 'updated_at': int(time.time())}
        tmp = _fp_file(data_dir) + '.tmp'
        fd = os.open(tmp, os.O_WRONLY | os.O_CREAT | os.O_TRUNC, 0o600)
        with os.fdopen(fd, 'w', encoding='utf-8') as f:
            json.dump(d, f)
        os.replace(tmp, _fp_file(data_dir))
    return True


# ── 綜合判定 ─────────────────────────────────────────
def resolve(af_name='', af_bytes=None, schools=None, pdf_texts=None,
            data_dir=None, content=None):
    """回傳 {'key', 'source', 'school_name', 'legacy_keys', 'ids'}。"""
    schools = schools or []
    names = dict(schools)
    content = content if content is not None else af_content(af_bytes, af_name)
    ids = content.get('ids') or []
    old = legacy_key(af_name)

    def done(key, source, school_name=''):
        legacy = [k for k in (old,) if k and k != key]
        return {'key': key, 'source': source,
                'school_name': school_name or names.get(key, ''),
                'legacy_keys': legacy, 'ids': ids}

    if content.get('code'):
        return done(content['code'], 'af_code')
    code = filename_code(af_name)
    if code:
        return done(code, 'filename_code')
    if content.get('name'):
        c, n = lookup_name(content['name'], schools)
        return done(c or n, 'af_name', n)
    if pdf_texts:
        c, n = find_school_in_text(pdf_texts, schools)
        if c:
            return done(c, 'pdf_name', n)
    if data_dir and ids:
        k, _rate = match_fingerprint(ids, data_dir)
        if k:
            return done(k, 'fingerprint')
    return done(clean_filename(af_name), 'filename')


def title_texts_from_text(text, limit=8):
    """PDF 文字層：取含學校字樣的前幾行（不含其他內容），供 find_school_in_text 使用。"""
    out = []
    for line in (text or '').splitlines()[:40]:
        if re.search(r'國民|國小|國中|中學|小學|學校|高中|高職', line):
            out.append(line.strip()[:80])
        if len(out) >= limit:
            break
    return out


def pick(d, key, legacy_keys=()):
    """讀取記憶：新鍵優先，查不到退回舊的檔名鍵。回傳 (實際鍵, 記錄)。"""
    if key in d:
        return key, d[key]
    for k in legacy_keys or ():
        if k in d:
            return k, d[k]
    return key, None
