"""AF 校對清冊加密 ZIP 解包（v1.3.20 新增）。

AF 系統「報表列印 → 校對清冊(A) → 下載報表」產出的是以「身分證號密碼」欄位
加密的 ZIP（傳統 ZipCrypto），裡面只有一個 .xls。這裡負責在記憶體中解開：

* 預設先用公版下載密碼 PUBLIC_PASSWORD（大家下載時一律填這組），使用者另外
  提供的密碼優先嘗試；密碼只用於本次解密，不寫檔、不記錄。
* 解開的檔案不落地，只回傳 bytes 與 ZIP 內原始檔名（Big5／CP950 還原），
  讓既有的 parse_af_filename()／school_identity 照舊用檔名中的機關代碼。
* 防護：僅接受恰好一個 .xls／.xlsx 項目、限制解壓後大小與壓縮比（zip bomb）、
  拒絕路徑跳脫檔名、檢查內容的檔頭確實是 Excel。
"""
import io
import posixpath
import zipfile
import zlib

PUBLIC_PASSWORD = '!@Abcdefg1234567'

MAX_ENTRY_BYTES = 30 * 1024 * 1024     # 單一 Excel 解壓後上限
MAX_RATIO = 200                        # 壓縮比上限（正常 xls 約 5～20 倍）
MAX_ENTRIES = 20                       # ZIP 目錄項目上限（含忽略的系統檔）
EXCEL_EXTS = ('.xls', '.xlsx')
OLE_MAGIC = b'\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1'
ZIP_MAGIC = b'PK\x03\x04'


class AFZipError(Exception):
    """可直接顯示給使用者的錯誤；code 供前端判斷要不要請使用者輸入密碼。"""

    def __init__(self, code, message):
        super().__init__(message)
        self.code = code
        self.message = message


def looks_like_zip(data):
    return bool(data) and data[:4] == ZIP_MAGIC


def _decode_name(info):
    """ZIP 未標 UTF-8 旗標時，Python 會用 CP437 解讀檔名；AF 系統用的是 Big5。"""
    name = info.filename
    if info.flag_bits & 0x800:
        return name
    try:
        raw = name.encode('cp437')
    except UnicodeEncodeError:
        return name
    for enc in ('utf-8', 'cp950', 'big5hkscs'):
        try:
            return raw.decode(enc)
        except UnicodeDecodeError:
            continue
    return name


def _is_ignorable(name):
    base = posixpath.basename(name.rstrip('/'))
    return (name.endswith('/') or name.startswith('__MACOSX/') or base.startswith('._')
            or base in ('.DS_Store', 'Thumbs.db', 'desktop.ini'))


def _safe_basename(name):
    name = name.replace('\\', '/')
    if name.startswith('/') or '..' in name.split('/'):
        raise AFZipError('unsafe_name', 'ZIP 內的檔名不合法，請重新從 AF 系統下載。')
    base = posixpath.basename(name)
    base = ''.join(ch for ch in base if ch >= ' ' and ch not in '<>:"|?*')
    return base.strip() or 'AF校對清冊.xls'


def _read_entry(zf, info, pwd):
    limit = min(MAX_ENTRY_BYTES, max(info.compress_size, 1) * MAX_RATIO)
    out = io.BytesIO()
    with zf.open(info, pwd=pwd) as fh:
        while True:
            chunk = fh.read(64 * 1024)
            if not chunk:
                break
            out.write(chunk)
            if out.tell() > limit:
                raise AFZipError('too_large', 'ZIP 解開後的檔案過大或壓縮比異常，已停止處理。')
    return out.getvalue()


def unpack_af_zip(data, password=None):
    """回傳 (excel_bytes, 原始檔名, 使用的密碼來源 'none'|'public'|'custom')。"""
    if not looks_like_zip(data):
        raise AFZipError('not_zip', '這不是 ZIP 壓縮檔。')
    try:
        zf = zipfile.ZipFile(io.BytesIO(data))
        infos = zf.infolist()
    except (zipfile.BadZipFile, zipfile.LargeZipFile, ValueError, EOFError):
        raise AFZipError('bad_zip', 'ZIP 檔已損毀或不完整，請重新從 AF 系統下載。')
    if len(infos) > MAX_ENTRIES:
        raise AFZipError('too_many', 'ZIP 內檔案過多，這不是 AF 校對清冊的下載檔。')
    names = [i.filename for i in infos]
    if '[Content_Types].xml' in names:
        raise AFZipError('is_xlsx', '這是 Excel（.xlsx）檔，請改用 .xlsx 副檔名直接上傳。')
    entries = [i for i in infos if not _is_ignorable(i.filename)]
    excel = [i for i in entries if _decode_name(i).lower().endswith(EXCEL_EXTS)]
    if not excel:
        raise AFZipError('no_excel', 'ZIP 內找不到 .xls／.xlsx 檔，請確認報表格式選的是 EXCEL。')
    if len(excel) > 1:
        raise AFZipError('multi_excel', 'ZIP 內有多個 Excel 檔，請一次只上傳一所學校、一份校對清冊。')
    info = excel[0]
    if info.file_size > MAX_ENTRY_BYTES:
        raise AFZipError('too_large', 'ZIP 解開後的檔案過大，已停止處理。')
    if info.compress_type not in (zipfile.ZIP_STORED, zipfile.ZIP_DEFLATED):
        # 99 = AES 加密（WinZip 格式），標準函式庫無法解開
        raise AFZipError('unsupported', '不支援這種 ZIP 壓縮／加密方式，請直接從 AF 系統重新下載。')
    name = _safe_basename(_decode_name(info))

    encrypted = bool(info.flag_bits & 0x1)
    if not encrypted:
        attempts = [(None, 'none')]
    else:
        attempts = []
        if password:
            attempts.append((password, 'custom'))
        if password != PUBLIC_PASSWORD:
            attempts.append((PUBLIC_PASSWORD, 'public'))

    content, used = None, None
    for pwd, label in attempts:
        try:
            content = _read_entry(zf, info, pwd.encode('utf-8') if pwd else None)
            used = label
            break
        except AFZipError:
            raise
        except (RuntimeError, zipfile.BadZipFile, zlib.error, EOFError, ValueError,
                NotImplementedError):
            # 密碼錯誤：ZipCrypto 檢查碼不符（RuntimeError）或 1/256 機率通過檢查碼後
            # 解壓失敗／CRC 不符（zlib.error／BadZipFile）
            continue
    if content is None:
        if password:
            raise AFZipError('bad_password', '密碼不正確，無法解開這個 ZIP，請再確認下載時填的密碼。')
        raise AFZipError('password_required',
                         '這個 ZIP 不是用公版密碼加密的，請輸入您下載時在「身分證號密碼」欄填的密碼。')

    if not (content.startswith(OLE_MAGIC) or content.startswith(ZIP_MAGIC)):
        raise AFZipError('not_excel', 'ZIP 內的檔案不是有效的 Excel 檔。')
    ext = '.xlsx' if content.startswith(ZIP_MAGIC) else '.xls'
    if not name.lower().endswith(ext):
        name = posixpath.splitext(name)[0] + ext
    return content, name, used
