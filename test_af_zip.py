"""AF 校對清冊加密 ZIP 匯入（v1.3.20）。全部使用虛構資料自製的 ZIP。"""
import io
import os
import re
import struct
import unittest
import urllib.parse
import zipfile
import zlib
from unittest.mock import patch

import af_zip
from app import app

FAKE_NAME = '校對清冊_Query_PayData_Print_000000000X_115112300000000.xls'
OLE = b'\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1' + b'\x00' * 504 + b'fake-af-content' * 50


def _crc_table():
    t = []
    for n in range(256):
        c = n
        for _ in range(8):
            c = (c >> 1) ^ 0xEDB88320 if c & 1 else c >> 1
        t.append(c)
    return t


_CRC = _crc_table()


class _ZipCrypto:
    def __init__(self, pwd):
        self.k = [0x12345678, 0x23456789, 0x34567890]
        for b in pwd:
            self._update(b)

    def _crc(self, c, b):
        return (c >> 8) ^ _CRC[(c ^ b) & 0xFF]

    def _update(self, b):
        k0, k1, k2 = self.k
        k0 = self._crc(k0, b)
        k1 = (k1 + (k0 & 0xFF)) & 0xFFFFFFFF
        k1 = (k1 * 134775813 + 1) & 0xFFFFFFFF
        k2 = self._crc(k2, (k1 >> 24) & 0xFF)
        self.k = [k0, k1, k2]

    def encrypt(self, data):
        out = bytearray()
        for b in data:
            t = (self.k[2] | 2) & 0xFFFF
            out.append(b ^ (((t * (t ^ 1)) >> 8) & 0xFF))
            self._update(b)
        return bytes(out)


def make_zip(entries, password=None, big5=True):
    """entries: [(name, bytes)]；以 ZipCrypto + deflate 加密（同 AF 系統），檔名預設 Big5 不標 UTF-8。"""
    buf = io.BytesIO()
    central = []
    for name, data in entries:
        raw_name = name.encode('cp950') if big5 else name.encode('utf-8')
        flags = (0 if big5 else 0x800)
        crc = zlib.crc32(data) & 0xFFFFFFFF
        co = zlib.compressobj(9, zlib.DEFLATED, -15)
        comp = co.compress(data) + co.flush()
        if password:
            flags |= 1
            zc = _ZipCrypto(password.encode('utf-8'))
            header = os.urandom(11) + bytes([crc >> 24])
            comp = zc.encrypt(header + comp)
        offset = buf.tell()
        buf.write(struct.pack('<IHHHHHIIIHH', 0x04034B50, 20, flags, 8, 0, 0, crc,
                              len(comp), len(data), len(raw_name), 0))
        buf.write(raw_name)
        buf.write(comp)
        central.append(struct.pack('<IHHHHHHIIIHHHHHII', 0x02014B50, 20, 20, flags, 8, 0, 0,
                                   crc, len(comp), len(data), len(raw_name), 0, 0, 0, 0, 0,
                                   offset) + raw_name)
    cd_start = buf.tell()
    for c in central:
        buf.write(c)
    cd_size = buf.tell() - cd_start
    buf.write(struct.pack('<IHHHHIIH', 0x06054B50, 0, 0, len(central), len(central),
                          cd_size, cd_start, 0))
    return buf.getvalue()


class UnpackTests(unittest.TestCase):
    def test_public_password_and_big5_name(self):
        data = make_zip([(FAKE_NAME, OLE)], af_zip.PUBLIC_PASSWORD)
        content, name, used = af_zip.unpack_af_zip(data)
        self.assertEqual(content, OLE)
        self.assertEqual(name, FAKE_NAME)
        self.assertEqual(used, 'public')

    def test_custom_password(self):
        data = make_zip([(FAKE_NAME, OLE)], 'Other#Pass99')
        with self.assertRaises(af_zip.AFZipError) as cm:
            af_zip.unpack_af_zip(data)
        self.assertEqual(cm.exception.code, 'password_required')
        with self.assertRaises(af_zip.AFZipError) as cm:
            af_zip.unpack_af_zip(data, 'wrong')
        self.assertEqual(cm.exception.code, 'bad_password')
        content, name, used = af_zip.unpack_af_zip(data, 'Other#Pass99')
        self.assertEqual((content, used), (OLE, 'custom'))

    def test_custom_password_falls_back_to_public(self):
        data = make_zip([(FAKE_NAME, OLE)], af_zip.PUBLIC_PASSWORD)
        content, _, used = af_zip.unpack_af_zip(data, 'remembered-old-one')
        self.assertEqual((content, used), (OLE, 'public'))

    def test_unencrypted_and_utf8_name(self):
        data = make_zip([(FAKE_NAME, OLE)], None, big5=False)
        content, name, used = af_zip.unpack_af_zip(data)
        self.assertEqual((content, name, used), (OLE, FAKE_NAME, 'none'))

    def test_ignores_system_files(self):
        data = make_zip([('__MACOSX/._x.xls', b'junk'), ('.DS_Store', b'x'), (FAKE_NAME, OLE)],
                        af_zip.PUBLIC_PASSWORD)
        self.assertEqual(af_zip.unpack_af_zip(data)[1], FAKE_NAME)

    def test_rejects_multiple_excel(self):
        data = make_zip([(FAKE_NAME, OLE), ('b.xls', OLE)], af_zip.PUBLIC_PASSWORD)
        with self.assertRaises(af_zip.AFZipError) as cm:
            af_zip.unpack_af_zip(data)
        self.assertEqual(cm.exception.code, 'multi_excel')

    def test_rejects_non_excel(self):
        data = make_zip([('readme.txt', b'hello')], af_zip.PUBLIC_PASSWORD)
        with self.assertRaises(af_zip.AFZipError) as cm:
            af_zip.unpack_af_zip(data)
        self.assertEqual(cm.exception.code, 'no_excel')

    def test_rejects_fake_excel_content(self):
        data = make_zip([(FAKE_NAME, b'MZ not an excel file' * 10)], af_zip.PUBLIC_PASSWORD)
        with self.assertRaises(af_zip.AFZipError) as cm:
            af_zip.unpack_af_zip(data)
        self.assertEqual(cm.exception.code, 'not_excel')

    def test_rejects_path_traversal(self):
        data = make_zip([('../../evil.xls', OLE)], af_zip.PUBLIC_PASSWORD)
        with self.assertRaises(af_zip.AFZipError) as cm:
            af_zip.unpack_af_zip(data)
        self.assertEqual(cm.exception.code, 'unsafe_name')

    def test_zip_bomb_is_stopped(self):
        bomb = OLE[:8] + b'\x00' * (40 * 1024 * 1024)
        data = make_zip([(FAKE_NAME, bomb)], af_zip.PUBLIC_PASSWORD)
        self.assertLess(len(data), 200 * 1024)
        with self.assertRaises(af_zip.AFZipError) as cm:
            af_zip.unpack_af_zip(data)
        self.assertEqual(cm.exception.code, 'too_large')

    def test_high_ratio_with_lying_header_is_stopped(self):
        bomb = OLE[:8] + b'\x00' * (8 * 1024 * 1024)
        data = bytearray(make_zip([(FAKE_NAME, bomb)], None, big5=False))
        # 解壓後大小未達上限，但壓縮比異常高，邊讀邊檢查時就應中止
        with patch.object(af_zip, 'MAX_RATIO', 50):
            with self.assertRaises(af_zip.AFZipError) as cm:
                af_zip.unpack_af_zip(bytes(data))
        self.assertEqual(cm.exception.code, 'too_large')

    def test_rejects_xlsx_and_garbage(self):
        xlsx = io.BytesIO()
        with zipfile.ZipFile(xlsx, 'w') as z:
            z.writestr('[Content_Types].xml', '<x/>')
        with self.assertRaises(af_zip.AFZipError) as cm:
            af_zip.unpack_af_zip(xlsx.getvalue())
        self.assertEqual(cm.exception.code, 'is_xlsx')
        with self.assertRaises(af_zip.AFZipError) as cm:
            af_zip.unpack_af_zip(b'PK\x03\x04garbage')
        self.assertEqual(cm.exception.code, 'bad_zip')
        with self.assertRaises(af_zip.AFZipError) as cm:
            af_zip.unpack_af_zip(b'hello')
        self.assertEqual(cm.exception.code, 'not_zip')


class EndpointTests(unittest.TestCase):
    def setUp(self):
        app.config.update(TESTING=True)
        self.client = app.test_client()
        self.h = {'X-Forwarded-Proto': 'https'}

    def post(self, data, password=None):
        form = {'file': (io.BytesIO(data), 'download.zip')}
        if password is not None:
            form['password'] = password
        return self.client.post('/af/unzip', data=form, headers=self.h,
                                content_type='multipart/form-data')

    def test_unzip_returns_excel_with_original_name(self):
        r = self.post(make_zip([(FAKE_NAME, OLE)], af_zip.PUBLIC_PASSWORD))
        self.assertEqual(r.status_code, 200)
        self.assertEqual(r.data, OLE)
        self.assertEqual(urllib.parse.unquote(r.headers['X-AF-Filename']), FAKE_NAME)
        self.assertEqual(r.headers['X-AF-Password'], 'public')
        self.assertIn('no-store', r.headers['Cache-Control'])

    def test_password_required_then_custom(self):
        data = make_zip([(FAKE_NAME, OLE)], 'Mine!2026')
        r = self.post(data)
        self.assertEqual(r.status_code, 422)
        self.assertEqual(r.get_json()['code'], 'password_required')
        r = self.post(data, 'nope')
        self.assertEqual(r.get_json()['code'], 'bad_password')
        r = self.post(data, 'Mine!2026')
        self.assertEqual(r.status_code, 200)
        self.assertEqual(r.headers['X-AF-Password'], 'custom')

    def test_password_not_logged(self):
        data = make_zip([(FAKE_NAME, OLE)], 'Secret!Pw1')
        with self.assertLogs(app.logger, level='DEBUG') as logs:
            app.logger.debug('marker')
            self.post(data, 'Secret!Pw1')
            self.post(b'PK\x03\x04broken', 'Secret!Pw1')
        self.assertFalse(any('Secret!Pw1' in line for line in logs.output))

    def test_errors_are_json(self):
        r = self.post(make_zip([('a.txt', b'x')], af_zip.PUBLIC_PASSWORD))
        self.assertEqual(r.status_code, 400)
        self.assertEqual(r.get_json()['code'], 'no_excel')
        r = self.client.post('/af/unzip', data={}, headers=self.h)
        self.assertEqual(r.get_json()['code'], 'no_file')

    @patch('app.touch_active')
    @patch('app.bump_counter', return_value={'visits': 1, 'sorts': 0, 'compares': 0})
    def test_home_has_zip_import_guide_and_bookmarklet(self, _c, _a):
        page = self.client.get('/', headers=self.h).get_data(as_text=True)
        self.assertIn('/static/af_import.js?v=v1.3.25', page)
        self.assertIn('AF:各機關學校用人費用管理資訊系統', page)
        self.assertIn('https://fund.dgpa.gov.tw/system/WebMain.aspx', page)
        self.assertIn('校對清冊(A)', page)
        self.assertIn(af_zip.PUBLIC_PASSWORD, page)
        self.assertIn('兼辦機關', page)

    def test_static_scripts_present(self):
        for p in ('/static/af_import.js', '/static/af_oneclick.js'):
            r = self.client.get(p, headers=self.h)
            self.assertEqual(r.status_code, 200, p)
            r.close()

    def test_scripts_do_not_contain_personal_secrets(self):
        root = os.path.dirname(os.path.abspath(__file__))
        for p in ('static/af_import.js', 'static/af_oneclick.js', 'af_zip.py',
                  'templates/_af_import.html'):
            text = open(os.path.join(root, p), encoding='utf-8').read()
            self.assertIsNone(re.search(r'[A-Z][12]\d{8}', text), p)


if __name__ == '__main__':
    unittest.main()
