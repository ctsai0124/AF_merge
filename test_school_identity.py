"""學校鍵判定測試：虛構學校、虛構機關代碼（999 開頭）、以檢查碼產生的假身分證；全部在暫存目錄。"""
import io
import json
import os
import tempfile
import time
import unittest
from unittest.mock import patch

import openpyxl

import app as A
import school_identity as SID

SCHOOLS = [
    ('999000101Y', '高雄市測試區星光國民小學'),
    ('999000102Y', '高雄市虛構區星光國民小學'),     # 與上一校同名不同區
    ('999000103X', '高雄市立月影國民中學'),
    ('999000104Y', '高雄市假想區晨曦國民小學'),
]
_LETTER = {c: v for c, v in zip('ABCDEFGHJKLMNPQRSTUVXYWZIO', range(10, 36))}


def fake_id(n, letter='A'):
    body = f'1{n:07d}'
    v = _LETTER[letter]
    total = v // 10 + (v % 10) * 9 + sum(int(d) * w for d, w in zip(body, range(8, 0, -1)))
    return f'{letter}{body}{(10 - total % 10) % 10}'


def make_af(ids, extra_cols=None):
    wb = openpyxl.Workbook()
    ws = wb.active
    head = ['單位', '身分證字號', '姓名', '薪俸表別', '支領數額', '專業加給表別', '支領數額',
            '職務加給表別', '支領數額']
    extra_cols = extra_cols or {}
    ws.append(head + list(extra_cols))
    for i, pid in enumerate(ids):
        ws.append(['教務處', pid, f'測試員{i:03d}', 'A1', '30000', '', '', '', '']
                  + list(extra_cols.values()))
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


class FilenameTests(unittest.TestCase):
    def test_noise_is_removed(self):
        for name in ('複本 校對清冊.xls', '校對清冊 (1).xls', '校對清冊（2）.xls',
                     '校對清冊 - Copy.xls', 'Copy of 校對清冊.xls', '校對清冊的副本.xls',
                     '校對清冊_複本.xls', '副本 校對清冊_1151120103012.xls'):
            self.assertEqual(SID.clean_filename(name), '校對清冊', name)

    def test_code_from_filename(self):
        self.assertEqual(SID.filename_code('AF_999000101Y_11511201030.xls'), '999000101Y')
        self.assertEqual(SID.filename_code('複本 999000101y 清冊.xlsx'), '999000101Y')
        self.assertEqual(SID.filename_code('複本 校對清冊.xls'), '')
        # 身分證格式（英文字母開頭）不可被當成機關代碼
        self.assertEqual(SID.filename_code(f'{fake_id(1)}.xls'), '')

    def test_legacy_key_matches_old_rule(self):
        self.assertEqual(SID.legacy_key('複本 校對清冊_1151120.xls'), '複本 校對清冊')
        self.assertEqual(SID.legacy_key('AF_999000101Y_115112010.xls'), '999000101Y')
        self.assertEqual(SID.legacy_key('清冊.xlsx'), '清冊.xlsx')


class SchoolNameTests(unittest.TestCase):
    def test_full_name_in_title(self):
        self.assertEqual(SID.find_school_in_text(
            ['高雄市立月影國民中學115年11月份'], SCHOOLS)[0], '999000103X')

    def test_short_form_unique_core(self):
        self.assertEqual(SID.find_school_in_text(['高雄市晨曦國小115年9月薪資印領清冊'],
                                                 SCHOOLS)[0], '999000104Y')

    def test_ambiguous_core_is_not_guessed(self):
        self.assertEqual(SID.find_school_in_text(['星光國小薪資清冊'], SCHOOLS), ('', ''))
        self.assertEqual(SID.find_school_in_text(['高雄市虛構區星光國民小學員工薪津'],
                                                 SCHOOLS)[0], '999000102Y')

    def test_misread_name_does_not_match(self):
        self.assertEqual(SID.find_school_in_text(['月彰國民中學'], SCHOOLS), ('', ''))


class ResolveTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.dir = self.tmp.name
        self.ids = [fake_id(i) for i in range(20)]

    def tearDown(self):
        self.tmp.cleanup()

    def res(self, name, af=None, texts=None):
        return SID.resolve(name, af, SCHOOLS, texts, self.dir)

    def test_content_code_beats_filename(self):
        af = make_af(self.ids, {'機關代碼': '999000104Y'})
        r = self.res('AF_999000101Y_1151120.xlsx', af)
        self.assertEqual((r['key'], r['source']), ('999000104Y', 'af_code'))

    def test_filename_code(self):
        r = self.res('AF_999000101Y_1151120.xlsx', make_af(self.ids))
        self.assertEqual((r['key'], r['source']), ('999000101Y', 'filename_code'))
        self.assertEqual(r['school_name'], '高雄市測試區星光國民小學')

    def test_content_school_name(self):
        af = make_af(self.ids, {'機關名稱': '高雄市立月影國民中學'})
        r = self.res('複本 校對清冊.xlsx', af)
        self.assertEqual((r['key'], r['source']), ('999000103X', 'af_name'))

    def test_pdf_title_name(self):
        r = self.res('複本 校對清冊.xlsx', make_af(self.ids), ['高雄市立月影國民中學115年11月份'])
        self.assertEqual((r['key'], r['source']), ('999000103X', 'pdf_name'))
        self.assertEqual(r['legacy_keys'], ['複本 校對清冊.xlsx'])

    def test_renamed_file_found_by_fingerprint(self):
        SID.remember_fingerprint('999000101Y', self.ids, self.dir)
        SID.remember_fingerprint('999000103X', [fake_id(500 + i) for i in range(20)], self.dir)
        # 人事異動：少了 3 人、多了 2 人，仍是同一校
        ids = self.ids[3:] + [fake_id(900), fake_id(901)]
        r = self.res('複本 校對清冊.xlsx', make_af(ids))
        self.assertEqual((r['key'], r['source']), ('999000101Y', 'fingerprint'))

    def test_unknown_school_falls_back_to_clean_filename(self):
        r = self.res('複本 校對清冊_1151120.xls', make_af(self.ids))
        self.assertEqual((r['key'], r['source']), ('校對清冊', 'filename'))
        self.assertEqual(r['legacy_keys'], ['複本 校對清冊'])

    def test_fingerprint_stores_no_raw_ids(self):
        SID.remember_fingerprint('999000101Y', self.ids, self.dir)
        raw = open(os.path.join(self.dir, 'school_fingerprints.json'), encoding='utf-8').read()
        for pid in self.ids:
            self.assertNotIn(pid, raw)
        mode = os.stat(os.path.join(self.dir, '.school_fp_secret')).st_mode & 0o777
        self.assertEqual(mode, 0o600)

    def test_lookup_without_fingerprints_writes_nothing(self):
        r = self.res('校對清冊.xlsx', make_af(self.ids))
        self.assertEqual(r['source'], 'filename')
        self.assertEqual(os.listdir(self.dir), [])

    def test_title_lines_from_text_layer(self):
        text = '高雄市立月影國民中學115年11月份薪資清冊\n姓名 職稱 月支俸額\n測試員 教師 30000'
        self.assertEqual(SID.title_texts_from_text(text), ['高雄市立月影國民中學115年11月份薪資清冊'])

    def test_too_few_or_ambiguous_fingerprints_are_ignored(self):
        SID.remember_fingerprint('999000101Y', self.ids, self.dir)
        SID.remember_fingerprint('999000102Y', self.ids, self.dir)    # 兩校相同 → 不猜
        r = self.res('校對清冊.xlsx', make_af(self.ids))
        self.assertEqual(r['source'], 'filename')
        self.assertEqual(SID.match_fingerprint(self.ids[:3], self.dir), ('', 0.0))


class AppMemoryTests(unittest.TestCase):
    """App 層：寫入新鍵、讀取退回舊檔名鍵；改名「複本 xxx」仍對回同一校。"""

    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.env = patch.dict(os.environ, {'DATA_DIR': self.tmp.name})
        self.env.start()
        self.tbl = patch('app.school_table', return_value=SCHOOLS)
        self.tbl.start()
        self.ids = [fake_id(100 + i) for i in range(15)]

    def tearDown(self):
        self.tbl.stop()
        self.env.stop()
        self.tmp.cleanup()

    def jread(self, name):
        with open(os.path.join(self.tmp.name, name), encoding='utf-8') as f:
            return json.load(f)

    def test_exclusions_fall_back_to_legacy_key_and_save_new_key(self):
        with open(os.path.join(self.tmp.name, 'exclusions.json'), 'w', encoding='utf-8') as f:
            json.dump({'複本 校對清冊': {'titles': ['測試兼任組長'], 'names': [],
                                      'use_default': True}}, f, ensure_ascii=False)
        ident = A.resolve_school('複本 校對清冊_1151120.xls', make_af(self.ids),
                                 pdf_texts=['高雄市立月影國民中學'])
        self.assertEqual(ident['key'], '999000103X')
        ex = A.load_exclusions(ident['key'], ident['legacy_keys'])
        self.assertEqual(ex['titles'], ['測試兼任組長'])
        A.save_exclusions(ident['key'], ex['titles'] + ['測試兼任主任'], [],
                          legacy_keys=ident['legacy_keys'])
        d = self.jread('exclusions.json')
        self.assertEqual(d['999000103X']['titles'], ['測試兼任主任', '測試兼任組長'])

    def test_layout_samples_continue_from_legacy_key(self):
        with open(os.path.join(self.tmp.name, 'layout_profiles.json'), 'w', encoding='utf-8') as f:
            json.dump({'複本 校對清冊': {'layout': 'horizontal', 'samples': 2,
                                      'name_match_rate': 1, 'updated_at': 1}}, f,
                      ensure_ascii=False)
        legacy = ['複本 校對清冊']
        self.assertEqual(A.load_layout_profile('999000103X', legacy)['layout'], 'horizontal')
        A.save_layout_profile('999000103X', 'horizontal', 0.9, None, legacy)
        self.assertEqual(self.jread('layout_profiles.json')['999000103X']['samples'], 3)

    def _scanned_flow(self, af_name, title_texts):
        client = A.app.test_client()
        af = make_af(self.ids)
        jid = 'schooltest' + str(int(time.time() * 1000))[-6:]
        ident = A.resolve_school(af_name, af)
        people = [{'姓名': f'測試員{i:03d}', '身分證': pid, '身分證有效': True,
                   '職稱': '教師', '薪俸': 30000, '專業加給': 0, '主管加給': 0,
                   '導師特教': 0, '其他加給': 0, '應發金額': 30000, '加總相符': True}
                  for i, pid in enumerate(self.ids)]
        with A._ocr_lock:
            A._ocr_jobs[jid] = {'status': 'pending', 'created': time.time(), 'pdf': b'%PDF',
                                'af': af, 'af_name': af_name, 'school_key': ident['key'],
                                'school': ident, 'people': None, 'layout': None, 'error': None}
        try:
            with patch('app._ocr_auth', return_value=True):
                claim = client.get('/ocr/claim').get_json()
                client.post('/ocr/result', json={
                    'job_id': jid, 'people': people, 'layout': 'vertical',
                    'layout_profile': {'style': 'vertical_column'},
                    'title_texts': title_texts})
            out = client.get(f'/ocr/status/{jid}').get_json()
        finally:
            A._ocr_jobs.pop(jid, None)
        return claim, out

    def test_renamed_af_maps_to_same_school_memory(self):
        # 第一次：原始檔名帶機關代碼 → 記住版面與人員組成指紋
        claim, out = self._scanned_flow('AF_999000104Y_1151120.xlsx', [])
        self.assertEqual(out['school_key'], '999000104Y')
        self.assertEqual(out['school_source'], 'filename_code')
        self.assertEqual(self.jread('layout_profiles.json')['999000104Y']['samples'], 1)
        # 第二次：同一份 AF 改名「複本 校對清冊.xlsx」，清冊標題也讀不到校名
        claim, out = self._scanned_flow('複本 校對清冊.xlsx', [])
        self.assertEqual(claim['school_key'], '999000104Y')
        self.assertEqual(claim['layout_hint'], 'vertical')
        self.assertEqual(out['school_key'], '999000104Y')
        self.assertEqual(out['school_source'], 'fingerprint')
        prof = self.jread('layout_profiles.json')
        self.assertEqual(prof['999000104Y']['samples'], 2)
        self.assertNotIn('複本 校對清冊.xlsx', prof)
        self.assertNotIn('校對清冊', prof)

    def test_ocr_title_upgrades_weak_filename_key(self):
        claim, out = self._scanned_flow('複本 校對清冊.xlsx', ['高雄市立月影國民中學115年11月份'])
        self.assertEqual(claim['school_key'], '校對清冊')
        self.assertEqual(out['school_key'], '999000103X')
        self.assertEqual(out['school_source'], 'pdf_name')
        self.assertIn('999000103X', self.jread('layout_profiles.json'))
        # 已驗證 → 指紋記住，下次改名也能對回
        r = A.resolve_school('又改名.xlsx', make_af(self.ids))
        self.assertEqual((r['key'], r['source']), ('999000103X', 'fingerprint'))


if __name__ == '__main__':
    unittest.main()
