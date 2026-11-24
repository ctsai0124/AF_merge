"""各校樣式記憶（layout_profiles.json）測試：全部在暫存目錄、使用虛構學校鍵。"""
import json
import os
import tempfile
import unittest
from unittest.mock import patch

import app as A


class LayoutProfileStyleTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.env = patch.dict(os.environ, {'DATA_DIR': self.tmp.name})
        self.env.start()
        self.path = os.path.join(self.tmp.name, 'layout_profiles.json')

    def tearDown(self):
        self.env.stop()
        self.tmp.cleanup()

    def write(self, d):
        with open(self.path, 'w', encoding='utf-8') as f:
            json.dump(d, f, ensure_ascii=False)

    def read(self):
        with open(self.path, encoding='utf-8') as f:
            return json.load(f)

    def test_legacy_record_is_still_readable(self):
        self.write({'999000001Y': {'layout': 'vertical', 'samples': 4,
                                   'name_match_rate': 0.8, 'updated_at': 1}})
        p = A.load_layout_profile('999000001Y')
        self.assertEqual(p['layout'], 'vertical')
        self.assertEqual(p['samples'], 4)
        self.assertNotIn('style', p)

    def test_save_records_style_traits_and_strategy(self):
        A.save_layout_profile('999000002Y', 'vertical', 0.9, {
            'style': 'vertical_column', 'style_label': '任意文字會被白名單取代',
            'traits': ['has_summary_pages', 'skewed_header', 'has_bufa_column', '王小明', 'evil'],
            'ocr_strategy': 'full'})
        rec = self.read()['999000002Y']
        self.assertEqual(rec['style'], 'vertical_column')
        self.assertEqual(rec['style_label'], '直式一人一欄')
        self.assertEqual(rec['traits'], ['has_bufa_column', 'has_summary_pages', 'skewed_header'])
        self.assertEqual(rec['ocr_strategy'], 'full')
        self.assertEqual(rec['samples'], 1)
        loaded = A.load_layout_profile('999000002Y')
        self.assertEqual(loaded['style'], 'vertical_column')

    def test_ledger_style_and_tiled_3x2_strategy_are_recorded(self):
        # v1.3.26：阿蓮國小型 A4 橫式薪資印領清冊
        A.save_layout_profile('999000005Y', 'horizontal', 1.0, {
            'style': 'horizontal_ledger',
            'traits': ['has_bufa_column', 'has_summary_pages', 'needs_tiled_ocr'],
            'ocr_strategy': 'tiled_3x2'})
        rec = self.read()['999000005Y']
        self.assertEqual(rec['style'], 'horizontal_ledger')
        self.assertEqual(rec['style_label'], A.STYLE_LABELS['horizontal_ledger'])
        self.assertEqual(rec['ocr_strategy'], 'tiled_3x2')
        self.assertEqual(rec['traits'], ['has_bufa_column', 'has_summary_pages', 'needs_tiled_ocr'])

    def test_style_contradicting_layout_is_not_saved(self):
        A.save_layout_profile('999000003Y', 'vertical', 0.9,
                              {'style': 'horizontal_grid', 'ocr_strategy': 'tiled_2x2'})
        self.assertNotIn('style', self.read()['999000003Y'])

    def test_missing_style_keeps_previous_style_and_legacy_fields(self):
        self.write({'999000004Y': {'layout': 'horizontal', 'samples': 2,
                                   'name_match_rate': 0.7, 'updated_at': 1,
                                   'style': 'horizontal_grid', 'traits': ['needs_tiled_ocr'],
                                   'ocr_strategy': 'tiled_2x2'},
                    '999000005Y': {'layout': 'vertical', 'samples': 1,
                                   'name_match_rate': 0.6, 'updated_at': 1}})
        A.save_layout_profile('999000004Y', 'horizontal', 0.8)
        d = self.read()
        self.assertEqual(d['999000004Y']['style'], 'horizontal_grid')
        self.assertEqual(d['999000004Y']['samples'], 3)
        self.assertEqual(d['999000005Y'], {'layout': 'vertical', 'samples': 1,
                                           'name_match_rate': 0.6, 'updated_at': 1})

    def test_school_switching_layout_replaces_old_style(self):
        A.save_layout_profile('999000006Y', 'horizontal', 0.8,
                              {'style': 'horizontal_grid', 'ocr_strategy': 'tiled_2x2'})
        A.save_layout_profile('999000006Y', 'vertical', 0.9)
        rec = self.read()['999000006Y']
        self.assertEqual(rec['layout'], 'vertical')
        self.assertNotIn('style', rec)          # 舊橫式樣式不可沿用到直式

    def test_schools_by_style_groups_across_schools(self):
        A.save_layout_profile('999000007Y', 'vertical', 0.9, {'style': 'vertical_column'})
        A.save_layout_profile('999000008Y', 'vertical', 0.9, {'style': 'vertical_column'})
        A.save_layout_profile('999000009Y', 'horizontal', 0.9, {'style': 'horizontal_row'})
        A.save_layout_profile('999000010Y', 'horizontal', 0.9)
        self.assertEqual(A.schools_by_style(), {
            'horizontal_row': ['999000009Y'],
            'vertical_column': ['999000007Y', '999000008Y']})

    def test_claim_sends_style_hint_and_result_stores_profile(self):
        A.save_layout_profile('999000011Y', 'horizontal', 0.9, {
            'style': 'horizontal_grid', 'traits': ['needs_tiled_ocr'],
            'ocr_strategy': 'tiled_2x2'})
        client = A.app.test_client()
        jid = 'testjob000001'
        with A._ocr_lock:
            A._ocr_jobs[jid] = {'status': 'pending', 'created': __import__('time').time(),
                                'pdf': b'%PDF', 'af': b'', 'af_name': 'x.xlsx',
                                'school_key': '999000011Y', 'people': None,
                                'layout': None, 'error': None}
        try:
            with patch('app._ocr_auth', return_value=True):
                r = client.get('/ocr/claim').get_json()
                self.assertEqual(r['job_id'], jid)
                self.assertEqual(r['layout_hint'], 'horizontal')
                self.assertEqual(r['style_hint']['style'], 'horizontal_grid')
                self.assertEqual(r['style_hint']['ocr_strategy'], 'tiled_2x2')
                client.post('/ocr/result', json={
                    'job_id': jid, 'people': [], 'layout': 'horizontal',
                    'layout_profile': {'style': 'horizontal_grid',
                                       'traits': ['needs_tiled_ocr', 'bogus'],
                                       'ocr_strategy': 'tiled_2x2'}})
            prof = A._ocr_jobs[jid]['layout_profile']
            self.assertEqual(prof['traits'], ['needs_tiled_ocr'])
        finally:
            A._ocr_jobs.pop(jid, None)

    def test_legacy_school_claim_has_no_style_hint(self):
        self.write({'999000012Y': {'layout': 'vertical', 'samples': 1,
                                   'name_match_rate': 0.6, 'updated_at': 1}})
        client = A.app.test_client()
        jid = 'testjob000002'
        with A._ocr_lock:
            A._ocr_jobs[jid] = {'status': 'pending', 'created': __import__('time').time(),
                                'pdf': b'%PDF', 'af': b'', 'af_name': 'x.xlsx',
                                'school_key': '999000012Y', 'people': None,
                                'layout': None, 'error': None}
        try:
            with patch('app._ocr_auth', return_value=True):
                r = client.get('/ocr/claim').get_json()
            self.assertEqual(r['layout_hint'], 'vertical')
            self.assertIsNone(r['style_hint'])
        finally:
            A._ocr_jobs.pop(jid, None)


if __name__ == '__main__':
    unittest.main()
