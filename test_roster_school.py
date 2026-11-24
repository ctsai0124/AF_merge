"""v1.3.23：AF 學校 → 自動帶入該校固定清冊、AF／固定清冊不同校防呆。全部使用虛構資料。"""
import io
import json
import os
import shutil
import tempfile
import unittest
from unittest.mock import patch

import openpyxl
import pandas as pd

import app as app_module
import roster_school as RS
from app import app

# 公開的機關代碼（static/school.xlsx），人員皆為虛構
CODE_A, NAME_A = '397085000Y', '高雄市湖內區三侯國民小學'
CODE_B, NAME_B = '397061200X', '高雄市立寶來國民中學'
PEOPLE_A = [('甲一一', 'A100000001'), ('甲二二', 'A100000002'), ('甲三三', 'A100000003'),
            ('甲四四', 'A100000004'), ('甲五五', 'A100000005')]
PEOPLE_B = [('乙一一', 'B100000001'), ('乙二二', 'B100000002'), ('乙三三', 'B100000003'),
            ('乙四四', 'B100000004')]
SCHOOLS = [{'schoolId': 'sch6', 'name': NAME_B}, {'schoolId': 'sch7', 'name': NAME_A},
           {'schoolId': 'sch8', 'name': '高雄市路竹區竹滬國民小學'}]


def xlsx(header, rows):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'input'
    ws.append(header)
    for r in rows:
        ws.append(list(r))
    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()


def af_bytes(people):
    return xlsx(['身分證字號', '姓名', '薪俸表別', '支領數額'],
                [(i, n, 'A1', '40000') for n, i in people])


def roster_bytes(people):
    return xlsx(['序號', '姓名', '身分證字號'],
                [(k + 1, n, i) for k, (n, i) in enumerate(people)])


def af_name(code):
    return f'校對清冊_Query_PayData_Print_{code}_115112400000000.xlsx'


class MatchAccountSchoolTests(unittest.TestCase):
    def test_org_code_wins(self):
        schools = [{'schoolId': 'x1', 'name': '別的名字', 'orgCode': CODE_A},
                   {'schoolId': 'x2', 'name': NAME_A}]
        self.assertEqual(RS.match_account_school(CODE_A, NAME_A, schools), 'x1')

    def test_full_and_short_names(self):
        self.assertEqual(RS.match_account_school(CODE_A, NAME_A, SCHOOLS), 'sch7')
        short = [{'schoolId': 's', 'name': '三侯國小'}]
        self.assertEqual(RS.match_account_school(CODE_A, NAME_A, short), 's')

    def test_unknown_or_ambiguous_school_is_not_guessed(self):
        self.assertEqual(RS.match_account_school('397084500Y', '高雄市路竹區三埤國民小學', SCHOOLS), '')
        dup = [{'schoolId': 'a', 'name': NAME_A}, {'schoolId': 'b', 'name': NAME_A}]
        self.assertEqual(RS.match_account_school(CODE_A, NAME_A, dup), '')
        self.assertEqual(RS.match_account_school('', '', SCHOOLS), '')
        self.assertEqual(RS.match_account_school(CODE_A, NAME_A, None), '')


class RosterOverlapTests(unittest.TestCase):
    def df(self, data):
        return pd.read_excel(io.BytesIO(data), dtype=str)

    def test_same_school(self):
        ov = RS.roster_overlap(self.df(roster_bytes(PEOPLE_A[:4] + [('新進者', 'C100000009')])),
                               self.df(af_bytes(PEOPLE_A)), app_module.PERSON_ID_COLS)
        self.assertEqual((ov['af_total'], ov['matched']), (5, 4))
        self.assertTrue(ov['same_school'])

    def test_other_school(self):
        ov = RS.roster_overlap(self.df(roster_bytes(PEOPLE_B)), self.df(af_bytes(PEOPLE_A)),
                               app_module.PERSON_ID_COLS)
        self.assertEqual(ov['matched'], 0)
        self.assertIs(ov['same_school'], False)

    def test_too_few_people_is_undecided(self):
        ov = RS.roster_overlap(self.df(roster_bytes(PEOPLE_B[:1])), self.df(af_bytes(PEOPLE_A[:2])),
                               app_module.PERSON_ID_COLS)
        self.assertIsNone(ov['same_school'])


class IdentifyEndpointTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.mkdtemp()
        self.env = patch.dict(os.environ, {'DATA_DIR': self.tmp})
        self.env.start()
        self.client = app.test_client()

    def tearDown(self):
        self.env.stop()
        shutil.rmtree(self.tmp, ignore_errors=True)

    def post(self, af=None, roster=None, schools=None):
        data = {}
        if af is not None:
            data['af'] = (io.BytesIO(af[1]), af[0])
        if roster is not None:
            data['roster'] = (io.BytesIO(roster), '固定清冊.xlsx')
        if schools is not None:
            data['schools'] = json.dumps(schools)
        return self.client.post('/af/identify', data=data, content_type='multipart/form-data')

    def test_af_school_maps_to_account_school(self):
        r = self.post((af_name(CODE_A), af_bytes(PEOPLE_A)), schools=SCHOOLS)
        self.assertEqual(r.status_code, 200)
        d = r.get_json()
        self.assertEqual(d['code'], CODE_A)
        self.assertEqual(d['school_name'], NAME_A)
        self.assertEqual(d['account_school_id'], 'sch7')
        self.assertIsNone(d['overlap'])
        self.assertIn('no-store', r.headers.get('Cache-Control', ''))

    def test_af_of_school_a_with_roster_of_school_b_is_flagged(self):
        r = self.post((af_name(CODE_A), af_bytes(PEOPLE_A)), roster_bytes(PEOPLE_B), SCHOOLS)
        d = r.get_json()
        self.assertIs(d['overlap']['same_school'], False)
        body = r.get_data(as_text=True)
        for n, i in PEOPLE_A + PEOPLE_B:      # 不回傳任何姓名或身分證
            self.assertNotIn(n, body)
            self.assertNotIn(i, body)

    def test_school_not_in_account(self):
        d = self.post((af_name('397084500Y'), af_bytes(PEOPLE_A)), schools=SCHOOLS).get_json()
        self.assertEqual(d['code'], '397084500Y')
        self.assertEqual(d['account_school_id'], '')

    def test_missing_file_and_bad_schools_json(self):
        self.assertEqual(self.post().status_code, 400)
        r = self.client.post('/af/identify', data={
            'af': (io.BytesIO(af_bytes(PEOPLE_A)), af_name(CODE_A)), 'schools': '{bad'},
            content_type='multipart/form-data')
        self.assertEqual(r.status_code, 200)
        self.assertEqual(r.get_json()['account_school_id'], '')

    def test_does_not_write_memory_files(self):
        self.post((af_name(CODE_A), af_bytes(PEOPLE_A)), roster_bytes(PEOPLE_A), SCHOOLS)
        self.assertEqual(os.listdir(self.tmp), [])

    def test_page_loads_school_script(self):
        page = self.client.get('/').get_data(as_text=True)
        self.assertIn('/static/af_school.js?v=v1.3.23', page)


if __name__ == '__main__':
    unittest.main()
