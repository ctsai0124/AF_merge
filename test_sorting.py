import io
import os
import tempfile
import unittest
from unittest.mock import patch

import openpyxl
import pandas as pd

from app import app, read_sheet, sort_af_by_roster


def af_rows(rows):
    return pd.DataFrame(rows, columns=['姓名', '身分證字號', '薪俸表別'])


def workbook_bytes(sheet_name, rows):
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = sheet_name
    headers = list(rows[0])
    sheet.append(headers)
    for row in rows:
        sheet.append([row.get(header, '') for header in headers])
    payload = io.BytesIO()
    workbook.save(payload)
    return payload.getvalue()


class SortAfByRosterTests(unittest.TestCase):
    def test_common_sequence_column_aliases_are_accepted(self):
        af = af_rows([
            {'姓名': '王小明', '身分證字號': 'A111111111', '薪俸表別': 'A'},
        ])

        for alias in ('編號', '號碼', '項次', '流水號', '序次', '次序', '排序', '排序號', 'No.', 'Number', '序 號'):
            with self.subTest(alias=alias):
                roster = pd.DataFrame([{alias: 1, '姓名': '王小明'}])
                result, warnings = sort_af_by_roster(roster, af)
                self.assertEqual(result.loc[0, '清冊序號'], 1)
                self.assertEqual(warnings, [])

    def test_unique_names_still_match_by_name_without_roster_ids(self):
        roster = pd.DataFrame([
            {'序號': 2, '姓名': '李小華'},
            {'序號': 1, '姓名': '王小明'},
        ])
        af = af_rows([
            {'姓名': '李小華', '身分證字號': 'B222222222', '薪俸表別': 'B'},
            {'姓名': '王小明', '身分證字號': 'A111111111', '薪俸表別': 'A'},
        ])

        result, warnings = sort_af_by_roster(roster, af)

        self.assertEqual(result['姓名'].tolist(), ['王小明', '李小華'])
        self.assertEqual(result['清冊序號'].tolist(), [1, 2])
        self.assertEqual(warnings, [])

    def test_same_names_match_by_id_without_duplicate_rows(self):
        roster = pd.DataFrame([
            {'序號': 1, '姓名': '王同名', '身分證字號': 'A111111111'},
            {'序號': 2, '姓名': '王同名', '身分證字號': 'B222222222'},
        ])
        af = af_rows([
            {'姓名': '王同名', '身分證字號': 'B222222222', '薪俸表別': 'B'},
            {'姓名': '王同名', '身分證字號': 'A111111111', '薪俸表別': 'A1'},
            {'姓名': '王同名', '身分證字號': 'A111111111', '薪俸表別': 'A2'},
        ])

        result, warnings = sort_af_by_roster(roster, af)

        self.assertEqual(len(result), 3)
        self.assertEqual(result['身分證字號'].tolist(), [
            'A111111111', 'A111111111', 'B222222222',
        ])
        self.assertEqual(result['清冊序號'].tolist(), [1, 1, 2])
        self.assertIn('同名同姓人員已依身分證字號正確配對：王同名', warnings)

    def test_same_names_require_roster_id_values(self):
        roster = pd.DataFrame([
            {'序號': 1, '姓名': '王同名', '身分證字號': 'A111111111'},
            {'序號': 2, '姓名': '王同名', '身分證字號': ''},
        ])
        af = af_rows([
            {'姓名': '王同名', '身分證字號': 'A111111111', '薪俸表別': 'A'},
            {'姓名': '王同名', '身分證字號': 'B222222222', '薪俸表別': 'B'},
        ])

        with self.assertRaisesRegex(ValueError, '固定清冊缺少身分證字號'):
            sort_af_by_roster(roster, af)

    def test_same_names_require_af_id_values(self):
        roster = pd.DataFrame([
            {'序號': 1, '姓名': '王同名', '身分證字號': 'A111111111'},
            {'序號': 2, '姓名': '王同名', '身分證字號': 'B222222222'},
        ])
        af = af_rows([
            {'姓名': '王同名', '身分證字號': 'A111111111', '薪俸表別': 'A'},
            {'姓名': '王同名', '身分證字號': '', '薪俸表別': 'B'},
        ])

        with self.assertRaisesRegex(ValueError, 'AF 資料缺少身分證字號'):
            sort_af_by_roster(roster, af)

    def test_roster_basename_does_not_affect_reading(self):
        workbook = openpyxl.Workbook()
        sheet = workbook.active
        sheet.title = 'input'
        sheet.append(['序號', '姓名', '身分證字號'])
        sheet.append([1, '王小明', 'A111111111'])
        payload = io.BytesIO()
        workbook.save(payload)

        result = read_sheet(
            payload.getvalue(),
            '學校基本資料_檔名可以自行修改.xlsx',
            'input',
        )

        self.assertEqual(result.loc[0, '姓名'], '王小明')

    def test_missing_input_sheet_has_clear_chinese_error(self):
        workbook = openpyxl.Workbook()
        workbook.active.title = '基本資料'
        workbook.active.append(['序號', '姓名'])
        workbook.active.append([1, '王小明'])
        payload = io.BytesIO()
        workbook.save(payload)

        with self.assertRaisesRegex(ValueError, '工作表分頁必須命名為 input'):
            read_sheet(payload.getvalue(), '任何檔名.xlsx', 'input')

    def test_missing_input_sheet_can_fallback_to_first_sheet(self):
        roster_bytes = workbook_bytes('基本資料', [
            {'序號': 1, '姓名': '王小明'},
        ])

        result = read_sheet(
            roster_bytes,
            '任何檔名.xlsx',
            'input',
            fallback_to_first=True,
        )

        self.assertEqual(result.loc[0, '姓名'], '王小明')
        self.assertEqual(result.attrs['fallback_sheet'], '基本資料')

    def test_process_endpoint_accepts_arbitrary_roster_filename(self):
        roster_bytes = workbook_bytes('input', [
            {'序號': 1, '姓名': '王同名', '身分證字號': 'A111111111'},
            {'序號': 2, '姓名': '王同名', '身分證字號': 'B222222222'},
        ])
        af_bytes = workbook_bytes('AF', [
            {'姓名': '王同名', '身分證字號': 'B222222222', '薪俸表別': 'B'},
            {'姓名': '王同名', '身分證字號': 'A111111111', '薪俸表別': 'A'},
        ])

        with tempfile.TemporaryDirectory() as data_dir, patch.dict(
            os.environ, {'DATA_DIR': data_dir}
        ):
            response = app.test_client().post('/process', data={
                'roster': (io.BytesIO(roster_bytes), '校內自訂基本資料檔名.xlsx'),
                'af': (io.BytesIO(af_bytes), 'AF_自訂檔名.xlsx'),
            })

        self.assertEqual(response.status_code, 200)
        payload = response.get_json()
        self.assertTrue(payload['success'])
        self.assertEqual(
            [row['身分證字號'] for row in payload['preview']],
            ['A111111111', 'B222222222'],
        )

    def test_process_endpoint_falls_back_and_reports_first_sheet(self):
        roster_bytes = workbook_bytes('學校名冊', [
            {'序號': 1, '姓名': '王小明'},
        ])
        af_bytes = workbook_bytes('AF', [
            {'姓名': '王小明', '身分證字號': 'A111111111', '薪俸表別': 'A'},
        ])

        with tempfile.TemporaryDirectory() as data_dir, patch.dict(
            os.environ, {'DATA_DIR': data_dir}
        ):
            response = app.test_client().post('/process', data={
                'roster': (io.BytesIO(roster_bytes), '基本資料.xlsx'),
                'af': (io.BytesIO(af_bytes), 'AF_自訂檔名.xlsx'),
            })

        self.assertEqual(response.status_code, 200)
        payload = response.get_json()
        self.assertIn(
            '固定清冊找不到 input 工作表，已自動改讀第一個工作表「學校名冊」',
            payload['warnings'],
        )

    def test_process_endpoint_accepts_sequence_alias_on_first_sheet(self):
        roster_bytes = workbook_bytes('基本資料', [
            {'編號': 1, '姓名': '王小明'},
        ])
        af_bytes = workbook_bytes('AF', [
            {'姓名': '王小明', '身分證字號': 'A111111111', '薪俸表別': 'A'},
        ])

        with tempfile.TemporaryDirectory() as data_dir, patch.dict(
            os.environ, {'DATA_DIR': data_dir}
        ):
            response = app.test_client().post('/process', data={
                'roster': (io.BytesIO(roster_bytes), '學校自己的基本資料.xlsx'),
                'af': (io.BytesIO(af_bytes), 'AF_自訂檔名.xlsx'),
            })

        self.assertEqual(response.status_code, 200)
        payload = response.get_json()
        self.assertTrue(payload['success'])
        self.assertEqual(payload['preview'][0]['清冊序號'], 1)


if __name__ == '__main__':
    unittest.main()
