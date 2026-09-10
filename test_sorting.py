import io
import os
import tempfile
import unittest
from unittest.mock import patch

import openpyxl
import pandas as pd

from app import app, read_sheet, sort_af_by_roster, build_excel


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

    def test_optional_roster_people_category_is_carried_to_result(self):
        roster = pd.DataFrame([{
            '序號': 1, '姓名': '王小明', '人員種類': '政務人員',
        }])
        af = af_rows([{
            '姓名': '王小明', '身分證字號': 'A111111111', '薪俸表別': 'A0001',
        }])

        result, warnings = sort_af_by_roster(roster, af)

        self.assertEqual(result.loc[0, '人員種類'], '政務人員')
        self.assertEqual(warnings, [])

    def test_invalid_roster_people_category_has_clear_error(self):
        roster = pd.DataFrame([{
            '序號': 1, '姓名': '王小明', '人員種類': '主任',
        }])
        af = af_rows([{
            '姓名': '王小明', '身分證字號': 'A111111111', '薪俸表別': 'A0001',
        }])

        # 舊版曾使用細職稱，仍會相容歸入大類，不中斷使用。
        result, _warnings = sort_af_by_roster(roster, af)
        self.assertEqual(result.loc[0, '人員種類'], '一般人員')

        roster.loc[0, '人員種類'] = '無此分類'
        with self.assertRaisesRegex(ValueError, '人員種類僅可填'):
            sort_af_by_roster(roster, af)

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
            [row['姓名'] for row in payload['preview']],
            ['王同名', '王同名'],
        )
        self.assertNotIn('身分證字號', payload['columns'])
        self.assertNotIn('身分證字號', payload['preview'][0])

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

    def test_fullwidth_sequence_digits_are_normalized(self):
        roster = pd.DataFrame([{'序號': '１', '姓名': '王小明'}])
        af = af_rows([{'姓名': '王小明', '身分證字號': 'A111111111', '薪俸表別': 'A'}])

        result, warnings = sort_af_by_roster(roster, af)

        self.assertEqual(result.loc[0, '清冊序號'], 1)
        self.assertEqual(warnings, [])

    def test_duplicate_sequence_numbers_trigger_warning(self):
        roster = pd.DataFrame([
            {'序號': 1, '姓名': '王小明'},
            {'序號': 1, '姓名': '陳大華'},
        ])
        af = af_rows([
            {'姓名': '王小明', '身分證字號': 'A111111111', '薪俸表別': 'A'},
            {'姓名': '陳大華', '身分證字號': 'B222222222', '薪俸表別': 'B'},
        ])

        result, warnings = sort_af_by_roster(roster, af)

        self.assertTrue(any('序號重複' in w and '1' in w for w in warnings))

    def test_build_excel_strips_comma_formatted_amounts(self):
        data = [{'清冊序號': 1, '姓名': '王小明', '總金額': '50,000'}]
        out = build_excel(data, ['清冊序號', '姓名', '總金額'])
        wb = openpyxl.load_workbook(out)
        ws = wb.active

        self.assertEqual(ws.cell(row=2, column=3).value, 50000)

    def test_build_excel_writes_formula_like_text_as_literal_text(self):
        values = ['=1+1', '+2+2', '-3+3', '@SUM(1,1)']
        data = [
            {'清冊序號': index, '姓名': value}
            for index, value in enumerate(values, 1)
        ]

        out = build_excel(data, ['清冊序號', '姓名'])
        wb = openpyxl.load_workbook(out, data_only=False)
        ws = wb.active

        for row, value in enumerate(values, 2):
            with self.subTest(value=value):
                cell = ws.cell(row=row, column=2)
                self.assertEqual(cell.value, "'" + value)
                self.assertEqual(cell.data_type, 's')

    def test_process_endpoint_rejects_corrupt_file_with_friendly_message(self):
        with tempfile.TemporaryDirectory() as data_dir, patch.dict(
            os.environ, {'DATA_DIR': data_dir}
        ):
            response = app.test_client().post('/process', data={
                'roster': (io.BytesIO(b'not an excel file'), 'roster.xlsx'),
                'af': (io.BytesIO(b'not an excel file'), 'af.xlsx'),
            })

        self.assertEqual(response.status_code, 400)
        payload = response.get_json()
        self.assertIn('檔案格式錯誤', payload['error'])
        self.assertNotIn('zip file', payload['error'])

    def test_download_results_do_not_leak_between_users(self):
        """兩個不同使用者（各自獨立的 session）前後處理不同資料時，
        各自下載到的必須是自己的結果，不會被對方蓋掉。"""
        roster_a = workbook_bytes('input', [{'序號': 1, '姓名': '王小明'}])
        af_a = workbook_bytes('AF', [
            {'姓名': '王小明', '身分證字號': 'A111111111', '薪俸表別': 'A'},
        ])
        roster_b = workbook_bytes('input', [{'序號': 1, '姓名': '陳大華'}])
        af_b = workbook_bytes('AF', [
            {'姓名': '陳大華', '身分證字號': 'B222222222', '薪俸表別': 'B'},
        ])

        with tempfile.TemporaryDirectory() as data_dir, patch.dict(
            os.environ, {'DATA_DIR': data_dir}
        ):
            client_a = app.test_client()
            client_b = app.test_client()

            resp_a = client_a.post('/process', data={
                'roster': (io.BytesIO(roster_a), 'roster_a.xlsx'),
                'af': (io.BytesIO(af_a), 'af_a.xlsx'),
            })
            self.assertTrue(resp_a.get_json()['success'])

            # 模擬第二位使用者在第一位使用者下載結果之前，也上傳處理了另一份資料。
            resp_b = client_b.post('/process', data={
                'roster': (io.BytesIO(roster_b), 'roster_b.xlsx'),
                'af': (io.BytesIO(af_b), 'af_b.xlsx'),
            })
            self.assertTrue(resp_b.get_json()['success'])

            download_a = client_a.get('/download-result')
            wb_a = openpyxl.load_workbook(io.BytesIO(download_a.data))
            names_a = [row[1] for row in wb_a.active.iter_rows(min_row=2, values_only=True)]

            download_b = client_b.get('/download-result')
            wb_b = openpyxl.load_workbook(io.BytesIO(download_b.data))
            names_b = [row[1] for row in wb_b.active.iter_rows(min_row=2, values_only=True)]

        self.assertEqual(names_a, ['王小明'])
        self.assertEqual(names_b, ['陳大華'])

    def test_process_endpoint_escapes_names_in_preview_html_paths(self):
        """後端不需要對 JSON preview 裡的姓名做跳脫（前端已改用 textContent），
        但這裡先確認惡意內容確實會原封不動地流到回應裡——用來佐證前端
        必須自行做安全處理，同時避免以後有人不小心在後端 HTML 頁面
        又直接把這欄位塞進字串。"""
        payload_name = '<img src=x onerror=alert(1)>'
        roster_bytes = workbook_bytes('input', [{'序號': 1, '姓名': payload_name}])
        af_bytes = workbook_bytes('AF', [
            {'姓名': payload_name, '身分證字號': 'A111111111', '薪俸表別': 'A'},
        ])

        with tempfile.TemporaryDirectory() as data_dir, patch.dict(
            os.environ, {'DATA_DIR': data_dir}
        ):
            response = app.test_client().post('/process', data={
                'roster': (io.BytesIO(roster_bytes), 'roster.xlsx'),
                'af': (io.BytesIO(af_bytes), 'af.xlsx'),
            })

        payload = response.get_json()
        self.assertEqual(payload['preview'][0]['姓名'], payload_name)

    def test_print_simple_escapes_malicious_names(self):
        roster_bytes = workbook_bytes('input', [
            {'序號': 1, '姓名': '<img src=x onerror=alert(1)>'},
        ])
        af_bytes = workbook_bytes('AF', [
            {'姓名': '<img src=x onerror=alert(1)>', '身分證字號': 'A111111111',
             '薪俸表別': 'A'},
        ])

        with tempfile.TemporaryDirectory() as data_dir, patch.dict(
            os.environ, {'DATA_DIR': data_dir}
        ):
            client = app.test_client()
            resp = client.post('/process', data={
                'roster': (io.BytesIO(roster_bytes), 'roster.xlsx'),
                'af': (io.BytesIO(af_bytes), 'af.xlsx'),
            })
            self.assertTrue(resp.get_json()['success'])

            html_resp = client.get('/print-simple')

        body = html_resp.get_data(as_text=True)
        self.assertNotIn('<img src=x onerror=alert(1)>', body)
        self.assertIn('&lt;img src=x onerror=alert(1)&gt;', body)

    def test_results_cache_evicts_previous_result_for_same_session(self):
        roster_1 = workbook_bytes('input', [{'序號': 1, '姓名': '王小明'}])
        af_1 = workbook_bytes('AF', [
            {'姓名': '王小明', '身分證字號': 'A111111111', '薪俸表別': 'A'},
        ])
        roster_2 = workbook_bytes('input', [{'序號': 1, '姓名': '陳大華'}])
        af_2 = workbook_bytes('AF', [
            {'姓名': '陳大華', '身分證字號': 'B222222222', '薪俸表別': 'B'},
        ])

        import app as app_module

        with tempfile.TemporaryDirectory() as data_dir, patch.dict(
            os.environ, {'DATA_DIR': data_dir}
        ):
            client = app_module.app.test_client()
            client.post('/process', data={
                'roster': (io.BytesIO(roster_1), 'roster1.xlsx'),
                'af': (io.BytesIO(af_1), 'af1.xlsx'),
            })
            before = len(app_module._results)

            client.post('/process', data={
                'roster': (io.BytesIO(roster_2), 'roster2.xlsx'),
                'af': (io.BytesIO(af_2), 'af2.xlsx'),
            })
            after = len(app_module._results)

        # 同一個 session 再次處理，暫存筆數不應該累加（舊的要被清掉）。
        self.assertEqual(before, after)

    def test_results_cache_evicts_oldest_result_when_byte_limit_is_reached(self):
        import app as app_module

        result = pd.DataFrame([{'清冊序號': 1, '姓名': '測試人員' * 20}])
        result_size = len(
            result.fillna('').to_json(orient='records', force_ascii=False).encode('utf-8')
        )

        with patch.object(app_module, 'MAX_RESULTS', 100), patch.object(
            app_module, 'MAX_RESULT_CACHE_BYTES', result_size * 2 - 1
        ), patch.object(app_module, 'MAX_SINGLE_RESULT_BYTES', result_size * 2 - 1):
            with app_module._results_lock:
                app_module._results.clear()
            try:
                with app_module.app.test_request_context('/'):
                    app_module.save_result(result, '', '', '')
                    first_rid = app_module.session['result_id']

                with app_module.app.test_request_context('/'):
                    app_module.save_result(result, '', '', '')
                    second_rid = app_module.session['result_id']

                self.assertNotIn(first_rid, app_module._results)
                self.assertIn(second_rid, app_module._results)
            finally:
                with app_module._results_lock:
                    app_module._results.clear()

    def test_single_oversized_result_is_rejected_before_caching(self):
        import app as app_module

        result = pd.DataFrame([{'清冊序號': 1, '姓名': '資料量測試'}])
        with app_module.app.test_request_context('/'), patch.object(
            app_module, 'MAX_SINGLE_RESULT_BYTES', 1
        ):
            with self.assertRaisesRegex(ValueError, '排序結果資料量過大'):
                app_module.save_result(result, '', '', '')

    def test_decimal_sequence_number_is_rejected_instead_of_truncated(self):
        roster = pd.DataFrame([
            {'序號': 1.5, '姓名': '王小明'},
            {'序號': 2, '姓名': '陳大華'},
        ])
        af = af_rows([
            {'姓名': '王小明', '身分證字號': 'A111111111', '薪俸表別': 'A'},
            {'姓名': '陳大華', '身分證字號': 'B222222222', '薪俸表別': 'B'},
        ])

        with self.assertRaisesRegex(ValueError, '序號必須是正整數'):
            sort_af_by_roster(roster, af)

    def test_non_numeric_sequence_number_is_rejected_instead_of_dropped(self):
        roster = pd.DataFrame([
            {'序號': 1, '姓名': '王小明'},
            {'序號': '1O', '姓名': '李小華'},
        ])
        af = af_rows([
            {'姓名': '王小明', '身分證字號': 'A111111111', '薪俸表別': 'A'},
            {'姓名': '李小華', '身分證字號': 'B222222222', '薪俸表別': 'B'},
        ])

        with self.assertRaisesRegex(ValueError, '李小華（序號 1O）'):
            sort_af_by_roster(roster, af)

    def test_blank_sequence_with_name_is_rejected(self):
        roster = pd.DataFrame([{'序號': '', '姓名': '王小明'}])
        af = af_rows([
            {'姓名': '王小明', '身分證字號': 'A111111111', '薪俸表別': 'A'},
        ])

        with self.assertRaisesRegex(ValueError, '王小明（序號 空白）'):
            sort_af_by_roster(roster, af)

    def test_300_row_roster_stays_within_normal_processing_range(self):
        roster = pd.DataFrame([
            {'序號': index, '姓名': f'測試人員{index:03d}'}
            for index in range(1, 301)
        ])
        af = af_rows([
            {
                '姓名': f'測試人員{index:03d}',
                '身分證字號': f'T{index:09d}',
                '薪俸表別': 'A',
            }
            for index in range(1, 301)
        ])

        result, warnings = sort_af_by_roster(roster, af)

        self.assertEqual(len(result), 300)
        self.assertEqual(result.iloc[-1]['清冊序號'], 300)
        self.assertEqual(warnings, [])


if __name__ == '__main__':
    unittest.main()
