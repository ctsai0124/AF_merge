import unittest

import paycheck


def af_person(name='范小茜'):
    return {
        '姓名': name,
        '身分證': '',
        '薪俸': 29270,
        '專業加給': 26560,
        '主管加給': 0,
        '導師特教': 6800,
    }


def inferred_person(name='范小茜'):
    return {
        '姓名': name,
        '職稱': '資源班',
        '薪俸': 29270,
        '專業加給': 26560,
        '主管加給': 0,
        '導師特教': 6800,
        '其他加給': 0,
        '應發金額': 62630,
        '應發金額推算': True,
        '最低信心': 0.5,
    }


class InferredGrossTests(unittest.TestCase):
    def test_exact_af_fields_allow_missing_printed_gross(self):
        good, need = paycheck.from_ocr(
            [inferred_person()], [af_person()])
        self.assertEqual(len(good), 1)
        self.assertEqual(need, [])
        self.assertEqual(good[0]['應發金額'], 62630)

    def test_af_amount_difference_still_requires_review(self):
        af = af_person()
        af['導師特教'] = 4000
        good, need = paycheck.from_ocr([inferred_person()], [af])
        self.assertEqual(good, [])
        self.assertEqual(len(need), 1)
        self.assertTrue(need[0]['金額可疑'])

    def test_two_matching_name_chars_and_unique_af_fields_can_approve(self):
        good, need = paycheck.from_ocr(
            [inferred_person('范小菁')], [af_person()])
        self.assertEqual(len(good), 1)
        self.assertEqual(good[0]['姓名'], '范小茜')
        self.assertEqual(need, [])

    def test_ambiguous_af_fingerprint_still_requires_review(self):
        duplicate = af_person('范小菁')
        good, need = paycheck.from_ocr(
            [inferred_person('范小菁')], [af_person(), duplicate])
        self.assertEqual(good, [])
        self.assertEqual(len(need), 1)
        self.assertTrue(need[0]['金額可疑'])

    def test_unverifiable_other_allowance_requires_review(self):
        person = inferred_person()
        person['其他加給'] = 4000
        person['應發金額'] = 66630
        good, need = paycheck.from_ocr([person], [af_person()])
        self.assertEqual(good, [])
        self.assertEqual(len(need), 1)
        self.assertTrue(need[0]['金額可疑'])



class OcrPlausibleTests(unittest.TestCase):
    def test_single_char_errors(self):
        P = paycheck.ocr_plausible
        self.assertTrue(P(35180, 35780))     # 7→1
        self.assertTrue(P(55890, 55690))     # 6→8
        self.assertTrue(P(53690, 55690))     # 5→3 單字替換
        self.assertTrue(P(59630, 59360))     # 相鄰互換
        self.assertTrue(P(559301, 55930))    # 尾端多讀 1
        self.assertTrue(P(5930, 59300))      # 漏一位
        self.assertFalse(P(55690, 55690))
        self.assertFalse(P(45690, 55600))    # 兩位不同
        self.assertFalse(P(5930, 35780))     # 完全不同
        self.assertFalse(P(100, 0))


class AfProofCorrectionTests(unittest.TestCase):
    """AF 正確時兩邊數字完全相同；OCR 單字誤差須有印列應發佐證才自動校正（虛構資料）。"""

    AF = [{'姓名': '王小明', '身分證': 'X', '薪俸': 55690, '專業加給': 35780,
           '主管加給': 5930, '導師特教': 0}]

    def ocr(self, **kw):
        p = {'姓名': '王小明', '職稱': '教師', '身分證': '', '身分證有效': False,
             '薪俸': 55690, '專業加給': 35780, '主管加給': 5930, '導師特教': 0,
             '其他加給': 0, '應發金額': 97400, '應發金額推算': False,
             '最低信心': 1.0}
        p.update(kw)
        return p

    def test_digit_misread_with_total_proof_is_auto_corrected(self):
        good, need = paycheck.from_ocr([self.ocr(專業加給=35180)], self.AF)
        self.assertEqual(need, [])
        self.assertEqual(good[0]['專業加給'], 35780)
        self.assertTrue(good[0]['_ocr_corrected'])
        out = paycheck.compare(good, self.AF)
        self.assertEqual(out['summary']['一致'], 1)
        self.assertEqual(out['summary']['OCR自動校正'], 1)

    def test_missed_field_with_total_proof_is_filled(self):
        good, need = paycheck.from_ocr([self.ocr(主管加給=0)], self.AF)
        self.assertEqual(need, [])
        self.assertEqual(good[0]['主管加給'], 5930)

    def test_two_missed_fields_stay_in_review(self):
        good, need = paycheck.from_ocr([self.ocr(薪俸=0, 專業加給=0)], self.AF)
        self.assertEqual(good, [])

    def test_total_misread_when_fields_equal_af(self):
        good, need = paycheck.from_ocr([self.ocr(應發金額=97100)], self.AF)
        self.assertEqual(need, [])
        self.assertEqual(good[0]['應發金額'], 97400)

    def test_real_difference_stays_in_review(self):
        # 清冊真的少發職務加給：印列應發與 AF 代入後不符，不可校正
        good, need = paycheck.from_ocr(
            [self.ocr(主管加給=0, 應發金額=91470)], self.AF)
        out = paycheck.compare(good, self.AF)
        self.assertEqual(out['summary']['有差異'], 1)
        self.assertEqual(out['summary']['OCR自動校正'], 0)
        # 清冊自身不平衡又與 AF 差距大：維持人工確認
        good, need = paycheck.from_ocr(
            [self.ocr(主管加給=0, 應發金額=95000)], self.AF)
        self.assertEqual(good, [])
        self.assertEqual(len(need), 1)

    def test_large_difference_stays_in_review(self):
        good, need = paycheck.from_ocr(
            [self.ocr(專業加給=26830, 應發金額=97400)], self.AF)
        self.assertEqual(good, [])

    def test_no_printed_total_stays_in_review(self):
        good, need = paycheck.from_ocr(
            [self.ocr(專業加給=35180, 應發金額=0)], self.AF)
        self.assertEqual(good, [])
        good, need = paycheck.from_ocr(
            [self.ocr(專業加給=35180, 應發金額=97400, 應發金額推算=True)], self.AF)
        self.assertEqual(good, [])

    def test_af_typo_is_not_hidden(self):
        # AF 打錯一位數：清冊讀值正確且自洽，必須呈現為差異而非被 AF 覆蓋
        af = [dict(self.AF[0], 專業加給=35180)]
        good, need = paycheck.from_ocr([self.ocr()], af)
        out = paycheck.compare(good, af)
        self.assertEqual(out['summary']['有差異'], 1)
        self.assertEqual(out['summary']['OCR自動校正'], 0)

    def test_name_misread_plus_digit_misread(self):
        good, need = paycheck.from_ocr(
            [self.ocr(姓名='王小朋', 薪俸=55890)], self.AF)
        self.assertEqual(need, [])
        self.assertEqual(good[0]['姓名'], '王小明')

if __name__ == '__main__':
    unittest.main()


def _vtable(header_labels, people_cols, data_rows):
    """輔助函式：組出 pdfplumber.extract_tables() 風格的表格結構
    （list of rows, each row a list of cell string）。
    header_labels: 表頭第一欄印的標籤（用 \\n 分隔的多行文字）。
    people_cols: 每個人的表頭欄位內容，list of 字串（\\n 分隔多行）。
    data_rows: [(row_label, [每人數值...]), ...]。
    """
    header = [header_labels] + list(people_cols)
    rows = [header]
    for label, values in data_rows:
        rows.append([label] + list(values))
    return rows


class ParseVerticalRealWorldFormatTests(unittest.TestCase):
    """涵蓋 2026-09 修正的兩個真實 bug：五行表頭（含身分證字號）的
    欄位錯位、以及多子表各自從 1 重新編號時的跨表撞號覆蓋問題。
    這兩個 bug 先前只用 /tmp 底下的一次性腳本驗證過，沒有留在正式
    測試套件裡，這裡補上永久回歸測試。"""

    def test_five_line_header_extracts_id_into_correct_field(self):
        tbl = _vtable(
            '序號\n身分證字號\n姓名\n職稱\n支薪俸級',
            ['1\nA123456789\n王小明\n教師\n300',
             '2\nB223456789\n林小美\n教師\n310'],
            [('薪俸', ['30000', '32000']),
             ('應發數合計', ['35000', '37000'])],
        )
        people = paycheck._parse_vertical([tbl])
        self.assertEqual(len(people), 2)
        self.assertEqual(people[0]['身分證'], 'A123456789')
        self.assertEqual(people[0]['姓名'], '王小明')
        self.assertEqual(people[0]['職稱'], '教師')
        self.assertEqual(people[1]['身分證'], 'B223456789')
        self.assertEqual(people[1]['姓名'], '林小美')

    def test_three_line_header_still_works_without_id(self):
        tbl = _vtable(
            '序號\n姓名\n職稱',
            ['1\n王小明\n教師', '2\n林小美\n教師'],
            [('薪俸', ['30000', '32000'])],
        )
        people = paycheck._parse_vertical([tbl])
        self.assertEqual(len(people), 2)
        self.assertEqual(people[0]['身分證'], '')
        self.assertEqual(people[0]['姓名'], '王小明')

    def test_cross_table_numbering_does_not_collide(self):
        """兩張子表都各自從序號1開始編號，修正前會互相覆蓋、掉資料。"""
        tbl1 = _vtable(
            '序號\n身分證字號\n姓名\n職稱\n支薪俸級',
            ['1\nA123456789\n王小明\n教師\n300'],
            [('薪俸', ['30000'])],
        )
        tbl2 = _vtable(
            '序號\n身分證字號\n姓名\n職稱\n支薪俸級',
            ['1\nC345678901\n陳大同\n幹事\n200'],
            [('薪俸', ['28000'])],
        )
        people = paycheck._parse_vertical([tbl1, tbl2])
        self.assertEqual(len(people), 2)
        names = {p['姓名'] for p in people}
        self.assertEqual(names, {'王小明', '陳大同'})
        ids = {p['身分證'] for p in people}
        self.assertEqual(ids, {'A123456789', 'C345678901'})


class ParsePdfSanityCheckTests(unittest.TestCase):
    """Leo 問「換一份完全陌生格式的清冊會怎樣」之後補的合理性檢查：
    偵測「姓名/職稱欄位長得像身分證字號」（欄位錯位跡象）跟「金額
    欄位全部是 0」（項目名稱對不到跡象），格式異常時要回傳警告文字，
    不能靜默產出一份看似正常、其實錯位或抓不到金額的資料。"""

    def test_normal_five_line_pdf_has_no_sanity_warning(self):
        tbl = _vtable(
            '序號\n身分證字號\n姓名\n職稱\n支薪俸級',
            ['1\nC123456789\n王小明\n教師\n300'],
            [('薪俸', ['30000']), ('應發數合計', ['35000'])],
        )
        people = paycheck._parse_vertical([tbl])
        self.assertEqual(paycheck._sanity_check(people), [])

    def test_id_shaped_name_field_triggers_warning(self):
        """模擬陌生表頭順序（序號/姓名/身分證字號，姓名剛好還留在對的
        位置，但身分證字號被錯塞進職稱欄）——職稱長得像身分證字號。"""
        tbl = _vtable(
            '序號\n姓名\n身分證字號',
            ['1\n王小明\nA123456789'],
            [('薪俸', ['30000'])],
        )
        people = paycheck._parse_vertical([tbl])
        self.assertEqual(people[0]['職稱'], 'A123456789')
        warnings = paycheck._sanity_check(people)
        self.assertTrue(warnings)
        self.assertIn('職稱', warnings[0])
        self.assertIn('身分證字號', warnings[0])

    def test_unknown_row_labels_zero_out_amounts_and_trigger_warning(self):
        """模擬項目名稱文字跟 FIELD_ALIAS 對不到的陌生格式，金額欄位
        會全部變 0，必須被合理性檢查抓出來，不能悄悄當成「這個人薪水
        是0元」就算了。"""
        tbl = _vtable(
            '序號\n身分證字號\n姓名\n職稱\n支薪俸級',
            ['1\nC123456789\n王小明\n教師\n300'],
            [('基本薪給', ['30000']), ('應發合計', ['35000'])],
        )
        people = paycheck._parse_vertical([tbl])
        self.assertEqual(people[0]['薪俸'], 0)
        warnings = paycheck._sanity_check(people)
        self.assertTrue(warnings)
        self.assertIn('0', warnings[0])

    def test_empty_people_list_has_no_warning(self):
        self.assertEqual(paycheck._sanity_check([]), [])


class LooksLikeIdTests(unittest.TestCase):
    def test_valid_shape(self):
        self.assertTrue(paycheck._looks_like_id('A123456789'))
        self.assertTrue(paycheck._looks_like_id('a123456789'))

    def test_invalid_shape(self):
        self.assertFalse(paycheck._looks_like_id('王小明'))
        self.assertFalse(paycheck._looks_like_id('A12345678'))  # 少一碼
        self.assertFalse(paycheck._looks_like_id(''))
        self.assertFalse(paycheck._looks_like_id(None))


class FromOcrClaimedNameTests(unittest.TestCase):
    """模糊校正不可把清冊其他人搶到已被本人精確占用的 AF 姓名（虛構資料）。"""

    def ocr(self, name, title, pay):
        return {'姓名': name, '職稱': title, '身分證': '', '身分證有效': False,
                '薪俸': pay, '專業加給': 30000, '主管加給': 0, '導師特教': 0,
                '其他加給': 0, '應發金額': pay + 30000, '應發金額推算': False,
                '加總相符': True, '最低信心': 1.0}

    def af(self, name, pay):
        return {'姓名': name, '身分證': 'X', '薪俸': pay, '專業加給': 30000,
                '主管加給': 0, '導師特教': 0}

    def test_close_name_already_matched_exactly_goes_to_review(self):
        people = [self.ocr('林甲乙', '教務主任', 50000),
                  self.ocr('林乙丙', '代理教師', 40000)]
        good, need = paycheck.from_ocr(people, [self.af('林甲乙', 50000)])
        self.assertEqual([g['姓名'] for g in good], ['林甲乙'])
        self.assertEqual(len(need), 1)
        self.assertEqual(need[0]['原始姓名'], '林乙丙')
        self.assertEqual(need[0]['建議姓名'], '')
        out = paycheck.compare(good, [self.af('林甲乙', 50000)])
        self.assertEqual(out['summary']['一致'], 1)

    def test_unclaimed_close_name_is_still_corrected(self):
        good, need = paycheck.from_ocr([self.ocr('林甲丁', '教師', 50000)],
                                       [self.af('林甲乙', 50000)])
        self.assertEqual([g['姓名'] for g in good], ['林甲乙'])
        self.assertEqual(need, [])

    def test_two_rows_cannot_both_claim_one_af_name(self):
        people = [self.ocr('林甲丁', '教師', 50000),
                  self.ocr('林甲戊', '教師', 50000)]
        good, need = paycheck.from_ocr(people, [self.af('林甲乙', 50000)])
        self.assertEqual([g['姓名'] for g in good], ['林甲乙'])
        self.assertEqual([n['原始姓名'] for n in need], ['林甲戊'])
