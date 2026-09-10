import unittest

import parse_tokens


def token(text, x, y=0.2, page=0, conf=0.9):
    return {
        'text': text,
        'x': x,
        'y': y,
        'w': 0.03,
        'conf': conf,
        'page': page,
    }


class HorizontalParserTests(unittest.TestCase):
    def parse(self, *items):
        return parse_tokens.parse_horizontal_row(
            [token(text, x) for text, x in items])

    def test_role_like_name_and_optional_allowances(self):
        person = self.parse(
            ('資源班', .03), ('650', .07), ('裴淑茵', .11),
            ('55690', .16), ('2800', .20), ('35780', .24),
            ('4000', .28), ('98270', .37), ('2614', .42))
        self.assertEqual(person['姓名'], '裴淑茵')
        self.assertEqual(person['主管加給'], 0)
        self.assertEqual(person['專業加給'], 35780)
        self.assertEqual(person['導師特教'], 6800)
        self.assertTrue(person['加總相符'])

    def fixture(self):
        return [
            token('單位', .04, .10), token('職稱', .08, .10),
            token('月支薪額', .17, .10), token('專業加給', .22, .10),
            token('主管加給', .26, .10), token('導師職加', .31, .10),
            token('特教職加', .35, .10), token('應發金額', .435, .10),
            token('補助金額', .485, .10),
            token('姓名', .05, .13), token('薪額職等', .11, .13),
            token('公提勞退', .22, .13), token('補助公保', .26, .13),
            token('補助勞保', .31, .13), token('補助健保', .35, .13),
            token('第一處', .04, .20), token('級任導師', .08, .20),
            token('50,000', .17, .20), token('30,000', .22, .20),
            token('4,000', .31, .20), token('84,000', .445, .205),
            token('王小明', .05, .214), token('2,500', .22, .214),
            token('5,000', .31, .214),
            token('第二處', .04, .23), token('主任', .08, .23),
            token('45,000', .17, .23), token('30,000', .22, .23),
            token('5,000', .26, .23), token('80,000', .445, .235),
            token('李小華', .05, .244), token('2,200', .22, .244),
            token('4,500', .31, .244),
            token('本頁小計', .10, .84), token('95,000', .17, .84),
            token('60,000', .22, .84), token('5,000', .26, .84),
            token('4,000', .31, .84),
        ]

    def test_detects_and_separates_salary_from_subsidy_line(self):
        people, layout = parse_tokens.parse_tokens(self.fixture())
        self.assertEqual(layout, 'horizontal')
        self.assertEqual(len(people), 2)
        self.assertEqual(people[0]['姓名'], '王小明')
        self.assertEqual(people[0]['專業加給'], 30000)
        self.assertEqual(people[0]['導師特教'], 4000)
        self.assertEqual(people[0]['應發金額'], 84000)
        self.assertTrue(people[0]['加總相符'])
        self.assertEqual(people[1]['主管加給'], 5000)
        self.assertTrue(people[1]['加總相符'])

    def test_repeated_ocr_separator_is_repaired(self):
        self.assertTrue(parse_tokens.is_num('40.，760'))
        self.assertEqual(parse_tokens.to_int('40.，760'), 40760)

    def test_manager_row(self):
        person = self.parse(
            ('教師兼主任', .03), ('650', .07), ('徐盛旺', .11),
            ('55690', .16), ('5930', .20), ('35780', .24),
            ('97400', .37))
        self.assertEqual(person['姓名'], '徐盛旺')
        self.assertEqual(person['主管加給'], 5930)
        self.assertEqual(person['專業加給'], 35780)
        self.assertTrue(person['加總相符'])

    def test_truncated_manager_title_keeps_manager_allowance(self):
        person = self.parse(
            ('導師兼主1650', .03), ('廖雅蘭', .11),
            ('55690', .16), ('5930', .20), ('35780', .24),
            ('4000', .28), ('101400', .37))
        self.assertEqual(person['主管加給'], 5930)
        self.assertEqual(person['導師特教'], 4000)
        self.assertTrue(person['加總相符'])

    def test_teacher_allowance_after_professional(self):
        person = self.parse(
            ('導師', .03), ('350', .07), ('邱淑娟', .11),
            ('36160', .16), ('30140', .24), ('4000', .28),
            ('70300', .37))
        self.assertEqual(person['姓名'], '邱淑娟')
        self.assertEqual(person['導師特教'], 4000)
        self.assertTrue(person['加總相符'])

    def test_fallback_keeps_small_special_allowance_before_professional(self):
        # 模擬導師加給漏讀，前綴和因而無法辨識；其餘欄位仍不可左移。
        person = self.parse(
            ('資源班', .03), ('245', .07), ('洪一茜', .11),
            ('29270', .16), ('2800', .20), ('26560', .24),
            ('62630', .37))
        self.assertEqual(person['姓名'], '洪一茜')
        self.assertEqual(person['主管加給'], 0)
        self.assertEqual(person['專業加給'], 26560)
        self.assertEqual(person['導師特教'], 2800)
        self.assertEqual(person['應發金額'], 62630)
        self.assertFalse(person['應發金額推算'])
        self.assertEqual(person['加總差額'], 4000)

    def test_insurance_amount_cannot_replace_missing_gross_total(self):
        person = self.parse(
            ('資源班', .03), ('245', .07), ('洪一茜', .11),
            ('29270', .16), ('2800', .20), ('26560', .24),
            ('4000', .28), ('3107', .42))
        self.assertEqual(person['主管加給'], 0)
        self.assertEqual(person['專業加給'], 26560)
        self.assertEqual(person['導師特教'], 6800)
        self.assertEqual(person['應發金額'], 62630)
        self.assertTrue(person['應發金額推算'])
        self.assertFalse(person['加總相符'])

    def test_special_and_teacher_allowance_merge_for_af(self):
        person = self.parse(
            ('資源班', .03), ('245', .07), ('洪一茜', .11),
            ('29270', .16), ('2800', .20), ('26560', .24),
            ('4000', .28), ('62630', .37), ('3107', .42))
        self.assertEqual(person['主管加給'], 0)
        self.assertEqual(person['導師特教'], 6800)
        self.assertEqual(person['應發金額'], 62630)
        self.assertFalse(person['應發金額推算'])
        self.assertTrue(person['加總相符'])

    def test_childcare_allowance_is_other(self):
        person = self.parse(
            ('教保員', .03), ('學士', .07), ('蕭惠心', .11),
            ('51747', .16), ('4000', .22), ('55747', .36))
        self.assertEqual(person['姓名'], '蕭惠心')
        self.assertEqual(person['其他加給'], 4000)
        self.assertEqual(person['專業加給'], 0)
        self.assertTrue(person['加總相符'])


class VerticalParserTests(unittest.TestCase):
    def row(self, label, y, *values):
        return [token(label, .02, y), *[
            token(str(value), x, y) for value, x in zip(values, (.2, .4))
        ]]

    def test_exact_labels_disable_shifted_positional_fallback(self):
        rows = [
            self.row('姓名', .10, '王小明', '李小華'),
            self.row('職稱', .11, '教師', '教師'),
            self.row('本俸', .20, 50000, 45000),
            self.row('專業加給', .21, 30000, 30000),
            self.row('未辨識甲', .22, 1000, 1000),
            self.row('未辨識乙', .23, 1000, 1000),
            self.row('未辨識丙', .24, 1000, 1000),
            self.row('未辨識丁', .25, 1000, 1000),
            self.row('未辨識戊', .26, 1000, 1000),
            self.row('未辨識己', .27, 1000, 1000),
            # 舊位置備援會把 base_idx + 8 誤當地域加給，再重複計入此總額。
            self.row('應發數合計', .28, 80000, 75000),
        ]
        people = parse_tokens.parse_vertical(rows)
        self.assertEqual(len(people), 2)
        self.assertTrue(all(p['加總相符'] for p in people))
        self.assertTrue(all(p['加總差額'] == 0 for p in people))


if __name__ == '__main__':
    unittest.main()
