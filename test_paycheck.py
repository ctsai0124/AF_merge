import unittest

import paycheck


def af_person(name='洪一茜'):
    return {
        '姓名': name,
        '身分證': '',
        '薪俸': 29270,
        '專業加給': 26560,
        '主管加給': 0,
        '導師特教': 6800,
    }


def inferred_person(name='洪一茜'):
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
            [inferred_person('洪一菁')], [af_person()])
        self.assertEqual(len(good), 1)
        self.assertEqual(good[0]['姓名'], '洪一茜')
        self.assertEqual(need, [])

    def test_ambiguous_af_fingerprint_still_requires_review(self):
        duplicate = af_person('洪一菁')
        good, need = paycheck.from_ocr(
            [inferred_person('洪一菁')], [af_person(), duplicate])
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


if __name__ == '__main__':
    unittest.main()
