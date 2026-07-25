import unittest

import parse_tokens


def token(text, x):
    return {
        'text': text,
        'x': x,
        'y': 0.2,
        'w': 0.03,
        'conf': 0.9,
        'page': 0,
    }


class HorizontalRowTests(unittest.TestCase):
    def parse(self, *items):
        return parse_tokens.parse_horizontal_row(
            [token(text, x) for text, x in items])

    def test_role_like_name_and_optional_allowances(self):
        person = self.parse(
            ('資源班', .03), ('650', .07), ('裴淑茵', .11),
            ('55690', .16), ('2800', .20), ('35780', .24),
            ('4000', .28), ('98270', .37), ('2614', .42))
        self.assertEqual(person['姓名'], '裴淑茵')
        self.assertEqual(person['主管加給'], 2800)
        self.assertEqual(person['專業加給'], 35780)
        self.assertEqual(person['導師特教'], 4000)
        self.assertTrue(person['加總相符'])

    def test_manager_row(self):
        person = self.parse(
            ('教師兼主任', .03), ('650', .07), ('徐盛旺', .11),
            ('55690', .16), ('5930', .20), ('35780', .24),
            ('97400', .37))
        self.assertEqual(person['姓名'], '徐盛旺')
        self.assertEqual(person['主管加給'], 5930)
        self.assertEqual(person['專業加給'], 35780)
        self.assertTrue(person['加總相符'])

    def test_teacher_allowance_after_professional(self):
        person = self.parse(
            ('導師', .03), ('350', .07), ('邱淑娟', .11),
            ('36160', .16), ('30140', .24), ('4000', .28),
            ('70300', .37))
        self.assertEqual(person['姓名'], '邱淑娟')
        self.assertEqual(person['導師特教'], 4000)
        self.assertTrue(person['加總相符'])

    def test_childcare_allowance_is_other(self):
        person = self.parse(
            ('教保員', .03), ('學士', .07), ('蕭惠心', .11),
            ('51747', .16), ('4000', .22), ('55747', .36))
        self.assertEqual(person['姓名'], '蕭惠心')
        self.assertEqual(person['其他加給'], 4000)
        self.assertEqual(person['專業加給'], 0)
        self.assertTrue(person['加總相符'])


if __name__ == '__main__':
    unittest.main()
