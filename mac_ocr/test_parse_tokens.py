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
            ('資源班', .03), ('650', .07), ('丁小茵', .11),
            ('55690', .16), ('2800', .20), ('35780', .24),
            ('4000', .28), ('98270', .37), ('2614', .42))
        self.assertEqual(person['姓名'], '丁小茵')
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
            ('教師兼主任', .03), ('650', .07), ('丁大旺', .11),
            ('55690', .16), ('5930', .20), ('35780', .24),
            ('97400', .37))
        self.assertEqual(person['姓名'], '丁大旺')
        self.assertEqual(person['主管加給'], 5930)
        self.assertEqual(person['專業加給'], 35780)
        self.assertTrue(person['加總相符'])

    def test_truncated_manager_title_keeps_manager_allowance(self):
        person = self.parse(
            ('導師兼主1650', .03), ('丁雅文', .11),
            ('55690', .16), ('5930', .20), ('35780', .24),
            ('4000', .28), ('101400', .37))
        self.assertEqual(person['主管加給'], 5930)
        self.assertEqual(person['導師特教'], 4000)
        self.assertTrue(person['加總相符'])

    def test_teacher_allowance_after_professional(self):
        person = self.parse(
            ('導師', .03), ('350', .07), ('丁淑文', .11),
            ('36160', .16), ('30140', .24), ('4000', .28),
            ('70300', .37))
        self.assertEqual(person['姓名'], '丁淑文')
        self.assertEqual(person['導師特教'], 4000)
        self.assertTrue(person['加總相符'])

    def test_fallback_keeps_small_special_allowance_before_professional(self):
        # 模擬導師加給漏讀，前綴和因而無法辨識；其餘欄位仍不可左移。
        person = self.parse(
            ('資源班', .03), ('245', .07), ('范小茜', .11),
            ('29270', .16), ('2800', .20), ('26560', .24),
            ('62630', .37))
        self.assertEqual(person['姓名'], '范小茜')
        self.assertEqual(person['主管加給'], 0)
        self.assertEqual(person['專業加給'], 26560)
        self.assertEqual(person['導師特教'], 2800)
        self.assertEqual(person['應發金額'], 62630)
        self.assertFalse(person['應發金額推算'])
        self.assertEqual(person['加總差額'], 4000)

    def test_insurance_amount_cannot_replace_missing_gross_total(self):
        person = self.parse(
            ('資源班', .03), ('245', .07), ('范小茜', .11),
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
            ('資源班', .03), ('245', .07), ('范小茜', .11),
            ('29270', .16), ('2800', .20), ('26560', .24),
            ('4000', .28), ('62630', .37), ('3107', .42))
        self.assertEqual(person['主管加給'], 0)
        self.assertEqual(person['導師特教'], 6800)
        self.assertEqual(person['應發金額'], 62630)
        self.assertFalse(person['應發金額推算'])
        self.assertTrue(person['加總相符'])

    def test_childcare_allowance_is_other(self):
        person = self.parse(
            ('教保員', .03), ('學士', .07), ('丁惠文', .11),
            ('51747', .16), ('4000', .22), ('55747', .36))
        self.assertEqual(person['姓名'], '丁惠文')
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



class GridTwoLineParserTests(unittest.TestCase):
    """A4 橫式「每人兩行、上下欄共用表頭」清冊；全部為虛構資料。"""

    HEADER = [
        ('員工編號', .039), ('員工姓名', .094), ('帳', .142), ('號', .170),
        ('月支俸額', .207), ('職務加給', .253), ('專業加給', .295),
        ('學術研究', .340), ('導師費', .388), ('公保費', .433),
        ('應發合計', .478), ('公勞保自付', .526), ('所得稅額', .659),
        ('應扣合計', .841), ('實發總額', .885), ('備註', .938),
    ]
    HEADER2 = [('職', .079), ('稱', .107), ('俸點', .143), ('特教津貼', .206),
               ('補發薪津', .251), ('外縣市交通', .295), ('輔導津貼', .340),
               ('儲蓄存款', .526)]

    def person(self, y, emp, name, title, point, first, second=(), page=0):
        out = [token(emp, .036, y, page), token(name, .077, y, page),
               token('0123456789012', .153, y, page),
               token(title, .079, y + .015, page),
               token(point, .143, y + .015, page)]
        out += [token(text, x, y, page) for text, x in first]
        out += [token(text, x, y + .015, page) for text, x in second]
        return out

    def page(self, page=0, subject='科目：教師'):
        toks = [token('虛構市立測試國民中學 115 年 1 月份', .49, .041, page),
                token(subject, .044, .062, page)]
        toks += [token(t, x, .085, page) for t, x in self.HEADER]
        toks += [token(t, x, .103, page) for t, x in self.HEADER2]
        return toks

    def fixture(self):
        toks = self.page()
        # 甲：校長，月支俸額＋職務加給＋學術研究，扣款欄不得混入。
        toks += self.person(.121, '090001', '王大明', '校長', '650', [
            ('55690', .213), ('10010', .259), ('35780', .349),
            ('101480', .484), ('1407', .533), ('4281', .668),
            ('21909', .849), ('79571', .893)], [('10000', .531)])
        # 乙：導師＋第二行特教津貼；月支俸額前有框線雜訊「：」。
        toks += self.person(.157, '090002', '李小華', '國中導師', '680', [
            ('：57220', .214), ('35780', .349), ('4000', .397),
            ('99800', .486), ('1446', .533), ('19764', .849),
            ('80036', .894)], [('2800', .217)])
        # 丙：學術研究尾端黏到右框線（301401），應發合計可修正回 30140。
        toks += self.person(.192, '090003', '陳一二', '專任教師', '625', [
            ('54160', .213), ('301401', .352), ('84300', .487),
            ('1369', .533), ('8415', .853), ('75885', .894)])
        # 丁：學術研究整格漏讀；差額 35780 是同份文件該欄已出現的金額。
        toks += self.person(.227, '090004', '林三四', '專任教師', '650', [
            ('55690', .213), ('91470', .486), ('1407', .533),
            ('15810', .849), ('75660', .894)])
        # 戊：應發合計漏讀，以應扣合計＋實發總額確認。
        toks += self.person(.262, '090005', '張五六', '專任教師', '650', [
            ('55690', .213), ('35780', .349), ('1407', .533),
            ('15810', .849), ('75660', .894)])
        toks += [token('本頁小計', .158, .828), token('278450', .210, .828),
                 token('10010', .261, .828), token('0', .317, .828),
                 token('173260', .350, .828), token('4000', .397, .828),
                 token('0', .453, .828), token('468520', .482, .828),
                 token('2800', .219, .847), token('Q', .272, .847),
                 token('群組累計', .157, .865), token('278450', .211, .865),
                 token('468520', .481, .865), token('2800', .220, .884),
                 token('單位人數累計：5人', .145, .906),
                 token('機關人數累計：5人', .283, .906)]
        return toks

    def test_detects_layout_and_parses_each_person(self):
        people, layout = parse_tokens.parse_tokens(self.fixture())
        self.assertEqual(layout, 'horizontal')
        self.assertEqual([p['姓名'] for p in people],
                         ['王大明', '李小華', '陳一二', '林三四', '張五六'])
        self.assertTrue(all(p['加總相符'] for p in people))
        boss, mentor, glued, missing, nogross = people
        self.assertEqual(boss['職稱'], '校長')
        self.assertEqual((boss['薪俸'], boss['主管加給'], boss['專業加給'],
                          boss['導師特教'], boss['應發金額']),
                         (55690, 10010, 35780, 0, 101480))
        self.assertEqual((mentor['薪俸'], mentor['導師特教']), (57220, 6800))
        self.assertEqual(glued['專業加給'], 30140)
        self.assertEqual(missing['專業加給'], 35780)
        self.assertLessEqual(missing['最低信心'], 0.3)
        self.assertEqual(nogross['應發金額'], 91470)
        self.assertFalse(nogross['應發金額推算'])

    def test_subtotal_and_footer_rows_are_not_people(self):
        people, _ = parse_tokens.parse_tokens(self.fixture())
        names = {p['姓名'] for p in people}
        self.assertFalse(any(w in n for n in names
                             for w in ('小計', '累計', '人數')))
        self.assertEqual(len(people), 5)

    def test_page_subtotals_reconcile(self):
        checks = parse_tokens.grid_page_checks(self.fixture())
        self.assertEqual(checks, [{'page': 0, 'people': 5,
                                   'has_subtotal': True, 'diffs': {}}])

    def test_header_missing_on_one_page_uses_other_pages(self):
        toks = self.fixture()
        page2 = [dict(t, page=1) for t in self.page(1)
                 if t['text'] not in ('月支俸額', '學術研究', '應發合計')]
        page2 += self.person(.121, '090006', '周七八', '專任教師', '650', [
            ('55690', .213), ('35780', .349), ('91470', .486)], page=1)
        people, _ = parse_tokens.parse_tokens(toks + page2)
        last = people[-1]
        self.assertEqual((last['姓名'], last['薪俸'], last['專業加給'],
                          last['應發金額'], last['加總相符']),
                         ('周七八', 55690, 35780, 91470, True))

    def test_existing_two_line_layout_is_not_grid(self):
        rows = parse_tokens.group_rows(HorizontalParserTests().fixture())
        self.assertFalse(parse_tokens._is_grid_two_line(rows))


class MergeTokenSetsTests(unittest.TestCase):
    def test_keeps_one_copy_of_overlapping_reads(self):
        full = [token('5569', .213, .121, conf=.3)]
        tiles = [token('55690', .213, .121, conf=.5),
                 token('55690', .214, .122, conf=.5),
                 token('35780', .349, .121)]
        merged = parse_tokens.merge_token_sets(tiles, full)
        self.assertEqual(sorted(t['text'] for t in merged), ['35780', '55690'])

    def test_different_rows_or_pages_are_kept(self):
        a = [token('4000', .397, .121), token('4000', .397, .121, page=1)]
        b = [token('4000', .397, .157)]
        self.assertEqual(len(parse_tokens.merge_token_sets(a, b)), 3)



class VerticalRobustnessTests(unittest.TestCase):
    """直式清冊：拆字標籤、上下顛倒頁面、標籤誤判列與以算術找應發列（虛構資料）。"""

    NAMES = ('王小明', '李大華', '陳美玲')
    PAY = (55690, 44970, 35010)
    PRO = (35780, 26830, 26830)
    DUTY = (5930, 0, 5930)

    def tok(self, text, x, y):
        return {'text': text, 'x': x, 'y': y, 'w': 0.04, 'conf': 1.0, 'page': 0}

    def build(self, total_label=('應發數', '合計')):
        xs = (0.3, 0.5, 0.7)
        out = [self.tok('職稱', 0.05, 0.08)]
        out += [self.tok('教師', x, 0.08) for x in xs]
        out.append(self.tok('姓名', 0.05, 0.12))
        out += [self.tok(n, x, 0.12) for n, x in zip(self.NAMES, xs)]
        rows = [(('本俸',), self.PAY, 0.20), (('專業', '加給'), self.PRO, 0.25),
                (('職務', '加給'), self.DUTY, 0.30)]
        totals = tuple(a + b + c for a, b, c in zip(self.PAY, self.PRO, self.DUTY))
        rows.append((total_label, totals, 0.40))
        for labels, vals, y in rows:
            for i, lab in enumerate(labels):
                out.append(self.tok(lab, 0.05 + 0.03 * i, y))
            out += [self.tok(f'{v:,}', x + 0.02, y) for v, x in zip(vals, xs) if v]
        return out

    def check(self, tokens):
        people, layout = parse_tokens.parse_tokens(tokens)
        self.assertEqual(layout, 'vertical')
        self.assertEqual([p['姓名'] for p in people], list(self.NAMES))
        self.assertTrue(all(p['加總相符'] for p in people))
        self.assertEqual([p['專業加給'] for p in people], list(self.PRO))
        self.assertEqual([p['主管加給'] for p in people], list(self.DUTY))
        return people

    def test_split_labels(self):
        self.check(self.build())

    def test_upside_down_page_is_mirrored(self):
        toks = [dict(t, y=1 - t['y']) for t in self.build()]
        self.check(toks)

    def test_total_row_found_by_arithmetic(self):
        # 標題文字擠在標籤欄，「應發數合計」無法辨認；改以加總恰等於該列判定
        self.check(self.build(total_label=('高雄市某某國小員工每月薪津發放清冊',)))

    def test_label_rows_without_total_are_dropped(self):
        kept = parse_tokens._drop_label_rows([
            {'姓名': '代扣數', '應發金額': 0}, {'姓名': '總務單位', '應發金額': 0},
            {'姓名': '王小明', '應發金額': 0}, {'姓名': '王小明', '應發金額': 100}])
        self.assertEqual([k['姓名'] for k in kept], ['王小明', '王小明'])


def fake_id(letter, body8):
    """以檢查碼產生虛構但格式正確的身分證字號（測試用，非真實證號）。"""
    for d in range(10):
        cand = f'{letter}{body8}{d}'
        if parse_tokens.valid_id(cand):
            return cand
    raise AssertionError('no checksum digit')


class VerticalColumnHeaderTests(unittest.TestCase):
    """直式一人一欄清冊（明正國小樣式）：表頭歪斜、多行職稱、總計表頁、扣薪列。"""

    COLS = (0.29, 0.42, 0.55)       # 各人金額欄（靠右）
    NAME_X = (0.23, 0.36, 0.49)     # 各人姓名欄（靠左）

    def people(self):
        return [
            dict(name='王小明', id=fake_id('E', '12345678'), title='幹事',
                 pay=44970, prof=26830, duty=0, other=(), total=71800),
            dict(name='李小華', id=fake_id('S', '22345678'), title='教保員',
                 pay=48305, prof=0, duty=0, other=(('教保費', 3000),
                                                    ('事病假扣薪', -805)),
                 total=50500),
            dict(name='陳大同', id=fake_id('T', '12233445'), title='組長兼教師',
                 pay=55690, prof=35780, duty=5930, other=(), total=97400),
        ]

    def build(self, skew=False, page=0, duty_label='職務加給',
              id_override=None):
        ppl = self.people()
        t = [token('某某國小員工每月薪津發放清冊', 0.5, 0.06, page),
             token('序號', 0.08, 0.10, page), token('身分證字號', 0.11, 0.12, page),
             token('姓名', 0.08, 0.14, page), token('職稱', 0.08, 0.16, page),
             token('支薪俸級', 0.10, 0.175, page), token('小計', 0.90, 0.14, page)]
        for i, p in enumerate(ppl):
            dy = -0.008 if (skew and i == 0) else 0.0
            pid = (id_override or {}).get(i, p['id'])
            t += [token(str(i + 1), self.NAME_X[i] - 0.02, 0.10 + dy, page),
                  token(pid, self.NAME_X[i] + 0.01, 0.12 + dy, page),
                  token(p['name'], self.NAME_X[i], 0.14 + dy, page),
                  token(p['title'], self.NAME_X[i], 0.155 + dy, page)]
            if skew and i == 0:
                t.append(token('薦任七等功六級', self.NAME_X[i] + 0.02, 0.165 + dy, page))
        rows = [('本俸', 0.19, 'pay'), ('專業加給', 0.21, 'prof'),
                (duty_label, 0.23, 'duty'), ('應發數合計', 0.43, 'total')]
        for label, y, key in rows:
            t.append(token(label, 0.09, y, page))
            for i, p in enumerate(ppl):
                if p[key]:
                    t.append(token(f'{p[key]:,}', self.COLS[i], y, page))
            t.append(token(f'{sum(p[key] for p in ppl):,}', 0.92, y, page))
        for j, label in enumerate(('教保費', '事病假扣薪')):
            y = 0.29 + 0.02 * j
            t.append(token(label, 0.09, y, page))
            for i, p in enumerate(ppl):
                for lab, v in p['other']:
                    if lab == label:
                        t.append(token(f'{v:,}', self.COLS[i], y, page))
        return t

    def check(self, people):
        want = {p['name']: p for p in self.people()}
        self.assertEqual(sorted(p['姓名'] for p in people), sorted(want))
        for p in people:
            w = want[p['姓名']]
            self.assertEqual(p['薪俸'], w['pay'], p['姓名'])
            self.assertEqual(p['專業加給'], w['prof'], p['姓名'])
            self.assertEqual(p['主管加給'], w['duty'], p['姓名'])
            self.assertEqual(p['應發金額'], w['total'], p['姓名'])
            self.assertTrue(p['加總相符'], p['姓名'])
            self.assertEqual(p['身分證'], w['id'])
            self.assertTrue(p['身分證有效'])

    def test_plain_page(self):
        people, layout = parse_tokens.parse_tokens(self.build())
        self.assertEqual(layout, 'vertical')
        self.check(people)

    def test_skewed_header_with_wrapped_title(self):
        # 第一欄多一行職稱而整欄上移，表頭各列被併成同一列
        people, _ = parse_tokens.parse_tokens(self.build(skew=True))
        self.check(people)
        first = next(p for p in people if p['姓名'] == '王小明')
        self.assertEqual(first['職稱'], '幹事')

    def test_misread_label_glyph(self):
        people, _ = parse_tokens.parse_tokens(self.build(duty_label='職務加紿'))
        self.check(people)

    def test_leave_deduction_and_other_allowance_both_counted(self):
        people, _ = parse_tokens.parse_tokens(self.build())
        p = next(p for p in people if p['姓名'] == '李小華')
        self.assertEqual(p['其他加給'], 3000 - 805)

    def test_summary_page_is_skipped(self):
        summary = [token('某某國小員工每月薪津發放清冊（總計表）', 0.5, 0.06, 0),
                   token('序號', 0.08, 0.10, 0), token('科室組別', 0.08, 0.12, 0),
                   token('1', 0.2, 0.10, 0), token('教師', 0.22, 0.12, 0),
                   token('2', 0.33, 0.10, 0), token('職員', 0.35, 0.12, 0),
                   token('員工人數', 0.08, 0.15, 0), token('5', 0.29, 0.15, 0),
                   token('本俸', 0.08, 0.19, 0), token('264,660', 0.29, 0.19, 0),
                   token('224,850', 0.42, 0.19, 0),
                   token('專業加給', 0.09, 0.21, 0), token('173,260', 0.29, 0.21, 0),
                   token('133,250', 0.42, 0.21, 0),
                   token('應發數合計', 0.09, 0.43, 0), token('437,920', 0.29, 0.43, 0),
                   token('358,100', 0.42, 0.43, 0)]
        people, _ = parse_tokens.parse_tokens(summary + self.build(page=1))
        self.check(people)

    def test_id_first_letter_repaired_by_checksum(self):
        good = self.people()
        bad_e = 'B' + good[0]['id'][1:]           # E 讀成 B
        people, _ = parse_tokens.parse_tokens(self.build(id_override={0: bad_e}))
        self.check(people)
        self.assertEqual(parse_tokens.repair_id('0' + fake_id('O', '20016148')[1:]),
                         fake_id('O', '20016148'))
        # 無法唯一還原者原樣保留
        self.assertEqual(parse_tokens.repair_id('A123'), 'A123')

    def test_id_with_gridline_punctuation_repaired(self):
        # 框線被讀成夾在字號中的冒號（S2:20744494 型）：去掉後檢查碼正確才採用
        good = self.people()
        noisy = good[1]['id'][:2] + ':' + good[1]['id'][2:]
        people, _ = parse_tokens.parse_tokens(self.build(id_override={1: noisy}))
        self.check(people)
        self.assertEqual(parse_tokens.repair_id(noisy), good[1]['id'])
        # 去掉雜訊後檢查碼仍不符：原樣保留，不猜
        wrong = good[1]['id'][:9] + str((int(good[1]['id'][9]) + 1) % 10)
        self.assertEqual(parse_tokens.repair_id(wrong[:2] + ':' + wrong[2:]),
                         wrong[:2] + ':' + wrong[2:])

    def test_id_grouped_into_serial_row_is_recovered(self):
        # 寶來國中型：某欄表頭上移，字號被分到「序號」列而非「身分證字號」列
        good = self.people()
        toks = self.build()
        for i in (0, 2):
            t = next(t for t in toks if t['text'] == good[i]['id'])
            t['y'] = 0.103
        people, _ = parse_tokens.parse_tokens(toks)
        self.check(people)

    def test_split_total_does_not_shift_people_into_subtotal(self):
        toks = self.build()
        # 第二人的應發數合計被拆成「50,」「500_」兩段
        toks = [t for t in toks if not (t['text'] == '50,500')]
        toks += [token('50,', 0.40, 0.43), token('500_', 0.425, 0.43)]
        people, _ = parse_tokens.parse_tokens(toks)
        self.check(people)

    def test_missing_total_keeps_columns_aligned(self):
        # 第二人的應發與本俸都漏讀：其他人不可整批錯位到小計欄
        toks = [t for t in self.build() if t['text'] not in ('50,500', '48,305')]
        people, _ = parse_tokens.parse_tokens(toks)
        by = {p['姓名']: p for p in people}
        self.assertEqual(by['陳大同']['薪俸'], 55690)
        self.assertEqual(by['陳大同']['應發金額'], 97400)
        self.assertTrue(by['陳大同']['加總相符'])
        self.assertNotIn('李小華', [p['姓名'] for p in people if p['加總相符']])

    def test_layout_variant_names_structure(self):
        summary = [token('員工人數', 0.08, 0.15, 0), token('科室組別', 0.08, 0.12, 0),
                   token('本俸', 0.08, 0.19, 0), token('264,660', 0.29, 0.19, 0)]
        code, desc, traits = parse_tokens.layout_variant(summary + self.build(page=1))
        self.assertEqual(code, 'vertical_column')
        self.assertIn('含總計表頁', traits)
        self.assertNotIn('王小明', desc + ''.join(traits))

    def test_style_profile_and_hint(self):
        prof = parse_tokens.style_profile(self.build(), 'tiled_2x2')
        self.assertEqual(prof['style'], 'vertical_column')
        self.assertIn('needs_tiled_ocr', prof['traits'])
        self.assertEqual(prof['ocr_strategy'], 'tiled_2x2')
        self.assertEqual(parse_tokens.style_hint_layout({'style': 'horizontal_grid'}), 'horizontal')
        self.assertEqual(parse_tokens.style_hint_layout({'style': '??'}, 'vertical'), 'vertical')
        self.assertIsNone(parse_tokens.style_hint_layout(None))

    def test_stale_school_hint_does_not_override_structure(self):
        # 該校記憶為橫式，但本次上傳的是直式清冊：仍須正確解析為直式
        people, layout = parse_tokens.parse_tokens(self.build(), preferred_layout='horizontal')
        self.assertEqual(layout, 'vertical')
        self.check(people)


def htoken(text, x, y, w=None, page=0):
    t = token(text, x, y, page)
    t['w'] = w if w is not None else 0.01 * len(text)
    return t


class HorizontalBufaColumnTests(unittest.TestCase):
    """三侯國小型掃描橫式清冊：導師費之後有「補發薪資」欄（晉級補發），表頭被讀成
    幾個長 token。補發不可併入導師特教，需另以「補發」回報。全部為虛構資料。"""

    def header(self, y, bufa_label='補發薪資'):
        # 各欄中心：主管 .256 學術 .302 導師 .342 補發 .381 應發 .426
        return [htoken('編號', .052, y), htoken('職稱', .078, y), htoken('員工姓名', .137, y),
                htoken('俸點薪俸', .185, y, w=.049),
                htoken('主管加給學術研究導師費', .2955, y, w=.1265),
                htoken(bufa_label + '應發金額公保自付健保自付', .449, y, w=.18)]

    def person(self, y, no, title, name, point, pay, mgr=0, prof=0, teach=0, bufa=0):
        total = pay + mgr + prof + teach + bufa
        t = [htoken(f'{no}{title}', .072, y), htoken(name, .129, y), htoken(str(point), .176, y),
             htoken(str(pay), .217, y)]
        for v, x in ((mgr, .260), (prof, .304), (teach, .345), (bufa, .389)):
            if v:
                t.append(htoken(str(v), x, y))
        t += [htoken(str(total), .432, y), htoken('1446', .480, y), htoken('900', .526, y)]
        return t

    def build(self):
        t = [htoken('某某國小115年11月薪資印領清冊', .5, .07)] + self.header(.108)
        t += self.person(.137, 1, '校長', '王小明', 680, 57220, mgr=10010, prof=35780)
        t += self.person(.157, 2, '教師', '李小華', 290, 32710, prof=26560, bufa=3450)
        t += self.person(.177, 3, '教師', '陳大同', 650, 55690, prof=35780, teach=4000, bufa=4590)
        t += self.person(.197, 4, '教師', '林美玲', 650, 55690, prof=35780, teach=4000)
        t += self.person(.217, 5, '組長', '張志強', 575, 51100, mgr=5930, prof=35780, bufa=4620)
        # 第二個子表：補發欄名為「補發金額」
        t += self.header(.30, '補發金額')
        t += self.person(.33, 6, '幹事', '黃淑芬', 590, 44970, prof=26830, bufa=1500)
        # 第三個子表：沒有補發欄
        t += [htoken('編號', .052, .40), htoken('職稱', .078, .40), htoken('員工姓名', .137, .40),
              htoken('俸點薪俸', .185, .40, w=.049), htoken('主管加給專業加給導師費', .2955, .40, w=.1265),
              htoken('應發金額公保自付健保自付', .449, .40, w=.18)]
        t += self.person(.43, 7, '工友', '吳建國', 150, 18230, prof=18280)
        return t

    def parse(self):
        rows = parse_tokens.group_rows(self.build())
        return {p['姓名']: p for p in parse_tokens._parse_as(rows, 'horizontal')}

    def test_bufa_is_not_teacher_allowance(self):
        got = self.parse()
        self.assertEqual(len(got), 7)
        for p in got.values():
            self.assertTrue(p['加總相符'], p['姓名'])
        self.assertEqual((got['李小華']['導師特教'], got['李小華']['補發']), (0, 3450))
        self.assertEqual((got['陳大同']['導師特教'], got['陳大同']['補發']), (4000, 4590))
        self.assertEqual((got['林美玲']['導師特教'], got['林美玲']['補發']), (4000, 0))
        self.assertEqual((got['張志強']['主管加給'], got['張志強']['導師特教'],
                          got['張志強']['補發']), (5930, 0, 4620))
        self.assertEqual(got['王小明']['補發'], 0)
        # 補發仍計入其他加給，讓伺服器端驗算等式不變
        self.assertEqual(got['陳大同']['其他加給'], 4590)

    def test_each_subtable_uses_its_own_header(self):
        got = self.parse()
        self.assertEqual((got['黃淑芬']['導師特教'], got['黃淑芬']['補發']), (0, 1500))
        self.assertNotIn('補發', got['吳建國'])       # 該子表沒有補發欄：不比對

    def test_header_token_is_split_by_character_position(self):
        cols = dict(parse_tokens._split_header_token(
            htoken('補發薪資應發金額', .40, .1, w=.08)))
        self.assertAlmostEqual(cols['補發薪資'], .38, places=3)
        self.assertAlmostEqual(cols['應發金額'], .42, places=3)

    def test_style_trait_records_bufa_column(self):
        self.assertIn('has_bufa_column', parse_tokens.style_profile(self.build())['traits'])


class VerticalBufaRowTests(unittest.TestCase):
    def test_vertical_bufa_row_reported_separately(self):
        t = [token('某某國小員工每月薪津發放清冊', .5, .06),
             token('序號', .08, .10), token('姓名', .08, .14), token('小計', .9, .14)]
        names = (('甲老師', .23, 32710, 26560, 3450), ('乙老師', .36, 55690, 35780, 0))
        for i, (n, x, *_r) in enumerate(names):
            t += [token(str(i + 1), x - .02, .10), token(n, x, .14)]
        rows = (('本俸', .19, 2), ('學術研究費', .21, 3), ('補發薪資', .23, 4), ('應發數合計', .43, None))
        for label, y, k in rows:
            t.append(token(label, .09, y))
            for n, x, pay, prof, bufa in names:
                v = {2: pay, 3: prof, 4: bufa}.get(k, pay + prof + bufa)
                if v:
                    t.append(token(f'{v:,}', x + .06, y))
        people, layout = parse_tokens.parse_tokens(t)
        self.assertEqual(layout, 'vertical')
        got = {p['姓名']: p for p in people}
        self.assertEqual((got['甲老師']['導師特教'], got['甲老師']['補發']), (0, 3450))
        self.assertEqual(got['乙老師']['補發'], 0)
        self.assertTrue(all(p['加總相符'] for p in people))
        self.assertIn('has_bufa_column', parse_tokens.style_profile(t)['traits'])

# ── v1.3.26：A4 橫式薪資印領清冊（阿蓮國小型）────────────────────
LC0, LSTEP = 0.164, 0.026          # 第 0 欄（實領金額）中心與欄寬
LEDGER_HEADER = ('實領金額', '薪俸', '職務加給', '工作或', '導師費', '特教津貼', '其他補發',
                 '教保費', '加項合計', '公保費', '勞保費', None, None, None, None, '所得稅',
                 '優惠存款', '自費代課', '勞退自提', '退撫自提', '扣項合計', '機補公保費',
                 '機補勞保費', '公提勞退金', None, None, None, None, '助款合計', None)


def ledger_num(text, col, y, page=1, skew=0.0, conf=0.5, span=1):
    """金額靠右對齊放在第 col 欄（span>1 表示跨數格的黏格 token）。"""
    right = LC0 + col * LSTEP + LSTEP / 2 - 0.002
    w = max(0.004 * len(text), 0.004) if span == 1 else LSTEP * span - 0.004
    x = right - w / 2
    return {'text': text, 'x': x, 'y': y + skew * (x - 0.5), 'w': w, 'conf': conf, 'page': page}


def ledger_page(people, page=1, y0=0.22, skew=0.0, category='教員', total=True):
    """people：[(姓名, 職稱, {欄序: 文字})]；回傳一頁 token。"""
    t = [htoken('某某國小', 0.47, 0.045, page=page),
         htoken('115年11月份薪資印領清冊', 0.47, 0.076, page=page),
         htoken(category, 0.023, 0.12, page=page),
         htoken('序編薪職稱', 0.059, 0.14, page=page), htoken('姓名', 0.116, 0.14, page=page),
         htoken('日數', 0.142, 0.14, page=page)]
    for i, label in enumerate(LEDGER_HEADER):
        if label:
            t.append(htoken(label[:3] if len(label) > 3 else label, LC0 + i * LSTEP, 0.14,
                            w=0.012, page=page))
    y = y0 - 0.026
    if total:
        t.append(htoken('總計：', 0.086, y, page=page))
        t.append(htoken(f'{len(people)}人', 0.126, y, page=page))
        t.append(ledger_num('9,999,999', 0, y, page))
    for n, (name, title, cells) in enumerate(people):
        y = y0 + n * 0.026
        t += [htoken(str(n + 1), 0.031, y + skew * (0.031 - 0.5), page=page),
              htoken(title, 0.075, y + skew * (0.075 - 0.5), page=page),
              htoken(name, 0.116, y + skew * (0.116 - 0.5), w=0.02, page=page),
              htoken('全月', 0.143, y + skew * (0.143 - 0.5), w=0.012, page=page)]
        for col, text in cells.items():
            if isinstance(text, tuple):
                t.append(ledger_num(text[0], col, y, page, skew, span=text[1]))
            else:
                t.append(ledger_num(text, col, y, page, skew))
    y = y0 + len(people) * 0.026 + 0.01
    t.append(htoken('頁小計：', 0.06, y, page=page))
    t.append(htoken(f'{len(people)}人', 0.126, y, page=page))
    return t


def ledger_cells(pay, mgr, prof, teach, sped, bufa, edu, deduct, zeros=True):
    gross = pay + mgr + prof + teach + sped + bufa + edu
    fmt = lambda v: f'{v:,}'
    cells = {0: fmt(gross - deduct), 1: fmt(pay), 2: fmt(mgr), 3: fmt(prof), 4: fmt(teach),
             5: fmt(sped), 6: fmt(bufa), 7: fmt(edu), 8: fmt(gross), 9: '1,407',
             14: fmt(deduct - 1407), 20: fmt(deduct), 21: '2,614', 28: '19,436'}
    if not zeros:
        cells = {k: v for k, v in cells.items() if v != '0'}
    return cells


def summary_page():
    t = [htoken('115年11月份薪資印領清冊彙總表', 0.47, 0.076, page=0)]
    for i, label in enumerate(('實領金額', '薪俸', '職務加給', '加項合計', '公保費', '扣項合計')):
        y = 0.16 + i * 0.02
        t += [htoken(label, 0.043, y, page=0), htoken('1,234,567', 0.18, y, page=0)]
    return t


LEDGER_PEOPLE = [
    ('王小明', '校長', ledger_cells(55690, 10010, 35780, 0, 0, 0, 0, 12168)),
    ('李小華', '輔導主任', ledger_cells(55690, 5930, 35780, 0, 0, 0, 0, 13103)),
    ('陳大同', '導師', ledger_cells(55690, 0, 35780, 4000, 0, 0, 0, 10250)),
    ('林美玲', '特教組長', ledger_cells(54160, 5930, 30140, 0, 2800, 0, 0, 9982)),
    ('張志強', '教師', ledger_cells(46500, 0, 35780, 0, 0, 3450, 0, 8000)),
    ('黃淑芬', '教保員', ledger_cells(40760, 0, 0, 0, 0, 0, 4000, 4603)),
]


class LedgerLayoutTests(unittest.TestCase):
    """阿蓮國小型 A4 橫式薪資印領清冊：一人一列 30 個等寬金額欄（實領金額在前、含扣項與
    機關補助款），第一頁為彙總表。全部為虛構資料。"""

    def parse(self, tokens):
        people, layout = parse_tokens.parse_tokens(tokens)
        return {p['姓名']: p for p in people}, layout

    def test_clean_page_parses_every_column(self):
        got, layout = self.parse(summary_page() + ledger_page(LEDGER_PEOPLE))
        self.assertEqual(layout, 'horizontal')
        self.assertEqual(sorted(got), sorted(n for n, _, _ in LEDGER_PEOPLE))
        self.assertTrue(all(p['加總相符'] for p in got.values()))
        p = got['林美玲']
        self.assertEqual((p['薪俸'], p['主管加給'], p['專業加給'], p['導師特教'], p['應發金額']),
                         (54160, 5930, 30140, 2800, 93030))
        self.assertEqual((got['張志強']['補發'], got['張志強']['其他加給']), (3450, 3450))
        self.assertEqual((got['黃淑芬']['其他加給'], got['黃淑芬']['補發']), (4000, 0))
        self.assertEqual(got['王小明']['職稱'], '校長')

    def test_layout_signature(self):
        prof = parse_tokens.style_profile(summary_page() + ledger_page(LEDGER_PEOPLE))
        self.assertEqual(prof['style'], 'horizontal_ledger')
        self.assertIn('has_bufa_column', prof['traits'])
        self.assertIn('has_summary_pages', prof['traits'])
        self.assertEqual(parse_tokens.style_layout('horizontal_ledger'), 'horizontal')

    def test_skewed_scan(self):
        # 掃描歪斜：左端比右端低約 0.02（將近一列高）
        got, _ = self.parse(ledger_page(LEDGER_PEOPLE, skew=-0.025))
        self.assertEqual(len(got), len(LEDGER_PEOPLE))
        self.assertTrue(all(p['加總相符'] for p in got.values()))
        self.assertEqual(got['陳大同']['導師特教'], 4000)

    def test_missing_zero_cells_and_category_pages(self):
        pages = (ledger_page([(n, t, {k: v for k, v in c.items() if v != '0'})
                              for n, t, c in LEDGER_PEOPLE[:3]], page=1)
                 + ledger_page(LEDGER_PEOPLE[3:], page=2, category='職員', total=False))
        got, _ = self.parse(pages)
        self.assertEqual(len(got), 6)
        self.assertTrue(all(p['加總相符'] for p in got.values()))

    def test_scan_noise_is_resolved_by_printed_totals(self):
        people = [list(x) for x in LEDGER_PEOPLE]
        c = dict(people[0][2]); c[7] = '01'; c[8] = ('0101.480|', 2)      # 鄰格 0＋框線黏到加項合計
        people[0][2] = c
        c = dict(people[1][2]); del c[0]; c[1] = ('84,297|55,690', 2)     # 實領與薪俸黏成一格
        c[8] = '97,40'                                                    # 加項合計末位被截斷
        people[1][2] = c
        c = dict(people[2][2]); c[3] = '135,780'                          # 左框線讀成開頭 1
        people[2][2] = c
        c = dict(people[3][2]); c[2] = '5.9301'; c[1] = '54:160'          # 右框線 1、千分位讀成冒號
        people[3][2] = c
        got, _ = self.parse(ledger_page([tuple(x) for x in people]))
        self.assertTrue(all(p['加總相符'] for p in got.values()), {n: p['加總差額'] for n, p in got.items()})
        self.assertEqual((got['王小明']['其他加給'], got['王小明']['應發金額']), (0, 101480))
        self.assertEqual((got['李小華']['薪俸'], got['李小華']['應發金額']), (55690, 97400))
        self.assertEqual(got['陳大同']['專業加給'], 35780)
        self.assertEqual((got['林美玲']['薪俸'], got['林美玲']['主管加給']), (54160, 5930))

    def test_multiple_ocr_passes_vote(self):
        clean = ledger_page(LEDGER_PEOPLE)
        noisy = []
        for t in clean:
            t = dict(t)
            if t['text'] == '35,780':
                t['text'] = '35,7B0'        # 另一次辨識讀壞，需靠其他讀法
            noisy.append(t)
        got, _ = self.parse(noisy + clean)
        self.assertEqual(len(got), len(LEDGER_PEOPLE))
        self.assertTrue(all(p['加總相符'] for p in got.values()))
        self.assertEqual(got['王小明']['專業加給'], 35780)

    def test_unverifiable_row_is_not_accepted(self):
        people = [list(x) for x in LEDGER_PEOPLE[:3]]
        c = dict(people[2][2]); c[1] = '55,600'; del c[0]                 # 薪俸讀錯且無實領可驗算
        people[2][2] = c
        got, _ = self.parse(ledger_page([tuple(x) for x in people]))
        self.assertFalse(got['陳大同']['加總相符'])
        self.assertTrue(got['王小明']['加總相符'])

    def test_wrong_column_amount_is_rejected(self):
        # 薪俸金額被歸到特教津貼欄：超過該欄合理上限，不可湊成相符
        people = [list(x) for x in LEDGER_PEOPLE[:2]]
        c = dict(people[1][2]); c[5] = c.pop(1)
        people[1][2] = c
        got, _ = self.parse(ledger_page([tuple(x) for x in people]))
        self.assertEqual(got['李小華']['導師特教'], 0)
        self.assertFalse(got['李小華']['加總相符'] and got['李小華']['薪俸'] == 0)

    def test_full_month_column_is_not_a_name(self):
        page = [t for t in ledger_page(LEDGER_PEOPLE[:2]) if t['text'] != '李小華']
        got, _ = self.parse(page)
        self.assertNotIn('全月', got)

    def test_other_layouts_are_not_ledger(self):
        # 三侯型表頭（應發／應扣／實發）不是本版型
        rows = parse_tokens.group_rows([
            htoken('編號', .05, .1), htoken('員工姓名', .13, .1), htoken('應發金額', .45, .1),
            htoken('所得稅', .6, .1), htoken('勞退自提', .65, .1), htoken('實發金額', .8, .1),
            htoken('應扣金額', .75, .1)])
        self.assertFalse(parse_tokens._is_ledger(rows))

    def test_several_missing_pay_cells_are_not_guessed(self):
        # 薪俸與專業加給同時漏讀：雖然實領＋扣項＝加項合計，也不可把差額全塞給薪俸
        people = [list(x) for x in LEDGER_PEOPLE[:3]]
        c = dict(people[2][2]); del c[1]; del c[3]
        people[2][2] = c
        got, _ = self.parse(ledger_page([tuple(x) for x in people]))
        self.assertFalse(got['陳大同']['加總相符'])
        self.assertTrue(got['王小明']['加總相符'])

    def test_second_check_must_hold(self):
        # 加項合計＝薪資項目和，但實領＋扣項合計對不上：送人工確認
        people = [list(x) for x in LEDGER_PEOPLE[:2]]
        c = dict(people[1][2]); c[0] = '84,111'
        people[1][2] = c
        got, _ = self.parse(ledger_page([tuple(x) for x in people]))
        self.assertFalse(got['李小華']['加總相符'])
        self.assertTrue(got['王小明']['加總相符'])

    def test_row_without_name_is_not_accepted(self):
        page = [t for t in ledger_page(LEDGER_PEOPLE[:2]) if t['text'] not in ('李小華', '全月')]
        people, _ = parse_tokens.parse_tokens(page)
        self.assertEqual(len(people), 2)
        self.assertFalse(any(p['加總相符'] for p in people if not p['姓名']))

    def test_degenerate_input_does_not_crash(self):
        self.assertEqual(parse_tokens.parse_horizontal_ledger([]), [])
        people, _ = parse_tokens.parse_tokens(summary_page())
        self.assertEqual(people, [])
        # 只有表頭沒有資料列
        header_only = [t for t in ledger_page(LEDGER_PEOPLE[:1]) if t['y'] < 0.18]
        parse_tokens.parse_tokens(header_only)

    def test_header_words_without_grid_fall_back(self):
        # 表頭詞彙黏成長 token、無法擬合等距欄格：不視為本版型，交回既有解析器
        rows = parse_tokens.group_rows([
            htoken('姓名', .1, .1), htoken('實領金額薪俸職務加給', .3, .1),
            htoken('加項合計公保費勞保費', .5, .1), htoken('扣項合計機補公保費', .7, .1)])
        self.assertFalse(parse_tokens._is_ledger(rows))

    def test_cell_reader(self):
        read = parse_tokens._ledger_read
        self.assertEqual(read('0101.480|')[0][0], 101480)
        self.assertIn(87650, read('0187,650')[0])
        self.assertEqual(read('01')[0][0], 0)
        self.assertIn(4320, read('4.3201')[0])
        self.assertIn(86540, read('86,54')[0])
        here, left = read('23,456|34.567')
        self.assertEqual((here[0], left[0]), (34567, 23456))
        self.assertNotIn(5690, read('55,690')[0])     # 不可任意丟掉開頭數字
        self.assertIn(31250, read('2,00031,250')[0])
        self.assertIn(2000, read('2,00031,250')[1])
        self.assertEqual(read('o')[0], [0])
        self.assertEqual(read('115年1')[0], [])


if __name__ == '__main__':
    unittest.main()
