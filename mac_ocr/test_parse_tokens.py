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

if __name__ == '__main__':
    unittest.main()
