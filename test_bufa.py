"""v1.3.24：清冊「補發薪資」欄與 AF 補發金額比對，不再算成導師＋特教加給。全部虛構資料。"""
import io
import unittest

import openpyxl

import paycheck

AF_HEADER = ['單位', '身分證字號', '姓名', '薪俸表別', '總金額', '支領數額', '待遇差額', '補發金額',
             '專業加給表別', '總金額.1', '支領數額.1', '待遇差額.1', '補發金額.1',
             '職務加給表別', '總金額.2', '支領數額.2', '待遇差額.2', '補發金額.2']


def af_xlsx(rows):
    wb = openpyxl.Workbook()
    ws = wb.active
    # pandas 會把重複欄名自動加上 .1/.2；直接寫入同名欄位模擬 AF 匯出
    ws.append([h.split('.')[0] for h in AF_HEADER])
    for r in rows:
        ws.append(r)
    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()


def af_row(pid, name, pay, bufa=0, prof=26560, duty_code='', duty=0):
    return ['教務處', pid, name, 'A00011', pay + bufa, pay, 0, bufa,
            'B2008', prof, prof, 0, 0,
            duty_code, duty or '', duty or '', 0, 0]


def person(name, pid, pay, prof=26560, duty=0, bufa=None, mgr=0):
    p = {'姓名': name, '職稱': '教師', '身分證': pid, '身分證有效': False,
         '薪俸': pay, '主管加給': mgr, '專業加給': prof, '導師特教': duty,
         '其他加給': bufa or 0, '應發金額': pay + prof + duty + mgr + (bufa or 0),
         '加總相符': True, '最低信心': 0.9}
    if bufa is not None:
        p['補發'] = bufa
    return p


class BufaCompareTests(unittest.TestCase):
    def setUp(self):
        self.af, _ = paycheck.load_af(af_xlsx([
            af_row('A100000001', '王小明', 32710, bufa=3450),
            af_row('A100000002', '李小華', 55690, prof=35780, duty_code='C1014', duty=4000),
            af_row('A100000003', '陳大同', 40760, prof=30140),
        ]), 'af.xlsx')
        self.by = {a['姓名']: a for a in self.af}

    def test_load_af_reads_bufa(self):
        self.assertEqual(self.by['王小明']['補發'], 3450)
        self.assertEqual(self.by['王小明']['薪俸'], 32710)       # 支領數額，不含補發
        self.assertEqual(self.by['李小華']['導師特教'], 4000)
        self.assertEqual(self.by['陳大同']['補發'], 0)

    def run_compare(self, people):
        good, need = paycheck.from_ocr(people, self.af)
        self.assertEqual(need, [])
        out = paycheck.compare(good, self.af)
        return {r['姓名']: r for r in out['results']}, out['summary']

    def test_bufa_matches_and_real_bufa_diff_is_its_own_row(self):
        res, summary = self.run_compare([
            person('王小明', 'A100000001', 32710, bufa=3450),
            person('李小華', 'A100000002', 55690, prof=35780, duty=4000, bufa=4590),
            person('陳大同', 'A100000003', 40760, prof=30140, bufa=0),
        ])
        self.assertEqual(res['王小明']['狀態'], 'ok')
        self.assertEqual(res['陳大同']['狀態'], 'ok')
        d = res['李小華']['差異']
        self.assertEqual([(x['欄位'], x['清冊'], x['AF']) for x in d],
                         [(paycheck.BUFA_LABEL, 4590, 0)])       # 導師加給 4000 一致
        self.assertEqual((summary['一致'], summary['有差異']), (2, 1))

    def test_roster_without_bufa_column_is_not_compared(self):
        res, _ = self.run_compare([person('王小明', 'A100000001', 32710)])
        self.assertEqual(res['王小明']['狀態'], 'ok')

    def test_user_fixed_rows_keep_bufa(self):
        out = paycheck.compare_with_fixed([], [{
            '姓名': '王小明', '薪俸': 32710, '專業加給': 26560, '主管加給': 0, '導師特教': 0,
            '其他加給': 3450, '補發': 3450, '應發金額': 62720}], self.af)
        r = next(r for r in out['results'] if r['姓名'] == '王小明')
        self.assertEqual(r['狀態'], 'ok')

    def test_text_layer_header_aliases(self):
        for label in ('補發薪資', '補發金額', '補發薪俸'):
            self.assertEqual(paycheck.FIELD_ALIAS[label], paycheck.BUFA_FIELD)
        people = paycheck._parse_horizontal([[
            ['編號', '職稱', '員工姓名', '薪俸', '學術研究', '導師費', '補發薪資', '應發金額'],
            ['1', '教師', '王小明', '32,710', '26,560', '', '3,450', '62,720'],
        ]])
        p = people[0]
        self.assertEqual((p['導師特教'], p['補發']), (0, 3450))
        self.assertEqual(paycheck.arith_check(p), (True, 0))


if __name__ == '__main__':
    unittest.main()
