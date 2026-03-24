from flask import Flask, request, jsonify, render_template, send_file
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import io, os, json, re
from docx import Document
from docx.oxml.ns import qn

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024

SIMPLE_COLS = ['清冊序號', '姓名', '薪俸表別', '總金額', '支領數額',
               '專業加給表別', '總金額.1', '支領數額.1',
               '職務加給表別', '總金額.2', '支領數額.2']

SIMPLE_REGION_COLS = ['清冊序號', '姓名', '薪俸表別', '總金額', '支領數額',
                      '專業加給表別', '總金額.1', '支領數額.1',
                      '職務加給表別', '總金額.2', '支領數額.2',
                      '地域加給表別', '總金額.3', '支領數額.3']

NUM_COLS = {'清冊序號', '總金額', '支領數額', '待遇差額', '補發金額', '增支',
            '總金額.1', '支領數額.1', '待遇差額.1', '補發金額.1',
            '總金額.2', '支領數額.2', '待遇差額.2', '補發金額.2',
            '總金額.3', '支領數額.3', '待遇差額.3', '補發金額.3'}


def load_school_df():
    """載入學校對照表"""
    path = os.path.join(app.root_path, 'static', 'school.xlsx')
    if not os.path.exists(path):
        return pd.DataFrame(columns=['sn1', 'sn2', 'school'])
    return pd.read_excel(path, engine='openpyxl', dtype=str)


def parse_af_filename(filename):
    """從 AF 檔案名稱擷取 sn1（機關代碼）和年月份"""
    base = os.path.splitext(filename)[0]
    parts = base.split('_')
    sn1 = ''
    yearmonth = ''
    if len(parts) >= 2:
        sn1 = parts[-2]           # 倒數第二段，例如 397085000Y
    if len(parts) >= 1:
        ts = parts[-1]            # 最後一段，例如 1150320101641635
        if len(ts) >= 5:
            yearmonth = ts[:5]    # 前5碼，例如 11503
    return sn1, yearmonth


def lookup_school(sn1):
    """根據 sn1 查詢學校名稱和 sn2"""
    df = load_school_df()
    if df.empty or not sn1:
        return '', ''
    row = df[df['sn1'].str.strip() == sn1.strip()]
    if row.empty:
        return '', ''
    return row.iloc[0]['school'], row.iloc[0]['sn2']


def read_sheet(file_bytes, filename, sheet_name):
    ext = os.path.splitext(filename)[1].lower()
    buf = io.BytesIO(file_bytes)
    engine = 'xlrd' if ext == '.xls' else 'openpyxl'
    return pd.read_excel(buf, sheet_name=sheet_name, engine=engine, dtype=str, header=0)


def build_excel(data, columns):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = '排序結果'

    hdr_fill = PatternFill('solid', start_color='1a3a5c')
    hdr_font = Font(bold=True, color='FFFFFF', name='Microsoft JhengHei', size=11)
    center = Alignment(horizontal='center', vertical='center')
    left = Alignment(horizontal='left', vertical='center')
    thin = Side(style='thin', color='AAAAAA')
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    alt_fill = PatternFill('solid', start_color='EEF3F9')
    data_font = Font(name='Microsoft JhengHei', size=10)

    for ci, col in enumerate(columns, 1):
        cell = ws.cell(row=1, column=ci, value=col)
        cell.fill = hdr_fill
        cell.font = hdr_font
        cell.alignment = center
        cell.border = border

    for ri, row in enumerate(data, 2):
        row_fill = alt_fill if ri % 2 == 0 else None
        for ci, col in enumerate(columns, 1):
            val = row.get(col, '')
            try:
                if col in NUM_COLS and val != '':
                    val = int(float(val))
            except (ValueError, TypeError):
                pass
            cell = ws.cell(row=ri, column=ci, value=val)
            cell.font = data_font
            cell.border = border
            cell.alignment = center if col != '姓名' else left
            if row_fill:
                cell.fill = row_fill

    col_widths = {'清冊序號': 9, '姓名': 10, '薪俸表別': 12, '專業加給表別': 14,
                  '職務加給表別': 14, '地域加給表別': 14, '增支': 8,
                  '總金額': 10, '支領數額': 10, '待遇差額': 10, '補發金額': 10,
                  '總金額.1': 10, '支領數額.1': 10, '待遇差額.1': 10, '補發金額.1': 10,
                  '總金額.2': 10, '支領數額.2': 10, '待遇差額.2': 10, '補發金額.2': 10,
                  '總金額.3': 10, '支領數額.3': 10}
    for ci, col in enumerate(columns, 1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(ci)].width = col_widths.get(col, 11)

    ws.row_dimensions[1].height = 22
    ws.freeze_panes = 'A2'

    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return out


def replace_in_element(element, old, new):
    """替換 XML element 內所有含 old 的文字，支援跨 run"""
    for para in element.findall('.//' + qn('w:p')):
        runs = para.findall(qn('w:r'))
        texts = []
        for r in runs:
            t = r.find(qn('w:t'))
            texts.append(t.text if t is not None else '')
        full = ''.join(texts)
        if old in full:
            new_full = full.replace(old, new)
            assigned = False
            for r in runs:
                t = r.find(qn('w:t'))
                if t is not None:
                    if not assigned:
                        t.text = new_full
                        assigned = True
                    else:
                        t.text = ''


def build_audit_docx(school_name, sn2, yearmonth):
    """產製填好資料的稽核表 Word 檔"""
    template_path = os.path.join(app.root_path, 'static', '稽核表-.docx')
    doc = Document(template_path)

    # 替換表格內的佔位符
    replace_in_element(doc.element.body, '<學校名稱>', school_name)
    replace_in_element(doc.element.body, '<年月份>', yearmonth)

    # 替換文字方塊內的 <編號>
    for txbx in doc.element.body.findall('.//' + qn('w:txbxContent')):
        para_texts = ''.join(t.text or '' for t in txbx.findall('.//' + qn('w:t')))
        if '<編號>' in para_texts:
            replace_in_element(txbx, '<編號>', sn2)

    out = io.BytesIO()
    doc.save(out)
    out.seek(0)
    return out


@app.route('/')
def index():
    return render_template('index.html')


@app.route('/download-template')
def download_template():
    path = os.path.join(app.root_path, 'static', '固定清冊範例.xlsx')
    return send_file(path, as_attachment=True, download_name='固定清冊範例.xlsx',
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')


@app.route('/download-standalone')
def download_standalone():
    path = os.path.join(app.root_path, 'static', 'AF欄位調整工具_單機d槽版.xlsm')
    return send_file(path, as_attachment=True, download_name='AF欄位調整工具_單機d槽版.xlsm',
                     mimetype='application/vnd.ms-excel.sheet.macroEnabled.12')


@app.route('/process', methods=['POST'])
def process():
    if 'roster' not in request.files or 'af' not in request.files:
        return jsonify({'error': '請上傳兩個檔案'}), 400

    roster_f = request.files['roster']
    af_f = request.files['af']
    roster_bytes = roster_f.read()
    af_bytes = af_f.read()

    try:
        roster_df = read_sheet(roster_bytes, roster_f.filename, 'input')
        roster_df.columns = [str(c).strip() for c in roster_df.columns]
        roster_df = roster_df[['序號', '姓名']].copy()
        roster_df['姓名'] = roster_df['姓名'].str.strip()
        roster_df['序號'] = pd.to_numeric(roster_df['序號'], errors='coerce')
        roster_df = roster_df.dropna(subset=['序號', '姓名'])
        roster_df = roster_df[roster_df['姓名'] != '']
        roster_df = roster_df.sort_values('序號')
        ordered_names = roster_df['姓名'].tolist()
        name_to_seq = {row['姓名']: int(row['序號']) for _, row in roster_df.iterrows()}

        af_df = read_sheet(af_bytes, af_f.filename, 0)
        af_df.columns = [str(c).strip() for c in af_df.columns]
        af_df['姓名'] = af_df['姓名'].str.strip()

        from collections import defaultdict
        af_map = defaultdict(list)
        for _, row in af_df.iterrows():
            af_map[row['姓名']].append(row)

        sorted_rows = []
        not_found = []
        found = set()

        for name in ordered_names:
            if name in af_map:
                seq = name_to_seq[name]
                for row in af_map[name]:
                    r = row.copy()
                    r['清冊序號'] = seq
                    sorted_rows.append(r)
                found.add(name)
            else:
                not_found.append(name)

        extra_seq = len(ordered_names) + 1
        extra_names = []
        for _, row in af_df.iterrows():
            if row['姓名'] not in found:
                r = row.copy()
                r['清冊序號'] = extra_seq
                sorted_rows.append(r)
                if row['姓名'] not in extra_names:
                    extra_names.append(row['姓名'])

        result_df = pd.DataFrame(sorted_rows).reset_index(drop=True)

        cols = result_df.columns.tolist()
        if '清冊序號' in cols:
            cols.remove('清冊序號')
        cols = ['清冊序號'] + cols
        result_df = result_df[cols]

        warnings = []
        if not_found:
            warnings.append('清冊中以下人員在 AF 找不到對應：' + '、'.join(not_found))
        if extra_names:
            warnings.append('AF 中以下人員不在清冊內，已附加至末尾：' + '、'.join(extra_names))

        # 從 AF 檔名查詢學校資訊
        sn1, yearmonth = parse_af_filename(af_f.filename)
        school_name, sn2 = lookup_school(sn1)

        app.config['LAST_RESULT'] = result_df.fillna('').to_json(orient='records', force_ascii=False)
        app.config['LAST_COLUMNS'] = result_df.columns.tolist()
        app.config['LAST_SCHOOL'] = school_name
        app.config['LAST_SN2'] = sn2
        app.config['LAST_YEARMONTH'] = yearmonth

        return jsonify({
            'success': True,
            'preview': result_df.head(20).fillna('').to_dict(orient='records'),
            'columns': result_df.columns.tolist(),
            'total': len(result_df),
            'warnings': warnings,
            'school_name': school_name,
            'sn2': sn2,
            'yearmonth': yearmonth
        })

    except KeyError as e:
        return jsonify({'error': '找不到欄位或工作表：' + str(e) + '，請確認檔案格式與範例相符'}), 400
    except Exception as e:
        return jsonify({'error': '處理錯誤：' + str(e)}), 500


@app.route('/download-simple')
def download_simple():
    if 'LAST_RESULT' not in app.config:
        return '尚無資料可下載', 400
    data = json.loads(app.config['LAST_RESULT'])
    all_cols = app.config['LAST_COLUMNS']
    cols = [c for c in SIMPLE_COLS if c in all_cols]
    out = build_excel(data, cols)
    return send_file(out, as_attachment=True, download_name='排序結果(簡單版).xlsx',
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')


@app.route('/download-simple-region')
def download_simple_region():
    if 'LAST_RESULT' not in app.config:
        return '尚無資料可下載', 400
    data = json.loads(app.config['LAST_RESULT'])
    all_cols = app.config['LAST_COLUMNS']
    cols = [c for c in SIMPLE_REGION_COLS if c in all_cols]
    out = build_excel(data, cols)
    return send_file(out, as_attachment=True, download_name='排序結果(簡單地域加給版).xlsx',
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')


@app.route('/download-result')
def download_result():
    if 'LAST_RESULT' not in app.config:
        return '尚無資料可下載', 400
    data = json.loads(app.config['LAST_RESULT'])
    columns = app.config['LAST_COLUMNS']
    out = build_excel(data, columns)
    return send_file(out, as_attachment=True, download_name='排序結果(完整版).xlsx',
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')


def _docx_to_pdf():
    """將填好資料的稽核表 Word 檔轉成 PDF，回傳 bytes，失敗回傳 None"""
    import subprocess, tempfile, shutil
    school_name = app.config.get('LAST_SCHOOL', '')
    sn2 = app.config.get('LAST_SN2', '')
    yearmonth = app.config.get('LAST_YEARMONTH', '')
    try:
        docx_buf = build_audit_docx(school_name, sn2, yearmonth)
        docx_buf.seek(0)  # 確保從頭讀取
        tmpdir = tempfile.mkdtemp()
        docx_path = os.path.join(tmpdir, 'audit.docx')
        pdf_path = os.path.join(tmpdir, 'audit.pdf')
        with open(docx_path, 'wb') as f:
            f.write(docx_buf.read())
        result = subprocess.run(
            ['libreoffice', '--headless', '--convert-to', 'pdf',
             '--outdir', tmpdir, docx_path],
            capture_output=True, timeout=60
        )
        if result.returncode != 0 or not os.path.exists(pdf_path):
            app.logger.error(f'LibreOffice error: {result.stderr}')
            return None
        with open(pdf_path, 'rb') as f:
            pdf_bytes = f.read()
        shutil.rmtree(tmpdir, ignore_errors=True)
        return pdf_bytes
    except Exception as e:
        app.logger.error(f'_docx_to_pdf exception: {e}')
        return None


def _build_audit_html(school_name, sn2, yearmonth, data):
    """產生稽核表的 HTML 片段（供列印用）"""
    rows_html = ''
    for row in data:
        rows_html += f'''<tr>
      <td style="text-align:left">{row.get('姓名','')}</td>
      <td>{row.get('薪俸表別','')}</td><td>{row.get('支領數額','')}</td>
      <td>{row.get('專業加給表別','')}</td><td>{row.get('支領數額.1','')}</td>
      <td>{row.get('職務加給表別','')}</td><td>{row.get('支領數額.2','')}</td>
      <td>{row.get('地域加給表別','')}</td><td>{row.get('支領數額.3','')}</td>
    </tr>\n'''
    return f'''<h2 style="text-align:center;font-size:15px;margin-bottom:12px">高雄市政府教育局所屬機關學校 待遇稽核情形紀錄表</h2>
<div style="display:flex;gap:40px;margin-bottom:12px;font-size:13px">
  <span><strong style="color:#1a3a5c">學校名稱：</strong>{school_name or '（未對應）'}</span>
  <span><strong style="color:#1a3a5c">編號：</strong>{sn2 or '—'}</span>
  <span><strong style="color:#1a3a5c">稽核月份：</strong>{yearmonth or '—'}</span>
</div>
<table style="border-collapse:collapse;width:100%;font-size:12px">
  <thead>
    <tr>
      <th rowspan="2" style="border:1px solid #333;padding:6px 10px;background:#1a3a5c;color:#fff;-webkit-print-color-adjust:exact;print-color-adjust:exact">姓名</th>
      <th colspan="2" style="border:1px solid #333;padding:6px 10px;background:#1a3a5c;color:#fff;-webkit-print-color-adjust:exact;print-color-adjust:exact">薪俸</th>
      <th colspan="2" style="border:1px solid #333;padding:6px 10px;background:#1a3a5c;color:#fff;-webkit-print-color-adjust:exact;print-color-adjust:exact">專業加給</th>
      <th colspan="2" style="border:1px solid #333;padding:6px 10px;background:#1a3a5c;color:#fff;-webkit-print-color-adjust:exact;print-color-adjust:exact">職務加給</th>
      <th colspan="2" style="border:1px solid #333;padding:6px 10px;background:#1a3a5c;color:#fff;-webkit-print-color-adjust:exact;print-color-adjust:exact">地域加給</th>
    </tr>
    <tr>
      <th style="border:1px solid #333;padding:6px 8px;background:#1a3a5c;color:#fff;-webkit-print-color-adjust:exact;print-color-adjust:exact">表別</th>
      <th style="border:1px solid #333;padding:6px 8px;background:#1a3a5c;color:#fff;-webkit-print-color-adjust:exact;print-color-adjust:exact">支領數額</th>
      <th style="border:1px solid #333;padding:6px 8px;background:#1a3a5c;color:#fff;-webkit-print-color-adjust:exact;print-color-adjust:exact">表別</th>
      <th style="border:1px solid #333;padding:6px 8px;background:#1a3a5c;color:#fff;-webkit-print-color-adjust:exact;print-color-adjust:exact">支領數額</th>
      <th style="border:1px solid #333;padding:6px 8px;background:#1a3a5c;color:#fff;-webkit-print-color-adjust:exact;print-color-adjust:exact">表別</th>
      <th style="border:1px solid #333;padding:6px 8px;background:#1a3a5c;color:#fff;-webkit-print-color-adjust:exact;print-color-adjust:exact">支領數額</th>
      <th style="border:1px solid #333;padding:6px 8px;background:#1a3a5c;color:#fff;-webkit-print-color-adjust:exact;print-color-adjust:exact">表別</th>
      <th style="border:1px solid #333;padding:6px 8px;background:#1a3a5c;color:#fff;-webkit-print-color-adjust:exact;print-color-adjust:exact">支領數額</th>
    </tr>
  </thead>
  <tbody style="font-size:12px">
    {rows_html}
  </tbody>
</table>'''


@app.route('/print-simple')
def print_simple():
    if 'LAST_RESULT' not in app.config:
        return '尚無資料可列印', 400
    school_name = app.config.get('LAST_SCHOOL', '')
    data = json.loads(app.config.get('LAST_RESULT', '[]'))
    all_cols = app.config.get('LAST_COLUMNS', [])
    simple_cols = [c for c in SIMPLE_COLS if c in all_cols]

    thead = ''.join(f'<th>{c}</th>' for c in simple_cols)
    tbody = ''
    for row in data:
        cells = ''.join(
            f'<td{"" if c != "姓名" else " class=\"name\""} >{row.get(c, "")}</td>'
            for c in simple_cols
        )
        tbody += f'<tr>{cells}</tr>\n'

    return f'''<!DOCTYPE html>
<html lang="zh-TW"><head><meta charset="UTF-8">
<title>排序結果（簡單版）</title>
<style>
  body{{font-family:"Microsoft JhengHei",Arial,sans-serif;font-size:12px;margin:20px}}
  .school{{font-size:15px;font-weight:bold;color:#1a3a5c;margin-bottom:10px}}
  h3{{font-size:13px;color:#1a3a5c;margin-bottom:8px}}
  table{{border-collapse:collapse;width:100%}}
  th{{border:1px solid #333;padding:6px 10px;background:#1a3a5c;color:#fff;text-align:center;-webkit-print-color-adjust:exact;print-color-adjust:exact}}
  td{{border:1px solid #aaa;padding:5px 8px;text-align:center}}
  td.name{{text-align:left;font-weight:600}}
  tr:nth-child(even){{background:#f0f4f9;-webkit-print-color-adjust:exact;print-color-adjust:exact}}
  @media print{{body{{margin:8px}}}}
</style>
</head>
<body onload="window.print()">
<div class="school">{school_name}</div>
<h3>排序結果（簡單版）</h3>
<table><thead><tr>{thead}</tr></thead><tbody>{tbody}</tbody></table>
</body></html>'''


@app.route('/print-audit')
def print_audit():
    if 'LAST_SCHOOL' not in app.config:
        return '尚無資料可列印', 400
    import subprocess, tempfile, shutil
    school_name = app.config.get('LAST_SCHOOL', '')
    sn2 = app.config.get('LAST_SN2', '')
    yearmonth = app.config.get('LAST_YEARMONTH', '')
    try:
        docx_buf = build_audit_docx(school_name, sn2, yearmonth)
        docx_buf.seek(0)
        tmpdir = tempfile.mkdtemp()
        env = os.environ.copy()
        env['HOME'] = tmpdir
        docx_path = os.path.join(tmpdir, 'audit.docx')
        html_path = os.path.join(tmpdir, 'audit.html')
        with open(docx_path, 'wb') as f:
            f.write(docx_buf.read())
        r = subprocess.run(
            ['libreoffice', '--headless', '--convert-to', 'html',
             '--outdir', tmpdir, docx_path],
            capture_output=True, timeout=60, env=env
        )
        if r.returncode != 0 or not os.path.exists(html_path):
            app.logger.error('LO error: ' + str(r.stderr))
            return '稽核表產製失敗，請改用下載 Word 後自行列印。', 500
        with open(html_path, 'r', encoding='utf-8', errors='replace') as f:
            html_content = f.read()
        shutil.rmtree(tmpdir, ignore_errors=True)
        html_content = html_content.replace('<body', '<body onload="window.print()"', 1)
        return html_content, 200, {'Content-Type': 'text/html; charset=utf-8'}
    except Exception as e:
        app.logger.error('print_audit error: ' + str(e))
        return '錯誤：' + str(e), 500


@app.route('/print-all')
def print_all():
    if 'LAST_RESULT' not in app.config:
        return '尚無資料可列印', 400
    import subprocess, tempfile, shutil, re as _re
    school_name = app.config.get('LAST_SCHOOL', '')
    sn2 = app.config.get('LAST_SN2', '')
    yearmonth = app.config.get('LAST_YEARMONTH', '')
    data = json.loads(app.config.get('LAST_RESULT', '[]'))
    all_cols = app.config.get('LAST_COLUMNS', [])

    simple_cols = [c for c in SIMPLE_COLS if c in all_cols]
    thead = ''.join('<th>' + c + '</th>' for c in simple_cols)
    tbody = ''
    for row in data:
        cells = ''
        for c in simple_cols:
            cls = ' class="name"' if c == '姓名' else ''
            cells += '<td' + cls + '>' + str(row.get(c, '')) + '</td>'
        tbody += '<tr>' + cells + '</tr>\n'

    simple_html = (
        '<div class="school-name">' + school_name + '</div>'
        '<h3>排序結果（簡單版）</h3>'
        '<table><thead><tr>' + thead + '</tr></thead><tbody>' + tbody + '</tbody></table>'
    )

    audit_html = ''
    if school_name:
        try:
            docx_buf = build_audit_docx(school_name, sn2, yearmonth)
            docx_buf.seek(0)
            tmpdir = tempfile.mkdtemp()
            env = os.environ.copy()
            env['HOME'] = tmpdir
            docx_path = os.path.join(tmpdir, 'audit.docx')
            html_path = os.path.join(tmpdir, 'audit.html')
            with open(docx_path, 'wb') as f:
                f.write(docx_buf.read())
            r = subprocess.run(
                ['libreoffice', '--headless', '--convert-to', 'html',
                 '--outdir', tmpdir, docx_path],
                capture_output=True, timeout=60, env=env
            )
            if r.returncode == 0 and os.path.exists(html_path):
                with open(html_path, 'r', encoding='utf-8', errors='replace') as hf:
                    raw = hf.read()
                m = _re.search(r'<body[^>]*>(.*?)</body>', raw, _re.DOTALL | _re.IGNORECASE)
                if m:
                    audit_html = m.group(1)
            shutil.rmtree(tmpdir, ignore_errors=True)
        except Exception as e:
            app.logger.error('print_all audit error: ' + str(e))

    audit_section = '<div class="audit-section">' + audit_html + '</div>' if audit_html else ''

    page = (
        '<!DOCTYPE html><html lang="zh-TW"><head><meta charset="UTF-8">'
        '<title>完整列印－' + school_name + '</title>'
        '<style>'
        'body{font-family:"Microsoft JhengHei",Arial,sans-serif;margin:20px;font-size:12px}'
        'h3{color:#1a3a5c;margin-bottom:8px;font-size:13px}'
        'table{border-collapse:collapse;width:100%;font-size:11px}'
        'th{border:1px solid #333;padding:5px 8px;background:#1a3a5c;color:#fff;text-align:center;-webkit-print-color-adjust:exact;print-color-adjust:exact}'
        'td{border:1px solid #aaa;padding:4px 8px;text-align:center}'
        'td.name{text-align:left;font-weight:600}'
        'tr:nth-child(even){background:#f0f4f9;-webkit-print-color-adjust:exact;print-color-adjust:exact}'
        '.school-name{font-size:14px;font-weight:700;color:#1a3a5c;margin-bottom:8px}'
        '.audit-section{page-break-before:right}'
        '@media print{body{margin:8px}}'
        '</style></head>'
        '<body onload="window.print()">'
        + simple_html + audit_section +
        '</body></html>'
    )
    return page


@app.route('/download-audit')
def download_audit():
    if 'LAST_SCHOOL' not in app.config:
        return '尚無資料可下載', 400
    school_name = app.config.get('LAST_SCHOOL', '')
    sn2 = app.config.get('LAST_SN2', '')
    yearmonth = app.config.get('LAST_YEARMONTH', '')
    out = build_audit_docx(school_name, sn2, yearmonth)
    filename = f'稽核表_{school_name or "未知學校"}.docx'
    return send_file(out, as_attachment=True, download_name=filename,
                     mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document')


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)
