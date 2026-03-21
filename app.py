from flask import Flask, request, jsonify, render_template, send_file
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import io, os, json

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024

def read_sheet(file_bytes, filename, sheet_name):
    ext = os.path.splitext(filename)[1].lower()
    buf = io.BytesIO(file_bytes)
    engine = 'xlrd' if ext == '.xls' else 'openpyxl'
    return pd.read_excel(buf, sheet_name=sheet_name, engine=engine, dtype=str)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/download-template')
def download_template():
    path = os.path.join(app.root_path, 'static', '固定清冊範例.xlsx')
    return send_file(path, as_attachment=True, download_name='固定清冊範例.xlsx',
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

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

        af_df = read_sheet(af_bytes, af_f.filename, 'output')
        af_df.columns = [str(c).strip() for c in af_df.columns]
        af_df['姓名'] = af_df['姓名'].str.strip()

        af_map = {row['姓名']: row for _, row in af_df.iterrows()}
        sorted_rows, not_found, found = [], [], set()

        for name in ordered_names:
            if name in af_map:
                sorted_rows.append(af_map[name])
                found.add(name)
            else:
                not_found.append(name)

        extra = [row for _, row in af_df.iterrows() if row['姓名'] not in found]
        result_df = pd.DataFrame(sorted_rows + extra).reset_index(drop=True)
        result_df.insert(0, '清冊序號', range(1, len(result_df) + 1))

        warnings = []
        if not_found:
            warnings.append('清冊中以下人員在 AF 找不到對應：' + '、'.join(not_found))
        if extra:
            warnings.append('AF 中以下人員不在清冊內，已附加至末尾：' + '、'.join([r['姓名'] for r in extra]))

        app.config['LAST_RESULT'] = result_df.fillna('').to_json(orient='records', force_ascii=False)
        app.config['LAST_COLUMNS'] = result_df.columns.tolist()

        return jsonify({
            'success': True,
            'preview': result_df.head(20).fillna('').to_dict(orient='records'),
            'columns': result_df.columns.tolist(),
            'total': len(result_df),
            'warnings': warnings
        })

    except KeyError as e:
        return jsonify({'error': '找不到欄位或工作表：' + str(e) + '，請確認檔案格式與範例相符'}), 400
    except Exception as e:
        return jsonify({'error': '處理錯誤：' + str(e)}), 500

@app.route('/download-result')
def download_result():
    if 'LAST_RESULT' not in app.config:
        return '尚無資料可下載', 400

    data = json.loads(app.config['LAST_RESULT'])
    columns = app.config['LAST_COLUMNS']

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

    num_cols = {'清冊序號', '總金額', '支領數額', '待遇差額', '補發金額', '增支'}

    for ri, row in enumerate(data, 2):
        row_fill = alt_fill if ri % 2 == 0 else None
        for ci, col in enumerate(columns, 1):
            val = row.get(col, '')
            try:
                if col in num_cols and val != '':
                    val = int(float(val))
            except (ValueError, TypeError):
                pass
            cell = ws.cell(row=ri, column=ci, value=val)
            cell.font = data_font
            cell.border = border
            cell.alignment = center if col != '姓名' else left
            if row_fill:
                cell.fill = row_fill

    col_widths = {'清冊序號': 9, '姓名': 10, '薪俸表別': 12, '總金額': 10,
                  '支領數額': 10, '待遇差額': 10, '補發金額': 10,
                  '專業加給表別': 14, '職務加給表別': 14, '增支': 8}
    for ci, col in enumerate(columns, 1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(ci)].width = col_widths.get(col, 11)

    ws.row_dimensions[1].height = 22
    ws.freeze_panes = 'A2'

    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return send_file(out, as_attachment=True, download_name='AF排序結果.xlsx',
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)
