"""模擬 AF 系統（fund.dgpa.gov.tw）的主畫面與「報表列印」彈窗（ASP.NET WebForms 風格，全虛構資料）。
fake_af_public.zip 以公版密碼加密，內容為假的 .xls（非真實資料）。"""
import sys, os, re, urllib.parse, tempfile
sys.path.insert(0, '/Users/cheng-changtsai/AF_merge')
FAKE_NAME = '校對清冊_Query_PayData_Print_000000000X_115112300000000.xls'
class af_zip: PUBLIC_PASSWORD = '!@Abcdefg1234567'
from playwright.sync_api import sync_playwright

VARIANT = 'popup'
STAGING = 'http://127.0.0.1:5061'
ZIP = open(os.path.join(os.path.dirname(os.path.abspath(__file__)),'fake_af_public.zip'),'rb').read()
log = []

CPA = '''<html><head><meta charset="utf-8"></head><body><form id="form1" method="post" action="./CPA.aspx">
<input type="hidden" name="__VIEWSTATE" value="vs1"><input type="hidden" name="__EVENTTARGET" value=""><input type="hidden" name="__EVENTARGUMENT" value="">
<span>兼辦機關</span><select name="ddlOrg"><option value="1">某某國小</option></select>
<h3>資料校對-依個人分列印</h3>
<input type="button" value="報表列印" onclick="window.open('ReportPrint.aspx','rpt','width=600,height=500')">
</form></body></html>'''

def report_page(autopost=False, rtype='B'):
    onch = ' onchange="javascript:setTimeout(\'__doPostBack(\\\'ddlType\\\',\\\'\\\')\', 0)"' if autopost else ''
    sel = lambda v: ' selected="selected"' if v == rtype else ''
    return f'''<html><head><meta charset="utf-8"></head><body>
<form method="post" action="./ReportPrint.aspx{'?ap=1' if autopost else ''}" id="form1">
<input type="hidden" name="__VIEWSTATE" value="vs2"><input type="hidden" name="__EVENTVALIDATION" value="ev2">
<input type="hidden" name="__EVENTTARGET" value=""><input type="hidden" name="__EVENTARGUMENT" value="">
<script>function __doPostBack(t,a){{var f=document.forms[0];f.__EVENTTARGET.value=t;f.__EVENTARGUMENT.value=a;f.submit();}}</script>
<table>
<tr><td>憑證機關</td><td><input type="text" name="txtOrg" value="000000000X"></td></tr>
<tr><td>清冊年月</td><td><input type="text" name="txtYm" value="11511"></td></tr>
<tr><td>報表種類</td><td><select name="ddlType"{onch}><option value="B"{sel('B')}>待遇明細表</option><option value="A"{sel('A')}>校對清冊（A）</option></select></td></tr>
<tr><td>人員類別</td><td><select name="ddlKind"><option value="1">一般人員</option><option value="0">全部</option></select></td></tr>
<tr><td>包含離職退伍否</td><td><input type="checkbox" name="chkLeave" checked></td></tr>
<tr><td>報表格式</td><td><input type="radio" name="fmt" id="fmtPdf" value="PDF" checked><label for="fmtPdf">PDF</label>
<input type="radio" name="fmt" id="fmtXls" value="XLS"><label for="fmtXls">EXCEL</label></td></tr>
<tr><td>身分證號密碼</td><td><input type="password" name="txtPwd"></td></tr>
</table>
<input type="submit" name="btnDownload" value="下載報表">
</form></body></html>'''

def fund_handler(route, request):
    u = urllib.parse.urlparse(request.url)
    path = u.path
    if path.endswith('/CPA.aspx'):
        return route.fulfill(status=200, content_type='text/html; charset=utf-8', body=CPA)
    if path.endswith('/ReportPrint.aspx'):
        autopost = 'ap=1' in (u.query or '') or VARIANT == 'autopost'
        if request.method == 'GET':
            return route.fulfill(status=200, content_type='text/html; charset=utf-8', body=report_page(autopost))
        form = dict(urllib.parse.parse_qsl(request.post_data or ''))
        log.append({k: (v if k != 'txtPwd' else ('PUBLIC' if v == af_zip.PUBLIC_PASSWORD else 'OTHER')) for k, v in form.items()})
        if form.get('__EVENTTARGET') == 'ddlType':
            return route.fulfill(status=200, content_type='text/html; charset=utf-8', body=report_page(autopost, form.get('ddlType')))
        ok = (form.get('btnDownload') == '下載報表' and form.get('ddlType') == 'A' and form.get('fmt') == 'XLS'
              and form.get('ddlKind') == '0' and form.get('txtPwd') == af_zip.PUBLIC_PASSWORD and form.get('__VIEWSTATE'))
        if not ok:
            return route.fulfill(status=200, content_type='text/html; charset=utf-8',
                                 body="<script>alert('請輸入身分證號密碼');</script>")
        return route.fulfill(status=200, body=ZIP, headers={'Content-Type': 'application/octet-stream',
                             'Content-Disposition': "attachment; filename*=UTF-8''" + urllib.parse.quote('校對清冊.zip')})
    return route.fulfill(status=404, body='nf')

