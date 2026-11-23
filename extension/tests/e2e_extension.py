"""Edge 擴充功能端對端測試（模擬 AF 頁面；不連真正的 fund.dgpa.gov.tw）。
用法：CHROMIUM=<Chrome for Testing 執行檔> python e2e_extension.py
（需 playwright；aftool 端預設導向 staging 127.0.0.1:5061，設 PROD=1 改用正式站）"""
import os, sys, tempfile, urllib.parse
HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import mock_fund as mockfund
from playwright.sync_api import sync_playwright

EXT = os.path.join(os.path.dirname(HERE), 'edge')
STAGING = os.environ.get('STAGING', 'http://127.0.0.1:5061')
PROD = bool(os.environ.get('PROD'))
SHOTS = os.environ.get('SHOTS')

def aftool_handler(route, request):
    resp = route.fetch(url=request.url.replace('https://aftool.leotsai.me', STAGING),
                       headers={**request.headers, 'x-forwarded-proto': 'https'})
    route.fulfill(response=resp)

def aftool_tabs(ctx):
    return [p for p in ctx.pages if p.url.startswith('https://aftool.leotsai.me/')]

def wait_banner(ctx, page, needle, before=None, ms=20000):
    for _ in range(ms // 250):
        for p in aftool_tabs(ctx):
            try:
                t = p.inner_text('#afi-banner', timeout=100)
                if needle in t and t != before:
                    return p, t
            except Exception:
                pass
        page.wait_for_timeout(250)
    return None, None

results = {}
with sync_playwright() as pw:
    ud = tempfile.mkdtemp()
    ctx = pw.chromium.launch_persistent_context(ud, executable_path=os.environ['CHROMIUM'], headless=True,
        viewport={'width': 1280, 'height': 800},
        args=[f'--disable-extensions-except={EXT}', f'--load-extension={EXT}'])
    ctx.route('https://fund.dgpa.gov.tw/**', mockfund.fund_handler)
    if not PROD:
        ctx.route('https://aftool.leotsai.me/**', aftool_handler)
    sw = ctx.service_workers[0] if ctx.service_workers else ctx.wait_for_event('serviceworker', timeout=10000)
    results['ext_loaded'] = sw.url.startswith('chrome-extension://')

    # 1) 主畫面按按鈕 → 彈出報表列印 → 由彈窗接手 → 開 aftool 分頁 → 匯入
    mockfund.VARIANT = 'popup'
    fund = ctx.new_page()
    fund.goto('https://fund.dgpa.gov.tw/system/CPA.aspx')
    fund.wait_for_selector('#af-aftool-btn', timeout=10000)
    if SHOTS:
        fund.screenshot(path=os.path.join(SHOTS, 'shot-af-button.png'))
    fund.click('#af-aftool-btn')
    aft, banner = wait_banner(ctx, fund, '已從 AF 系統匯入')
    results['run1_banner'] = banner
    results['run1_name'] = aft.inner_text('#name-af') if aft else None
    results['run1_posts'] = [{k: d.get(k) for k in ('ddlType', 'fmt', 'ddlKind', 'txtPwd', 'btnDownload')} for d in mockfund.log]
    popup = [p for p in ctx.pages if p.url.endswith('/ReportPrint.aspx')]
    results['popup_toast'] = popup[0].inner_text('#af-aftool-toast') if popup else None

    # 2) 多校：再按一次（同一個 aftool 分頁）
    aft.evaluate("()=>{document.getElementById('afi-banner').textContent='x'}")
    fund.bring_to_front(); fund.wait_for_timeout(500)
    fund.click('#af-aftool-btn')
    _, banner2 = wait_banner(ctx, fund, '已從 AF 系統匯入', before='x')
    results['run2_banner'] = banner2
    results['aftool_tabs'] = len(aftool_tabs(ctx))
    results['posts_total'] = len(mockfund.log)

    # 3) 直接在報表列印視窗按按鈕
    if popup:
        aft.evaluate("()=>{document.getElementById('afi-banner').textContent='y'}")
        popup[0].bring_to_front(); popup[0].click('#af-aftool-btn')
        _, banner3 = wait_banner(ctx, popup[0], '已從 AF 系統匯入', before='y')
        results['run3_in_popup'] = banner3
        if SHOTS:
            popup[0].wait_for_timeout(300)
            popup[0].set_viewport_size({'width': 1280, 'height': 800})
            popup[0].screenshot(path=os.path.join(SHOTS, 'shot-popup-done.png'))

    # 4) 工具列圖示路徑（背景送 run 給目前 AF 分頁）
    aft.evaluate("()=>{document.getElementById('afi-banner').textContent='z'}")
    fund.bring_to_front()
    sw.evaluate("""async () => { const [t] = await chrome.tabs.query({url: 'https://fund.dgpa.gov.tw/system/CPA.aspx'});
        await chrome.tabs.sendMessage(t.id, {type: 'run'}, {frameId: 0}); }""")
    _, banner4 = wait_banner(ctx, fund, '已從 AF 系統匯入', before='z')
    results['run4_action'] = banner4
    seq4 = []
    for i in range(8):
        seq4.append(fund.inner_text('#af-aftool-toast')[:30]); fund.wait_for_timeout(250)
    results['run4_main_seq'] = seq4

    # 5) AutoPostBack 版型
    mockfund.VARIANT = 'autopost'; n0 = len(mockfund.log)
    aft.evaluate("()=>{document.getElementById('afi-banner').textContent='w'}")
    if popup: popup[0].close()
    newpages = []; ctx.on('page', lambda pg: newpages.append(pg))
    fund.bring_to_front(); fund.click('#af-aftool-btn')
    results['run5_new'] = [(pg.url, pg.is_closed()) for pg in newpages]
    for pg in newpages:
        try: results['run5_newtoast'] = pg.evaluate("()=>{const t=document.getElementById('af-aftool-toast');return [location.href, t?t.textContent:'(none)', document.body.innerText.slice(0,80)]}")
        except Exception as e: results['run5_newtoast'] = str(e)[:200]
    _, banner5 = wait_banner(ctx, fund, '已從 AF 系統匯入', before='w')
    results['run5_autopost'] = banner5
    results['run5_posts'] = [d.get('__EVENTTARGET', '') for d in mockfund.log[n0:]]
    pp = [p for p in ctx.pages if '/ReportPrint.aspx' in p.url]
    results['run5_pages'] = [p.url for p in ctx.pages]
    results['run5_main_toast'] = fund.inner_text('#af-aftool-toast')
    results['run5_popup_toast'] = [p.evaluate("()=>{const t=document.getElementById('af-aftool-toast');return t?t.textContent:'(none)'}") for p in pp]

    # 6) 找不到報表列印的頁面 → 錯誤提示
    other = ctx.new_page(); other.goto('https://fund.dgpa.gov.tw/system/Other.aspx')
    other.wait_for_selector('#af-aftool-btn'); other.click('#af-aftool-btn'); other.wait_for_timeout(800)
    results['run6_error'] = other.inner_text('#af-aftool-toast')

    if SHOTS:
        aft.bring_to_front(); aft.set_viewport_size({'width': 1280, 'height': 800})
        aft.evaluate("()=>{const b=document.getElementById('afi-banner'); b.className='afi-banner'; b.textContent='✅ 已從 AF 系統匯入校對清冊：'+files.af.name+'。接著上傳固定清冊或薪資清冊後即可執行。'; window.scrollTo(0,0)}")
        aft.wait_for_timeout(300)
        aft.screenshot(path=os.path.join(SHOTS, 'shot-aftool-imported.png'))
    ctx.close()

for k, v in results.items():
    print(k, '=>', v)

ok = '已從 AF 系統匯入'
checks = {
    'extension loaded': results.get('ext_loaded'),
    'run1 popup flow': ok in (results.get('run1_banner') or ''),
    'run1 form values': results.get('run1_posts') == [{'ddlType': 'A', 'fmt': 'XLS', 'ddlKind': '0', 'txtPwd': 'PUBLIC', 'btnDownload': '下載報表'}],
    'run2 second school': ok in (results.get('run2_banner') or ''),
    'one aftool tab reused': results.get('aftool_tabs') == 1,
    'run3 button in popup': ok in (results.get('run3_in_popup') or ''),
    'run4 toolbar action': ok in (results.get('run4_action') or ''),
    'run5 autopostback': ok in (results.get('run5_autopost') or '') and results.get('run5_posts') == ['ddlType', ''],
    'run6 error message': '找不到「報表列印」' in (results.get('run6_error') or ''),
}
for k, v in checks.items():
    print(('PASS ' if v else 'FAIL ') + k)
sys.exit(0 if all(checks.values()) else 1)
