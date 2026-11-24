#!/usr/bin/env python3
"""薪資清冊 OCR 工作程式：輪詢伺服器、辨識、回傳"""
import json, os, re, subprocess, sys, tempfile, time, base64
import urllib.request, urllib.error

HERE = os.path.dirname(os.path.abspath(__file__))
CFG = json.load(open(os.path.join(HERE, 'config.json'), encoding='utf-8'))
SERVER, KEY = CFG['server'].rstrip('/'), CFG['key']
POLL_MIN, POLL_MAX = 3, 60
WORKER_VERSION = 'v23'

sys.path.insert(0, HERE)
from parse_tokens import (
    parse_tokens as parse_all,
    layout_diagnostics,
    _suspicious_name,
    _is_two_line_horizontal,
    _is_grid_two_line,
    _is_ledger,
    group_rows,
    merge_token_sets,
    ensemble_people,
    layout_variant,
    style_profile,
    style_hint_layout,
)

DIAG_DIR = os.path.join(HERE, 'diagnostics')
DIAG_KEEP = 10
TITLE_RE = re.compile(r'國民|國小|國中|中學|小學|學校|高中|高職')


def title_texts(tokens, limit=10):
    """各頁頂端（上方 15%）或含「清冊」字樣、且含學校字樣的標題文字，供伺服器以
    校名對回機關代碼。只在伺服器記憶體中使用，不寫入記錄檔或記憶檔。"""
    out = []
    for t in sorted(tokens, key=lambda t: (t.get('page', 0), t.get('y', 1))):
        text = (t.get('text') or '').strip()
        top = t.get('y', 1) < 0.15 or '清冊' in text      # 旋轉頁的標題不在頂端
        if top and len(text) >= 4 and TITLE_RE.search(text):
            if text[:80] not in out:
                out.append(text[:80])
        if len(out) >= limit:
            break
    return out


def req(path, data=None, timeout=60):
    url = SERVER + path
    body = json.dumps(data).encode() if data is not None else None
    r = urllib.request.Request(url, data=body, method='POST' if body else 'GET')
    r.add_header('X-OCR-KEY', KEY)
    if body:
        r.add_header('Content-Type', 'application/json')
    try:
        with urllib.request.urlopen(r, timeout=timeout) as resp:
            if resp.status == 204:
                return {}
            return json.loads(resp.read() or b'{}')
    except urllib.error.HTTPError as e:
        print(f'  HTTP {e.code}：{e.read()[:200]}', flush=True)
        return None
    except Exception as e:
        print(f'  連線失敗：{e}', flush=True)
        return None


def save_diagnostics(tokens, job_id, layout):
    """保存異常工作的原始 token；最多保留 10 份，檔名不含學校或個資。"""
    os.makedirs(DIAG_DIR, mode=0o700, exist_ok=True)
    try:
        os.chmod(DIAG_DIR, 0o700)
    except OSError:
        pass
    safe_jid = ''.join(c for c in (job_id or 'unknown') if c.isalnum())[:20]
    stamp = time.strftime('%Y%m%d-%H%M%S')
    path = os.path.join(DIAG_DIR, f'{stamp}_{safe_jid}_{layout}.json')
    with open(path, 'w', encoding='utf-8') as f:
        json.dump(tokens, f, ensure_ascii=False)
    try:
        os.chmod(path, 0o600)
    except OSError:
        pass

    files = sorted(
        (os.path.join(DIAG_DIR, n) for n in os.listdir(DIAG_DIR)
         if n.endswith('.json')),
        key=os.path.getmtime)
    for old in files[:-DIAG_KEEP]:
        try:
            os.unlink(old)
        except OSError:
            pass
    return path


def ocr(pdf_bytes, preferred_layout=None, job_id='', style_hint=None):
    with tempfile.NamedTemporaryFile(suffix='.pdf', delete=False) as f:
        f.write(pdf_bytes)
        pdf_path = f.name
    try:
        swift = os.path.join(HERE, 'ocr_extract.swift')

        def recognize(scale, tiles=None):
            cmd = ['swift', swift, pdf_path, '--scale', str(scale)]
            if tiles:
                cmd += ['--tiles', tiles]
            out = subprocess.run(cmd, capture_output=True, timeout=300)
            if out.returncode != 0:
                raise RuntimeError(out.stderr.decode()[:300])
            return json.loads(out.stdout)

        def quality(found):
            return (sum(1 for p in found if p.get('加總相符')),
                    -sum(1 for p in found if _suspicious_name(p.get('姓名'))),
                    len(found))

        strategy = 'full'
        tokens = recognize(3)
        rows = group_rows(tokens)
        if _is_two_line_horizontal(rows):
            strategy = 'upscale_4x'
            # A3 雙行表在不同縮放倍率各有少數姓名／粗體金額失真。第二次
            # 只補進姓名、職稱與薪資項目欄，避免把薪俸起始列重複計數。
            standard = tokens
            tokens = recognize(4)
            tokens.extend(t for t in standard
                          if t.get('x', 0) < 0.16
                          or 0.19 <= t.get('x', 0) <= 0.46)
            people, layout = parse_all(tokens, preferred_layout=preferred_layout)
        elif _is_ledger(rows):
            # v23：A4 橫式薪資印領清冊（一人一列 30 個金額欄，字小、格線密）。整頁
            # 辨識常整格漏讀或黏格，再以 5 倍 3×2、3 倍與 4 倍 2×2 分塊各辨識一次；
            # 全部結果交給解析器（不去重），由各格候選值與印列加項合計／實領／
            # 扣項合計驗算挑選，姓名逐字投票。阿蓮國小實檔約 60 秒。
            strategy = 'tiled_3x2'
            extra = []
            for scale, tiles in ((5, '3x2'), (3, '2x2'), (4, '2x2')):
                try:
                    extra.extend(recognize(scale, tiles))
                except Exception as e:
                    print(f'  分塊 {scale}倍 {tiles} 辨識失敗，略過：{e}', flush=True)
            tokens = tokens + extra
            people, layout = parse_all(tokens, preferred_layout=preferred_layout)
        else:
            people, layout = parse_all(tokens, preferred_layout=preferred_layout)
            # 整頁辨識偶爾整欄漏字（例：A4 橫式每人兩行清冊的月支俸額／學術研究／
            # 應發合計欄）。遇到該版面或結果不完整時，再以 2×2 分塊補辨並去重合併。
            # 只有合併後確認是「每人兩行共用表頭」版面、且算術相符人數變多時才採用；
            # 其他既有版面（直式、一人一列、A3 雙行）一律維持整頁結果，避免回歸。
            if _is_grid_two_line(rows) or not people or any(
                    not p.get('加總相符') for p in people):
                try:
                    merged = merge_token_sets(recognize(3, '2x2'), tokens)
                    tiled_people, tiled_layout = parse_all(
                        merged, preferred_layout=preferred_layout)
                    if (_is_grid_two_line(group_rows(merged))
                            and quality(tiled_people) > quality(people)):
                        print('  採用分塊補辨結果', flush=True)
                        tokens, people, layout = merged, tiled_people, tiled_layout
                        strategy = 'tiled_2x2'
                    elif tiled_layout == layout:
                        # 其他版面：以人為單位合併，整頁已相符的列不受影響。
                        combined = ensemble_people(people, tiled_people)
                        if quality(combined) > quality(people):
                            print('  以分塊結果補強未相符／漏讀人員', flush=True)
                            people = combined
                            strategy = 'tiled_2x2'
                except Exception as e:
                    print(f'  分塊補辨失敗，沿用整頁結果：{e}', flush=True)
        # 安全退回：該校記憶只是加分提示。若帶提示的結果仍有加總不符，
        # 改用「不帶提示」的通用判斷重解析，品質較好就採用，避免舊記憶誤判換版面的學校。
        if preferred_layout and any(not p.get('加總相符') for p in people):
            alt_people, alt_layout = parse_all(tokens)
            if quality(alt_people) > quality(people):
                print(f'  該校記憶（{preferred_layout}）不適用本檔，已退回通用判斷：{alt_layout}',
                      flush=True)
                people, layout = alt_people, alt_layout
        print(f'  版面判定：{layout}', flush=True)
        profile = None
        try:
            code, desc, traits = layout_variant(tokens)
            print(f'  版型特徵：{code}｜{desc}｜{"、".join(traits)}', flush=True)
            profile = style_profile(tokens, strategy)
            if isinstance(style_hint, dict) and style_hint.get('style'):
                same = style_hint.get('style') == profile['style']
                print('  與該校記憶樣式' + ('相同' if same else
                      f'不同（記憶 {style_hint.get("style")} → 本次 {profile["style"]}），'
                      '已依本次結構判斷'), flush=True)
        except Exception as e:      # 只供紀錄，失敗不影響辨識
            print(f'  版型特徵判讀失敗：{e}', flush=True)
        ok = sum(1 for p in people if p.get('加總相符'))
        suspicious = sum(1 for p in people if _suspicious_name(p.get('姓名')))
        page_count = len({t.get('page') for t in tokens})
        too_few_for_pages = page_count >= 3 and len(people) < page_count * 2
        abnormal = (
            not people
            or ok < len(people)
            or suspicious >= max(2, len(people) * 0.1)
            or too_few_for_pages
        )
        if abnormal:
            diag_path = save_diagnostics(tokens, job_id, layout)
            summary = layout_diagnostics(tokens)
            print(f'  ⚠ 已保存診斷：{diag_path}', flush=True)
            print(f'  雙版面診斷：{summary}', flush=True)
        return people, layout, profile, title_texts(tokens)
    finally:
        os.unlink(pdf_path)


def main():
    print(f'OCR 工作程式啟動 {WORKER_VERSION}｜伺服器 {SERVER}', flush=True)
    while True:
        resp = req('/ocr/claim')
        if resp is None:
            time.sleep(POLL_MAX)
            continue

        wait = max(POLL_MIN, min(POLL_MAX, int(resp.get('next_poll', POLL_MAX))))
        if not resp.get('job_id'):
            time.sleep(wait)
            continue

        jid = resp['job_id']
        style_hint = resp.get('style_hint') if isinstance(resp.get('style_hint'), dict) else None
        preferred = resp.get('layout_hint') or style_hint_layout(style_hint)
        hint_text = f'（該校記憶：{preferred}）' if preferred else ''
        if style_hint and style_hint.get('style'):
            hint_text += f'（樣式：{style_hint.get("style_label") or style_hint["style"]}）'
        print(f'領到工作 {jid}{hint_text}', flush=True)
        try:
            people, layout, profile, titles = ocr(
                base64.b64decode(resp['pdf_b64']), preferred, jid, style_hint)
            ok = sum(1 for p in people if p['加總相符'])
            print(f'  解析 {len(people)} 人，加總相符 {ok}', flush=True)
            req('/ocr/result', {'job_id': jid, 'people': people, 'layout': layout,
                                'layout_profile': profile, 'title_texts': titles})
        except Exception as e:
            print(f'  辨識失敗：{e}', flush=True)
            req('/ocr/result', {'job_id': jid, 'error': str(e)[:300]})


if __name__ == '__main__':
    main()
