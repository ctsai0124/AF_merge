#!/usr/bin/env python3
"""薪資清冊 OCR 工作程式：輪詢伺服器、辨識、回傳"""
import json, os, subprocess, sys, tempfile, time, base64
import urllib.request, urllib.error

HERE = os.path.dirname(os.path.abspath(__file__))
CFG = json.load(open(os.path.join(HERE, 'config.json'), encoding='utf-8'))
SERVER, KEY = CFG['server'].rstrip('/'), CFG['key']
POLL_MIN, POLL_MAX = 3, 60
WORKER_VERSION = 'v10'

sys.path.insert(0, HERE)
from parse_tokens import (
    parse_tokens as parse_all,
    layout_diagnostics,
    _suspicious_name,
)

DIAG_DIR = os.path.join(HERE, 'diagnostics')
DIAG_KEEP = 10


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


def ocr(pdf_bytes, preferred_layout=None, job_id=''):
    with tempfile.NamedTemporaryFile(suffix='.pdf', delete=False) as f:
        f.write(pdf_bytes)
        pdf_path = f.name
    try:
        out = subprocess.run(
            ['swift', os.path.join(HERE, 'ocr_extract.swift'), pdf_path],
            capture_output=True, timeout=300)
        if out.returncode != 0:
            raise RuntimeError(out.stderr.decode()[:300])
        tokens = json.loads(out.stdout)
        people, layout = parse_all(tokens, preferred_layout=preferred_layout)
        print(f'  版面判定：{layout}', flush=True)
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
        return people, layout
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
        preferred = resp.get('layout_hint')
        hint_text = f'（該校記憶：{preferred}）' if preferred else ''
        print(f'領到工作 {jid}{hint_text}', flush=True)
        try:
            people, layout = ocr(
                base64.b64decode(resp['pdf_b64']), preferred, jid)
            ok = sum(1 for p in people if p['加總相符'])
            print(f'  解析 {len(people)} 人，加總相符 {ok}', flush=True)
            req('/ocr/result', {'job_id': jid, 'people': people, 'layout': layout})
        except Exception as e:
            print(f'  辨識失敗：{e}', flush=True)
            req('/ocr/result', {'job_id': jid, 'error': str(e)[:300]})


if __name__ == '__main__':
    main()
