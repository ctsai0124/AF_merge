#!/usr/bin/env python3
"""薪資清冊 OCR 工作程式：輪詢伺服器、辨識、回傳"""
import json, os, subprocess, sys, tempfile, time, base64
import urllib.request, urllib.error

HERE = os.path.dirname(os.path.abspath(__file__))
CFG = json.load(open(os.path.join(HERE, 'config.json'), encoding='utf-8'))
SERVER, KEY = CFG['server'].rstrip('/'), CFG['key']
POLL_MIN, POLL_MAX = 3, 60

sys.path.insert(0, HERE)
from parse_tokens import parse_tokens as parse_all


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


def ocr(pdf_bytes, preferred_layout=None):
    with tempfile.NamedTemporaryFile(suffix='.pdf', delete=False) as f:
        f.write(pdf_bytes)
        path = f.name
    try:
        out = subprocess.run(
            ['swift', os.path.join(HERE, 'ocr_extract.swift'), path],
            capture_output=True, timeout=300)
        if out.returncode != 0:
            raise RuntimeError(out.stderr.decode()[:300])
        tokens = json.loads(out.stdout)
        people, layout = parse_all(tokens, preferred_layout=preferred_layout)
        print(f'  版面判定：{layout}', flush=True)
        return people, layout
    finally:
        os.unlink(path)


def main():
    print(f'OCR 工作程式啟動｜伺服器 {SERVER}', flush=True)
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
            people, layout = ocr(base64.b64decode(resp['pdf_b64']), preferred)
            ok = sum(1 for p in people if p['加總相符'])
            print(f'  解析 {len(people)} 人，加總相符 {ok}', flush=True)
            req('/ocr/result', {'job_id': jid, 'people': people, 'layout': layout})
        except Exception as e:
            print(f'  辨識失敗：{e}', flush=True)
            req('/ocr/result', {'job_id': jid, 'error': str(e)[:300]})


if __name__ == '__main__':
    main()
