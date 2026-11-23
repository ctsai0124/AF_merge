"""把 extension/edge 打包成商店／開發人員模式用的 ZIP（manifest.json 在 ZIP 根目錄）。
輸出：static/extension/AF-aftool-extension-<版本>.zip（aftool 網站提供下載，也用於上傳 Partner Center）。"""
import json, os, zipfile

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
SRC = os.path.join(ROOT, 'extension', 'edge')


def build():
    version = json.load(open(os.path.join(SRC, 'manifest.json'), encoding='utf-8'))['version']
    out_dir = os.path.join(ROOT, 'static', 'extension')
    os.makedirs(out_dir, exist_ok=True)
    out = os.path.join(out_dir, f'AF-aftool-extension-{version}.zip')
    files = []
    for base, _dirs, names in os.walk(SRC):
        for n in names:
            if n.startswith('.'):
                continue
            files.append(os.path.relpath(os.path.join(base, n), SRC))
    with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as z:
        for rel in sorted(files):
            info = zipfile.ZipInfo(rel.replace(os.sep, '/'), date_time=(2026, 11, 23, 0, 0, 0))
            info.compress_type = zipfile.ZIP_DEFLATED
            info.external_attr = 0o644 << 16
            z.writestr(info, open(os.path.join(SRC, rel), 'rb').read())
    return out


if __name__ == '__main__':
    print(build())
