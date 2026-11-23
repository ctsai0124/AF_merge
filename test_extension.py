"""Edge 擴充功能「AF → aftool 一鍵匯入」（v1.3.21；v1.3.22 起首頁入口預設隱藏）：權限最小化、無遠端程式碼、打包檔與網站頁面。"""
import io
import json
import os
import re
import unittest
import zipfile
from unittest.mock import patch

from app import app, AF_EXTENSION_VERSION

ROOT = os.path.dirname(os.path.abspath(__file__))
EXT = os.path.join(ROOT, 'extension', 'edge')
ZIP = os.path.join(ROOT, 'static', 'extension', f'AF-aftool-extension-{AF_EXTENSION_VERSION}.zip')
HOSTS = {'https://fund.dgpa.gov.tw/*', 'https://aftool.leotsai.me/*'}


def manifest():
    with open(os.path.join(EXT, 'manifest.json'), encoding='utf-8') as fh:
        return json.load(fh)


class ManifestTests(unittest.TestCase):
    def test_minimal_permissions(self):
        m = manifest()
        self.assertEqual(m['manifest_version'], 3)
        self.assertEqual(m['version'], AF_EXTENSION_VERSION)
        self.assertEqual(set(m['host_permissions']), HOSTS)
        self.assertFalse(m.get('permissions'))
        self.assertNotIn('optional_permissions', m)
        self.assertNotIn('optional_host_permissions', m)
        matches = {u for cs in m['content_scripts'] for u in cs['matches']}
        self.assertEqual(matches, HOSTS)
        self.assertLessEqual(len(m['description']), 132)

    def test_files_exist(self):
        m = manifest()
        paths = [m['background']['service_worker']] + list(m['icons'].values())
        paths += [js for cs in m['content_scripts'] for js in cs['js']]
        for p in paths:
            self.assertTrue(os.path.exists(os.path.join(EXT, p)), p)

    def test_no_remote_code(self):
        for name in os.listdir(EXT):
            if not name.endswith('.js'):
                continue
            text = open(os.path.join(EXT, name), encoding='utf-8').read()
            self.assertNotRegex(text, r'\beval\s*\(|new\s+Function\s*\(|importScripts|createElement\(\s*[\'"]script', name)
            self.assertNotIn('<all_urls>', text)
            for url in re.findall(r'https?://[\w.-]+', text):
                self.assertIn(url, ('https://aftool.leotsai.me', 'https://fund.dgpa.gov.tw'), name)

    def test_no_personal_data(self):
        for base, _d, names in os.walk(EXT):
            for n in names:
                if n.endswith(('.js', '.json')):
                    text = open(os.path.join(base, n), encoding='utf-8').read()
                    self.assertIsNone(re.search(r'[A-Z][12]\d{8}', text), n)

    def test_package_matches_source(self):
        self.assertTrue(os.path.exists(ZIP), '請執行 extension/tools/build_zip.py')
        with zipfile.ZipFile(ZIP) as z:
            names = set(z.namelist())
            self.assertIn('manifest.json', names)
            for name in names:
                with open(os.path.join(EXT, name), 'rb') as fh:
                    self.assertEqual(z.read(name), fh.read(), f'{name} 與原始碼不同，請重新打包')
        src = {os.path.relpath(os.path.join(b, n), EXT).replace(os.sep, '/')
               for b, _d, ns in os.walk(EXT) for n in ns if not n.startswith('.')}
        self.assertEqual(names, src)


class SiteTests(unittest.TestCase):
    def setUp(self):
        app.config.update(TESTING=True)
        self.client = app.test_client()
        self.h = {'X-Forwarded-Proto': 'https'}

    def test_privacy_page(self):
        r = self.client.get('/privacy-extension', headers=self.h)
        self.assertEqual(r.status_code, 200)
        page = r.get_data(as_text=True)
        self.assertIn('隱私權聲明', page)
        self.assertIn('fund.dgpa.gov.tw', page)
        self.assertIn('aftool.leotsai.me', page)

    @patch('app.AF_ONECLICK_ENABLED', True)
    @patch('app.touch_active')
    @patch('app.bump_counter', return_value={'visits': 1, 'sorts': 0, 'compares': 0})
    def test_home_links_extension_download_when_enabled(self, _c, _a):
        page = self.client.get('/', headers=self.h).get_data(as_text=True)
        self.assertIn('data-oneclick="1"', page)
        self.assertIn('afi-bookmarklet', page)
        self.assertIn(f'/static/extension/AF-aftool-extension-{AF_EXTENSION_VERSION}.zip', page)
        self.assertIn('edge://extensions', page)
        self.assertIn('/privacy-extension', page)
        self.assertIn('about:blank#blocked', page)

    @patch('app.touch_active')
    @patch('app.bump_counter', return_value={'visits': 1, 'sorts': 0, 'compares': 0})
    def test_oneclick_hidden_by_default_but_zip_import_kept(self, _c, _a):
        """2026-11-24 起預設隱藏一鍵匯入入口（AF_ONECLICK_ENABLED = False），ZIP 匯入照常。"""
        import app as A
        self.assertFalse(A.AF_ONECLICK_ENABLED)
        page = self.client.get('/', headers=self.h).get_data(as_text=True)
        self.assertIn('data-oneclick="0"', page)
        for hidden in ('AF-aftool-extension-', 'edge://extensions', 'afi-bookmarklet', 'afi-copy-bm',
                       '商店審核中', '可一鍵匯入', '/privacy-extension'):
            self.assertNotIn(hidden, page)
        # ZIP 拖放／公版密碼解密入口仍在
        self.assertIn('/static/af_import.js?v=', page)
        self.assertIn('如何從 AF 系統取得檔案', page)
        self.assertIn(A.af_zip.PUBLIC_PASSWORD, page)
        self.assertIn('afi-pw-tpl', page)
        # 外掛、書籤程式與隱私權頁仍保留，可隨時恢復
        self.assertEqual(self.client.get('/privacy-extension', headers=self.h).status_code, 200)
        r = self.client.get('/static/af_oneclick.js', headers=self.h); self.assertEqual(r.status_code, 200); r.close()

    def test_zip_download(self):
        r = self.client.get(f'/static/extension/AF-aftool-extension-{AF_EXTENSION_VERSION}.zip', headers=self.h)
        self.assertEqual(r.status_code, 200)
        self.assertTrue(zipfile.is_zipfile(io.BytesIO(r.data)))
        r.close()


if __name__ == '__main__':
    unittest.main()
