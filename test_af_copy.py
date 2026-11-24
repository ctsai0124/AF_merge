"""v1.3.25：首頁「如何從 AF 系統取得檔案」的公版密碼點一下即複製。"""
import re
import unittest
from unittest.mock import patch

import af_zip
import app as app_module
from app import app


class PublicPasswordCopyTests(unittest.TestCase):
    def setUp(self):
        app.config.update(TESTING=True)
        self.client = app.test_client()
        self.h = {'X-Forwarded-Proto': 'https'}

    @patch('app.touch_active')
    @patch('app.bump_counter', return_value={'visits': 1, 'sorts': 0, 'compares': 0})
    def home(self, _c, _a):
        return self.client.get('/', headers=self.h).get_data(as_text=True)

    def test_password_has_copy_hook(self):
        page = self.home()
        m = re.search(r'<code class="afi-copy"[^>]*>', page)
        self.assertIsNotNone(m)
        tag = m.group(0)
        self.assertIn('data-copy="%s"' % af_zip.PUBLIC_PASSWORD, tag)
        for attr in ('role="button"', 'tabindex="0"', 'title="點一下複製"'):
            self.assertIn(attr, tag)
        self.assertIn('/static/af_copy.js?v=%s' % app_module.APP_VERSION, page)

    def test_copy_hook_present_when_oneclick_hidden(self):
        with patch.object(app_module, 'AF_ONECLICK_ENABLED', False):
            self.assertIn('class="afi-copy"', self.home())

    def test_copy_script(self):
        r = self.client.get('/static/af_copy.js', headers=self.h)
        self.assertEqual(r.status_code, 200)
        js = r.get_data(as_text=True)
        r.close()
        for needle in ('navigator.clipboard.writeText', "execCommand('copy')", '已複製',
                       'stopPropagation', "'Enter'", "' '"):
            self.assertIn(needle, js)


if __name__ == '__main__':
    unittest.main()
