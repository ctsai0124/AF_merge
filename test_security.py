import io
import os
import unittest
from unittest.mock import patch

from app import app


class SecurityAndUiTests(unittest.TestCase):
    def setUp(self):
        app.config.update(TESTING=True)
        self.client = app.test_client()

    @patch('app.touch_active')
    @patch('app.bump_counter', return_value={'visits': 1, 'sorts': 0, 'compares': 0})
    def test_home_has_current_version_and_no_external_counter(self, _counter, _active):
        response = self.client.get('/', headers={'X-Forwarded-Proto': 'https'})

        self.assertEqual(response.status_code, 200)
        page = response.get_data(as_text=True)
        self.assertIn('系統版本：v1.3.12', page)
        self.assertIn('版本 2026.09.10.2', page)
        self.assertIn('<main class="container">', page)
        self.assertNotIn('site-counter.js', page)

    def test_http_forwarded_request_redirects_to_https_and_keeps_method(self):
        response = self.client.post(
            '/process?source=test',
            headers={'X-Forwarded-Proto': 'http', 'Host': 'aftool.leotsai.me'},
        )

        self.assertEqual(response.status_code, 308)
        self.assertEqual(
            response.headers['Location'],
            'https://aftool.leotsai.me/process?source=test',
        )

    def test_security_headers_are_present(self):
        response = self.client.get('/stats', headers={'X-Forwarded-Proto': 'https'})

        self.assertEqual(response.headers['X-Frame-Options'], 'DENY')
        self.assertEqual(response.headers['X-Content-Type-Options'], 'nosniff')
        self.assertIn("frame-ancestors 'none'", response.headers['Content-Security-Policy'])
        self.assertEqual(response.headers['X-Robots-Tag'], 'noindex, nofollow, noarchive')
        self.assertIn('max-age=31536000', response.headers['Strict-Transport-Security'])

    def test_sensitive_response_is_not_cached(self):
        response = self.client.post('/process', data={})

        self.assertEqual(response.status_code, 400)
        self.assertEqual(response.headers['Cache-Control'], 'no-store, max-age=0')
        self.assertEqual(response.headers['Pragma'], 'no-cache')

    def test_large_upload_returns_json_instead_of_html(self):
        old_limit = app.config['MAX_CONTENT_LENGTH']
        app.config['MAX_CONTENT_LENGTH'] = 32
        try:
            response = self.client.post(
                '/process',
                data={'roster': (io.BytesIO(b'x' * 64), 'too-large.xlsx')},
            )
        finally:
            app.config['MAX_CONTENT_LENGTH'] = old_limit

        self.assertEqual(response.status_code, 413)
        self.assertIn('上傳檔案合計超過', response.get_json()['error'])

    def test_admin_endpoint_is_hidden_until_admin_key_is_configured(self):
        with patch.dict(os.environ, {}, clear=False):
            os.environ.pop('ADMIN_KEY', None)
            response = self.client.get('/stats/diag')

        self.assertEqual(response.status_code, 404)

    def test_admin_endpoint_requires_and_accepts_configured_key(self):
        with patch.dict(os.environ, {'ADMIN_KEY': 'test-admin-key'}):
            denied = self.client.get('/stats/diag')
            allowed = self.client.get(
                '/stats/diag', headers={'X-Admin-Key': 'test-admin-key'})

        self.assertEqual(denied.status_code, 401)
        self.assertEqual(allowed.status_code, 200)


if __name__ == '__main__':
    unittest.main()
