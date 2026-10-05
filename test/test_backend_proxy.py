import io
import json
import os
from pathlib import Path
import tempfile
import unittest
from unittest.mock import MagicMock, patch

import backend_proxy
import ui_server as ui


class BackendTests(unittest.TestCase):
    def handler(self, path, method='GET', body=b''):
        h = ui.Handler.__new__(ui.Handler)
        h.path, h.command = path, method
        h.headers = {'Host': 'localhost', 'Content-Length': str(len(body)), 'Content-Type': 'application/json'}
        h.rfile, h.wfile = io.BytesIO(body), io.BytesIO()
        h.send_response = MagicMock()
        h.send_header = MagicMock()
        h.end_headers = MagicMock()
        return h

    def test_vercel_without_durable_backend_does_not_create_local_jobs(self):
        h = self.handler('/api/jobs', 'POST', b'{}')
        with patch.dict(os.environ, {'VERCEL': '1', 'SHIPMENT_BACKEND_URL': '', 'SHIPMENT_API_TOKEN': ''}), \
             patch.object(ui, 'create_job') as create:
            h.do_POST()
        create.assert_not_called()
        h.send_response.assert_called_with(503)
        self.assertEqual(json.loads(h.wfile.getvalue())['code'], 'BACKEND_NOT_CONFIGURED')

    def test_all_session_routes_use_same_backend_and_preserve_bytes(self):
        job = 'a'*32
        routes = [('GET', '/api/config'), ('POST', '/api/jobs'),
                  ('GET', f'/api/jobs/{job}/report'), ('GET', f'/api/jobs/{job}/preview?file=-1&sheet=2'),
                  ('POST', f'/api/jobs/{job}/edit/-1'), ('GET', f'/api/jobs/{job}/download?file=-1'),
                  ('GET', f'/api/jobs/{job}/zip'), ('DELETE', f'/api/jobs/{job}')]
        for method, path in routes:
            with self.subTest(path=path), patch.dict(os.environ, {
                    'VERCEL': '1', 'SHIPMENT_BACKEND_URL': 'https://private.example',
                    'SHIPMENT_BACKEND_TOKEN': 'x'*48, 'SHIPMENT_API_TOKEN': ''}), \
                 patch.object(backend_proxy.http.client, 'HTTPSConnection') as connect, \
                 patch.object(ui, 'directory', side_effect=AssertionError('must not use Vercel /tmp')):
                response = connect.return_value.getresponse.return_value
                response.status = 200
                response.getheader.side_effect = lambda name: 'application/octet-stream' if name == 'Content-Type' else None
                response.read1.side_effect = [b'PK\x00\xff', b'final', b'']
                h = self.handler(path, method, b'{"edits":[]}')
                getattr(h, 'do_'+method)()
                connect.assert_called_once_with('private.example', 443, timeout=300)
                args, kwargs = connect.return_value.request.call_args
                self.assertEqual(args, (method, path))
                self.assertEqual(kwargs['headers']['Authorization'], 'Bearer '+'x'*48)
                if method in ('POST', 'DELETE'):
                    self.assertEqual(kwargs['body'], b'{"edits":[]}')
                self.assertEqual(h.wfile.getvalue(), b'PK\x00\xfffinal')

    def test_backend_requires_secret_even_for_localhost(self):
        with patch.dict(os.environ, {'SHIPMENT_API_TOKEN': 'x'*48}):
            h = self.handler('/api/config')
            with self.assertRaises(PermissionError):
                h.guard()
            h.headers['Authorization'] = 'Bearer '+'x'*48
            h.guard()

    def test_missing_session_has_distinct_code_for_reads_and_edits(self):
        with tempfile.TemporaryDirectory() as folder, patch.object(ui, 'STORE', Path(folder)), \
             patch.dict(os.environ, {'VERCEL': '', 'SHIPMENT_BACKEND_URL': '', 'SHIPMENT_API_TOKEN': ''}):
            for method, suffix, body in [('GET', '', b''), ('POST', '/edit/-1', b'{"edits":[]}')]:
                h = self.handler('/api/jobs/'+'a'*32+suffix, method, body)
                getattr(h, 'do_'+method)()
                h.send_response.assert_called_with(404)
                self.assertEqual(json.loads(h.wfile.getvalue())['code'], 'SESSION_NOT_FOUND')

    def test_no_redirect_or_plain_http_with_credentials(self):
        with patch.dict(os.environ, {'SHIPMENT_BACKEND_URL': 'http://private.example', 'SHIPMENT_BACKEND_TOKEN': 'x'*48}), \
             patch.object(backend_proxy.http.client, 'HTTPSConnection') as connect:
            h = self.handler('/api/config')
            backend_proxy.forward_api(h)
            connect.assert_not_called()
            h.send_response.assert_called_with(503)
        with patch.dict(os.environ, {'SHIPMENT_BACKEND_URL': 'https://private.example', 'SHIPMENT_BACKEND_TOKEN': 'x'*48}), \
             patch.object(backend_proxy.http.client, 'HTTPSConnection') as connect:
            connect.return_value.getresponse.return_value.status = 302
            h = self.handler('/api/config')
            backend_proxy.forward_api(h)
            self.assertEqual(connect.return_value.request.call_count, 1)
            h.send_response.assert_called_with(502)


if __name__ == '__main__':
    unittest.main()
