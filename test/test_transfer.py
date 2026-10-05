"""Transport optimizations must preserve uploaded bytes and response contents."""
import base64
import gzip
import io
import json
from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch
from zipfile import ZipFile

import ui_server as ui


class TransferTests(unittest.TestCase):
    def test_vercel_completes_job_before_returning(self):
        with tempfile.TemporaryDirectory() as folder:
            root=Path(folder)
            source=root/'source.xlsx';source.write_bytes(b'input')
            template=root/'template.xlsx';template.write_bytes(b'template')
            def finish(job_id,args):
                ui.update_job(job_id,status='ready')
            with patch.object(ui,'STORE',root/'jobs'), patch.object(ui,'local_input',return_value=source), \
                 patch.object(ui.process,'find_default_template',return_value=template), \
                 patch.dict(ui.os.environ,{'VERCEL':'1'}), \
                 patch.object(ui,'generate',side_effect=finish) as generate, \
                 patch.object(ui.POOL,'submit') as submit:
                job=ui.create_job({'use_local':True})
                self.assertEqual(job['status'],'ready')
                generate.assert_called_once()
                submit.assert_not_called()

    def test_multipart_matches_legacy_upload_byte_for_byte(self):
        output = io.BytesIO()
        with ZipFile(output, 'w') as archive:
            archive.writestr('[Content_Types].xml', '<Types/>')
            archive.writestr('test.bin', bytes(range(256)))
        raw = output.getvalue()
        # Construct raw binary parts exactly as a browser sends them (no MIME line normalization).
        boundary = 'upload-test-boundary'
        body = (f'--{boundary}\r\nContent-Disposition: form-data; name="options"\r\n\r\n'
                '{"week":40,"use_local":false}\r\n'
                f'--{boundary}\r\nContent-Disposition: form-data; name="input"; filename="Kế hoạch.xlsx"\r\n'
                'Content-Type: application/octet-stream\r\n\r\n').encode() + raw + f'\r\n--{boundary}--\r\n'.encode()
        parsed = ui.parse_payload(f'multipart/form-data; boundary={boundary}', body)
        self.assertEqual(parsed['input']['raw'], raw)
        self.assertEqual(parsed['input']['name'], 'Kế hoạch.xlsx')
        self.assertFalse(parsed['use_local'])
        with tempfile.TemporaryDirectory() as folder:
            a,b = Path(folder)/'a.xlsx',Path(folder)/'b.xlsx'
            ui.save_upload(parsed['input'],a)
            ui.save_upload({'name':'input.xlsx','data':base64.b64encode(raw).decode()},b)
            self.assertEqual(a.read_bytes(),b.read_bytes())

    def test_reject_malformed_form(self):
        with self.assertRaises(ValueError):
            ui.parse_payload('multipart/form-data',b'broken')
        with self.assertRaises(ValueError):
            ui.save_upload({'name':'input.xlsx','raw':b'not excel'},Path('unused.xlsx'))

    def test_gzip_preserves_json_and_leaves_excel_bytes_unchanged(self):
        handler = ui.Handler.__new__(ui.Handler)
        handler.headers = {'Accept-Encoding':'gzip, deflate, br'}
        headers = {}
        handler.send_response = lambda status: None
        handler.send_header = lambda name,value: headers.update({name:value})
        handler.end_headers = lambda: None
        handler.wfile = io.BytesIO()
        data = {'source':'Kế hoạch', 'values':list(range(1000))}
        handler.send(data)
        self.assertEqual(headers['Content-Encoding'],'gzip')
        self.assertEqual(json.loads(gzip.decompress(handler.wfile.getvalue())),data)
        handler.wfile = io.BytesIO();headers.clear()
        binary = b'PK' * 2000
        handler.send(binary,content_type='application/zip',filename='output.zip')
        self.assertNotIn('Content-Encoding',headers)
        self.assertEqual(handler.wfile.getvalue(),binary)
