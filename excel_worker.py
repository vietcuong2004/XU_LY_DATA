"""Private Windows Excel worker. Place behind an HTTPS reverse proxy.

Set EXCEL_WORKER_TOKEN to a long random secret shared only with Vercel.
Run: python excel_worker.py --host 127.0.0.1 --port 8767
"""
import argparse
import hmac
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
import io
import json
import os
from pathlib import Path
import tempfile
import threading
from zipfile import ZipFile, BadZipFile

import excel_engine

LIMIT = 46 * 1024 * 1024
BUSY = threading.Lock()


class Handler(BaseHTTPRequestHandler):
    def log_message(self, *args):
        pass

    def reply(self, status, body, mime='application/json'):
        if isinstance(body, dict):
            body = json.dumps(body).encode('utf-8')
        self.send_response(status)
        self.send_header('Content-Type', mime)
        self.send_header('Content-Length', str(len(body)))
        self.send_header('Cache-Control', 'no-store')
        self.end_headers()
        self.wfile.write(body)

    def do_POST(self):
        token = os.environ.get('EXCEL_WORKER_TOKEN', '')
        if not token or not hmac.compare_digest(self.headers.get('Authorization', ''), 'Bearer ' + token):
            return self.reply(401, {'error': 'Unauthorized'})
        if self.path != '/recalculate':
            return self.reply(404, {'error': 'Not found'})
        try:
            length = int(self.headers.get('Content-Length', '0'))
            if not 0 < length <= LIMIT:
                return self.reply(413, {'error': 'Payload too large'})
        except ValueError:
            return self.reply(400, {'error': 'Invalid length'})
        if not BUSY.acquire(blocking=False):
            return self.reply(503, {'error': 'Excel is busy. Retry shortly.'})
        try:
            self.connection.settimeout(60)
            raw = self.rfile.read(length)
            with ZipFile(io.BytesIO(raw)) as archive:
                if set(archive.namelist()) != {'input.xlsx', 'edits.json'}:
                    raise ValueError('Invalid package')
                if sum(x.file_size for x in archive.infolist()) > LIMIT:
                    raise ValueError('Package too large')
                edits = json.loads(archive.read('edits.json'))
                if not isinstance(edits, list) or len(edits) > 10000:
                    raise ValueError('Invalid edits')
                with tempfile.TemporaryDirectory(prefix='shipment-excel-') as temp:
                    path = Path(temp)/'input.xlsx'
                    path.write_bytes(archive.read('input.xlsx'))
                    excel_engine.edit_and_calculate_local(path, edits)
                    result = path.read_bytes()
            self.reply(200, result, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        except (ValueError, OSError, BadZipFile, KeyError, TypeError):
            self.reply(422, {'error': 'Excel could not process this workbook. Original session unchanged.'})
        finally:
            BUSY.release()


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--host', default='127.0.0.1')
    parser.add_argument('--port', type=int, default=8767)
    args = parser.parse_args()
    if len(os.environ.get('EXCEL_WORKER_TOKEN', '')) < 32:
        parser.error('Set EXCEL_WORKER_TOKEN to a random secret of at least 32 characters.')
    if not excel_engine.local_available():
        parser.error('Microsoft Excel is required on this Windows machine.')
    ThreadingHTTPServer((args.host, args.port), Handler).serve_forever()
