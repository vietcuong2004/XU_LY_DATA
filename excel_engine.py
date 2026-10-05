"""Recalculate and edit a staged workbook with desktop Microsoft Excel."""
import json
import io
import os
from pathlib import Path
import shutil
import subprocess
import threading
from urllib.request import Request, build_opener, HTTPRedirectHandler
from urllib.parse import urlparse
from urllib.error import HTTPError, URLError
from zipfile import ZipFile, ZIP_STORED, BadZipFile

LOCK = threading.Lock()
SCRIPT = Path(__file__).with_name('excel_recalculate.ps1')


class NoRedirect(HTTPRedirectHandler):
    def redirect_request(self, req, fp, code, msg, headers, newurl):
        raise ValueError('Worker không được chuyển hướng sang địa chỉ khác.')


def local_available():
    available = False
    if os.name == 'nt' and not os.environ.get('VERCEL'):
        import winreg
        try:
            with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, r'Excel.Application\CLSID'):
                available = bool(shutil.which('powershell'))
        except OSError:
            pass
    return available


def capability():
    remote = bool(os.environ.get('EXCEL_WORKER_URL') and os.environ.get('EXCEL_WORKER_TOKEN'))
    available = remote or local_available()
    return {'available': available, 'engine': 'Microsoft Excel (Windows worker)' if remote else 'Microsoft Excel' if available else None,
            'message': ('Sửa mọi sheet, tính lại toàn bộ workbook và tạo lại các Family.' if available else
                        'Chưa kết nối máy Windows xử lý Excel. Cần cấu hình EXCEL_WORKER_URL và EXCEL_WORKER_TOKEN trên máy chủ.')}


def edit_and_calculate(path, edits):
    if not capability()['available']:
        raise ValueError(capability()['message'])
    path = Path(path).resolve()
    worker = os.environ.get('EXCEL_WORKER_URL', '').rstrip('/')
    if worker:
        parsed = urlparse(worker)
        if parsed.scheme != 'https' and not (parsed.scheme == 'http' and parsed.hostname in ('localhost', '127.0.0.1')):
            raise ValueError('EXCEL_WORKER_URL phải dùng HTTPS (ngoại trừ localhost).')
        payload = io.BytesIO()
        with ZipFile(payload, 'w', ZIP_STORED) as archive:
            archive.write(path, 'input.xlsx')
            archive.writestr('edits.json', json.dumps(edits, ensure_ascii=False))
        request = Request(worker + '/recalculate', data=payload.getvalue(), headers={
            'Authorization': 'Bearer ' + os.environ['EXCEL_WORKER_TOKEN'],
            'Content-Type': 'application/zip'}, method='POST')
        try:
            with build_opener(NoRedirect).open(request, timeout=240) as response:
                result = response.read(46 * 1024 * 1024)
            with ZipFile(io.BytesIO(result)) as archive:
                if 'xl/workbook.xml' not in archive.namelist():
                    raise ValueError('Worker không trả về workbook Excel hợp lệ.')
            path.write_bytes(result)
            return
        except (HTTPError, URLError, TimeoutError, BadZipFile) as exc:
            raise ValueError('Không nhận được kết quả từ máy Windows. Phiên trước vẫn giữ nguyên.') from exc
    edit_and_calculate_local(path, edits)


def edit_and_calculate_local(path, edits):
    if not local_available():
        raise ValueError('Windows worker chưa cài hoặc chưa đăng ký Microsoft Excel.')
    path = Path(path).resolve()
    request = path.with_suffix('.edit.json')
    request.write_text(json.dumps({'path': str(path), 'edits': edits}, ensure_ascii=False), encoding='utf-8')
    try:
        with LOCK:
            try:
                result = subprocess.run(
                    ['powershell', '-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass',
                     '-File', str(SCRIPT), '-RequestPath', str(request)],
                    capture_output=True, text=True, encoding='utf-8', errors='replace', timeout=240,
                    creationflags=subprocess.CREATE_NO_WINDOW)
            except subprocess.TimeoutExpired as exc:
                raise ValueError('Excel tính quá lâu. Chưa công bố thay đổi; phiên trước vẫn giữ nguyên.') from exc
            if result.returncode:
                raise ValueError('Excel chưa tính và lưu được workbook: ' + (result.stderr or result.stdout)[-2000:])
    finally:
        record = Path(str(request) + '.process.json')
        if record.exists():
            # Only the newly created Excel PID with the same start time may be
            # cleaned up. Never stop an Excel process already open by the user.
            subprocess.run(['powershell', '-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass',
                            '-File', str(SCRIPT.with_name('excel_cleanup.ps1')), '-RecordPath', str(record)],
                           capture_output=True, timeout=15, creationflags=subprocess.CREATE_NO_WINDOW)
            record.unlink(missing_ok=True)
        request.unlink(missing_ok=True)
