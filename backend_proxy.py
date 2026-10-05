"""Forward all session operations to one durable Shipment Studio backend."""
import http.client
import os
import cloud_store
from urllib.parse import urlsplit


def forward_api(handler):
    if not handler.path.startswith('/api/'):
        return False
    if cloud_store.enabled():
        return False
    address = os.environ.get('SHIPMENT_BACKEND_URL', '').strip()
    if not address:
        if not os.environ.get('VERCEL'):
            return False
        handler.send({'code': 'BACKEND_NOT_CONFIGURED', 'error':
                      'Chưa cấu hình kho lưu trữ cloud. Cần SHIPMENT_S3_BUCKET và thông tin kết nối S3 '
                      'trên Vercel để giữ phiên và các file Excel.'}, 503)
        return True

    endpoint = urlsplit(address)
    token = os.environ.get('SHIPMENT_BACKEND_TOKEN', '')
    if (endpoint.scheme != 'https' or not endpoint.hostname or endpoint.username or endpoint.password
            or endpoint.query or endpoint.fragment or endpoint.path not in ('', '/') or len(token) < 32):
        handler.send({'code': 'BACKEND_CONFIG_INVALID', 'error':
                      'Cấu hình máy xử lý không hợp lệ: cần URL gốc HTTPS và khóa ít nhất 32 ký tự.'}, 503)
        return True

    connection = None
    started = False
    try:
        body = None
        if handler.command in ('POST', 'DELETE'):
            length = int(handler.headers.get('Content-Length', '0'))
            if not 0 <= length <= 125 * 1024 * 1024:
                raise ValueError('Dữ liệu tải lên quá lớn.')
            body = handler.rfile.read(length)
            if len(body) != length:
                raise ValueError('Dữ liệu tải lên chưa đầy đủ.')
        headers = {'Authorization': 'Bearer ' + token}
        for name in ('Content-Type', 'Accept', 'Accept-Encoding'):
            if handler.headers.get(name):
                headers[name] = handler.headers[name]
        connection = http.client.HTTPSConnection(endpoint.hostname, endpoint.port or 443, timeout=300)
        connection.request(handler.command, handler.path, body=body, headers=headers)
        response = connection.getresponse()
        # Never follow redirects with the backend credential.
        if 300 <= response.status < 400:
            raise ValueError('Máy xử lý trả chuyển hướng. Hãy cấu hình URL HTTPS cuối cùng.')
        handler.send_response(response.status)
        for name in ('Content-Type', 'Content-Length', 'Content-Encoding', 'Content-Disposition', 'Vary'):
            value = response.getheader(name)
            if value is not None:
                handler.send_header(name, value)
        handler.send_header('Cache-Control', 'no-store, no-transform')
        handler.send_header('X-Content-Type-Options', 'nosniff')
        handler.end_headers()
        started = True
        while True:
            chunk = response.read1(64 * 1024)
            if not chunk:
                break
            handler.wfile.write(chunk)
            handler.wfile.flush()
    except (OSError, ValueError, http.client.HTTPException):
        if not started:
            handler.send({'code': 'BACKEND_UNAVAILABLE', 'error':
                          'Không liên lạc được máy Windows xử lý. Kiểm tra máy đang bật, '
                          'dịch vụ Shipment Studio và kết nối HTTPS rồi thử lại.'}, 502)
        # After headers, close the stream. The client treats truncation as failure.
    finally:
        if connection is not None:
            connection.close()
        handler.close_connection = True
    return True
