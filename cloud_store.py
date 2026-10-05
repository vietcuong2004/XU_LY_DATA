"""Private S3 storage for immutable session snapshots and direct transfers.

Cloud is optional locally; Vercel requires a configured bucket. AWS credentials
use boto3's normal provider chain and never reach the browser.
"""
import io
import json
import os
from pathlib import Path
import re
import shutil
import tempfile
import time
import uuid
from contextlib import contextmanager
from contextvars import ContextVar
import hashlib
import hmac
from http.cookies import SimpleCookie

WORKSPACE = ContextVar('shipment_workspace', default='local')


def bind_workspace(handler):
    secret = os.environ.get('SHIPMENT_SESSION_SECRET', '')
    if len(secret) < 32:
        raise ValueError('Cần cấu hình SHIPMENT_SESSION_SECRET ngẫu nhiên ít nhất 32 ký tự để tách phiên người dùng.')
    cookies = SimpleCookie()
    try:
        cookies.load(handler.headers.get('Cookie', ''))
        token = cookies['shipment_workspace'].value if 'shipment_workspace' in cookies else ''
    except Exception:
        token = ''
    parts = token.split('.')
    valid = len(parts) == 2 and re.fullmatch('[a-f0-9]{32}', parts[0]) and hmac.compare_digest(
        hmac.new(secret.encode(), parts[0].encode(), hashlib.sha256).hexdigest(), parts[1])
    workspace = parts[0] if valid else uuid.uuid4().hex
    WORKSPACE.set(workspace)
    if not valid:
        signature = hmac.new(secret.encode(), workspace.encode(), hashlib.sha256).hexdigest()
        secure = '; Secure' if os.environ.get('VERCEL') or handler.headers.get('X-Forwarded-Proto') == 'https' else ''
        handler.workspace_cookie = f'shipment_workspace={workspace}.{signature}; HttpOnly; SameSite=Strict; Path=/; Max-Age=31536000{secure}'
from zipfile import ZipFile, ZIP_DEFLATED


def enabled():
    return bool(os.environ.get('SHIPMENT_S3_BUCKET'))


def client():
    upload_origin()
    import boto3
    from botocore.config import Config
    access = os.environ.get('SHIPMENT_S3_ACCESS_KEY_ID')
    secret = os.environ.get('SHIPMENT_S3_SECRET_ACCESS_KEY')
    if os.environ.get('VERCEL') and not (access and secret):
        raise ValueError('Cần SHIPMENT_S3_ACCESS_KEY_ID và SHIPMENT_S3_SECRET_ACCESS_KEY trên Vercel.')
    return boto3.client('s3', endpoint_url=os.environ.get('SHIPMENT_S3_ENDPOINT') or None,
                        region_name=os.environ.get('SHIPMENT_S3_REGION', 'us-east-1'),
                        aws_access_key_id=access, aws_secret_access_key=secret,
                        aws_session_token=os.environ.get('SHIPMENT_S3_SESSION_TOKEN'),
                        config=Config(signature_version='s3v4', retries={'max_attempts': 3},
                                      s3={'addressing_style': 'path' if os.environ.get('SHIPMENT_S3_ENDPOINT') else 'virtual'}))


def upload_origin():
    from urllib.parse import urlsplit
    endpoint = os.environ.get('SHIPMENT_S3_ENDPOINT', '')
    if endpoint:
        parsed = urlsplit(endpoint)
        if parsed.scheme != 'https' or not parsed.hostname or parsed.username or parsed.password:
            raise ValueError('SHIPMENT_S3_ENDPOINT phải là địa chỉ HTTPS hợp lệ.')
        return f'https://{parsed.netloc}'
    return 'https://*.amazonaws.com'


def bucket():
    return os.environ['SHIPMENT_S3_BUCKET']


def key(name):
    category, _, rest = name.partition('/')
    return os.environ.get('SHIPMENT_S3_PREFIX', 'shipment').strip('/') + '/' + category + '/' + WORKSPACE.get() + '/' + rest


def job_key(job):
    if not re.fullmatch('[a-f0-9]{32}', job):
        raise ValueError('Mã phiên không hợp lệ.')
    return key('jobs/'+job+'.json')


def read(job):
    from botocore.exceptions import ClientError
    try:
        response = client().get_object(Bucket=bucket(), Key=job_key(job))
        try:
            state = json.loads(response['Body'].read())
        finally:
            response['Body'].close()
        if state.get('status') in ('queued', 'processing') and time.time()-state.get('heartbeat', time.time()) > 360:
            state = {**state, 'status': 'error', 'message':
                     'Phiên xử lý bị gián đoạn hoặc vượt thời gian cho phép. Hãy tạo lại từ file nguồn; phiên đã lưu trước đó vẫn được giữ.'}
        return state
    except ClientError as exc:
        if exc.response['Error']['Code'] in ('NoSuchKey', '404'):
            return None
        raise


def publish(folder, state):
    """Upload the complete snapshot first; only then expose a ready manifest."""
    data = io.BytesIO()
    with ZipFile(data, 'w', ZIP_DEFLATED) as archive:
        for path in folder.rglob('*'):
            if path.is_file() and not path.name.endswith('.tmp') and path.name != '.cloud-version':
                archive.write(path, path.relative_to(folder).as_posix())
    archive_key = key('snapshots/'+state['id']+'/'+uuid.uuid4().hex+'.zip')
    s3 = client()
    s3.put_object(Bucket=bucket(), Key=archive_key, Body=data.getvalue(), ContentType='application/zip')
    manifest = {**state, 'snapshot': archive_key}
    s3.put_object(Bucket=bucket(), Key=job_key(state['id']), Body=json.dumps(manifest).encode(), ContentType='application/json')


def progress(state):
    """Ready states require a committed snapshot; other states are resumable."""
    if state.get('status') == 'ready':
        return
    client().put_object(Bucket=bucket(), Key=job_key(state['id']),
                        Body=json.dumps({**state, 'heartbeat': time.time()}).encode(), ContentType='application/json')


def restore(job, folder):
    manifest = read(job)
    if not manifest or not manifest.get('snapshot'):
        return False
    # Ready snapshots are immutable for a given generation. Refresh when changed.
    marker = folder/'.cloud-version'
    if marker.exists() and marker.read_text() == manifest['snapshot'] and (folder/'job.json').exists():
        return True
    data = client().get_object(Bucket=bucket(), Key=manifest['snapshot'])['Body'].read()
    folder.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory(dir=folder.parent) as temp:
        staging = Path(temp)
        with ZipFile(io.BytesIO(data)) as archive:
            if sum(item.file_size for item in archive.infolist()) > 1024*1024*1024:
                raise ValueError('Phiên vượt giới hạn giải nén.')
            for item in archive.infolist():
                dest = (staging/item.filename).resolve()
                if not dest.is_relative_to(staging.resolve()) or '\\' in item.filename:
                    raise ValueError('Đường dẫn phiên không hợp lệ.')
            archive.extractall(staging)
        if not (staging/'job.json').exists():
            raise ValueError('Snapshot phiên chưa đầy đủ.')
        folder.mkdir(parents=True, exist_ok=True)
        shutil.copytree(staging, folder, dirs_exist_ok=True)
        marker.write_text(manifest['snapshot'])
    return True


def list_jobs():
    s3 = client(); jobs = []
    for page in s3.get_paginator('list_objects_v2').paginate(Bucket=bucket(), Prefix=key('jobs/')):
        for obj in page.get('Contents', []):
            job = obj['Key'].rsplit('/', 1)[-1].removesuffix('.json')
            if re.fullmatch('[a-f0-9]{32}', job):
                state = read(job)
                if state:
                    jobs.append(state)
    return sorted(jobs, key=lambda j: j.get('created', ''), reverse=True)


def delete(job):
    # Remove visibility first; snapshots can be lifecycle-expired in the bucket.
    client().delete_object(Bucket=bucket(), Key=job_key(job))


@contextmanager
def edit_lock(job):
    if not enabled():
        yield
        return
    from botocore.exceptions import ClientError
    job_key(job)
    s3 = client(); lock = key('locks/'+job)
    # The deployed function is capped at 300 seconds; a crashed invocation's
    # lease is safely recoverable after 600 seconds.
    body = json.dumps({'owner': uuid.uuid4().hex, 'until': time.time()+600}).encode()
    try:
        result = s3.put_object(Bucket=bucket(), Key=lock, Body=body, IfNoneMatch='*')
    except ClientError as exc:
        if exc.response['Error']['Code'] in ('PreconditionFailed', 'ConditionalRequestConflict', '412', '409'):
            current = s3.get_object(Bucket=bucket(), Key=lock)
            if json.loads(current['Body'].read())['until'] > time.time():
                raise ValueError('Phiên đang được lưu ở cửa sổ khác. Hãy đợi rồi mở phiên mới nhất.') from exc
            try:
                result = s3.put_object(Bucket=bucket(), Key=lock, Body=body, IfMatch=current['ETag'])
            except ClientError as conflict:
                raise ValueError('Một cửa sổ khác vừa bắt đầu lưu. Hãy đợi rồi thử lại.') from conflict
        else:
            raise
    try:
        yield
    finally:
        try:
            s3.delete_object(Bucket=bucket(), Key=lock, IfMatch=result['ETag'])
        except (ClientError, OSError):
            # Never delete a newer owner's lease; expired leases are recoverable.
            pass


def sign_upload(name, size):
    if not str(name).lower().endswith('.xlsx') or type(size) is not int or not 0 < size <= 45*1024*1024:
        raise ValueError('Chọn workbook .xlsx tối đa 45 MB.')
    upload_key = key('uploads/'+uuid.uuid4().hex+'.xlsx')
    url = client().generate_presigned_url('put_object', Params={
        'Bucket': bucket(), 'Key': upload_key, 'ContentType': 'application/octet-stream'}, ExpiresIn=600)
    return {'key': upload_key, 'url': url}


def get_upload(upload_key, name):
    if not isinstance(upload_key, str) or not re.fullmatch(re.escape(key('uploads/'))+'[a-f0-9]{32}\\.xlsx', upload_key):
        raise ValueError('Mã tải lên không hợp lệ.')
    response = client().get_object(Bucket=bucket(), Key=upload_key)
    if response['ContentLength'] > 45*1024*1024:
        response['Body'].close()
        raise ValueError('File vượt 45 MB.')
    try:
        raw = response['Body'].read(45*1024*1024+1)
    finally:
        response['Body'].close()
    if len(raw) > 45*1024*1024:
        raise ValueError('File vượt 45 MB.')
    return {'name': name, 'raw': raw}


def download(data, name, content_type):
    from urllib.parse import quote
    download_key = key('downloads/'+uuid.uuid4().hex)
    s3 = client()
    disposition = "attachment; filename*=UTF-8''" + quote(name)
    s3.put_object(Bucket=bucket(), Key=download_key, Body=data, ContentType=content_type,
                  ContentDisposition=disposition)
    return s3.generate_presigned_url('get_object', Params={'Bucket': bucket(), 'Key': download_key}, ExpiresIn=600)
