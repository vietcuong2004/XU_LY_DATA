import io
import json
import hashlib
import os
from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch
from botocore.exceptions import ClientError

import cloud_store as cloud


class MemoryS3:
    def __init__(self): self.objects={}; self.fail_manifest=False
    def put_object(self, **kw):
        key=kw['Key']
        if self.fail_manifest and '/jobs/' in key: raise OSError('connection lost')
        if kw.get('IfNoneMatch')=='*' and key in self.objects:
            raise ClientError({'Error':{'Code':'PreconditionFailed'}},'PutObject')
        if 'IfMatch' in kw and (key not in self.objects or hashlib.md5(self.objects[key]).hexdigest()!=kw['IfMatch']):
            raise ClientError({'Error':{'Code':'PreconditionFailed'}},'PutObject')
        self.objects[key]=kw['Body']
        return {'ETag':hashlib.md5(kw['Body']).hexdigest()}
    def get_object(self, **kw):
        if kw['Key'] not in self.objects:
            raise ClientError({'Error':{'Code':'NoSuchKey'}},'GetObject')
        data=self.objects[kw['Key']]
        return {'Body':io.BytesIO(data),'ContentLength':len(data),'ETag':hashlib.md5(data).hexdigest()}
    def delete_object(self, **kw):
        if 'IfMatch' in kw and kw['Key'] in self.objects and hashlib.md5(self.objects[kw['Key']]).hexdigest()!=kw['IfMatch']:
            raise ClientError({'Error':{'Code':'PreconditionFailed'}},'DeleteObject')
        self.objects.pop(kw['Key'],None)


class CloudTests(unittest.TestCase):
    def setUp(self):
        self.s3=MemoryS3()
        p=patch.object(cloud,'client',return_value=self.s3);p.start();self.addCleanup(p.stop)
        p=patch.dict(os.environ,{'SHIPMENT_S3_BUCKET':'test','SHIPMENT_SESSION_SECRET':'x'*48});p.start();self.addCleanup(p.stop)
        token=cloud.WORKSPACE.set('test');self.addCleanup(cloud.WORKSPACE.reset,token)

    def test_cold_instance_restores_source_outputs_and_report(self):
        with tempfile.TemporaryDirectory() as temp:
            a=Path(temp)/'one';a.mkdir()
            state={'id':'a'*32,'status':'ready'}
            (a/'job.json').write_text(json.dumps(state))
            (a/'input.xlsx').write_bytes(b'original-excel-bytes')
            (a/'review.json').write_text('{"families":[]}')
            cloud.publish(a,state)
            b=Path(temp)/'two'
            self.assertTrue(cloud.restore(state['id'],b))
            self.assertEqual((b/'input.xlsx').read_bytes(),b'original-excel-bytes')
            cloud.delete(state['id'])
            self.assertFalse(cloud.restore(state['id'],b))

    def test_failed_upload_does_not_publish_ready_session(self):
        with tempfile.TemporaryDirectory() as temp:
            a=Path(temp);state={'id':'b'*32,'status':'ready'}
            (a/'job.json').write_text(json.dumps(state))
            self.s3.fail_manifest=True
            with self.assertRaises(OSError):cloud.publish(a,state)
            self.assertIsNone(cloud.read(state['id']))

    def test_processing_progress_survives_cold_worker_without_snapshot(self):
        state={'id':'c'*32,'status':'processing','completed':5,'total':26}
        cloud.progress(state)
        self.assertEqual({k:cloud.read(state['id'])[k] for k in state},state)
        cloud.progress({**state,'status':'ready'})
        self.assertEqual(cloud.read(state['id'])['status'],'processing')

    def test_stopped_worker_does_not_leave_an_infinite_progress_screen(self):
        with patch.object(cloud.time,'time',return_value=100):
            cloud.progress({'id':'c'*32,'status':'processing','completed':5,'total':26})
        with patch.object(cloud.time,'time',return_value=500):
            self.assertEqual(cloud.read('c'*32)['status'],'error')

    def test_browser_upload_cannot_read_another_workspace_key(self):
        first=cloud.key('uploads/'+'a'*32+'.xlsx')
        token=cloud.WORKSPACE.set('other')
        try:
            with self.assertRaises(ValueError):cloud.get_upload(first,'file.xlsx')
        finally:cloud.WORKSPACE.reset(token)

    def test_two_editors_cannot_write_same_parent_concurrently(self):
        with cloud.edit_lock('a'*32):
            with self.assertRaises(ValueError):
                with cloud.edit_lock('a'*32):pass
        with cloud.edit_lock('a'*32):pass

    def test_crashed_worker_lease_can_be_recovered_after_expiry(self):
        self.s3.objects[cloud.key('locks/'+'a'*32)]=b'{"owner":"crashed","until":0}'
        with cloud.edit_lock('a'*32):
            self.assertNotIn(b'crashed',self.s3.objects[cloud.key('locks/'+'a'*32)])
        self.assertNotIn(cloud.key('locks/'+'a'*32),self.s3.objects)

    def test_workspace_cookie_is_signed_and_scopes_storage(self):
        class Handler: headers={}
        h=Handler();cloud.bind_workspace(h)
        first=cloud.key('jobs/a');cookie=h.workspace_cookie.split(';')[0]
        h=Handler();h.headers={'Cookie':cookie};cloud.bind_workspace(h)
        self.assertEqual(cloud.key('jobs/a'),first)
        h=Handler();h.headers={'Cookie':cookie+'tampered'};cloud.bind_workspace(h)
        self.assertNotEqual(cloud.key('jobs/a'),first)


if __name__=='__main__':unittest.main()
