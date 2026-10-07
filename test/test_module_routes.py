import re
import threading
import unittest
from http.server import ThreadingHTTPServer
from urllib.error import HTTPError
from urllib.request import urlopen

import ui_server


class ModuleRouteTests(unittest.TestCase):
    def test_entrypoint_and_all_imports_are_served_as_javascript(self):
        server = ThreadingHTTPServer(('127.0.0.1', 0), ui_server.Handler)
        thread = threading.Thread(target=server.serve_forever, daemon=True)
        thread.start()
        base = f'http://127.0.0.1:{server.server_port}'
        try:
            with urlopen(base + '/') as response:
                html = response.read().decode()
                self.assertIn('/app.js?v=28', html)
                self.assertNotIn('type="module"', html)
                self.assertLess(html.index('id="grid-container"'), html.index('id="sheet-tabs"'))
                self.assertLess(html.index('id="sheet-tabs"'), html.index('class="grid-footer"'))
            with urlopen(base + '/app.js') as response:
                entry = response.read().decode()
                self.assertIn('createWorkbookController', entry)
                self.assertNotIn("from './modules/", entry)
            modules = ['state.mjs', 'api.mjs', 'file-list.mjs', 'grid-view.mjs',
                       'preview-view.mjs', 'workbook-controller.mjs', 'week-picker.mjs']
            for module in modules:
                with urlopen(base + '/modules/' + module) as response:
                    self.assertEqual(response.headers.get_content_type(), 'text/javascript')
                    self.assertIn('export ', response.read().decode())
            with self.assertRaises(HTTPError):
                urlopen(base + '/modules/../../process.py')
        finally:
            server.shutdown()
            server.server_close()
            thread.join()
