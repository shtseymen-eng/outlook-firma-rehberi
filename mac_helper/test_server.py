import importlib.util
import json
from pathlib import Path
import tempfile
import threading
import unittest
import urllib.request
import urllib.error
from unittest.mock import patch

from server import BridgeServer, MailReader

class BridgeTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        Path(self.temp.name, 'index.html').write_text('safe ui')
        self.reader = lambda command, args: {'accounts': []}
        self.server = BridgeServer(('127.0.0.1', 0), self.temp.name, self.reader)
        self.origin = f'http://127.0.0.1:{self.server.server_port}'
        self.thread = threading.Thread(target=self.server.serve_forever, daemon=True)
        self.thread.start()
    def tearDown(self):
        self.server.shutdown(); self.server.server_close(); self.temp.cleanup()
    def request(self, path='/api/accounts', token=None, origin=None, host=None, payload=None):
        headers = {'Origin': origin or self.origin, 'Content-Type': 'application/json'}
        if token is not None: headers['Authorization'] = 'Bearer ' + token
        if host: headers['Host'] = host
        request = urllib.request.Request(self.origin+path, headers=headers, data=json.dumps(payload or {}).encode(), method='POST')
        try:
            with urllib.request.urlopen(request) as response: return response.status, json.load(response)
        except urllib.error.HTTPError as error:
            with error: return error.code, json.load(error)
    def test_requires_secret_and_exact_origin(self):
        self.assertEqual(self.request()[0], 401)
        self.assertEqual(self.request(token=self.server.token, origin='https://evil.example')[0], 403)
        self.assertEqual(self.request(token=self.server.token, host='evil.example')[0], 403)
        self.assertEqual(self.request(token=self.server.token), (200, {'accounts': []}))
    def test_only_fixed_commands_and_bounded_args(self):
        self.assertEqual(self.request('/api/execute', self.server.token)[0], 404)
        self.assertEqual(self.request('/api/messages', self.server.token, payload={'account':'a','path':['inbox'],'ids':list(range(101))})[0],400)
        self.assertEqual(self.request('/api/folders', self.server.token, payload={'account':5})[0],400)
    def test_errors_do_not_leak_mail_or_files(self):
        self.server.reader = lambda *_: (_ for _ in ()).throw(RuntimeError('private@example.org /Users/private'))
        status, body = self.request(token=self.server.token)
        self.assertEqual(status, 502)
        self.assertNotIn('private', json.dumps(body))
    def test_reader_uses_fixed_script_and_stdin_no_shell(self):
        with patch('server.subprocess.run') as run:
            run.return_value.stdout = '{"accounts":[]}'
            run.return_value.returncode = 0
            self.assertEqual(MailReader()('accounts', {}), {'accounts': []})
            args, kwargs = run.call_args
            self.assertEqual(args[0][:3], ['/usr/bin/osascript','-l','JavaScript'])
            self.assertNotIn('shell', kwargs)
            self.assertEqual(json.loads(kwargs['input'])['command'], 'accounts')

if __name__ == '__main__': unittest.main()
