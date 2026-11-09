"""Loopback-only Mac Mail reader. No external dependencies, disk mailbox cache or writes."""
import functools
import hmac
from http.server import ThreadingHTTPServer, SimpleHTTPRequestHandler
import json
from pathlib import Path
import secrets
import subprocess
import sys
import threading
import webbrowser

class MailError(Exception):
    pass

class MailReader:
    def __init__(self):
        self.lock = threading.Lock()
    def __call__(self, command, args):
        # Script and command set are fixed; data goes over stdin, never shell interpolation.
        with self.lock:
            try:
                result = subprocess.run(['/usr/bin/osascript', '-l', 'JavaScript',
                    str(Path(__file__).with_name('mail_reader.js'))],
                    input=json.dumps({'command': command, **args}), text=True,
                    capture_output=True, timeout=180)
            except subprocess.TimeoutExpired:
                raise MailError('Mail yanıt vermedi. Mail’i açıp yeniden deneyin; gerekirse klasörü küçültün.')
        if result.returncode:
            if '-1743' in result.stderr:
                raise MailError('Mail erişimine izin verilmedi. Sistem Ayarları → Gizlilik ve Güvenlik → Otomasyon bölümünden Mac Mail Rehberi için Mail iznini açın.')
            raise MailError('Mail okunamadı. Mail’i açın, hesabın yüklendiğini kontrol edin ve yeniden deneyin.')
        try:
            value = json.loads(result.stdout)
        except (ValueError, TypeError):
            raise MailError('Mail yanıtı okunamadı. Klasör taramasını yeniden deneyin.')
        return value

class BridgeServer(ThreadingHTTPServer):
    daemon_threads = True
    def __init__(self, address, directory, reader):
        if address[0] != '127.0.0.1': raise ValueError('Only loopback is allowed')
        self.token = secrets.token_hex(32)
        self.reader = reader
        super().__init__(address, functools.partial(Handler, directory=str(directory)))
        self.origin = f'http://127.0.0.1:{self.server_port}'

class Handler(SimpleHTTPRequestHandler):
    def log_message(self, *_): pass  # Never log message data, paths or pairing secrets.
    def end_headers(self):
        self.send_header('Cache-Control', 'no-store')
        self.send_header('Referrer-Policy', 'no-referrer')
        self.send_header('X-Content-Type-Options', 'nosniff')
        super().end_headers()
    def host_ok(self): return self.headers.get('Host') == self.server.origin.removeprefix('http://')
    def reply(self, status, value):
        data = json.dumps(value, ensure_ascii=False).encode()
        self.send_response(status); self.send_header('Content-Type','application/json; charset=utf-8')
        self.send_header('Content-Length',str(len(data))); self.end_headers(); self.wfile.write(data)
    def do_GET(self):
        if not self.host_ok(): return self.reply(403, {'error':'Geçersiz bağlantı.'})
        # Static bundle only, no directory indexes, hidden files, source or symlinks.
        target = (Path(self.directory) / self.path.split('?',1)[0].lstrip('/')).resolve()
        root = Path(self.directory).resolve()
        if target == root: self.path = '/index.html'
        elif not target.is_relative_to(root) or not target.is_file() or any(p.startswith('.') for p in target.relative_to(root).parts):
            return self.reply(404, {'error':'Bulunamadı.'})
        return super().do_GET()
    def do_OPTIONS(self): self.reply(403, {'error':'Harici sayfalara erişim kapalı.'})
    def do_POST(self):
        if not self.host_ok() or self.headers.get('Origin') != self.server.origin:
            return self.reply(403, {'error':'Yalnızca yerel rehber ekranı bağlanabilir.'})
        authorization = self.headers.get('Authorization','')
        if not hmac.compare_digest(authorization, 'Bearer '+self.server.token):
            return self.reply(401, {'error':'Yardımcı bağlantısı geçersiz. Yardımcıdan rehber ekranını tekrar açın.'})
        command = {'/api/accounts':'accounts','/api/folders':'folders','/api/snapshot':'snapshot','/api/messages':'messages'}.get(self.path)
        if not command: return self.reply(404, {'error':'İşlem bulunamadı.'})
        try:
            size = int(self.headers.get('Content-Length','0'))
            if not 0 < size <= 16384 or self.headers.get('Content-Type') != 'application/json': raise ValueError()
            payload = json.loads(self.rfile.read(size))
            args = validate(command, payload)
        except (ValueError, TypeError, KeyError): return self.reply(400, {'error':'Geçersiz istek.'})
        try: self.reply(200, self.server.reader(command, args))
        except MailError as error: self.reply(502, {'error':str(error)})
        except Exception: self.reply(502, {'error':'Mail okunamadı. Mail’i açıp yeniden deneyin.'})

def validate(command, payload):
    if not isinstance(payload,dict): raise ValueError()
    if command == 'accounts': return {}
    account = payload['account']
    if not isinstance(account,str) or not 0 < len(account) <= 500: raise ValueError()
    args = {'account':account}
    if command == 'folders': return args
    path = payload['path']
    if not isinstance(path,list) or not 0 < len(path) <= 50 or any(not isinstance(p,str) or not 0 < len(p) <= 1000 for p in path): raise ValueError()
    args['path'] = path
    if command == 'messages':
        ids = payload['ids']
        if not isinstance(ids,list) or not 0 < len(ids) <= 50 or any(type(i) is not int or i < 0 for i in ids): raise ValueError()
        args.update(ids=ids, signatures=payload.get('signatures') is True)
    return args

def main():
    root = Path(__file__).resolve().parent
    server = BridgeServer(('127.0.0.1',0), root/'web', MailReader())
    url = server.origin+'/#mail='+server.token
    # Pipe is read privately by the native launcher, never logged.
    print(url, flush=True)
    server.serve_forever()

if __name__ == '__main__': main()
