"""SchoolFlow website — local server.

Serves the landing page at /, the app at /app/, and the stateless API at
POST /api/call (the same code the online version runs on Vercel). Each
browser keeps its own school data; the server stores nothing.

    python server.py                 # http://127.0.0.1:8000
    python server.py --port 9000
    python server.py --host 0.0.0.0  # let other computers on the network in
"""

import argparse
import json
import mimetypes
import os
import sys
import threading
import urllib.parse
import webbrowser
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

HERE = os.path.dirname(os.path.abspath(__file__))
SITE = os.path.join(HERE, "site")
sys.path.insert(0, os.path.join(HERE, "engine"))

from sfcore import handle  # noqa: E402

MAX_BODY = 8 * 1024 * 1024
mimetypes.add_type("text/javascript", ".js")
mimetypes.add_type("text/css", ".css")
mimetypes.add_type("image/svg+xml", ".svg")


class Handler(BaseHTTPRequestHandler):
    server_version = "SchoolFlow"

    def log_message(self, fmt, *args):
        if os.environ.get("SCHOOLFLOW_VERBOSE"):
            super().log_message(fmt, *args)

    def _send(self, status, body, content_type, extra=None):
        self.send_response(status)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        for key, value in (extra or {}).items():
            self.send_header(key, value)
        self.end_headers()
        if self.command != "HEAD":
            self.wfile.write(body)

    def do_GET(self):
        path = urllib.parse.unquote(urllib.parse.urlparse(self.path).path)
        if path == "/app":
            return self._send(HTTPStatus.MOVED_PERMANENTLY, b"", "text/plain", {"Location": "/app/"})
        relative = path.lstrip("/")
        if relative == "" or relative.endswith("/"):
            relative += "index.html"
        file = os.path.realpath(os.path.join(SITE, *relative.split("/")))
        if not file.startswith(os.path.realpath(SITE) + os.sep) or not os.path.isfile(file):
            return self._send(HTTPStatus.NOT_FOUND, b"Not found", "text/plain")
        with open(file, "rb") as handle_:
            body = handle_.read()
        self._send(HTTPStatus.OK, body, mimetypes.guess_type(file)[0] or "application/octet-stream")

    do_HEAD = do_GET

    def do_POST(self):
        if urllib.parse.urlparse(self.path).path != "/api/call":
            return self._send(HTTPStatus.NOT_FOUND, b"Not found", "text/plain")
        try:
            length = int(self.headers.get("Content-Length") or 0)
            if length > MAX_BODY:
                raise ValueError("That request is too large.")
            result = handle(json.loads(self.rfile.read(length) or b"{}"))
        except Exception as error:
            result = {"ok": False, "error": str(error) or error.__class__.__name__}
        self._send(HTTPStatus.OK, json.dumps(result, default=str).encode("utf-8"), "application/json; charset=utf-8")


def main():
    parser = argparse.ArgumentParser(description="Run the SchoolFlow website locally.")
    parser.add_argument("--host", default=os.environ.get("HOST", "127.0.0.1"))
    parser.add_argument("--port", type=int, default=int(os.environ.get("PORT", "8000")))
    parser.add_argument("--no-browser", action="store_true", help="don't open a browser tab")
    options = parser.parse_args()

    server = ThreadingHTTPServer((options.host, options.port), Handler)
    address = "http://" + ("127.0.0.1" if options.host in ("0.0.0.0", "") else options.host) + ":" + str(options.port) + "/"
    print("SchoolFlow is running at " + address)
    print("Press Ctrl+C to stop.")
    if not options.no_browser:
        threading.Timer(0.8, lambda: webbrowser.open(address)).start()
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\nStopped.")
    finally:
        server.server_close()


if __name__ == "__main__":
    main()
