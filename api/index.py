# Vercel serverless entry point: POST /api/call runs one stateless SchoolFlow API call.

import json
import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "engine"))

from sfcore import handle  # noqa: E402

MAX_BODY = 4 * 1024 * 1024


def app(environ, start_response):
    if environ.get("REQUEST_METHOD") != "POST":
        start_response("405 Method Not Allowed", [("Content-Type", "text/plain"), ("Allow", "POST")])
        return [b"POST only"]
    try:
        length = int(environ.get("CONTENT_LENGTH") or 0)
        if length > MAX_BODY:
            raise ValueError("That request is too large.")
        payload = json.loads(environ["wsgi.input"].read(length) or b"{}")
        result = handle(payload)
    except Exception as error:  # report every failure to the app as a readable message
        result = {"ok": False, "error": str(error) or error.__class__.__name__}
    body = json.dumps(result, default=str).encode("utf-8")
    start_response("200 OK", [("Content-Type", "application/json; charset=utf-8"),
                              ("Cache-Control", "no-store"), ("Content-Length", str(len(body)))])
    return [body]
