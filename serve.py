"""Tiny static server for previewing the site locally (mimics GitHub Pages).

Serves the repo root so the page loads at http://127.0.0.1:8778/ne-ice/, the
same path it will have on GitHub Pages. Computes the root from __file__ because
os.getcwd() is not always permitted in the preview sandbox.
"""

import functools
import http.server
import os
import socketserver

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
PORT = 8778


class Handler(http.server.SimpleHTTPRequestHandler):
    def end_headers(self):
        self.send_header("Cache-Control", "no-store")
        super().end_headers()

    def log_message(self, fmt, *args):
        print("%s - %s" % (self.address_string(), fmt % args), flush=True)


if __name__ == "__main__":
    socketserver.TCPServer.allow_reuse_address = True
    handler = functools.partial(Handler, directory=ROOT)
    with socketserver.TCPServer(("127.0.0.1", PORT), handler) as httpd:
        print(f"serving {ROOT} at http://127.0.0.1:{PORT}/ne-ice/", flush=True)
        httpd.serve_forever()
