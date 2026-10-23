"""A local, synthetic, three-page API that deliberately returns 429 and 503."""
import argparse
import csv
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
import json
from pathlib import Path
from threading import Thread
from urllib.parse import parse_qs, urlsplit

from .cli import main as pipeline


def serve(path):
    with path.open(newline='', encoding='utf-8') as stream:
        rows = list(csv.DictReader(stream))
    counts = {}

    class Handler(BaseHTTPRequestHandler):
        def log_message(self, *_args):
            pass

        def do_GET(self):
            parsed = urlsplit(self.path)
            if parsed.path != '/orders':
                self.send_error(404)
                return
            try:
                page = int(parse_qs(parsed.query).get('page', ['1'])[0])
            except ValueError:
                self.send_error(400)
                return
            if page not in (1, 2, 3):
                self.send_error(400)
                return
            counts[page] = counts.get(page, 0) + 1
            if counts[page] == 1 and page in (1, 2):
                self.send_response(429 if page == 1 else 503)
                self.send_header('Retry-After', '0')
                self.end_headers()
                return
            body = json.dumps({'items': rows[(page - 1) * 6:page * 6], 'next_page': page + 1 if page < 3 else None}).encode()
            self.send_response(200)
            self.send_header('Content-Type', 'application/json')
            self.send_header('Content-Length', str(len(body)))
            self.end_headers()
            self.wfile.write(body)

    return ThreadingHTTPServer(('127.0.0.1', 0), Handler)


def main():
    parser = argparse.ArgumentParser(description='Run the synthetic API demo twice to prove duplicate protection')
    parser.add_argument('--output', type=Path, default=Path('output/demo'))
    args = parser.parse_args()
    sample = Path(__file__).resolve().parent.parent / 'examples/orders.csv'
    if not sample.exists():
        parser.error('Run the demo from the source checkout installed with pip install -e .')
    server = serve(sample)
    thread = Thread(target=server.serve_forever, daemon=True)
    thread.start()
    try:
        common = ['--db', str(args.output / 'orders.sqlite3'), '--output', str(args.output / 'reports'), '--log', str(args.output / 'pipeline.log'),
                  'import', '--api', f'http://127.0.0.1:{server.server_port}/orders']
        print('FIRST RUN: paginated API with simulated 429 and 503 errors')
        if pipeline(common):
            return 1
        print('SECOND RUN: same input, no new orders should be inserted')
        return pipeline(common)
    finally:
        server.shutdown()
        server.server_close()
        thread.join()


if __name__ == '__main__':
    raise SystemExit(main())
