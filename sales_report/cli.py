"""Command-line entry point; credentials come from the environment only."""
import argparse
import csv
import json
import logging
from pathlib import Path
import re
import sqlite3
import sys
from zipfile import BadZipFile

from openpyxl.utils.exceptions import InvalidFileException

from .report import generate
from .sources import APIClient, load_file
from .store import connect, import_rows


def main(argv=None):
    parser = argparse.ArgumentParser(description='Import, validate and report sales orders')
    parser.add_argument('--db', type=Path, default=Path('output/orders.sqlite3'))
    parser.add_argument('--output', type=Path, default=Path('output/reports'))
    parser.add_argument('--log', type=Path, default=Path('output/pipeline.log'))
    sub = parser.add_subparsers(dest='command', required=True)
    ingest = sub.add_parser('import', help='Import a file or a paginated REST endpoint')
    source = ingest.add_mutually_exclusive_group(required=True)
    source.add_argument('--file', type=Path)
    source.add_argument('--api', help='Endpoint URL with no query parameters')
    ingest.add_argument('--sheet', help='Excel worksheet name; first worksheet by default')
    ingest.add_argument('--currency', default='USD')
    ingest.add_argument('--token-env', help='Name of an environment variable holding the bearer token')
    ingest.add_argument('--attempts', type=int, default=4)
    ingest.add_argument('--timeout', type=float, default=15)
    report = sub.add_parser('report', help='Regenerate a report using existing data')
    report.add_argument('--run-id', required=True)
    args = parser.parse_args(argv)
    args.log.parent.mkdir(parents=True, exist_ok=True)
    handler = logging.FileHandler(args.log, encoding='utf-8')
    handler.setFormatter(logging.Formatter('%(asctime)s %(levelname)s %(message)s'))
    log = logging.getLogger('sales_report')
    log.setLevel(logging.INFO)
    log.addHandler(handler)
    db, run_id = None, None
    try:
        if args.command == 'import':
            currency = args.currency.upper()
            if not re.fullmatch('[A-Z]{3}', currency):
                raise ValueError('Currency must be a three-letter uppercase code')
            if args.api:
                client = APIClient(args.api, token_env=args.token_env, attempts=args.attempts, timeout=args.timeout)
                rows, source_name = client.fetch(), 'REST API'
                retries = client.retries
            else:
                rows, source_name = load_file(args.file, args.sheet), args.file.name
                retries = 0
            db = connect(args.db)
            run_id = import_rows(db, rows, currency, source_name, retries)
            log.info('import_committed run_id=%s', run_id)
        else:
            if not args.db.is_file():
                raise ValueError('Database does not exist')
            db = connect(args.db)
            run_id = args.run_id
        path, result = generate(db, run_id, args.output)
        log.info('report_completed run_id=%s received=%s inserted=%s duplicates=%s rejected=%s',
                 run_id, *(result['import'][key] for key in ('received', 'inserted', 'duplicates', 'rejected')))
        print(json.dumps({'report_directory': str(path.resolve()), **result}, indent=2))
        return 0
    except (ValueError, OSError, sqlite3.Error, KeyError, csv.Error, BadZipFile, InvalidFileException) as exc:
        # No URLs, tokens, HTTP bodies, customer names or raw rows enter logs.
        log.error('run_failed type=%s run_id=%s', type(exc).__name__, run_id or 'not_committed')
        print(f'Error: {exc}', file=sys.stderr)
        if run_id:
            print(f'Data is committed. Regenerate with report --run-id {run_id} and another --output directory.', file=sys.stderr)
        return 1
    finally:
        if db:
            db.close()
        log.removeHandler(handler)
        handler.close()


if __name__ == '__main__':
    raise SystemExit(main())
