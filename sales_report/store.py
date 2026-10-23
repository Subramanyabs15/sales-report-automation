"""Transactional imports and a persistent order-ID uniqueness constraint."""
from collections import Counter
from datetime import datetime, timezone
from decimal import Decimal
import json
from pathlib import Path
import sqlite3
from uuid import uuid4

from .model import Order, normalize


def connect(path: Path):
    path.parent.mkdir(parents=True, exist_ok=True)
    db = sqlite3.connect(path, timeout=1)
    db.row_factory = sqlite3.Row
    db.executescript('''
      CREATE TABLE IF NOT EXISTS metadata (key TEXT PRIMARY KEY, value TEXT NOT NULL);
      CREATE TABLE IF NOT EXISTS orders (order_id TEXT PRIMARY KEY, payload TEXT NOT NULL);
      CREATE TABLE IF NOT EXISTS runs (
        run_id TEXT PRIMARY KEY, created_at TEXT NOT NULL, source TEXT NOT NULL,
        currency TEXT NOT NULL, stats TEXT NOT NULL, issues TEXT NOT NULL
      );
    ''')
    return db


def import_rows(db, rows, currency, source, retries=0):
    stats = Counter(received=len(rows), inserted=0, duplicates=0, rejected=0, retries=retries)
    issues = []
    run_id = datetime.now(timezone.utc).strftime('%Y%m%dT%H%M%S') + '-' + uuid4().hex[:8]
    # BEGIN IMMEDIATE serializes readers that are about to write, including scheduled runs.
    db.execute('BEGIN IMMEDIATE')
    try:
        existing_currency = db.execute("SELECT value FROM metadata WHERE key='currency'").fetchone()
        if existing_currency and existing_currency[0] != currency:
            raise ValueError('Database currency differs; use another database for this currency')
        db.execute("INSERT OR IGNORE INTO metadata VALUES ('currency', ?)", (currency,))
        for index, raw in enumerate(rows, 1):
            try:
                order = normalize(raw, currency)
                payload = json.dumps(order.payload(), sort_keys=True)
                old = db.execute('SELECT payload FROM orders WHERE order_id=?', (order.order_id,)).fetchone()
                if old:
                    if old[0] != payload:
                        raise ValueError('order_id: conflicts with stored record; review before updating')
                    stats['duplicates'] += 1
                    issues.append({'record': index, 'order_id': order.order_id, 'reason': 'Exact duplicate skipped', 'decision': 'duplicate', 'raw': raw})
                else:
                    db.execute('INSERT INTO orders VALUES (?, ?)', (order.order_id, payload))
                    stats['inserted'] += 1
            except (ValueError, TypeError) as exc:
                stats['rejected'] += 1
                issues.append({'record': index, 'order_id': str(raw.get('order_id', '')), 'reason': str(exc), 'decision': 'rejected', 'raw': raw})
        total = sum((Order(**json.loads(row[0])).total for row in db.execute('SELECT payload FROM orders')), Decimal('0'))
        if total > Decimal('999999999999.99'):
            raise ValueError('Report total exceeds supported Excel numeric precision; import rolled back')
        stats['stored_orders'] = db.execute('SELECT count(*) FROM orders').fetchone()[0]
        created = datetime.now(timezone.utc).isoformat(timespec='seconds')
        db.execute('INSERT INTO runs VALUES (?, ?, ?, ?, ?, ?)',
                   (run_id, created, source, currency, json.dumps(dict(stats)), json.dumps(issues, default=str)))
        db.commit()
    except Exception:
        db.rollback()
        raise
    return run_id


def snapshot(db, run_id):
    run = db.execute('SELECT * FROM runs WHERE run_id=?', (run_id,)).fetchone()
    if not run:
        raise ValueError('Run ID does not exist')
    orders = [Order(**json.loads(row[0])) for row in db.execute('SELECT payload FROM orders ORDER BY order_id')]
    return dict(run), orders
