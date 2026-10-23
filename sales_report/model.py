"""Explicit sales rules: one order per row, one reporting currency per database."""
from dataclasses import asdict, dataclass
from datetime import date, datetime
from decimal import Decimal, InvalidOperation
import re

FIELDS = ('order_id', 'order_date', 'customer', 'product', 'quantity', 'unit_price', 'currency')


@dataclass(frozen=True)
class Order:
    order_id: str
    order_date: str
    customer: str
    product: str
    quantity: int
    unit_price: str
    currency: str

    @property
    def total(self) -> Decimal:
        return Decimal(self.unit_price) * self.quantity

    def payload(self) -> dict:
        return asdict(self)


def normalize(raw: dict, currency: str) -> Order:
    """Reject ambiguous values rather than silently guessing dates or prices."""
    def text(key):
        value = raw.get(key)
        if value is None or isinstance(value, (dict, list, bool)):
            raise ValueError(f'{key}: required scalar value')
        value = str(value).strip()
        if not value or len(value) > 200:
            raise ValueError(f'{key}: must contain 1–200 characters')
        if any(ord(char) < 32 for char in value):
            raise ValueError(f'{key}: control characters are not supported')
        return value

    identity = text('order_id')
    raw_day = raw.get('order_date')
    if isinstance(raw_day, datetime) and raw_day.time().isoformat() == '00:00:00':
        day = raw_day.date().isoformat()
    elif isinstance(raw_day, date) and not isinstance(raw_day, datetime):
        day = raw_day.isoformat()
    else:
        day = text('order_date')
    if not re.fullmatch(r'\d{4}-\d{2}-\d{2}', day):
        raise ValueError('order_date: use YYYY-MM-DD; ambiguous dates need review')
    try:
        date.fromisoformat(day)
    except ValueError as exc:
        raise ValueError('order_date: invalid calendar date') from exc
    qty = text('quantity')
    if not re.fullmatch(r'[1-9]\d{0,5}', qty):
        raise ValueError('quantity: positive integer, maximum 999999')
    try:
        amount = Decimal(text('unit_price'))
    except InvalidOperation as exc:
        raise ValueError('unit_price: expected a decimal number without currency symbols') from exc
    if not amount.is_finite() or not Decimal('0') <= amount <= Decimal('999999999.99'):
        raise ValueError('unit_price: expected a finite, non-negative amount below 1 billion')
    if amount != amount.quantize(Decimal('.01')):
        raise ValueError('unit_price: maximum two decimal places')
    if amount * int(qty) > Decimal('999999999999.99'):
        raise ValueError('order total: exceeds supported Excel numeric precision')
    code = text('currency').upper()
    if code != currency:
        raise ValueError(f'currency: expected {currency}; no currency conversion is performed')
    return Order(identity, day, text('customer'), text('product'), int(qty), f'{amount:.2f}', code)
