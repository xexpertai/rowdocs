"""Jinja filters available in templates and naming patterns."""

from __future__ import annotations

from datetime import date, datetime, timedelta
from decimal import ROUND_HALF_UP, Decimal, InvalidOperation
from typing import Any

from jinja2 import Environment

_EXCEL_EPOCH = datetime(1899, 12, 30)


def to_datetime(value: Any) -> datetime | None:
    """Read an Excel date cell, an Excel serial number or an ISO string. Empty -> None."""
    if value is None or (isinstance(value, str) and not value.strip()):
        return None
    if isinstance(value, datetime):
        return value
    if isinstance(value, date):
        return datetime(value.year, value.month, value.day)
    if isinstance(value, int | float) and not isinstance(value, bool):
        return _EXCEL_EPOCH + timedelta(days=float(value))
    if isinstance(value, str):
        text = value.strip()
        try:
            return datetime.fromisoformat(text.replace("Z", "+00:00"))
        except ValueError:
            pass
    raise ValueError(f"date: cannot read {value!r} as a date (use an Excel date or YYYY-MM-DD)")


def date_filter(value: Any, fmt: str = "%d %b %Y") -> str:
    """``{{ issued|date('%d %b %Y') }}`` -> "05 Mar 2026". Empty values render as ""."""
    parsed = to_datetime(value)
    return "" if parsed is None else parsed.strftime(fmt)


def money(value: Any, symbol: str = "", decimals: int = 2) -> str:
    """``{{ amount|money('₹') }}`` -> "₹1,234,567.50". Empty values render as ""."""
    if value is None or (isinstance(value, str) and not value.strip()):
        return ""
    if isinstance(value, bool):
        raise ValueError(f"money: cannot read {value!r} as a number")
    text = str(value).strip()
    if symbol:
        text = text.replace(symbol, "")
    text = text.replace(",", "").replace(" ", "").replace("_", "")
    try:
        number = Decimal(text)
    except InvalidOperation as exc:
        raise ValueError(f"money: cannot read {value!r} as a number") from exc
    if not number.is_finite():
        raise ValueError(f"money: cannot read {value!r} as a number")
    decimals = int(decimals)
    quantum = Decimal(1).scaleb(-decimals)
    rounded = number.quantize(quantum, rounding=ROUND_HALF_UP)
    sign = "-" if rounded < 0 else ""
    return f"{sign}{symbol}{abs(rounded):,.{decimals}f}"


def make_env() -> Environment:
    """The Jinja environment used for documents and file-name patterns."""
    env = Environment(autoescape=False)
    env.filters["date"] = date_filter
    env.filters["money"] = money
    return env
