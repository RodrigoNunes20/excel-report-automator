"""Cleaning and normalization helpers.

Everything in this module is pure (no file access) so it is easy to test.
"""
from __future__ import annotations

import math
import numbers
import re
import unicodedata
from dataclasses import dataclass, field
from datetime import date, datetime, timedelta
from typing import Any, Optional

import pandas as pd

# Canonical column name -> accepted aliases (already normalized, see normalize_column_name).
COLUMN_ALIASES: dict[str, set[str]] = {
    "date": {"date", "data", "order_date", "sale_date", "dt", "data_venda", "transaction_date"},
    "customer": {"customer", "client", "cliente", "customer_name", "buyer"},
    "category": {"category", "categoria", "product_category", "type", "tipo"},
    "product": {"product", "produto", "item", "product_name"},
    "quantity": {"quantity", "qty", "qtd", "quantidade", "units"},
    "amount": {"amount", "total", "valor", "value", "revenue", "total_amount", "valor_total"},
}
REQUIRED_COLUMNS = ("date", "amount")
OUTPUT_COLUMNS = ["date", "customer", "category", "product", "quantity", "amount"]
FRAME_COLUMNS = [*OUTPUT_COLUMNS, "source", "source_row"]

_ALIAS_LOOKUP = {alias: canonical for canonical, aliases in COLUMN_ALIASES.items() for alias in aliases}


@dataclass
class Issue:
    """A row (or file) that was rejected or removed, with the reason why."""

    source: str
    row: Optional[int]
    reason: str
    raw: str = ""


@dataclass
class CleanResult:
    frame: pd.DataFrame
    issues: list[Issue] = field(default_factory=list)
    rows_read: int = 0


def normalize_column_name(name: Any) -> str:
    """'  Valor Total ' -> 'valor_total'; 'Preço' -> 'preco'."""
    text = unicodedata.normalize("NFKD", str(name))
    text = "".join(ch for ch in text if not unicodedata.combining(ch))
    text = re.sub(r"[^0-9a-zA-Z]+", "_", text.strip().lower())
    return text.strip("_")


def is_blank(value: Any) -> bool:
    if value is None:
        return True
    if isinstance(value, str):
        return value.strip() == ""
    try:
        return bool(pd.isna(value))
    except (TypeError, ValueError):
        return False


_NON_NUMERIC_RE = re.compile(r"[^\d,.\-+()]")


def parse_amount(value: Any) -> Optional[float]:
    """Parse numbers written in many styles. Returns None when the value is not a number.

    Handles: 1234.5, "1,234.56", "1.234,56", "R$ 1.234,56", "$99.90", "(45.00)", "-12,5".
    Rule for a single separator: one "," or one "." is treated as the decimal mark;
    several of the same separator are treated as thousands marks.
    """
    if is_blank(value) or isinstance(value, bool):
        return None
    if isinstance(value, numbers.Real):
        number = float(value)
        return None if math.isnan(number) or math.isinf(number) else number

    text = str(value).strip()
    negative = text.startswith("(") and text.endswith(")")
    cleaned = _NON_NUMERIC_RE.sub("", text).replace("(", "").replace(")", "")
    if cleaned.startswith("-"):
        negative = True
    cleaned = cleaned.lstrip("+-")
    if not cleaned or not re.search(r"\d", cleaned) or "-" in cleaned or "+" in cleaned:
        return None

    has_dot, has_comma = "." in cleaned, "," in cleaned
    if has_dot and has_comma:
        decimal = "," if cleaned.rfind(",") > cleaned.rfind(".") else "."
        thousands = "." if decimal == "," else ","
        cleaned = cleaned.replace(thousands, "").replace(decimal, ".")
    elif has_comma:
        cleaned = cleaned.replace(",", "") if cleaned.count(",") > 1 else cleaned.replace(",", ".")
    elif has_dot and cleaned.count(".") > 1:
        cleaned = cleaned.replace(".", "")

    if not re.fullmatch(r"\d+(\.\d+)?|\.\d+", cleaned):
        return None
    number = float(cleaned)
    return -number if negative else number


def _date_formats(dayfirst: bool) -> list[str]:
    local = (
        ["%d/%m/%Y", "%d-%m-%Y", "%d.%m.%Y", "%d/%m/%y"]
        if dayfirst
        else ["%m/%d/%Y", "%m-%d-%Y", "%m/%d/%y"]
    )
    return ["%Y-%m-%d", "%Y/%m/%d", "%Y-%m-%d %H:%M:%S", *local, "%b %d, %Y", "%B %d, %Y", "%d %b %Y"]


def _plausible(day: date) -> bool:
    return 1990 <= day.year <= 2100


def parse_date(value: Any, dayfirst: bool = True) -> Optional[date]:
    """Parse dates from real date cells, Excel serial numbers, or text in common formats."""
    if is_blank(value) or isinstance(value, bool):
        return None
    if isinstance(value, datetime):  # includes pandas.Timestamp
        result = value.date()
        return result if _plausible(result) else None
    if isinstance(value, date):
        return value if _plausible(value) else None
    if isinstance(value, numbers.Real):
        # Excel stores dates as days since 1899-12-30.
        if 20000 <= float(value) <= 80000:
            return (datetime(1899, 12, 30) + timedelta(days=int(value))).date()
        return None

    text = str(value).strip()
    for fmt in _date_formats(dayfirst):
        try:
            result = datetime.strptime(text, fmt).date()
        except ValueError:
            continue
        return result if _plausible(result) else None
    try:
        result = datetime.fromisoformat(text).date()
    except ValueError:
        return None
    return result if _plausible(result) else None


def _clean_text(value: Any, default: str = "") -> str:
    if is_blank(value):
        return default
    return " ".join(str(value).split())


def _raw_repr(row: pd.Series) -> str:
    parts = [f"{key}={value}" for key, value in row.items() if not is_blank(value)]
    return "; ".join(parts)


def empty_frame() -> pd.DataFrame:
    return pd.DataFrame(columns=FRAME_COLUMNS)


def clean_frame(df: pd.DataFrame, source: str, dayfirst: bool = True) -> CleanResult:
    """Normalize one raw table into the standard schema and collect rejected rows."""
    if df is None or df.empty:
        return CleanResult(empty_frame(), [Issue(source, None, "No data rows found")], 0)

    df = df.copy()
    df.columns = [_ALIAS_LOOKUP.get(normalize_column_name(c), normalize_column_name(c)) for c in df.columns]
    df = df.loc[:, ~df.columns.duplicated()]  # keep the first of repeated columns

    missing = [c for c in REQUIRED_COLUMNS if c not in df.columns]
    if missing:
        found = ", ".join(str(c) for c in df.columns)
        reason = f"Missing required column(s): {', '.join(missing)} (found: {found})"
        return CleanResult(empty_frame(), [Issue(source, None, reason)], 0)

    issues: list[Issue] = []
    records: list[dict[str, Any]] = []
    rows_read = 0
    for position, (_, row) in enumerate(df.iterrows()):
        row_number = position + 2  # +1 for the header row, +1 because spreadsheets start at 1
        if all(is_blank(v) for v in row.values):
            continue
        rows_read += 1

        day = parse_date(row.get("date"), dayfirst=dayfirst)
        if day is None:
            issues.append(Issue(source, row_number, f"Invalid or missing date: {row.get('date')!r}", _raw_repr(row)))
            continue
        amount = parse_amount(row.get("amount"))
        if amount is None:
            issues.append(Issue(source, row_number, f"Invalid or missing amount: {row.get('amount')!r}", _raw_repr(row)))
            continue

        quantity = parse_amount(row.get("quantity")) if "quantity" in row.index else None
        records.append(
            {
                "date": day,
                "customer": _clean_text(row.get("customer"), "Unknown"),
                "category": _clean_text(row.get("category"), "Uncategorized").title(),
                "product": _clean_text(row.get("product")),
                "quantity": quantity if quantity is not None else float("nan"),
                "amount": amount,
                "source": source,
                "source_row": row_number,
            }
        )

    frame = pd.DataFrame.from_records(records, columns=FRAME_COLUMNS) if records else empty_frame()
    return CleanResult(frame, issues, rows_read)


DUPLICATE_KEY = ["date", "customer", "category", "product", "quantity", "amount"]


def remove_duplicates(frame: pd.DataFrame) -> tuple[pd.DataFrame, list[Issue]]:
    """Drop exact duplicate transactions (same day, customer, product, quantity and amount)."""
    if frame.empty:
        return frame, []
    mask = frame.duplicated(subset=DUPLICATE_KEY, keep="first")
    issues = [
        Issue(
            str(r.source),
            int(r.source_row),
            "Duplicate row removed",
            f"date={r.date}; customer={r.customer}; amount={r.amount}",
        )
        for r in frame[mask].itertuples()
    ]
    return frame[~mask].reset_index(drop=True), issues
