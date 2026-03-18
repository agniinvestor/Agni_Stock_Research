"""
Indian fiscal year utilities.
Indian FY runs April 1 → March 31.
  FY25 = April 1 2024 → March 31 2025  (labelled by the March year-end)
"""
from datetime import date, datetime
from typing import Optional


def fy_label(dt: date | datetime | None = None, year: Optional[int] = None,
             month: Optional[int] = None) -> str:
    """Return FY label (e.g. 'FY25') for a given date or (year, month) pair."""
    if dt is not None:
        if isinstance(dt, datetime):
            dt = dt.date()
        year, month = dt.year, dt.month
    if year is None or month is None:
        raise ValueError("Provide either dt or (year, month)")
    # If month is Apr–Dec, FY ends next March; if Jan–Mar, FY ends this March
    fy_year = year if month <= 3 else year + 1
    return f"FY{str(fy_year)[-2:]}"


def fy_full_label(dt: date | datetime | None = None, year: Optional[int] = None,
                  month: Optional[int] = None) -> str:
    """Return 4-digit FY label (e.g. 'FY2025')."""
    short = fy_label(dt, year, month)
    base = int(short[2:])
    century = 2000 if base < 50 else 1900
    return f"FY{century + base}"


def fy_end_year(fy_label_str: str) -> int:
    """'FY25' → 2025, 'FY2025' → 2025."""
    s = fy_label_str.upper().lstrip("FY")
    n = int(s)
    if n < 100:
        return (2000 + n) if n < 50 else (1900 + n)
    return n


def fy_date_range(fy_label_str: str) -> tuple[date, date]:
    """Return (start, end) date for a FY label."""
    end_year = fy_end_year(fy_label_str)
    start = date(end_year - 1, 4, 1)
    end   = date(end_year, 3, 31)
    return start, end


def fy_start_year(fy_label_str: str) -> int:
    """'FY25' → 2024 (April of that year)."""
    return fy_end_year(fy_label_str) - 1


def quarter_label(dt: date | datetime) -> str:
    """Return quarter label e.g. 'Q1FY25' for a date."""
    if isinstance(dt, datetime):
        dt = dt.date()
    m = dt.month
    if m in (4, 5, 6):
        q = 1
    elif m in (7, 8, 9):
        q = 2
    elif m in (10, 11, 12):
        q = 3
    else:  # 1, 2, 3
        q = 4
    return f"Q{q}{fy_label(dt)}"


def fy_range(start_fy: str, end_fy: str) -> list[str]:
    """Return list of FY labels from start to end inclusive.
    fy_range('FY16', 'FY25') → ['FY16', 'FY17', ..., 'FY25']
    """
    s, e = fy_end_year(start_fy), fy_end_year(end_fy)
    labels = []
    for y in range(s, e + 1):
        short = str(y)[-2:]
        labels.append(f"FY{short}")
    return labels


def estimate_suffix(fy_label_str: str, actuals_end_fy: str) -> str:
    """Return 'A' if fy is actual, 'E' if estimate.
    e.g. estimate_suffix('FY26', 'FY25') → 'E'
    """
    if fy_end_year(fy_label_str) <= fy_end_year(actuals_end_fy):
        return "A"
    return "E"


def label_with_suffix(fy_label_str: str, actuals_end_fy: str) -> str:
    """'FY25' + actuals_end='FY25' → 'FY25A'; 'FY26' → 'FY26E'."""
    return f"{fy_label_str}{estimate_suffix(fy_label_str, actuals_end_fy)}"


def current_fy() -> str:
    """Return the current Indian FY label."""
    return fy_label(date.today())
