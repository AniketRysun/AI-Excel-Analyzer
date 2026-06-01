"""
hhw_parser.py
-------------
Dedicated parser for the AmSpec "H&HW" (Headcount & Hours Worked) cross-tab
sheets. These sheets are irregular enough (a 2-column 'Year to Date Average'
block followed by 3-column month blocks, region banners, repeated sub-headers,
and total rows) that the generic reshaper mislabels them. This parser handles
that exact family deterministically.

Output columns:
    Region | Country | Month | Metric columns (Headcount, Work Week, Hrs. Worked)

Year-to-Date Average is treated as a summary, not a month, and excluded from the
granular grain by default (it has no Work Week and duplicates monthly figures).
"""
from __future__ import annotations

import re
import pandas as pd

import engine_core as E

MONTHS = {"january", "february", "march", "april", "may", "june", "july",
          "august", "september", "october", "november", "december"}
TOTAL_RE = re.compile(r"\btotal\b", re.IGNORECASE)
BANNER_RE = re.compile(r"business assurance|sustainability| - (apac|emea|latam|nam|global)",
                       re.IGNORECASE)


def looks_like_hhw(sheet: "E.Sheet") -> bool:
    """Heuristic: one row holds >=3 month names AND the row just below it holds
    the H&HW metric labels (Headcount / Work Week / Hrs. Worked) repeated."""
    grid = E.apply_merges(sheet)
    for r in range(min(15, len(grid) - 1)):
        row_strs = [str(v).strip().lower() for v in grid[r] if not E._is_blank(v)]
        month_hits = sum(1 for v in row_strs if v in MONTHS)
        if month_hits < 3:
            continue
        below = [str(v).strip().lower() for v in grid[r + 1] if not E._is_blank(v)]
        metric_hits = sum(1 for v in below if v in {"headcount", "work week", "hrs. worked"})
        if metric_hits >= 4:   # metrics repeat across multiple month blocks
            return True
    return False


def _find_header_rows(grid) -> tuple[int, int]:
    """Return (month_row_idx, metric_row_idx) 0-indexed."""
    month_row = None
    for r in range(len(grid)):
        row_strs = [str(v).strip().lower() for v in grid[r] if not E._is_blank(v)]
        if sum(1 for v in row_strs if v in MONTHS) >= 3:
            month_row = r
            break
    if month_row is None:
        raise ValueError("No month header row found — not an H&HW sheet.")
    return month_row, month_row + 1


def parse(sheet: "E.Sheet", include_ytd: bool = False) -> dict:
    grid = E.apply_merges(sheet)
    width = max((len(r) for r in grid), default=0)
    log: list[str] = []

    month_row, metric_row = _find_header_rows(grid)
    log.append(f"Detected month header on row {month_row+1}, metrics on row {metric_row+1}.")

    # Build per-column (month, metric); keep only columns whose month is a real month
    col_map: dict[int, tuple[str, str]] = {}
    for c in range(width):
        m = grid[month_row][c] if c < len(grid[month_row]) else None
        k = grid[metric_row][c] if c < len(grid[metric_row]) else None
        if E._is_blank(m) or E._is_blank(k):
            continue
        month = str(m).strip()
        metric = str(k).strip()
        if month.lower() in MONTHS:
            col_map[c] = (month, metric)
        elif include_ytd and "year to date" in month.lower():
            col_map[c] = ("Year to Date", metric)
    metrics = sorted({mt for _, mt in col_map.values()})
    months = list(dict.fromkeys(m for m, _ in col_map.values()))  # preserve order
    log.append(f"Months: {len(months)}; metrics: {', '.join(metrics)}.")

    # Walk data rows: region comes from banner rows; country from column A
    records = []
    region = None
    dropped_total = dropped_repeat = 0
    metric_label_set = {mt.lower() for _, mt in col_map.values()} | {"headcount", "work week", "hrs. worked"}

    for r in range(metric_row + 1, len(grid)):
        row = grid[r]
        label = str(row[0]).strip() if (len(row) and not E._is_blank(row[0])) else ""

        if E._nonempty_count(row) == 0:
            continue
        if TOTAL_RE.search(label):
            dropped_total += 1
            continue
        # banner row: label mentions a region/business unit and the rest is non-numeric
        numeric_in_row = sum(1 for c in col_map if c < len(row) and E.to_number(row[c]) is not None)
        if label and BANNER_RE.search(label) and numeric_in_row == 0:
            region = label
            dropped_repeat += 1
            continue
        # repeated sub-header row (cells equal metric labels, no numbers)
        if numeric_in_row == 0:
            dropped_repeat += 1
            continue
        if not label:  # data row must have a country name
            continue

        # emit one row per month for this country
        per_month: dict[str, dict] = {}
        for c, (month, metric) in col_map.items():
            val = E.to_number(row[c]) if c < len(row) else None
            per_month.setdefault(month, {})[metric] = val
        for month, mvals in per_month.items():
            if all(v is None for v in mvals.values()):
                continue
            rec = {"Region": region, "Country": label, "Month": month}
            for mt in metrics:
                rec[mt] = mvals.get(mt)
            records.append(rec)

    tidy = pd.DataFrame(records)
    log.append(f"Built {len(tidy)} granular rows "
               f"(dropped {dropped_total} total and {dropped_repeat} banner/sub-header rows).")
    return {"tidy": tidy, "log": log, "months": months, "metrics": metrics}
