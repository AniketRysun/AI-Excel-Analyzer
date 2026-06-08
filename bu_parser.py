"""
bu_parser.py
------------
Parser for the right-hand table on the "Global Business Units" sheet.

That table is a cross-tab: the COLUMN headers are business-unit / region labels
(AmSpec Globally, BA&S, BA&S - APAC, ... E&C, ENV, Food & Agri, Europe, ...) and
the ROW labels are categories (Countries / Territories, Established Locations,
Inspection Offices, Laboratories - E&C, ...).

The user wants it TRANSPOSED into a tidy table:
    Master Region | Region | <one column per category> | Year

  * Each business-unit/region column becomes a row.
  * "Master Region" = the part before " - " (e.g. "BA&S" from "BA&S - APAC");
    a roll-up column like "BA&S" with no region maps to Master Region "BA&S",
    Region "BA&S".
  * Category row-labels become the measure columns (Union across years).

The right table is located by finding the "AmSpec Globally" header cell; the row
labels sit in the column immediately to its left.
"""
from __future__ import annotations

import re
import pandas as pd

import engine_core as E

ANCHOR_RE = re.compile(r"amspec glob", re.IGNORECASE)


def has_bu_table(sheet: "E.Sheet") -> bool:
    grid = E.apply_merges(sheet)
    for row in grid[:3]:
        for v in row:
            if not E._is_blank(v) and ANCHOR_RE.search(str(v)):
                return True
    return False


def _clean(s) -> str:
    return " ".join(str(s).split()).strip()


def parse(sheet: "E.Sheet", year: str | None = None) -> dict:
    grid = E.apply_merges(sheet)
    width = max((len(r) for r in grid), default=0)
    log: list[str] = []

    # locate the anchor "AmSpec Globally" in the header row (row 1)
    hdr = 0
    start = None
    for c in range(width):
        v = grid[hdr][c] if c < len(grid[hdr]) else None
        if not E._is_blank(v) and ANCHOR_RE.search(str(v)):
            start = c
            break
    if start is None:
        raise ValueError("Right-hand Business Units table not found (no 'AmSpec Globally').")
    label_col = start - 1
    log.append(f"Right table starts at column {start+1}; category labels in column {label_col+1}.")

    # business-unit/region columns = header cells from `start` to the right
    bu_cols = []
    for c in range(start, width):
        v = grid[hdr][c] if c < len(grid[hdr]) else None
        if not E._is_blank(v):
            bu_cols.append((c, _clean(v)))

    # category rows = non-blank labels in label_col below the header
    cat_rows = []
    for r in range(hdr + 1, len(grid)):
        v = grid[r][label_col] if label_col < len(grid[r]) else None
        if not E._is_blank(v):
            cat_rows.append((r, _clean(v)))

    log.append(f"{len(bu_cols)} business-unit/region columns x {len(cat_rows)} categories.")

    # build one row per business-unit/region column (transpose)
    records = []
    for c, bu_label in bu_cols:
        # Master Region / Region split on ' - '
        if " - " in bu_label:
            master = bu_label.split(" - ", 1)[0].strip()
            region = bu_label
        else:
            master = bu_label
            region = bu_label
        rec = {"Master Region": master, "Region": region}
        for r, cat in cat_rows:
            val = grid[r][c] if c < len(grid[r]) else None
            rec[cat] = E.to_number(val) if val is not None else None
        records.append(rec)

    tidy = pd.DataFrame(records)
    # column order: identity then categories (in sheet order)
    cat_order = [cat for _, cat in cat_rows]
    tidy = tidy.reindex(columns=["Master Region", "Region"] + cat_order)
    if year is not None:
        tidy["Year"] = year
    log.append(f"Transposed into {len(tidy)} rows (one per business-unit/region), "
               f"{len(cat_order)} category columns.")
    return {"tidy": tidy, "log": log}
