"""
eventlog_parser.py
------------------
Parser for the HSSE Event / Incident Log sheets.

Unlike the cross-tab sheets, an event log is ALREADY granular — one row per
incident. The challenge across years is that the columns differ a lot:
  2018 has 'Stop Work Authority', 'Equipment Impact', ...
  2019 adds 'Region', 'Incident Site', ...
  2026 adds 'Description of Incident', 'Corrective Actions', 'Local File Number', ...

So this parser does NOT use a fixed schema. It:
  * finds the header row,
  * keeps EVERY column under its own (cleaned) name — including brand-new ones,
  * drops fully-blank rows,
  * adds a Year column (from the filename).

This is the Union rule the user asked for: whatever columns a file has are kept.
"""
from __future__ import annotations

import re
import pandas as pd

import engine_core as E

# the sheet is an event log if its header row contains several of these tokens
SIGNATURE = {"date of\nincident", "date of incident", "name of individual",
             "nature of incident", "root cause", "nature of illness/injury",
             "incident #", "hsse impact event", "age",
             "client(s) working on behalf of"}


def _norm(s) -> str:
    """Normalise a header for matching: collapse whitespace/newlines, lowercase."""
    return " ".join(str(s).split()).strip().lower()


def _clean_header(s) -> str:
    """Human-friendly column name: collapse internal newlines/whitespace."""
    return " ".join(str(s).split()).strip()


def _find_header_row(grid) -> int | None:
    """The header row is the first row that matches several signature tokens."""
    best_i, best_hits = None, 0
    for i in range(min(8, len(grid))):
        toks = {_norm(v) for v in grid[i] if not E._is_blank(v)}
        hits = sum(1 for t in toks if t in SIGNATURE)
        if hits > best_hits:
            best_i, best_hits = i, hits
    return best_i if best_hits >= 3 else None


def looks_like_eventlog(sheet: "E.Sheet") -> bool:
    return _find_header_row(E.apply_merges(sheet)) is not None


def parse_flat(sheet: "E.Sheet", year: str | None = None) -> dict:
    """Generic 'already granular' table: find the header row, keep every column
    by name, drop blank rows, add Year. Use for flat sheets (e.g. the D&A roster)
    that don't match the Event Log signature but are still one-row-per-record."""
    grid = E.apply_merges(sheet)
    width = max((len(r) for r in grid), default=0)
    log: list[str] = []
    # header = first row that is mostly non-blank text
    hdr = 0
    for i in range(min(8, len(grid))):
        nb = E._nonempty_count(grid[i])
        texty = sum(1 for v in grid[i] if not E._is_blank(v) and E.to_number(v) is None)
        if nb >= max(3, 0.4 * width) and texty >= 2:
            hdr = i
            break
    col_idx, col_names, seen = [], [], {}
    for c in range(width):
        v = grid[hdr][c] if c < len(grid[hdr]) else None
        if E._is_blank(v):
            continue
        name = _clean_header(v)
        if name in seen:
            seen[name] += 1
            name = f"{name} ({seen[name]})"
        else:
            seen[name] = 0
        col_idx.append(c)
        col_names.append(name)
    records, dropped = [], 0
    for r in range(hdr + 1, len(grid)):
        row = grid[r]
        vals = [row[c] if c < len(row) else None for c in col_idx]
        if all(E._is_blank(v) for v in vals):
            dropped += 1
            continue
        records.append(dict(zip(col_names, vals)))
    tidy = pd.DataFrame(records, columns=col_names)
    if year is not None:
        tidy["Year"] = year
    log.append(f"Header on row {hdr+1}; kept {len(col_names)} columns, {len(tidy)} rows. "
               f"Year={year}. All columns preserved (Union).")
    return {"tidy": tidy, "log": log}

    # build column names; keep EVERY non-blank header column (Union of all columns)
    col_idx, col_names, seen = [], [], {}
    for c in range(width):
        v = grid[hdr][c] if c < len(grid[hdr]) else None
        if E._is_blank(v):
            continue
        name = _clean_header(v)
        if name in seen:
            seen[name] += 1
            name = f"{name} ({seen[name]})"
        else:
            seen[name] = 0
        col_idx.append(c)
        col_names.append(name)
    log.append(f"Header on row {hdr+1}; kept {len(col_names)} columns: "
               f"{', '.join(col_names[:8])}{'...' if len(col_names) > 8 else ''}.")

    # read data rows: keep a row if it has any content in the kept columns
    records, dropped_blank = [], 0
    for r in range(hdr + 1, len(grid)):
        row = grid[r]
        vals = [row[c] if c < len(row) else None for c in col_idx]
        if all(E._is_blank(v) for v in vals):
            dropped_blank += 1
            continue
        records.append(dict(zip(col_names, vals)))

    tidy = pd.DataFrame(records, columns=col_names)
    if year is not None:
        tidy["Year"] = year
    log.append(f"Kept {len(tidy)} incident rows (dropped {dropped_blank} blank). "
               f"Every column in this file is preserved, including any new ones.")
    return {"tidy": tidy, "log": log}
