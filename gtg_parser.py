"""
gtg_parser.py
-------------
Dedicated parser for the "Global Technical Governance" sheet.

That sheet has two zones stacked vertically:
  * TOP  : aggregated region/business-unit tables (the agent must NOT use these).
  * BELOW: a granular block — one row per Location, grouped under region banners,
           with a header that REPEATS for every region group:
             Location | HSSE | Code of Conduct | Risk | GTG Attempts |
             Risk | Score | Year | Follow Up | Score | Year | Follow Up | Score | Year ...

This parser locates the granular block, reads its repeating header, captures the
region from each banner row, and emits one tidy row per Location.

Output columns:
    Region | Location | HSSE | Code of Conduct | Risk | GTG Attempts |
    GTG Score 1 | GTG Year 1 | GTG Follow Up 1 |
    GTG Score 2 | GTG Year 2 | GTG Follow Up 2 |
    GTG Score 3 | GTG Year 3 | GTG Follow Up 3 | Year

Region is the raw banner text (e.g. "Energy & Chemicals - Asia Pacific"); any
custom Master Region / Region Category hierarchy is left to be added downstream
(e.g. in Power BI) since those labels do not exist in the source file.
"""
from __future__ import annotations

import re
import pandas as pd

import engine_core as E

TOTAL_RE = re.compile(r"\btotal\b", re.IGNORECASE)
# the granular header is recognised by carrying these tokens across one row
HDR_TOKENS = {"hsse", "code of conduct", "risk", "gtg attempts", "score", "year", "follow up"}

# fixed output schema (3 GTG attempts; extend if a file ever has more)
OUT_COLS = ["Region", "Location", "HSSE", "Code of Conduct", "Risk", "GTG Attempts",
            "GTG Score 1", "GTG Year 1", "GTG Follow Up 1",
            "GTG Score 2", "GTG Year 2", "GTG Follow Up 2",
            "GTG Score 3", "GTG Year 3", "GTG Follow Up 3"]


def looks_like_gtg(sheet: "E.Sheet") -> bool:
    """True if any row carries the granular GTG header tokens (Risk + Follow Up +
    GTG Attempts together)."""
    grid = E.apply_merges(sheet)
    for row in grid:
        toks = {str(v).strip().lower() for v in row if not E._is_blank(v)}
        if {"risk", "follow up", "gtg attempts"}.issubset(toks):
            return True
    return False


def _is_header_row(row) -> bool:
    toks = {str(v).strip().lower() for v in row if not E._is_blank(v)}
    return {"risk", "follow up", "gtg attempts"}.issubset(toks)


def _map_columns(header_row) -> dict:
    """Map an instance of the repeating granular header to output fields.
    The header looks like (col0 = banner/region):
       0:<banner> 1:HSSE 2:Code of Conduct 3:Risk 4:GTG Attempts
       5:Risk 6:Score 7:Year 8:Follow Up 9:Score 10:Year 11:Follow Up 12:Score 13:Year...
    """
    cols = {"HSSE": None, "Code of Conduct": None, "Risk": None, "GTG Attempts": None}
    # find the anchor columns by label (first occurrence)
    labels = [str(v).strip().lower() if not E._is_blank(v) else "" for v in header_row]
    for c, lab in enumerate(labels):
        if lab == "hsse" and cols["HSSE"] is None:
            cols["HSSE"] = c
        elif lab == "code of conduct" and cols["Code of Conduct"] is None:
            cols["Code of Conduct"] = c
        elif lab == "risk" and cols["Risk"] is None:
            cols["Risk"] = c
        elif lab == "gtg attempts" and cols["GTG Attempts"] is None:
            cols["GTG Attempts"] = c
    # the attempts blocks: every (Score, Year, Follow Up) triple AFTER GTG Attempts
    attempts = []
    c = (cols["GTG Attempts"] or 0) + 1
    n = len(labels)
    while c < n:
        if labels[c] == "score":
            score_c = c
            year_c = c + 1 if c + 1 < n and labels[c + 1] == "year" else None
            fu_c = c + 2 if c + 2 < n and labels[c + 2] == "follow up" else None
            attempts.append((score_c, year_c, fu_c))
            c += 3
        else:
            c += 1
    cols["attempts"] = attempts
    return cols


def parse(sheet: "E.Sheet", year: str | None = None) -> dict:
    grid = E.apply_merges(sheet)
    width = max((len(r) for r in grid), default=0)
    log: list[str] = []

    # find the FIRST granular header row → start of the granular block
    first_hdr = None
    for i, row in enumerate(grid):
        if _is_header_row(row):
            first_hdr = i
            break
    if first_hdr is None:
        raise ValueError("No granular GTG block found (header with Risk/GTG Attempts/Follow Up).")
    log.append(f"Granular block starts at row {first_hdr+1} (skipped the aggregated tables above).")

    records = []
    region = None
    colmap = None
    dropped_total = dropped_banner = 0

    for r in range(first_hdr, len(grid)):
        row = grid[r]
        if E._nonempty_count(row) == 0:
            continue
        label = str(row[0]).strip() if (len(row) and not E._is_blank(row[0])) else ""

        if _is_header_row(row):
            # this row is both the region banner (col0) and the column header
            colmap = _map_columns(row)
            if label:
                region = label
            dropped_banner += 1
            continue
        if TOTAL_RE.search(label):
            dropped_total += 1
            continue
        if not label or colmap is None:
            continue
        # a banner-only row (region change without a repeated header) — rare, but
        # if the row has a label and no numbers, treat as a region update
        numeric = sum(1 for c in range(1, len(row)) if E.to_number(row[c]) is not None)
        if numeric == 0 and " - " in label:
            region = label
            dropped_banner += 1
            continue

        def gv(c):
            return row[c] if (c is not None and c < len(row)) else None

        rec = {
            "Region": region,
            "Location": label,
            "HSSE": gv(colmap["HSSE"]),
            "Code of Conduct": gv(colmap["Code of Conduct"]),
            "Risk": gv(colmap["Risk"]),
            "GTG Attempts": gv(colmap["GTG Attempts"]),
        }
        for idx in range(3):  # up to 3 attempts
            if idx < len(colmap["attempts"]):
                sc, yr, fu = colmap["attempts"][idx]
                rec[f"GTG Score {idx+1}"] = gv(sc)
                rec[f"GTG Year {idx+1}"] = gv(yr)
                rec[f"GTG Follow Up {idx+1}"] = gv(fu)
            else:
                rec[f"GTG Score {idx+1}"] = None
                rec[f"GTG Year {idx+1}"] = None
                rec[f"GTG Follow Up {idx+1}"] = None
        records.append(rec)

    tidy = pd.DataFrame(records, columns=OUT_COLS)
    if year is not None:
        tidy["Year"] = year
    log.append(f"Built {len(tidy)} Location rows "
               f"(dropped {dropped_total} total and {dropped_banner} banner/header rows).")
    return {"tidy": tidy, "log": log}
