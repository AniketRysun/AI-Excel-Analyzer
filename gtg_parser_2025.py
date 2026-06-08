"""
gtg_parser_2025.py
------------------
Parser for the 2025-style "GTG - <Unit>" sheets (one sheet per business unit).

Layout (differs from 2026):
  * 6-row preamble, then a TWO-ROW header (e.g. rows 13-14) that REPEATS per region.
  * Region label sits in the header row of **column C**; Locations sit in column C below.
  * Columns: C=Location, D=Inspection, E=Laboratory, F=Other, G=Risk,
    then repeating (Score, Year, Follow Up) triples for attempts 1..3.
  * 'Residents' rows are kept (user wants them).

Output columns (union-friendly; aligns with the 2026 GTG output plus 2025 extras):
  Region | Location | Inspection | Laboratory | Other |
  Risk | GTG Score 1 | GTG Year 1 | GTG Follow Up 1 |
  GTG Score 2 | GTG Year 2 | GTG Follow Up 2 |
  GTG Score 3 | GTG Year 3 | GTG Follow Up 3 | Year
"""
from __future__ import annotations

import re
import pandas as pd

import engine_core as E

TOTAL_RE = re.compile(r"\btotal\b", re.IGNORECASE)

OUT_COLS = ["Region", "Location", "Inspection", "Laboratory", "Other", "Risk",
            "GTG Score 1", "GTG Year 1", "GTG Follow Up 1",
            "GTG Score 2", "GTG Year 2", "GTG Follow Up 2",
            "GTG Score 3", "GTG Year 3", "GTG Follow Up 3"]


def _row_tokens(row):
    return {str(v).strip().lower() for v in row if not E._is_blank(v)}


def looks_like_gtg_2025(sheet: "E.Sheet") -> bool:
    """A 2025 GTG sheet has a two-row header whose lower row carries
    Risk + Score + Year + Follow Up AND a 'Services' band in the row directly
    above it. The 'Services' band is what distinguishes the 2025 per-unit layout
    from the 2026 single-sheet layout (which has no Services row)."""
    grid = E.apply_merges(sheet)
    for i in range(1, len(grid) - 1):
        toks = _row_tokens(grid[i])
        if {"risk", "score", "year", "follow up"}.issubset(toks):
            above = _row_tokens(grid[i - 1])
            if "services" in above:
                return True
    return False


def _is_metric_header(row) -> bool:
    toks = _row_tokens(row)
    return {"risk", "score", "year", "follow up"}.issubset(toks)


def _map_columns(metric_row) -> dict:
    """Map the lower header row to output fields. Service columns vary per unit,
    so we capture every column between Location and Risk as a dynamic 'service'
    flag using its own header label."""
    labels = [str(v).strip() if not E._is_blank(v) else "" for v in metric_row]
    low = [s.lower() for s in labels]
    cols = {"Location": None, "Risk": None, "services": [], "attempts": []}
    for c, lab in enumerate(low):
        if cols["Risk"] is None and lab == "risk":
            cols["Risk"] = c
    risk_c = cols["Risk"] if cols["Risk"] is not None else 6
    # Location: first column to the left of Risk that has a label band of services.
    # In practice the region/location column is col C (index 2); services sit
    # between it and Risk.
    cols["Location"] = 2
    # service flag columns = labelled columns strictly between Location and Risk
    for c in range(cols["Location"] + 1, risk_c):
        name = labels[c] or f"Service {c}"
        cols["services"].append((c, name))
    # attempt triples after Risk
    c = risk_c + 1
    n = len(low)
    while c < n and len(cols["attempts"]) < 3:
        if low[c] == "score":
            yr = c + 1 if c + 1 < n and low[c + 1] == "year" else None
            fu = c + 2 if c + 2 < n and low[c + 2] == "follow up" else None
            cols["attempts"].append((c, yr, fu))
            c += 3
        else:
            c += 1
    return cols


def parse(sheet: "E.Sheet", year: str | None = None) -> dict:
    grid = E.apply_merges(sheet)
    width = max((len(r) for r in grid), default=0)
    log: list[str] = []

    # find first metric header
    first = next((i for i in range(len(grid)) if _is_metric_header(grid[i])), None)
    if first is None:
        raise ValueError("No 2025 GTG granular header (Risk/Score/Year/Follow Up) found.")
    colmap = _map_columns(grid[first])
    loc_col = colmap["Location"]
    log.append(f"Granular block starts at row {first+1}; Location in column {loc_col+1}.")

    records = []
    region = None
    dropped = 0

    for r in range(first, len(grid)):
        row = grid[r]
        if E._nonempty_count(row) == 0:
            continue
        if _is_metric_header(row):
            # the header row carries the region label in the Location column
            colmap = _map_columns(row)
            loc_col = colmap["Location"]
            reg = row[loc_col] if loc_col < len(row) else None
            if not E._is_blank(reg):
                region = str(reg).strip()
            dropped += 1
            continue

        loc = row[loc_col] if loc_col < len(row) else None
        label = str(loc).strip() if not E._is_blank(loc) else ""
        if not label:
            continue
        if TOTAL_RE.search(label):
            dropped += 1
            continue

        def gv(c):
            return row[c] if (c is not None and c < len(row)) else None

        rec = {
            "Region": region,
            "Location": label,
            "Risk": gv(colmap["Risk"]),
        }
        for sc, sname in colmap["services"]:
            rec[sname] = gv(sc)
        for idx in range(3):
            if idx < len(colmap["attempts"]):
                s, yr, fu = colmap["attempts"][idx]
                rec[f"GTG Score {idx+1}"] = gv(s)
                rec[f"GTG Year {idx+1}"] = gv(yr)
                rec[f"GTG Follow Up {idx+1}"] = gv(fu)
            else:
                rec[f"GTG Score {idx+1}"] = None
                rec[f"GTG Year {idx+1}"] = None
                rec[f"GTG Follow Up {idx+1}"] = None
        records.append(rec)

    tidy = pd.DataFrame(records)
    # order columns: identity, services, then GTG metrics, then Year
    base = ["Region", "Location"]
    metric = ["Risk", "GTG Score 1", "GTG Year 1", "GTG Follow Up 1",
              "GTG Score 2", "GTG Year 2", "GTG Follow Up 2",
              "GTG Score 3", "GTG Year 3", "GTG Follow Up 3"]
    services = [c for c in tidy.columns if c not in base + metric]
    tidy = tidy.reindex(columns=base + services + metric)
    if year is not None:
        tidy["Year"] = year
    log.append(f"Built {len(tidy)} Location rows (dropped {dropped} header/total rows; "
               f"Residents rows kept). Service columns: {', '.join(services) or 'none'}.")
    return {"tidy": tidy, "log": log}
