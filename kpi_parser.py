"""
kpi_parser.py
-------------
Parser for the "Global KPI Workbook" sheet (HSSE KPI / stats).

Structure:
  * The sheet stacks MULTIPLE sections down the page. Each section starts with a
    row whose column A == "Period" and whose column C holds the section/region
    name (e.g. "The AmSpec Group", "E&C - APAC", "Food & Agriculture").
  * Each section has: a category banner row (categories placed at the START column
    of each group, blanks after), a sub-header row (metric names), then monthly
    rows plus Quarter/EOY subtotal rows.

Output (fully long / Option 2):
    Region | Year | Month | Category | Metric | Value

  * Region  = the section name (col C of the Period row).
  * Category = read from the banner: each category owns the columns from where its
    name appears until the next category name (adapts to any number of fields).
  * Quarter/EOY/Total subtotal rows are DROPPED.
  * Calculated columns are kept as plain metrics (their values pass through); we do
    not recompute anything. (User builds rates in Power BI.)
"""
from __future__ import annotations

import re
import pandas as pd

import engine_core as E

SUBTOTAL_RE = re.compile(r"^(q[1-4]|eoy|ytd|year[- ]?end|total)\b", re.IGNORECASE)
MONTHS = {"january", "february", "march", "april", "may", "june", "july", "august",
          "september", "october", "november", "december"}


def looks_like_kpi(sheet: "E.Sheet") -> bool:
    """A KPI Workbook sheet has a 'Period' header row and a 'Learnings from
    Incidents' category banner."""
    g = sheet.values
    saw_period = saw_learnings = False
    for row in g[:12]:
        for v in row:
            if E._is_blank(v):
                continue
            t = str(v).strip().lower()
            if t == "period":
                saw_period = True
            if "learnings from incidents" in t:
                saw_learnings = True
    return saw_period and saw_learnings


def _category_spans(banner_row, width):
    """Map each column -> category name, where a category owns columns from where
    its name appears until the next non-blank category name."""
    points = []
    for c in range(width):
        v = banner_row[c] if c < len(banner_row) else None
        if not E._is_blank(v):
            points.append((c, str(v).strip()))
    spans = {}
    for i, (c, name) in enumerate(points):
        end = points[i + 1][0] if i + 1 < len(points) else width
        for cc in range(c, end):
            spans[cc] = name
    return spans


def parse(sheet: "E.Sheet", drop_calculated: bool = True) -> dict:
    g = sheet.values  # raw values; banners sit at the group's start column
    width = max((len(r) for r in g), default=0)
    log: list[str] = []

    # categories that are computed rates/ratios — user builds these in Power BI
    CALC_CATEGORIES = {"trir & ltir", "emr", "various hsse ratios"}

    # find all section starts: col A == "Period"
    sections = []
    for i, row in enumerate(g):
        if len(row) and not E._is_blank(row[0]) and str(row[0]).strip().lower() == "period":
            region = row[2] if len(row) > 2 and not E._is_blank(row[2]) else None
            sections.append((i, str(region).strip() if region else None))
    if not sections:
        raise ValueError("No 'Period' section headers found — not a KPI Workbook sheet.")
    log.append(f"Found {len(sections)} sections: "
               f"{', '.join(s[1] or '?' for s in sections)}.")

    records = []
    dropped_sub = 0
    for si, (banner_row_idx, region) in enumerate(sections):
        cat_spans = _category_spans(g[banner_row_idx], width)
        header_idx = banner_row_idx + 1
        header = g[header_idx] if header_idx < len(g) else []
        # metric columns = those with a sub-header, excluding Year/Qu-Month (0,1)
        metric_cols = []
        for c in range(2, width):
            h = header[c] if c < len(header) else None
            if not E._is_blank(h):
                cat = cat_spans.get(c, "")
                if drop_calculated and cat.strip().lower() in CALC_CATEGORIES:
                    continue
                metric_cols.append((c, " ".join(str(h).split()).strip(), cat))
        # data rows: from header+1 until the next section start (or end)
        end_idx = sections[si + 1][0] if si + 1 < len(sections) else len(g)
        for r in range(header_idx + 1, end_idx):
            row = g[r]
            if E._nonempty_count(row) == 0:
                continue
            year = row[0] if len(row) and not E._is_blank(row[0]) else None
            period = str(row[1]).strip() if len(row) > 1 and not E._is_blank(row[1]) else ""
            # drop subtotal rows (Q1-Q4, EOY, Total); keep months
            if period.lower() not in MONTHS:
                if SUBTOTAL_RE.match(period) or period == "":
                    dropped_sub += 1
                    continue
            for c, metric, category in metric_cols:
                val = E.to_number(row[c]) if c < len(row) else None
                if val is None:
                    continue
                # Employees / Exposure Hours sit under the section banner (not a
                # real metric group) -> tag them as 'Period' consistently.
                cat = category
                if metric.lower() in ("employees", "exposure hours") or not category \
                        or category == region:
                    cat = "Period"
                records.append({
                    "Region": region, "Year": year, "Month": period,
                    "Category": cat, "Metric": metric, "Value": val,
                })

    tidy = pd.DataFrame(records, columns=["Region", "Year", "Month", "Category", "Metric", "Value"])
    log.append(f"Built {len(tidy)} rows (one per Region x Month x Metric); "
               f"dropped {dropped_sub} subtotal rows. Calculated columns kept as metrics, "
               f"not recomputed.")
    return {"tidy": tidy, "log": log}
