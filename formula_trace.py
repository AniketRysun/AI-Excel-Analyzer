"""
formula_trace.py
----------------
Bounded, honest formula-source tracing.

Given a cell that contains a formula, we:
  * extract the cell/range references it points at (incl. cross-sheet),
  * read the underlying source values for those references,
  * classify the formula as traceable or not.

We deliberately do NOT implement a full Excel calculation engine. Formulas that
rely on volatile / dynamic resolution cannot be statically traced and are
flagged "source not traceable" rather than guessed.
"""

from __future__ import annotations

import re
from typing import Any

from openpyxl.utils import range_boundaries, get_column_letter

# Functions we cannot statically resolve to a fixed source range.
UNTRACEABLE_FUNCS = {"INDIRECT", "OFFSET", "FILTER", "UNIQUE", "SORT", "SEQUENCE", "RAND", "NOW", "TODAY"}

# A1 / A1:B2 reference, optionally qualified by a sheet name ('My Sheet'! or Sheet1!)
_REF = re.compile(
    r"(?:(?P<sheet>'[^']+'|[A-Za-z_][A-Za-z0-9_.]*)!)?"
    r"(?P<rng>\$?[A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?)"
)
_FUNC = re.compile(r"([A-Z][A-Z0-9.]+)\s*\(")


def _funcs_in(formula: str) -> set[str]:
    return {m.group(1).upper() for m in _FUNC.finditer(formula)}


def extract_references(formula: str, current_sheet: str) -> list[dict]:
    """Return a list of {sheet, range} the formula points at (best effort)."""
    refs = []
    for m in _REF.finditer(formula):
        sheet = m.group("sheet")
        if sheet:
            sheet = sheet.strip("'")
        else:
            sheet = current_sheet
        rng = m.group("rng").replace("$", "")
        refs.append({"sheet": sheet, "range": rng})
    # de-duplicate
    seen, out = set(), []
    for r in refs:
        key = (r["sheet"], r["range"])
        if key not in seen:
            seen.add(key)
            out.append(r)
    return out


def _read_range(values_grid: list[list[Any]], rng: str) -> list[list[Any]]:
    min_col, min_row, max_col, max_row = range_boundaries(rng)
    out = []
    for r in range(min_row, max_row + 1):
        row = []
        for c in range(min_col, max_col + 1):
            v = values_grid[r - 1][c - 1] if (r - 1 < len(values_grid) and c - 1 < len(values_grid[r - 1])) else None
            row.append(v)
        out.append(row)
    return out


def trace_cell(book, sheet_name: str, coord: str) -> dict:
    """Trace one cell. `book` is engine_core.Book; coord like 'C5'.

    Returns dict with: formula, status ('traceable'|'untraceable'|'not_a_formula'),
    reason, functions, and sources (list of {sheet, range, values}).
    """
    sheet = book.sheets.get(sheet_name)
    if sheet is None or sheet.formulas is None:
        return {"formula": None, "status": "not_a_formula",
                "reason": "No formula information for this sheet/file type.",
                "functions": [], "sources": []}

    col_letter = re.match(r"([A-Z]+)(\d+)", coord)
    if not col_letter:
        return {"formula": None, "status": "not_a_formula", "reason": "Bad coord.",
                "functions": [], "sources": []}
    # locate the formula text
    min_col, min_row, _, _ = range_boundaries(coord)
    formula = None
    if min_row - 1 < len(sheet.formulas) and min_col - 1 < len(sheet.formulas[min_row - 1]):
        formula = sheet.formulas[min_row - 1][min_col - 1]

    if not (isinstance(formula, str) and formula.startswith("=")):
        return {"formula": formula, "status": "not_a_formula",
                "reason": "Cell is a literal value, not a formula.",
                "functions": [], "sources": []}

    funcs = _funcs_in(formula)
    blocked = funcs & UNTRACEABLE_FUNCS
    if blocked:
        return {"formula": formula, "status": "untraceable",
                "reason": f"Uses {', '.join(sorted(blocked))} — source not traceable.",
                "functions": sorted(funcs), "sources": []}

    refs = extract_references(formula, sheet_name)
    if not refs:
        return {"formula": formula, "status": "untraceable",
                "reason": "No resolvable cell/range references found.",
                "functions": sorted(funcs), "sources": []}

    sources = []
    for ref in refs:
        src = book.sheets.get(ref["sheet"])
        if src is None:
            sources.append({**ref, "values": None, "note": "referenced sheet not found"})
            continue
        try:
            vals = _read_range(src.values, ref["range"])
        except Exception as e:  # malformed range etc.
            vals = None
            ref = {**ref, "note": f"could not read range ({e})"}
        sources.append({**ref, "values": vals})

    lineage = f"{sheet_name}!{coord} <- " + " , ".join(f"{r['sheet']}!{r['range']}" for r in refs)
    return {"formula": formula, "status": "traceable", "reason": "Resolved references to source ranges.",
            "functions": sorted(funcs), "sources": sources, "lineage": lineage}


def trace_sheet(book, sheet_name: str, max_cells: int = 200) -> list[dict]:
    """Scan a sheet and trace every formula cell (capped for performance)."""
    sheet = book.sheets.get(sheet_name)
    if sheet is None or sheet.formulas is None:
        return []
    out = []
    for r, row in enumerate(sheet.formulas, start=1):
        for c, val in enumerate(row, start=1):
            if isinstance(val, str) and val.startswith("="):
                coord = f"{get_column_letter(c)}{r}"
                res = trace_cell(book, sheet_name, coord)
                res["cell"] = coord
                out.append(res)
                if len(out) >= max_cells:
                    return out
    return out
