"""
engine_core.py
--------------
Deterministic data engine for the Excel -> Power BI Granularizer.

Responsibilities:
  * load a workbook into plain grids (values AND formulas) for .xlsx/.xls/.csv
  * heuristically guess the layout (header band, id columns, total/banner rows)
  * deterministically reshape a sheet into a tidy / long granular table given a config
  * reconcile the output against the source (no values invented or lost)
  * export to CSV and to a two-sheet XLSX (clean data + transformation log)

Nothing here invents numbers. Detection only *guesses* structure; the actual
reshape is fully determined by the config the user confirms in the UI.
"""

from __future__ import annotations

import io
import re
from dataclasses import dataclass, field
from typing import Any

import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

TOTAL_RE = re.compile(r"\b(sub)?total\b|grand\s*total", re.IGNORECASE)


# --------------------------------------------------------------------------- #
# Loading
# --------------------------------------------------------------------------- #
@dataclass
class Sheet:
    name: str
    values: list[list[Any]]          # cached/computed values (numbers, text)
    formulas: list[list[Any]] | None # formula strings (=...) or None if N/A
    merged: list[tuple[int, int, int, int]] = field(default_factory=list)
    # merged ranges as (min_row, min_col, max_row, max_col), 1-indexed


@dataclass
class Book:
    sheets: dict[str, Sheet]

    @property
    def sheet_names(self) -> list[str]:
        return list(self.sheets.keys())


def _grid_from_ws(ws) -> list[list[Any]]:
    grid = []
    for row in ws.iter_rows(values_only=True):
        grid.append(list(row))
    # normalise ragged rows to equal width
    width = max((len(r) for r in grid), default=0)
    for r in grid:
        r.extend([None] * (width - len(r)))
    return grid


def load_book(file_bytes: bytes, filename: str) -> Book:
    """Load any supported file into a Book of Sheets.

    For .xlsx we read twice with openpyxl: once for cached values, once for
    formula text. For .xls / .csv we only get values (no formula tracing).
    """
    ext = filename.lower().rsplit(".", 1)[-1]

    if ext == "xlsx":
        wb_v = load_workbook(io.BytesIO(file_bytes), data_only=True, read_only=False)
        wb_f = load_workbook(io.BytesIO(file_bytes), data_only=False, read_only=False)
        sheets: dict[str, Sheet] = {}
        for name in wb_v.sheetnames:
            ws_v = wb_v[name]
            ws_f = wb_f[name]
            merged = [
                (m.min_row, m.min_col, m.max_row, m.max_col)
                for m in ws_v.merged_cells.ranges
            ]
            sheets[name] = Sheet(
                name=name,
                values=_grid_from_ws(ws_v),
                formulas=_grid_from_ws(ws_f),
                merged=merged,
            )
        return Book(sheets)

    if ext == "xls":
        xls = pd.read_excel(io.BytesIO(file_bytes), sheet_name=None, header=None)
        return Book({n: Sheet(n, df.values.tolist(), None) for n, df in xls.items()})

    if ext == "csv":
        df = pd.read_csv(io.BytesIO(file_bytes), header=None, dtype=object)
        return Book({"CSV": Sheet("CSV", df.values.tolist(), None)})

    raise ValueError(f"Unsupported file type: .{ext}")


def apply_merges(sheet: Sheet) -> list[list[Any]]:
    """Return a copy of the value grid with merged cells filled (top-left value
    propagated across the merged range). Helps header/banner detection."""
    grid = [row[:] for row in sheet.values]
    for (r0, c0, r1, c1) in sheet.merged:
        val = grid[r0 - 1][c0 - 1] if r0 - 1 < len(grid) else None
        for r in range(r0 - 1, min(r1, len(grid))):
            for c in range(c0 - 1, min(c1, len(grid[r]))):
                if grid[r][c] in (None, ""):
                    grid[r][c] = val
    return grid


# --------------------------------------------------------------------------- #
# Helpers
# --------------------------------------------------------------------------- #
def _is_blank(v: Any) -> bool:
    return v is None or (isinstance(v, str) and v.strip() == "")


def _nonempty_count(row: list[Any]) -> int:
    return sum(0 if _is_blank(v) else 1 for v in row)


def to_number(v: Any):
    """Best-effort numeric coercion. Returns float or None (never raises)."""
    if v is None:
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip()
    if s == "":
        return None
    pct = s.endswith("%")
    s = s.replace(",", "").replace("$", "").replace("€", "").replace("£", "").rstrip("%").strip()
    s = s.replace("(", "-").replace(")", "")
    try:
        n = float(s)
        return n / 100.0 if pct else n
    except ValueError:
        return None


def _looks_numeric_row(row: list[Any], min_frac: float = 0.4) -> bool:
    cells = [v for v in row if not _is_blank(v)]
    if not cells:
        return False
    nums = sum(1 for v in cells if to_number(v) is not None)
    return nums / len(cells) >= min_frac


# --------------------------------------------------------------------------- #
# Heuristic detection (only guesses; user confirms in the UI)
# --------------------------------------------------------------------------- #
@dataclass
class LayoutGuess:
    header_start: int          # 0-indexed row where the header band begins
    header_rows: int           # number of header rows
    data_start: int            # 0-indexed first data row
    id_columns: list[int]      # column indexes that look like identifiers
    total_rows: list[int]      # 0-indexed rows that look like totals/subtotals
    width: int


def detect_layout(grid: list[list[Any]]) -> LayoutGuess:
    width = max((len(r) for r in grid), default=0)
    n = len(grid)

    # 1) header band start: first reasonably "dense" row
    dense_threshold = max(3, int(0.5 * width))
    header_start = 0
    for i, row in enumerate(grid):
        if _nonempty_count(row) >= dense_threshold:
            header_start = i
            break

    # 2) data start: first numeric-looking row at/after header_start
    data_start = header_start + 1
    for i in range(header_start, n):
        if _looks_numeric_row(grid[i]):
            data_start = i
            break
    header_rows = max(1, data_start - header_start)

    # 3) id columns: leading columns whose data cells are mostly non-numeric
    id_columns = []
    for c in range(width):
        col_cells = [grid[r][c] for r in range(data_start, n) if c < len(grid[r])]
        col_cells = [v for v in col_cells if not _is_blank(v)]
        if not col_cells:
            continue
        non_num = sum(1 for v in col_cells if to_number(v) is None)
        if non_num / len(col_cells) >= 0.6:
            id_columns.append(c)
        else:
            break  # stop at the first numeric (measure) column
    if not id_columns:
        id_columns = [0]
    # never emit a column index outside the actual width
    id_columns = [c for c in id_columns if 0 <= c < width] or [0]

    # 4) total / subtotal rows (match in any id column)
    total_rows = []
    for r in range(data_start, n):
        joined = " ".join(str(grid[r][c]) for c in id_columns if c < len(grid[r]) and not _is_blank(grid[r][c]))
        if TOTAL_RE.search(joined):
            total_rows.append(r)

    return LayoutGuess(header_start, header_rows, data_start, id_columns, total_rows, width)


# --------------------------------------------------------------------------- #
# Reshape (deterministic, driven by confirmed config)
# --------------------------------------------------------------------------- #
@dataclass
class ReshapeConfig:
    header_start: int
    header_rows: int
    data_start: int
    id_columns: list[int]
    header_sep: str = " | "
    drop_totals: bool = True
    use_row_hierarchy: bool = False   # treat "only first col filled" rows as section banners
    hierarchy_name: str = "Section"
    split_variable: bool = False      # split the melted Variable by header_sep
    split_names: list[str] = field(default_factory=list)


def _flatten_headers(grid, cfg: ReshapeConfig) -> list[str]:
    width = max((len(r) for r in grid), default=0)
    band = [grid[cfg.header_start + k] if cfg.header_start + k < len(grid) else []
            for k in range(cfg.header_rows)]
    # forward-fill each header row horizontally (handles merged/blank header cells)
    filled = []
    for row in band:
        out, last = [], None
        for c in range(width):
            v = row[c] if c < len(row) else None
            if not _is_blank(v):
                last = str(v).strip()
            out.append(last)
        filled.append(out)
    names = []
    for c in range(width):
        parts = [filled[k][c] for k in range(len(filled)) if filled[k][c]]
        # de-dup consecutive identical parts
        dedup = []
        for p in parts:
            if not dedup or dedup[-1] != p:
                dedup.append(p)
        names.append(cfg.header_sep.join(dedup) if dedup else f"col_{get_column_letter(c+1)}")
    # ensure uniqueness
    seen, uniq = {}, []
    for nm in names:
        if nm in seen:
            seen[nm] += 1
            uniq.append(f"{nm}_{seen[nm]}")
        else:
            seen[nm] = 0
            uniq.append(nm)
    return uniq


def reshape(sheet: Sheet, cfg: ReshapeConfig) -> dict:
    grid = apply_merges(sheet)
    width = max((len(r) for r in grid), default=0)
    names = _flatten_headers(grid, cfg)

    log: list[str] = []
    log.append(f"Header band: rows {cfg.header_start+1}-{cfg.header_start+cfg.header_rows} "
               f"({cfg.header_rows} row(s)).")
    log.append(f"Data starts at row {cfg.data_start+1}.")

    id_names = [names[c] for c in cfg.id_columns]
    val_cols = [c for c in range(width) if c not in cfg.id_columns]

    rows, dropped_blank, dropped_total, set_aside_totals = [], 0, 0, []
    current_section = None
    for r in range(cfg.data_start, len(grid)):
        row = grid[r]
        if _nonempty_count(row) == 0:
            dropped_blank += 1
            continue
        first_vals = [row[c] for c in cfg.id_columns if c < len(row)]
        joined = " ".join(str(v) for v in first_vals if not _is_blank(v))

        if cfg.drop_totals and TOTAL_RE.search(joined):
            dropped_total += 1
            set_aside_totals.append([row[c] if c < len(row) else None for c in val_cols])
            continue

        # section banner: only the first id column filled, rest of row empty
        if cfg.use_row_hierarchy:
            only_first = (not _is_blank(row[cfg.id_columns[0]] if cfg.id_columns[0] < len(row) else None)
                          and _nonempty_count(row) == 1)
            if only_first:
                current_section = str(row[cfg.id_columns[0]]).strip()
                continue

        rec = {names[c]: (row[c] if c < len(row) else None) for c in cfg.id_columns}
        if cfg.use_row_hierarchy:
            rec[cfg.hierarchy_name] = current_section
        for c in val_cols:
            rec[names[c]] = row[c] if c < len(row) else None
        rows.append(rec)

    wide = pd.DataFrame(rows)
    log.append(f"Dropped {dropped_blank} blank row(s) and {dropped_total} total/subtotal row(s).")
    if cfg.use_row_hierarchy:
        log.append(f"Filled row hierarchy down into column '{cfg.hierarchy_name}'.")

    id_vars = list(id_names)
    if cfg.use_row_hierarchy:
        id_vars = [cfg.hierarchy_name] + id_vars
    value_vars = [names[c] for c in val_cols]

    tidy = wide.melt(id_vars=id_vars, value_vars=value_vars,
                     var_name="Variable", value_name="Value")
    tidy["Value"] = tidy["Value"].map(to_number)
    tidy = tidy.dropna(subset=["Value"]).reset_index(drop=True)
    log.append(f"Unpivoted {len(value_vars)} measure column(s) into long format "
               f"-> {len(tidy)} granular rows.")

    if cfg.split_variable:
        parts = tidy["Variable"].str.split(re.escape(cfg.header_sep), expand=True)
        ncols = parts.shape[1]
        col_names = (cfg.split_names + [f"Level_{i+1}" for i in range(ncols)])[:ncols]
        for i in range(ncols):
            tidy[col_names[i]] = parts[i].str.strip()
        tidy = tidy.drop(columns=["Variable"])
        log.append(f"Split header into columns: {', '.join(col_names)}.")

    return {
        "tidy": tidy,
        "wide": wide,
        "log": log,
        "id_vars": id_vars,
        "value_vars": value_vars,
        "set_aside_totals": set_aside_totals,
    }


# --------------------------------------------------------------------------- #
# Reconciliation
# --------------------------------------------------------------------------- #
def reconcile(result: dict, sheet: Sheet, cfg: ReshapeConfig) -> dict:
    """Global sum check: total of all output values vs total of the source
    measure cells (excluding total rows). Confirms nothing was added or lost."""
    grid = apply_merges(sheet)
    width = max((len(r) for r in grid), default=0)
    val_cols = [c for c in range(width) if c not in cfg.id_columns]

    src_total = 0.0
    for r in range(cfg.data_start, len(grid)):
        row = grid[r]
        joined = " ".join(str(row[c]) for c in cfg.id_columns
                          if c < len(row) and not _is_blank(row[c]))
        if cfg.drop_totals and TOTAL_RE.search(joined):
            continue
        if cfg.use_row_hierarchy and _nonempty_count(row) == 1:
            continue
        for c in val_cols:
            n = to_number(row[c] if c < len(row) else None)
            if n is not None:
                src_total += n

    out_total = float(result["tidy"]["Value"].sum())
    diff = out_total - src_total
    ok = abs(diff) < max(1e-6, abs(src_total) * 1e-9)
    return {"source_total": src_total, "output_total": out_total, "difference": diff, "ok": ok}


# --------------------------------------------------------------------------- #
# Export
# --------------------------------------------------------------------------- #
def to_csv_bytes(df: pd.DataFrame) -> bytes:
    return df.to_csv(index=False).encode("utf-8")


def to_xlsx_bytes(df: pd.DataFrame, log_lines: list[str]) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as xw:
        df.to_excel(xw, index=False, sheet_name="Granular")
        pd.DataFrame({"Transformation log": log_lines}).to_excel(
            xw, index=False, sheet_name="Transformation_Log")
    return buf.getvalue()
