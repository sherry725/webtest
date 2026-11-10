"""Read an Excel file into a DataFrame and save edits back, keeping all existing
formatting, merges and widths."""
import datetime as dt
import shutil
from copy import copy
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter as L, column_index_from_string as CI

from config import BLOCK, FIRST_DATA, GREEN, HEADER_ROWS, RED


def _num(v):
    try:
        if v is None or v == "" or pd.isna(v):
            return None
        return float(v)
    except (TypeError, ValueError):
        return None


def read_template(path, cfg):
    """Return (df, layout). df columns are Excel letters (A, B, ...).
    layout: header columns, row spans (from the file's own merges) and block_cols
    (columns merged once per product, which are merged again for newly added products)."""
    ws = load_workbook(path).active
    max_col = ws.max_column
    letters = [L(c) for c in range(1, max_col + 1)]
    rows = [[ws.cell(r, c).value for c in range(1, max_col + 1)]
            for r in range(FIRST_DATA, ws.max_row + 1)]
    df = pd.DataFrame(rows, columns=letters, dtype=object)
    df = df.astype(object).where(df.notna(), None)

    merges = list(ws.merged_cells.ranges)
    columns = []
    for c in range(1, max_col + 1):
        h1, h2 = ws.cell(1, c).value, ws.cell(2, c).value
        group, header = None, h2 or h1
        for m in merges:
            if m.min_row == 1 and m.min_col <= c <= m.max_col:
                if m.max_row >= HEADER_ROWS:
                    header = h1
                elif m.max_col > m.min_col:
                    group, header = (m.coord, h1), h2
        width = ws.column_dimensions[L(c)].width if L(c) in ws.column_dimensions else None
        columns.append({"field": L(c), "header": header, "group": group, "width": width})

    spans, block_cols = {}, []
    for m in merges:
        height = m.max_row - m.min_row + 1
        if m.min_row >= FIRST_DATA and height > 1 and m.min_col == m.max_col:
            col = L(m.min_col)
            spans.setdefault(col, {})[m.min_row - FIRST_DATA] = height
            if m.min_row == FIRST_DATA and height == BLOCK:
                block_cols.append(col)
    return df, {"columns": columns, "spans": spans, "block_cols": sorted(block_cols, key=CI)}


def add_product(df, layout, cfg):
    """Append a blank Goal/Actual pair; returns the new df and registers its merged cells."""
    start = len(df)
    goal, actual = {c: None for c in df.columns}, {c: None for c in df.columns}
    goal[cfg["metric"]], actual[cfg["metric"]] = "Goal", "Actual"
    for col in layout["block_cols"]:
        layout["spans"].setdefault(col, {})[start] = BLOCK
    new = pd.DataFrame([goal, actual], columns=df.columns, dtype=object)
    return pd.concat([df, new], ignore_index=True)


def save_to_template(path, df, layout, cfg, out=None):
    """Write df into the Excel file. With out=None the file itself is overwritten (after a
    backup, which is returned). With out=<file-like> the file is left alone and the edited
    workbook is written to out (used for downloads)."""
    path = Path(path)
    backup = None
    if out is None:
        backup_dir = path.parent / "backups"
        backup_dir.mkdir(exist_ok=True)
        backup = backup_dir / f"{path.stem}_{dt.datetime.now():%Y%m%d_%H%M%S}{path.suffix}"
        shutil.copy2(path, backup)

    metric, months, ytd, numeric = cfg["metric"], cfg["months"], cfg["ytd"], cfg["numeric"]
    wb = load_workbook(path)
    ws = wb.active
    max_col = ws.max_column

    # 1) products added in the app: copy template styles + merges for them
    have = (ws.max_row - HEADER_ROWS) // BLOCK
    want = len(df) // BLOCK
    for b in range(have, want):
        r0 = FIRST_DATA + b * BLOCK
        for off in range(BLOCK):
            for c in range(1, max_col + 1):
                src = ws.cell(FIRST_DATA + off, c)
                if isinstance(src, MergedCell):
                    src = ws.cell(FIRST_DATA, c)
                ws.cell(r0 + off, c)._style = copy(src._style)
        for col in layout["block_cols"]:
            ws.merge_cells(start_row=r0, start_column=CI(col), end_row=r0 + BLOCK - 1, end_column=CI(col))

    # 2) values (only the top-left cell of a merge holds a value)
    for i in range(len(df)):
        r = FIRST_DATA + i
        for c in range(1, max_col + 1):
            cell = ws.cell(r, c)
            if isinstance(cell, MergedCell):
                continue
            col = L(c)
            v = df.iloc[i][col]
            if col == ytd:
                vals = [x for x in (_num(df.iloc[i][m]) for m in months) if x is not None]
                v = round(sum(vals), 6) if vals else None
            if col in numeric:
                v = _num(v)
                if v is not None and v == int(v):
                    v = int(v)
            elif v is None or (isinstance(v, float) and pd.isna(v)) or v == "":
                v = None
            cell.value = v

    # 3) Actual-vs-Goal colors (replaces stale manual fills)
    def val(row, col):                      # YTD is the sum of the months
        if col == ytd:
            vals = [v for v in (_num(row[m]) for m in months) if v is not None]
            return sum(vals) if vals else None
        return _num(row[col])

    for i in range(1, len(df)):
        if df.iloc[i][metric] != "Actual":
            continue
        for col in months + [ytd]:
            cell = ws[f"{col}{FIRST_DATA + i}"]
            a, g = val(df.iloc[i], col), val(df.iloc[i - 1], col)
            if a is None or g is None:
                cell.fill = PatternFill(fill_type=None)
                continue
            fill, font = GREEN if a >= g else RED
            cell.fill = PatternFill("solid", start_color=fill, end_color=fill)
            f = copy(cell.font)
            cell.font = Font(name=f.name, sz=f.sz, b=f.b, i=f.i, color=font)

    wb.save(out if out is not None else path)
    return backup
