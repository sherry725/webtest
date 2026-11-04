"""Read the Excel template into a DataFrame and save edits back into the same file,
keeping all of the template's existing formatting, merges and widths."""
import datetime as dt
import shutil
from copy import copy
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter as L, column_index_from_string as CI

# ------------------------------------------------------------------ CONFIG (edit to match your sheet)
HEADER_ROWS = 2          # rows 1-2 are headers
FIRST_DATA = 3           # data starts at row 3
BLOCK = 2                # one product = 2 rows (Goal + Actual)
METRIC_COL = "D"         # column holding "Goal" / "Actual"
MONTH_COLS = ["E", "F", "G", "H", "I", "J"]    # Jan..Jun, used for YTD and Actual-vs-Goal colors
YTD_COL = "K"            # filled automatically on save (sum of MONTH_COLS)
NUMERIC_COLS = ["E", "F", "G", "H", "I", "J", "K", "L"]
GREEN = ("C6EFCE", "006100")   # (fill, font) when Actual >= Goal
RED = ("FFC7CE", "9C0006")     # when Actual < Goal
# ------------------------------------------------------------------


def _num(v):
    try:
        if v is None or v == "" or pd.isna(v):
            return None
        return float(v)
    except (TypeError, ValueError):
        return None


def read_template(path):
    """Return (df, layout). df columns are Excel letters (A, B, ...)."""
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
                    header = h1                       # header merged across both rows
                elif m.max_col > m.min_col:
                    group, header = (m.coord, h1), h2  # group header over several columns
        columns.append({"field": L(c), "header": header, "group": group,
                        "width": ws.column_dimensions[L(c)].width if L(c) in ws.column_dimensions else None})
    merge_cols = sorted({L(m.min_col) for m in merges
                         if m.min_row == FIRST_DATA and m.max_row > m.min_row and m.min_col == m.max_col},
                        key=CI)
    return df, {"columns": columns, "merge_cols": merge_cols}


def new_product_rows(df):
    """Blank Goal/Actual rows to append for a new product."""
    blank = {c: None for c in df.columns}
    goal, actual = dict(blank), dict(blank)
    goal[METRIC_COL], actual[METRIC_COL] = "Goal", "Actual"
    return pd.DataFrame([goal, actual], columns=df.columns, dtype=object)


def save_to_template(path, df, merge_cols):
    """Write df into the template file. Returns the backup path."""
    path = Path(path)
    backup_dir = path.parent / "backups"
    backup_dir.mkdir(exist_ok=True)
    backup = backup_dir / f"{path.stem}_{dt.datetime.now():%Y%m%d_%H%M%S}{path.suffix}"
    shutil.copy2(path, backup)

    wb = load_workbook(path)
    ws = wb.active
    max_col = ws.max_column

    # 1) extend the sheet if the user added products: copy template styles + merges
    have = (ws.max_row - HEADER_ROWS) // BLOCK
    want = len(df) // BLOCK
    for b in range(have, want):
        r0 = FIRST_DATA + b * BLOCK
        for off in range(BLOCK):
            for c in range(1, max_col + 1):
                src, dst = ws.cell(FIRST_DATA + off, c), ws.cell(r0 + off, c)
                if not isinstance(src, MergedCell):
                    dst._style = copy(src._style)
                else:   # style of merged-away cell: use top-left of its block
                    dst._style = copy(ws.cell(FIRST_DATA, c)._style)
        for col in merge_cols:
            ws.merge_cells(start_row=r0, start_column=CI(col), end_row=r0 + BLOCK - 1, end_column=CI(col))

    # 2) values
    for i in range(len(df)):
        r = FIRST_DATA + i
        for c in range(1, max_col + 1):
            cell = ws.cell(r, c)
            if isinstance(cell, MergedCell):
                continue                      # only the top-left cell of a merge holds the value
            col = L(c)
            v = df.iloc[i][col]
            if col in NUMERIC_COLS:
                v = _num(v)
                if v is not None and v == int(v):
                    v = int(v)
            elif v is None or (isinstance(v, float) and pd.isna(v)) or v == "":
                v = None
            if col == YTD_COL:
                vals = [_num(df.iloc[i][m]) for m in MONTH_COLS]
                vals = [x for x in vals if x is not None]
                v = round(sum(vals), 6) if vals else None
                if v is not None and v == int(v):
                    v = int(v)
            cell.value = v

    # 3) Actual-vs-Goal colors (replaces any stale manual fills)
    for i in range(1, len(df)):
        if df.iloc[i][METRIC_COL] != "Actual":
            continue
        for col in MONTH_COLS:
            cell = ws[f"{col}{FIRST_DATA + i}"]
            a, g = _num(df.iloc[i][col]), _num(df.iloc[i - 1][col])
            if a is None or g is None:
                cell.fill = PatternFill(fill_type=None)
                continue
            fill, font = GREEN if a >= g else RED
            cell.fill = PatternFill("solid", start_color=fill, end_color=fill)
            f = copy(cell.font)
            cell.font = Font(name=f.name, sz=f.sz, b=f.b, i=f.i, color=font)

    wb.save(path)
    return backup
