"""All per-file settings live here. Columns are Excel letters."""
from pathlib import Path

from openpyxl.utils import column_index_from_string as _CI, get_column_letter as _L

BASE = Path(__file__).parent
DATA_DIR = BASE / "data"          # put username_work.xlsx and username_life.xlsx here
USERS_FILE = BASE / "users.json"  # created with manage_users.py

HEADER_ROWS = 2     # rows 1-2 are headers
FIRST_DATA = 3      # data starts at row 3
BLOCK = 2           # one product = 2 rows (Goal + Actual)

GREEN = ("C6EFCE", "006100")   # (fill, font) when Actual >= Goal
RED = ("FFC7CE", "9C0006")     # when Actual < Goal

SHEETS = {
    "work": {
        "title": "Work",
        "icon": "💼",
        "suffix": "work",                       # file = <username>_work.xlsx
        "metric": "D",                          # column holding Goal / Actual
        "months": list("EFGHIJ"),               # Jan..Jun: used for YTD and colors
        "ytd": "K",                             # calculated automatically
        "numeric": list("EFGHIJKL"),
        "frozen": list("ABCD"),                 # pinned on the left like Excel freeze panes
        "locked": {"D", "K"},                   # users cannot edit these columns
        "locked_rows": ["Goal"],                # rows whose Metrics value is listed here are read-only ([] = none)
        "dropdowns": {"A": ["car", "bus"], "B": ["door", "window"]},
        # column widths in pixels: a column ("E"), a range ("A-D") or "*" for all others.
        # Columns not listed keep the width from the Excel file.
        "col_widths": {"A-D": 50, "E-F": 70},
    },
    "life": {
        "title": "Life",
        "icon": "🏡",
        "suffix": "life",                       # file = <username>_life.xlsx
        "metric": "C",
        "months": list("DEFGHI"),
        "ytd": "J",
        "numeric": list("DEFGHIJK"),
        "frozen": list("ABC"),
        "locked": {"C", "J"},
        "locked_rows": ["Goal"],
        "dropdowns": {},                        # e.g. {"B": ["food", "house"]}
        "col_widths": {"A-C": 50, "D-E": 70},
    },
}


def expand_widths(spec):
    """{"A-D": 50, "E": 70, "*": 60} -> {"A": 50, ..., "D": 50, "E": 70, "*": 60}.
    Keys are a column ("E"), a range ("A-D") or "*" (all columns not listed)."""
    out = {}
    for key, width in spec.items():
        if key == "*":
            out["*"] = width
            continue
        first, _, last = key.partition("-")
        for i in range(_CI(first.strip()), _CI((last or first).strip()) + 1):
            out[_L(i)] = width
    return out
