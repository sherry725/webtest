"""All per-file settings live here. Columns are Excel letters."""
from pathlib import Path

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
        "dropdowns": {"A": ["car", "bus"], "B": ["door", "window"]},
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
        "dropdowns": {},                        # e.g. {"B": ["food", "house"]}
    },
}
