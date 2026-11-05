"""
Run:   pip install streamlit streamlit-aggrid pandas openpyxl
       streamlit run app.py
Keep app.py, xlsx_io.py and template.xlsx in the same folder.
"""
import io
import json
from pathlib import Path

import pandas as pd
import streamlit as st
from st_aggrid import AgGrid, JsCode

import xlsx_io as X

TEMPLATE = Path(__file__).parent / "template.xlsx"
FROZEN = ["A", "B", "C", "D"]           # columns pinned on the left (like Excel freeze panes)
DROPDOWNS = {"A": ["car", "bus"],         # column letter -> allowed choices
             "B": ["door", "window"]}
LOCKED = {"D", "K"}                      # columns users cannot edit (K = YTD is calculated)

st.set_page_config(page_title="Tracking sheet", layout="wide")
st.title("Monthly tracking")

# ---------------------------------------------------------------- load
if "df" not in st.session_state:
    st.session_state.df, st.session_state.layout = X.read_template(TEMPLATE)
st.session_state.setdefault("ver", 0)   # grid key counter; must never reset, or the browser keeps old edits
df, layout = st.session_state.df, st.session_state.layout

# ---------------------------------------------------------------- JS helpers
spans = {c: {i: X.BLOCK for i in range(0, len(df), X.BLOCK)} for c in layout["merge_cols"]}
row_span = JsCode("function(p){ const s=%s[p.colDef.field]||{}; return s[p.node.rowIndex]||1; }" % json.dumps(spans))

color_rule = JsCode("""
function(p){
  if (!p.data || p.data.%(m)s !== 'Actual') return null;
  const goal = p.api.getDisplayedRowAtIndex(p.node.rowIndex - 1);
  if (!goal) return null;
  const a = Number(p.value), g = Number(goal.data[p.colDef.field]);
  if (p.value === null || p.value === '' || goal.data[p.colDef.field] === null || isNaN(a) || isNaN(g)) return null;
  return a >= g ? {backgroundColor:'#c6efce', color:'#006100', fontWeight:'600'}
                : {backgroundColor:'#ffc7ce', color:'#9c0006', fontWeight:'600'};
}""" % {"m": X.METRIC_COL})

ytd_getter = JsCode("""
function(p){ let s=0, any=false;
  for (const m of %s){ const v=p.data[m]; if (v!==null && v!=='' && !isNaN(Number(v))){ s+=Number(v); any=true; } }
  return any ? s : null; }""" % json.dumps(X.MONTH_COLS))

num_parser = JsCode("""
function(p){ const v=p.newValue; if (v===''||v===null||v===undefined) return null;
  const n=Number(v); return isNaN(n) ? p.oldValue : n; }""")

# Fallback for single-click editing: start the editor directly on click (works even if singleClickEdit is ignored)
click_to_edit = JsCode("""
function(e){
  if (!e.colDef || e.colDef.editable === false) return;
  const busy = e.api.getEditingCells().some(c => c.rowIndex === e.rowIndex && c.column === e.column);
  if (!busy) e.api.startEditingCell({rowIndex: e.rowIndex, colKey: e.column.getColId()});
}""")

repaint = JsCode("function(e){ e.api.refreshCells({force:true}); }")

# ---------------------------------------------------------------- columns
def col_def(c):
    f = c["field"]
    d = {"field": f, "headerName": c["header"] or "", "editable": f not in LOCKED,
         "width": int((c["width"] or 11) * 9 + 20), "minWidth": 70}
    if f in DROPDOWNS:
        d.update(cellEditor="agSelectCellEditor", cellEditorParams={"values": DROPDOWNS[f]})
    if f in FROZEN:
        d["pinned"] = "left"
    if f in layout["merge_cols"]:
        d.update(rowSpan=row_span, cellClass="cell-span")
    if f in X.NUMERIC_COLS:
        d.update(type="numericColumn", valueParser=num_parser)
    if f in X.MONTH_COLS:
        d["cellStyle"] = color_rule
    if f == X.YTD_COL:
        d.update(valueGetter=ytd_getter, cellStyle={"fontWeight": "600", "backgroundColor": "rgba(128,128,128,.12)"})
    elif f in LOCKED:
        d["cellStyle"] = {"backgroundColor": "rgba(128,128,128,.08)"}
    return d

columnDefs, groups = [], {}
for c in layout["columns"]:
    if c["group"]:
        key = c["group"][0]
        if key not in groups:
            groups[key] = {"headerName": c["group"][1], "marryChildren": True, "children": []}
            columnDefs.append(groups[key])
        groups[key]["children"].append(col_def(c))
    else:
        columnDefs.append(col_def(c))

gridOptions = {
    "columnDefs": columnDefs,
    "suppressRowTransform": True,       # required for rowSpan
    "suppressMovableColumns": True,
    "singleClickEdit": True,            # one click starts editing (default is double click)
    "rowHeight": 38,
    "headerHeight": 38,
    "groupHeaderHeight": 38,
    "onCellValueChanged": repaint,
    "onCellClicked": click_to_edit,
    "defaultColDef": {"resizable": True, "sortable": False,
                      "singleClickEdit": True},      # one click starts editing, on every column
}

# Full grid lines on every cell and header (like Excel with gridlines on)
GRID = "1px solid #b7bdc8"
custom_css = {
    ".ag-cell": {"border": GRID + " !important", "line-height": "normal !important"},
    ".ag-header-cell, .ag-header-group-cell": {"border": GRID + " !important", "font-weight": "700"},
    ".ag-header-group-cell": {"justify-content": "center"},
    ".cell-span": {"display": "flex !important", "align-items": "center",
                   "background-color": "var(--ag-background-color)", "font-weight": "600"},
    ".ag-root-wrapper": {"border": GRID + " !important"},
}

# ---------------------------------------------------------------- buttons
b1, b2, b3, b4, _ = st.columns([1, 1, 1, 1, 4])
if b1.button("➕ Add product"):
    st.session_state.df = pd.concat([df, X.new_product_rows(df)], ignore_index=True)
    st.session_state.ver += 1
    st.rerun()
if b2.button("↺ Reload from file"):
    st.session_state.df, st.session_state.layout = X.read_template(TEMPLATE)
    st.session_state.ver += 1           # new grid key = browser grid is rebuilt from the file
    st.rerun()

# ---------------------------------------------------------------- grid
cols = list(df.columns)
resp = AgGrid(
    df.copy(),                          # AgGrid adds an internal id column, so give it a copy
    gridOptions=gridOptions,
    height=84 + 38 * len(df) + 24,
    allow_unsafe_jscode=True,
    update_on=["cellValueChanged"],
    custom_css=custom_css,
    theme="streamlit",
    show_search=False,
    show_download_button=False,
    key=f"grid{st.session_state.ver}",
)
if resp["data"] is not None:
    edited = pd.DataFrame(resp["data"]).reindex(columns=cols).astype(object)
    st.session_state.df = edited.where(edited.notna(), None)

# ---------------------------------------------------------------- save
if b3.button("💾 Save to Excel", type="primary"):
    try:
        backup = X.save_to_template(TEMPLATE, st.session_state.df, layout["merge_cols"])
        st.success(f"Saved to {TEMPLATE.name}. Backup of the previous version: {backup.name}")
    except PermissionError:
        st.error("Could not write the file. Close template.xlsx in Excel and try again.")

# ---------------------------------------------------------------- download (current edits, even if not saved)
buf = io.BytesIO()
X.save_to_template(TEMPLATE, st.session_state.df, layout["merge_cols"], out=buf)
b4.download_button("⬇️ Download Excel", data=buf.getvalue(), file_name="tracking_edited.xlsx",
                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
