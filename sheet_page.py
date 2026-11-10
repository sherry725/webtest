"""One editable sheet page. render("work") / render("life") shows the logged-in user's own file."""
import io
import json

import pandas as pd
import streamlit as st
from st_aggrid import AgGrid, JsCode

import xlsx_io as X
from config import DATA_DIR, SHEETS, expand_widths


def render(kind):
    cfg = SHEETS[kind]
    user = st.session_state["user"]                       # set only by the login page
    path = DATA_DIR / f"{user}_{cfg['suffix']}.xlsx"      # built from the session user, never from user input
    st.title(f"{cfg['icon']} {cfg['title']} tracking")
    if not path.exists():
        st.error(f"File not found: data/{path.name}")
        st.stop()

    k = f"{user}_{kind}"                                  # session keys are per user and per page
    if f"{k}_df" not in st.session_state:
        st.session_state[f"{k}_df"], st.session_state[f"{k}_layout"] = X.read_template(path, cfg)
        st.session_state[f"{k}_lock_n"] = len(st.session_state[f"{k}_df"])   # rows from the file are locked; new products stay editable
    st.session_state.setdefault(f"{k}_ver", 0)            # grid key counter; never reset
    df, layout = st.session_state[f"{k}_df"], st.session_state[f"{k}_layout"]

    # ------------------------------------------------------------ JS helpers
    row_span = JsCode("function(p){ const s=%s[p.colDef.field]||{}; return s[p.node.rowIndex]||1; }"
                      % json.dumps(layout["spans"]))
    # Actual-vs-Goal rule, used by the month columns and the YTD column.
    # `fb` = style returned when no color applies (none for months, grey/bold for YTD).
    rule_tpl = """
    function(p){
      const MONTHS = %(months)s, YTD = '%(ytd)s';
      if (p.data && p.node.rowIndex < %(n)d && %(rows)s.includes(p.data.%(m)s))
        return {backgroundColor:'rgba(128,128,128,.12)'};           // locked row: grey, no color rule
      if (!p.data || p.data.%(m)s !== 'Actual') return %(fb)s;
      const goal = p.api.getDisplayedRowAtIndex(p.node.rowIndex - 1);
      if (!goal) return %(fb)s;
      const val = (node) => {                                      // YTD is calculated, so sum the months
        if (p.colDef.field === YTD) {
          let s=0, any=false;
          for (const m of MONTHS){ const v=node.data[m]; if (v!==null && v!=='' && !isNaN(Number(v))){ s+=Number(v); any=true; } }
          return any ? s : null;
        }
        const v = node.data[p.colDef.field];
        return (v===null || v===undefined || v==='' || isNaN(Number(v))) ? null : Number(v);
      };
      const a = val(p.node), g = val(goal);
      if (a === null || g === null) return %(fb)s;
      return a >= g ? {backgroundColor:'#c6efce', color:'#006100', fontWeight:'600'}
                    : {backgroundColor:'#ffc7ce', color:'#9c0006', fontWeight:'600'};
    }"""
    rule_args = {"m": cfg["metric"], "n": st.session_state[f"{k}_lock_n"], "rows": json.dumps(cfg["locked_rows"]),
                 "months": json.dumps(cfg["months"]), "ytd": cfg["ytd"]}
    color_rule = JsCode(rule_tpl % {**rule_args, "fb": "null"})
    ytd_rule = JsCode(rule_tpl % {**rule_args, "fb": "{fontWeight:'600', backgroundColor:'rgba(128,128,128,.12)'}"})
    ytd_getter = JsCode("""
    function(p){ let s=0, any=false;
      for (const m of %s){ const v=p.data[m]; if (v!==null && v!=='' && !isNaN(Number(v))){ s+=Number(v); any=true; } }
      return any ? s : null; }""" % json.dumps(cfg["months"]))
    num_parser = JsCode("""
    function(p){ const v=p.newValue; if (v===''||v===null||v===undefined) return null;
      const n=Number(v); return isNaN(n) ? p.oldValue : n; }""")
    int_fmt = JsCode("""
    function(p){ const v=p.value; if (v===null || v===undefined || v==='' || isNaN(Number(v))) return v;
      return Math.round(Number(v)); }""")
    click_to_edit = JsCode("""
    function(e){
      if (!e.colDef || e.colDef.editable === false) return;
      const busy = e.api.getEditingCells().some(c => c.rowIndex === e.rowIndex && c.column === e.column);
      if (!busy) e.api.startEditingCell({rowIndex: e.rowIndex, colKey: e.column.getColId()});
    }""")
    # Row lock: Goal rows (loaded from the file) are read-only, Actual rows stay editable
    lock_args = (st.session_state[f"{k}_lock_n"], json.dumps(cfg["locked_rows"]), cfg["metric"])
    row_editable = JsCode("function(p){ return !(p.node.rowIndex < %d && %s.includes(p.data.%s)); }" % lock_args)
    locked_look = JsCode("function(p){ return (p.node.rowIndex < %d && %s.includes(p.data.%s)) "
                         "? {backgroundColor:'rgba(128,128,128,.12)'} : null; }" % lock_args)
    repaint = JsCode("function(e){ e.api.refreshCells({force:true}); }")

    # ------------------------------------------------------------ columns
    widths = expand_widths(cfg.get("col_widths", {}))

    def col_def(c):
        f = c["field"]
        w = widths.get(f, widths.get("*"))                 # from config.py, else from the Excel file
        if w is None:
            w = int((c["width"] or 11) * 9 + 20)
        d = {"field": f, "headerName": c["header"] or "", "editable": f not in cfg["locked"],
             "width": w, "minWidth": min(w, 40)}
        if f in cfg["dropdowns"]:
            d.update(cellEditor="agSelectCellEditor", cellEditorParams={"values": cfg["dropdowns"][f]})
        if f in cfg["frozen"]:
            d["pinned"] = "left"
        if f in layout["spans"]:
            d.update(rowSpan=row_span, cellClass="cell-span")
        if f in cfg["numeric"]:
            d.update(type="numericColumn", valueParser=num_parser, valueFormatter=int_fmt)
            if cfg["locked_rows"] and f not in cfg["locked"]:
                d.update(editable=row_editable, cellStyle=locked_look)
        if f in cfg["months"]:
            d["cellStyle"] = color_rule
        if f == cfg["ytd"]:
            d.update(valueGetter=ytd_getter, cellStyle=ytd_rule)
        elif f in cfg["locked"]:
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
        "singleClickEdit": True,
        "rowHeight": 38, "headerHeight": 38, "groupHeaderHeight": 38,
        "onCellValueChanged": repaint,
        "onCellClicked": click_to_edit,
        "defaultColDef": {"resizable": True, "sortable": False, "singleClickEdit": True},
    }
    GRID = "1px solid #b7bdc8"
    custom_css = {
        ".ag-cell": {"border": GRID + " !important", "line-height": "normal !important"},
        ".ag-header-cell, .ag-header-group-cell": {"border": GRID + " !important", "font-weight": "700"},
        ".ag-header-group-cell": {"justify-content": "center"},
        ".cell-span": {"display": "flex !important", "align-items": "center",
                       "background-color": "var(--ag-background-color)", "font-weight": "600"},
        ".ag-root-wrapper": {"border": GRID + " !important"},
    }

    # ------------------------------------------------------------ buttons
    b1, b2, b3, b4, _ = st.columns([1, 1, 1, 1, 4])
    if b1.button("➕ Add product", key=f"{k}_add"):
        st.session_state[f"{k}_df"] = X.add_product(df, layout, cfg)
        st.session_state[f"{k}_ver"] += 1
        st.rerun()
    if b2.button("↺ Reload from file", key=f"{k}_reload"):
        st.session_state[f"{k}_df"], st.session_state[f"{k}_layout"] = X.read_template(path, cfg)
        st.session_state[f"{k}_lock_n"] = len(st.session_state[f"{k}_df"])
        st.session_state[f"{k}_ver"] += 1       # new grid key = browser grid rebuilt from the file
        st.rerun()

    # ------------------------------------------------------------ grid
    cols = list(df.columns)
    resp = AgGrid(
        df.copy(),                              # AgGrid adds an internal id column, so give it a copy
        gridOptions=gridOptions,
        height=84 + 38 * len(df) + 24,
        allow_unsafe_jscode=True,
        update_on=["cellValueChanged"],
        custom_css=custom_css,
        theme="streamlit",
        show_search=False,
        show_download_button=False,
        key=f"grid_{k}_{st.session_state[f'{k}_ver']}",
    )
    if resp["data"] is not None:
        edited = pd.DataFrame(resp["data"]).reindex(columns=cols).astype(object)
        st.session_state[f"{k}_df"] = edited.where(edited.notna(), None)
    df = st.session_state[f"{k}_df"]

    # ------------------------------------------------------------ save
    if b3.button("💾 Save to Excel", type="primary", key=f"{k}_save"):
        try:
            backup = X.save_to_template(path, df, layout, cfg)
            st.success(f"Saved to {path.name}. Backup of the previous version: {backup.name}")
        except PermissionError:
            st.error(f"Could not write {path.name}. Close it in Excel and try again.")

    # ------------------------------------------------------------ download (includes unsaved edits)
    buf = io.BytesIO()
    X.save_to_template(path, df, layout, cfg, out=buf)
    b4.download_button("⬇️ Download Excel", data=buf.getvalue(), file_name=f"{user}_{cfg['suffix']}_edited.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                       key=f"{k}_dl")
