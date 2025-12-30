import tkinter as tk
from tkinter import ttk
import pandas as pd
import operator
from typing import Optional, List, Dict, Any

# ---------------- Operators ----------------
OPS = {
    "==": operator.eq,
    "!=": operator.ne,
    ">": operator.gt,
    "<": operator.lt,
    ">=": operator.ge,
    "<=": operator.le,
}

COLUMN_OPS = ("== Column", "!= Column")

# ---------------- Helpers ----------------
def _clean_numeric_series(s: pd.Series) -> pd.Series:
    s2 = s.astype(str).str.strip()
    s2 = s2.str.replace("%", "", regex=False).str.replace(",", "", regex=False)
    return pd.to_numeric(s2, errors="coerce")


class FiltersFrame(ttk.Frame):
    """
    Multi-row filters with AND / OR per row
    Supports:
      - Value comparison
      - Column-to-column comparison
    """

    def __init__(self, parent, on_change_callback=None, df: Optional[pd.DataFrame] = None):
        super().__init__(parent)
        self.on_change = on_change_callback
        self.df = df
        self.rows: List[Optional[Dict[str, Any]]] = []
        self._build_ui()

    # ---------------- UI ----------------
    def _build_ui(self):
        top = ttk.Frame(self)
        top.pack(fill="x", padx=6, pady=6)

        ttk.Label(top, text="Filters (AND / OR per row)").pack(side="left")
        ttk.Button(top, text="Add Filter", command=self.add_row).pack(side="left", padx=8)
        ttk.Button(top, text="Clear", command=self.reset).pack(side="left")

        wrapper = ttk.Frame(self)
        wrapper.pack(fill="both", expand=True, padx=6, pady=(0, 6))

        self.canvas = tk.Canvas(wrapper, highlightthickness=0)
        self.inner = ttk.Frame(self.canvas)
        vs = ttk.Scrollbar(wrapper, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=vs.set)

        self.canvas.pack(side="left", fill="both", expand=True)
        vs.pack(side="right", fill="y")

        self.win = self.canvas.create_window((0, 0), window=self.inner, anchor="nw")
        self.inner.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.canvas.bind("<Configure>", lambda e: self.canvas.itemconfigure(self.win, width=e.width))

        hdr = ttk.Frame(self.inner)
        hdr.grid(row=0, column=0, sticky="w", pady=(2, 6))

        ttk.Label(hdr, text="Join", width=6).grid(row=0, column=0)
        ttk.Label(hdr, text="Column", width=22).grid(row=0, column=1)
        ttk.Label(hdr, text="Operator", width=14).grid(row=0, column=2)
        ttk.Label(hdr, text="Value", width=26).grid(row=0, column=3)
        ttk.Label(hdr, text="Compare Column", width=22).grid(row=0, column=4)

        self.add_row()

    # ---------------- Rows ----------------
    def _columns(self):
        return [] if self.df is None else list(self.df.columns)

    def add_row(self, preset: Dict[str, Any] | None = None):
        idx = len([r for r in self.rows if r])
        f = ttk.Frame(self.inner)
        f.grid(row=idx + 1, column=0, sticky="w", pady=2)

        join = tk.StringVar(value="AND" if idx else "")
        col = tk.StringVar()
        op = tk.StringVar(value="==")
        val = tk.StringVar()
        cmp_col = tk.StringVar()

        join_cb = ttk.Combobox(f, values=["AND", "OR"], textvariable=join, width=6, state="readonly")
        if idx == 0:
            join_cb.configure(state="disabled")
        join_cb.grid(row=0, column=0)

        col_cb = ttk.Combobox(f, values=self._columns(), textvariable=col, width=22, state="readonly")
        col_cb.grid(row=0, column=1, padx=4)

        op_cb = ttk.Combobox(
            f,
            values=list(OPS.keys()) + ["contains", "in"] + list(COLUMN_OPS),
            textvariable=op,
            width=14,
            state="readonly"
        )
        op_cb.grid(row=0, column=2, padx=4)

        val_cb = ttk.Combobox(f, textvariable=val, width=26)
        val_cb.grid(row=0, column=3, padx=4)

        cmp_cb = ttk.Combobox(f, values=self._columns(), textvariable=cmp_col, width=22, state="disabled")
        cmp_cb.grid(row=0, column=4, padx=4)

        ttk.Button(f, text="✖", width=3, command=lambda: self._remove_row(f)).grid(row=0, column=5, padx=4)

        row = {
            "frame": f,
            "join": join,
            "col": col,
            "op": op,
            "val": val,
            "cmp": cmp_col,
            "val_cb": val_cb,
            "cmp_cb": cmp_cb,
        }
        self.rows.append(row)

        col_cb.bind("<<ComboboxSelected>>", lambda e, r=row: self._populate_values(r))
        op_cb.bind("<<ComboboxSelected>>", lambda e, r=row: self._op_changed(r))
        val_cb.bind("<<ComboboxSelected>>", lambda e: self._changed())
        cmp_cb.bind("<<ComboboxSelected>>", lambda e: self._changed())

        if preset:
            join.set(preset.get("join", ""))
            col.set(preset.get("col", ""))
            op.set(preset.get("op", "=="))
            val.set(preset.get("val", ""))
            cmp_col.set(preset.get("cmp", ""))

        self._op_changed(row)
        self._changed()

    def _remove_row(self, frame):
        for i, r in enumerate(self.rows):
            if r and r["frame"] == frame:
                r["frame"].destroy()
                self.rows[i] = None
                break
        self.rows = [r for r in self.rows if r]
        self._changed()

    # ---------------- UI Logic ----------------
    def _op_changed(self, row):
        if row["op"].get() in COLUMN_OPS:
            row["val_cb"].configure(state="disabled")
            row["cmp_cb"].configure(state="readonly")
        else:
            row["val_cb"].configure(state="normal")
            row["cmp_cb"].configure(state="disabled")
        self._changed()

    def _populate_values(self, row):
        if self.df is None:
            return
        col = row["col"].get()
        if col in self.df.columns:
            vals = self.df[col].dropna().astype(str).unique().tolist()[:300]
            row["val_cb"]["values"] = ["Column Average"] + vals

    # ---------------- Apply ----------------
    def apply_filters(self, df: pd.DataFrame) -> pd.DataFrame:
        if df is None or df.empty:
            return df

        masks, joins = [], []

        for r in self.rows:
            col, op, val, cmp_col = r["col"].get(), r["op"].get(), r["val"].get(), r["cmp"].get()
            if not col or col not in df.columns:
                continue

            s = df[col]

            if op in COLUMN_OPS and cmp_col in df.columns:
                s2 = df[cmp_col]
                m = s.astype(str) != s2.astype(str) if op.startswith("!=") else s.astype(str) == s2.astype(str)

            elif op == "contains":
                m = s.astype(str).str.contains(str(val), case=False, na=False)

            elif op == "in":
                values = [v.strip() for v in str(val).split(",") if v.strip()]
                m = s.astype(str).isin(values)

            elif val.lower() == "column average":
                avg = _clean_numeric_series(s).mean()
                m = OPS[op](_clean_numeric_series(s), avg)

            else:
                try:
                    m = OPS[op](_clean_numeric_series(s), float(val))
                except Exception:
                    m = OPS[op](s.astype(str), str(val))

            masks.append(m.fillna(False))
            joins.append(r["join"].get())

        if not masks:
            return df

        final = masks[0]
        for i in range(1, len(masks)):
            final = final | masks[i] if joins[i] == "OR" else final & masks[i]

        return df[final]

    # ---------------- Presets API (FIX) ----------------
    def get_config(self) -> Dict[str, Any]:
        return {
            "filters": [
                {
                    "join": r["join"].get(),
                    "col": r["col"].get(),
                    "op": r["op"].get(),
                    "val": r["val"].get(),
                    "cmp": r["cmp"].get(),
                }
                for r in self.rows
            ]
        }

    def load_config(self, cfg: Dict[str, Any]):
        self.reset(silent=True)
        for f in cfg.get("filters", []):
            self.add_row(f)
        self._changed()

    # ---------------- Misc ----------------
    def reset(self, silent=False):
        for r in self.rows:
            r["frame"].destroy()
        self.rows = []
        self.add_row()
        if not silent:
            self._changed()

    def refresh_source_df(self, df):
        self.df = df
        for r in self.rows:
            r["cmp_cb"]["values"] = self._columns()
            self._populate_values(r)

    def _changed(self):
        if callable(self.on_change):
            self.on_change()
