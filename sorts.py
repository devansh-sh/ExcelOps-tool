import tkinter as tk
from tkinter import ttk, messagebox
import pandas as pd

class SortsFrame(ttk.Frame):
    """
    Multi-level sort with per-row Join (AND/OR).
    Semantics:
      • Rows are evaluated top→bottom.
      • 'AND' means "keep adding sort keys to the current group".
      • 'OR' starts a NEW group.
      • Apply groups RIGHT→LEFT with stable sort (kind='mergesort'):
          - Sort last group first (lowest precedence),
          - then previous group, …,
          - first group last (highest precedence).
        This gives OR real meaning while staying coherent with pandas sorting.
    API used by main.py:
      - apply_sorts(df) -> df
      - get_config() / load_config(cfg)
      - reset()
      - refresh_columns(df)
    """
    def __init__(self, parent, on_change_callback, df: pd.DataFrame | None):
        super().__init__(parent)
        self.on_change = on_change_callback
        self.df = df
        self.rows = []  # [{frame, join_var?, col_var, order_var, widgets...}]
        self._build_ui()

    # ---------- UI ----------
    def _build_ui(self):
        top = ttk.Frame(self)
        top.pack(fill="x", padx=6, pady=6)

        ttk.Button(top, text="Add Sort Level", command=self.add_row).pack(side="left")
        ttk.Button(top, text="Clear", command=self.reset).pack(side="left", padx=6)

        wrap = ttk.Frame(self)
        wrap.pack(fill="both", expand=True, padx=6, pady=(0,6))

        self.canvas = tk.Canvas(wrap, borderwidth=0, highlightthickness=0)
        self.inner = ttk.Frame(self.canvas)
        vs = ttk.Scrollbar(wrap, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=vs.set)

        self.canvas.pack(side="left", fill="both", expand=True)
        vs.pack(side="right", fill="y")
        self.win_id = self.canvas.create_window((0,0), window=self.inner, anchor="nw")
        self.inner.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.canvas.bind("<Configure>", self._stretch)

        # header
        hdr = ttk.Frame(self.inner)
        hdr.grid(row=0, column=0, sticky="ew", padx=2, pady=(2,6))
        ttk.Label(hdr, text="Join").grid(row=0, column=0, padx=4, sticky="w")
        ttk.Label(hdr, text="Column", width=28).grid(row=0, column=1, padx=4, sticky="w")
        ttk.Label(hdr, text="Order", width=12).grid(row=0, column=2, padx=4, sticky="w")

        self.grid_start = 1

    def _stretch(self, event):
        try:
            self.canvas.itemconfigure(self.win_id, width=event.width)
        except Exception:
            pass

    def _columns_list(self):
        return [] if self.df is None else list(self.df.columns)

    def add_row(self, preset=None):
        """
        preset: optional (join, col, order)
          join: 'AND'/'OR' (ignored for first row)
          order: 'Ascending'|'Descending'
        """
        idx_alive = len([r for r in self.rows if r is not None])
        rowf = ttk.Frame(self.inner)
        rowf.grid(row=self.grid_start + idx_alive, column=0, sticky="ew", padx=2, pady=2)

        join_var = tk.StringVar(value="AND")
        if idx_alive == 0:
            join_w = ttk.Label(rowf, text="—")
            join_ref = None
        else:
            join_w = ttk.Combobox(rowf, values=["AND","OR"], textvariable=join_var, state="readonly", width=6)
            join_w.bind("<<ComboboxSelected>>", lambda e: self._changed())

        col_var = tk.StringVar()
        order_var = tk.StringVar(value="Ascending")

        col_cb = ttk.Combobox(rowf, values=self._columns_list(), textvariable=col_var, state="readonly", width=28)
        ord_cb = ttk.Combobox(rowf, values=["Ascending","Descending"], textvariable=order_var, state="readonly", width=12)
        rem_btn = ttk.Button(rowf, text="✖", width=3, command=lambda rf=rowf: self._remove_row(rf))

        join_w.grid(row=0, column=0, padx=4, sticky="w")
        col_cb.grid(row=0, column=1, padx=4, sticky="w")
        ord_cb.grid(row=0, column=2, padx=4, sticky="w")
        rem_btn.grid(row=0, column=3, padx=4)

        row = {
            "frame": rowf,
            "join_widget": join_w,
            "join_var": (join_var if idx_alive > 0 else None),
            "col_var": col_var,
            "order_var": order_var,
            "col_cb": col_cb,
            "ord_cb": ord_cb,
        }
        self.rows.append(row)

        # preset populate
        if preset:
            j,c,o = preset
            if idx_alive > 0 and j in ("AND","OR"):
                join_var.set(j)
            if c:
                col_var.set(c)
            if o in ("Ascending","Descending"):
                order_var.set(o)

        col_cb.bind("<<ComboboxSelected>>", lambda e: self._changed())
        ord_cb.bind("<<ComboboxSelected>>", lambda e: self._changed())

        self._changed()

    def _remove_row(self, frame):
        for i,r in enumerate(self.rows):
            if r and r["frame"] is frame:
                try:
                    r["frame"].destroy()
                except Exception:
                    pass
                self.rows[i] = None
                break
        self._regrid_rows()
        self._changed()

    def _regrid_rows(self):
        alive = [r for r in self.rows if r is not None]
        self.rows = []
        for i, r in enumerate(alive):
            rf = r["frame"]
            for w in rf.grid_slaves():
                w.grid_forget()
            # rebuild join role
            if i == 0:
                if isinstance(r["join_widget"], ttk.Combobox):
                    try:
                        r["join_widget"].destroy()
                    except Exception:
                        pass
                    r["join_widget"] = ttk.Label(rf, text="—")
                    r["join_var"] = None
            else:
                if not isinstance(r["join_widget"], ttk.Combobox):
                    jv = tk.StringVar(value="AND")
                    jw = ttk.Combobox(rf, values=["AND","OR"], textvariable=jv, state="readonly", width=6)
                    jw.bind("<<ComboboxSelected>>", lambda e: self._changed())
                    r["join_widget"] = jw
                    r["join_var"] = jv

            r["join_widget"].grid(row=0, column=0, padx=4, sticky="w")
            r["col_cb"].grid(row=0, column=1, padx=4, sticky="w")
            r["ord_cb"].grid(row=0, column=2, padx=4, sticky="w")
            # restore/remove button
            btn = ttk.Button(rf, text="✖", width=3, command=lambda fr=rf: self._remove_row(fr))
            btn.grid(row=0, column=3, padx=4)
            self.rows.append(r)

    def _changed(self):
        if callable(self.on_change):
            self.on_change()

    # ---------- logic ----------
    def apply_sorts(self, df: pd.DataFrame) -> pd.DataFrame:
        if df is None or df.empty:
            return df

        # Build groups split by OR
        groups = []  # list of [(col, asc), ...]
        current = []
        alive = [r for r in self.rows if r is not None]
        for i, r in enumerate(alive):
            col = r["col_var"].get()
            if not col or col not in df.columns:
                continue
            asc = (r["order_var"].get() != "Descending")
            if i == 0:
                current.append((col, asc))
            else:
                join = r["join_var"].get() if r["join_var"] is not None else "AND"
                if join == "OR":
                    # close previous group, start new
                    if current:
                        groups.append(current)
                    current = [(col, asc)]
                else:
                    current.append((col, asc))
        if current:
            groups.append(current)

        if not groups:
            return df

        # Apply stable sorts: last group first → first group last (highest precedence)
        out = df
        for group in reversed(groups):
            cols = [c for c,_ in group]
            asc = [a for _,a in group]
            try:
                out = out.sort_values(by=cols, ascending=asc, na_position="last", kind="mergesort")
            except Exception:
                # fallback: ignore kind param
                out = out.sort_values(by=cols, ascending=asc, na_position="last")
        return out

    # ---------- config ----------
    def get_config(self):
        cfg = []
        alive = [r for r in self.rows if r is not None]
        for i, r in enumerate(alive):
            join = (r["join_var"].get() if r["join_var"] is not None else "")
            cfg.append((join, r["col_var"].get(), r["order_var"].get()))
        return {"sorts": cfg}

    def load_config(self, cfg: dict):
        self.reset()
        data = cfg.get("sorts", [])
        for tup in data:
            # backward compatibility: if old tuple (col, order), add default join
            if len(tup) == 2:
                j, c, o = ("AND", tup[0], tup[1])
            else:
                j, c, o = tup
            self.add_row(preset=(j, c, o))

    def reset(self):
        for r in list(self.rows):
            if r and r["frame"]:
                try:
                    r["frame"].destroy()
                except Exception:
                    pass
        self.rows = []
        # add a single empty row by default
        self.add_row()

    def refresh_columns(self, df: pd.DataFrame):
        self.df = df
        cols = self._columns_list()
        for r in [x for x in self.rows if x is not None]:
            try:
                r["col_cb"]["values"] = cols
            except Exception:
                pass
