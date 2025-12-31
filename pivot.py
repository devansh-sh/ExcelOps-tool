# pivot.py
import tkinter as tk
from tkinter import ttk, messagebox
import pandas as pd


class PivotFrame(ttk.Frame):
    """
    Simple, Excel-like Pivot Table UI.

    Exposed API (used by main.py & presets):
      - preview_pivot(df) -> DataFrame | None
      - apply_pivot_if_requested(df) -> DataFrame
      - get_config() / load_config(cfg)
      - reset()
    """

    def __init__(self, parent, on_preview_callback, df: pd.DataFrame | None, data_provider=None):
        super().__init__(parent)
        self.on_preview = on_preview_callback
        self.source_df = df
        self.data_provider = data_provider

        self.value_col = tk.StringVar()
        self.generated = False  # pivot locked for export
        self._build_ui()

    # ---------------- UI ----------------
    def _build_ui(self):
        cols = [] if self.source_df is None else list(self.source_df.columns)

        pad = {"padx": 6, "pady": 4}

        ttk.Label(
            self,
            text=(
                "How to use: 1) Pick Rows. 2) (Optional) Pick Columns. "
                "3) Pick one or more Values. 4) Choose Aggregation. "
                "5) Preview or Generate."
            ),
            wraplength=760,
            justify="left"
        ).pack(anchor="w", **pad)

        # ---- Rows ----
        ttk.Label(self, text="Rows").pack(anchor="w", **pad)
        self.rows_lb = tk.Listbox(self, selectmode="multiple", height=5)
        self.rows_lb.pack(fill="x", **pad)

        # ---- Columns ----
        ttk.Label(self, text="Columns (optional)").pack(anchor="w", **pad)
        self.cols_lb = tk.Listbox(self, selectmode="multiple", height=4)
        self.cols_lb.pack(fill="x", **pad)

        # ---- Values ----
        val_frame = ttk.Frame(self)
        val_frame.pack(fill="x", **pad)
        val_frame.columnconfigure(0, weight=1)
        val_frame.columnconfigure(1, weight=0)

        ttk.Label(val_frame, text="Values").grid(row=0, column=0, sticky="w")
        ttk.Label(val_frame, text="Aggregation").grid(row=0, column=1, sticky="w")

        self.values_lb = tk.Listbox(val_frame, selectmode="multiple", height=5, exportselection=False)
        self.values_lb.grid(row=1, column=0, padx=(0, 12), sticky="ew")

        self.agg_var = tk.StringVar(value="sum")
        self.values_lb.grid(row=1, column=0, padx=(0, 8), sticky="ew")

        self.agg_var = tk.StringVar(value="sum")
        self.value_cb = ttk.Combobox(
            val_frame,
            textvariable=self.value_col,
            values=cols,
            state="readonly",
            width=30
        )
        self.value_cb.grid(row=1, column=0, padx=(0, 8))

        ttk.Combobox(
            val_frame,
            textvariable=self.agg_var,
            values=["sum", "mean", "count", "min", "max"],
            state="readonly",
            width=12
        ).grid(row=1, column=1, sticky="w")

        # ---- Buttons ----
        btns = ttk.Frame(self)
        btns.pack(fill="x", pady=8)

        ttk.Button(btns, text="🔍 Preview Pivot", command=self._preview).pack(side="left", padx=4)
        ttk.Button(btns, text="✅ Generate Pivot Table", command=self._generate).pack(side="left", padx=4)
        ttk.Button(btns, text="❌ Clear Pivot", command=self.reset).pack(side="right", padx=4)

        # populate listboxes
        self._refresh_columns()

    # ---------------- helpers ----------------
    def _refresh_columns(self, keep_selection: bool = False):
        cols = [] if self.source_df is None else list(self.source_df.columns)
        selected_rows = self._selected(self.rows_lb) if keep_selection else []
        selected_cols = self._selected(self.cols_lb) if keep_selection else []
        selected_vals = self._selected(self.values_lb) if keep_selection else []
        selected_val = self.value_col.get() if keep_selection else ""
        self.rows_lb.delete(0, "end")
        self.cols_lb.delete(0, "end")
        self.values_lb.delete(0, "end")
        for c in cols:
            self.rows_lb.insert("end", c)
            self.cols_lb.insert("end", c)
            self.values_lb.insert("end", c)
        if keep_selection:
            for i, c in enumerate(cols):
                if c in selected_rows:
                    self.rows_lb.selection_set(i)
                if c in selected_cols:
                    self.cols_lb.selection_set(i)
                if c in selected_vals:
                    self.values_lb.selection_set(i)
            if selected_val in cols:
                self.value_col.set(selected_val)
            else:
                self.value_col.set("")
        if self.value_cb is not None:
            self.value_cb["values"] = cols

    def refresh_source_df(self, df: pd.DataFrame | None):
        self.source_df = df
        self._refresh_columns(keep_selection=True)

    def _selected(self, lb):
        return [lb.get(i) for i in lb.curselection()]

    # ---------------- logic ----------------
    def _build_pivot(self, df: pd.DataFrame) -> pd.DataFrame | None:
        rows = self._selected(self.rows_lb)
        cols = self._selected(self.cols_lb)
        vals = self._selected(self.values_lb)
        agg = self.agg_var.get()

        if not rows:
            return None

        try:
            if vals:
                pt = pd.pivot_table(
                    df,
                    index=rows,
                    columns=cols if cols else None,
                    values=vals,
                    aggfunc=agg,
                    fill_value=0
                )
            else:
                pt = pd.pivot_table(
                    df,
                    index=rows,
                    columns=cols if cols else None,
                    aggfunc="size",
                    fill_value=0
                )
                pt = pt.rename("Count")
            if isinstance(pt.columns, pd.MultiIndex):
                pt.columns = [" | ".join([str(c) for c in col if c != ""]) for col in pt.columns.values]
            elif isinstance(pt.columns, pd.Index):
                pt.columns = [str(c) for c in pt.columns]
            return pt.reset_index()
        except Exception as e:
            messagebox.showerror("Pivot error", str(e))
            return None

    def _preview(self):
        df = self._get_preview_df()
        if df is None:
            return
        df = self._build_pivot(df)
        if df is None or df.empty:
            messagebox.showinfo("Pivot", "No pivot result with current selection.")
            return
        self.on_preview(df)

    def _generate(self):
        df = self._get_preview_df()
        if df is None:
            return
        df = self._build_pivot(df)
        if df is None:
            return
        self.generated = True
        self.on_preview(df)

    def _get_preview_df(self) -> pd.DataFrame | None:
        if callable(self.data_provider):
            return self.data_provider()
        return self.source_df

    # ---------------- API used by main.py ----------------
    def apply_pivot_if_requested(self, df: pd.DataFrame) -> pd.DataFrame:
        if not self.generated:
            return df
        out = self._build_pivot(df)
        return out if out is not None else df

    def reset(self):
        self.generated = False
        self.rows_lb.selection_clear(0, "end")
        self.cols_lb.selection_clear(0, "end")
        self.values_lb.selection_clear(0, "end")
        self.value_col.set("")
        self.agg_var.set("sum")

    # ---------------- presets ----------------
    def get_config(self):
        return {
            "rows": self._selected(self.rows_lb),
            "columns": self._selected(self.cols_lb),
            "values": self._selected(self.values_lb),
            "agg": self.agg_var.get(),
            "generated": self.generated
        }

    def load_config(self, cfg: dict):
        self.reset()
        cols = [] if self.source_df is None else list(self.source_df.columns)

        for i, c in enumerate(cols):
            if c in cfg.get("rows", []):
                self.rows_lb.selection_set(i)
            if c in cfg.get("columns", []):
                self.cols_lb.selection_set(i)
            if c in cfg.get("values", []):
                self.values_lb.selection_set(i)
            if "values" not in cfg and c == cfg.get("value", ""):
                self.values_lb.selection_set(i)

        self.agg_var.set(cfg.get("agg", "sum"))
        self.generated = cfg.get("generated", False)
