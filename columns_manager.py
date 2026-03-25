# columns_manager.py
import tkinter as tk
from tkinter import ttk, messagebox
import pandas as pd
from typing import Optional, Dict, Any, List


class ColumnsManagerFrame(ttk.Frame):
    """
    Columns Manager:
      - Reorder columns
      - Hide / show columns (with visual indication)
      - Remove duplicates by column
      - Add calculated columns using formula expressions
    """

    def __init__(self, parent, on_change_callback=None, df: Optional[pd.DataFrame] = None):
        super().__init__(parent)
        self.on_change = on_change_callback
        self.df = df

        self.column_order: List[str] = list(df.columns) if df is not None else []
        self.column_visible: Dict[str, bool] = {c: True for c in self.column_order}

        self.remove_duplicates_var = tk.BooleanVar(value=False)
        self.duplicate_column_var = tk.StringVar()

        self.formulas: List[Dict[str, str]] = []
        self.formula_name_var = tk.StringVar()
        self.formula_expr_var = tk.StringVar()
        self.formula_column_picker_var = tk.StringVar()

        self._build_ui()

    # ------------------------------------------------------------------
    # UI
    # ------------------------------------------------------------------
    def _build_ui(self):
        # ---------- Edit Columns ----------
        cols_frame = ttk.LabelFrame(self, text="Edit Columns")
        cols_frame.pack(fill="both", expand=True, padx=6, pady=6)

        self.listbox = tk.Listbox(cols_frame, exportselection=False)
        self.listbox.pack(fill="both", expand=True, padx=4, pady=4)

        btns = ttk.Frame(cols_frame)
        btns.pack(fill="x", pady=4)

        ttk.Button(btns, text="Move Up", command=self._move_up).pack(side="left", padx=4)
        ttk.Button(btns, text="Move Down", command=self._move_down).pack(side="left", padx=4)
        ttk.Button(btns, text="Hide Column", command=self._toggle_visibility).pack(side="left", padx=4)
        ttk.Button(btns, text="Select All", command=self._show_all).pack(side="left", padx=4)
        ttk.Button(btns, text="Unselect All", command=self._hide_all).pack(side="left", padx=4)
        ttk.Button(btns, text="Apply Changes", command=self._apply).pack(side="right", padx=4)

        # ---------- Calculated Columns ----------
        calc = ttk.LabelFrame(self, text="Calculated Columns")
        calc.pack(fill="both", padx=6, pady=(0, 6))

        ttk.Label(
            calc,
            text="Formula examples: Buyqty * BuyAvg_net_rate, (Buyqty - Sellqty), BuyNetvalue + SellNetvalue",
        ).pack(anchor="w", padx=4, pady=(4, 2))

        form_row = ttk.Frame(calc)
        form_row.pack(fill="x", padx=4, pady=2)
        ttk.Label(form_row, text="Column Name").pack(side="left")
        ttk.Entry(form_row, textvariable=self.formula_name_var, width=24).pack(side="left", padx=(6, 16))
        ttk.Label(form_row, text="Formula").pack(side="left")
        ttk.Entry(form_row, textvariable=self.formula_expr_var, width=52).pack(side="left", padx=(6, 6), fill="x", expand=True)

        picker_row = ttk.Frame(calc)
        picker_row.pack(fill="x", padx=4, pady=(0, 2))
        ttk.Label(picker_row, text="Insert Column in Formula").pack(side="left")
        self.formula_col_cb = ttk.Combobox(
            picker_row,
            textvariable=self.formula_column_picker_var,
            state="readonly",
            values=self.column_order,
            width=28,
        )
        self.formula_col_cb.pack(side="left", padx=(6, 6))
        ttk.Button(picker_row, text="Insert", command=self._insert_selected_column_into_formula).pack(side="left")

        calc_btns = ttk.Frame(calc)
        calc_btns.pack(fill="x", padx=4, pady=2)
        ttk.Button(calc_btns, text="Add / Update Formula", command=self._upsert_formula).pack(side="left", padx=4)
        ttk.Button(calc_btns, text="Delete Formula", command=self._delete_formula).pack(side="left", padx=4)

        self.formula_list = tk.Listbox(calc, height=4, exportselection=False)
        self.formula_list.pack(fill="x", padx=4, pady=(2, 4))
        self.formula_list.bind("<<ListboxSelect>>", lambda e: self._on_formula_selected())

        # ---------- Duplicates ----------
        dup = ttk.LabelFrame(self, text="Duplicates")
        dup.pack(fill="x", padx=6, pady=(0, 6))

        ttk.Checkbutton(
            dup,
            text="Remove duplicate rows based on column",
            variable=self.remove_duplicates_var,
            command=self._changed,
        ).pack(anchor="w", padx=4, pady=2)

        self.dup_col_cb = ttk.Combobox(
            dup,
            textvariable=self.duplicate_column_var,
            state="readonly",
            values=self.column_order,
        )
        self.dup_col_cb.pack(fill="x", padx=4, pady=2)
        self.dup_col_cb.bind("<<ComboboxSelected>>", lambda e: self._changed())

        self.formula_col_cb["values"] = self.column_order
        self._refresh_listbox()
        self._refresh_formula_listbox()

    # ------------------------------------------------------------------
    # Listbox helpers
    # ------------------------------------------------------------------
    def _refresh_listbox(self, select_idx: int | None = None):
        self.listbox.delete(0, "end")
        for c in self.column_order:
            label = c if self.column_visible.get(c, True) else f"[HIDDEN] {c}"
            self.listbox.insert("end", label)

        if select_idx is not None:
            try:
                self.listbox.selection_set(select_idx)
            except Exception:
                pass

    def _refresh_formula_listbox(self):
        self.formula_list.delete(0, "end")
        for f in self.formulas:
            self.formula_list.insert("end", f"{f.get('name', '')} = {f.get('expr', '')}")

    def _selected_index(self) -> Optional[int]:
        sel = self.listbox.curselection()
        return sel[0] if sel else None

    # ------------------------------------------------------------------
    # Column actions
    # ------------------------------------------------------------------
    def _move_up(self):
        idx = self._selected_index()
        if idx is None or idx == 0:
            return
        self.column_order[idx - 1], self.column_order[idx] = (
            self.column_order[idx],
            self.column_order[idx - 1],
        )
        self._refresh_listbox(idx - 1)
        self._changed()

    def _move_down(self):
        idx = self._selected_index()
        if idx is None or idx >= len(self.column_order) - 1:
            return
        self.column_order[idx + 1], self.column_order[idx] = (
            self.column_order[idx],
            self.column_order[idx + 1],
        )
        self._refresh_listbox(idx + 1)
        self._changed()

    def _toggle_visibility(self):
        idx = self._selected_index()
        if idx is None:
            return
        col = self.column_order[idx]
        self.column_visible[col] = not self.column_visible.get(col, True)
        self._refresh_listbox(idx)
        self._changed()

    def _show_all(self):
        for c in self.column_visible:
            self.column_visible[c] = True
        self._refresh_listbox()
        self._changed()

    def _hide_all(self):
        for c in self.column_visible:
            self.column_visible[c] = False
        self._refresh_listbox()
        self._changed()

    def _apply(self):
        self._changed()

    # ------------------------------------------------------------------
    # Formula actions
    # ------------------------------------------------------------------
    def _on_formula_selected(self):
        sel = self.formula_list.curselection()
        if not sel:
            return
        i = sel[0]
        f = self.formulas[i]
        self.formula_name_var.set(f.get("name", ""))
        self.formula_expr_var.set(f.get("expr", ""))

    def _insert_selected_column_into_formula(self):
        col = self.formula_column_picker_var.get().strip()
        if not col:
            return
        current = self.formula_expr_var.get()
        col_token = f"`{col}`" if " " in col else col
        new_expr = f"{current} {col_token}".strip() if current else col_token
        self.formula_expr_var.set(new_expr)

    def _upsert_formula(self):
        name = self.formula_name_var.get().strip()
        expr = self.formula_expr_var.get().strip()
        if not name or not expr:
            messagebox.showwarning("Formula", "Both Column Name and Formula are required.")
            return

        existing = next((i for i, f in enumerate(self.formulas) if f.get("name") == name), None)
        entry = {"name": name, "expr": expr}
        if existing is None:
            self.formulas.append(entry)
            if name not in self.column_order:
                self.column_order.append(name)
                self.column_visible[name] = True
        else:
            self.formulas[existing] = entry

        self._refresh_formula_listbox()
        self._refresh_listbox()
        self._changed()

    def _delete_formula(self):
        sel = self.formula_list.curselection()
        if not sel:
            return
        i = sel[0]
        name = self.formulas[i].get("name", "")
        del self.formulas[i]
        self._refresh_formula_listbox()

        if name in self.column_order and name not in (list(self.df.columns) if self.df is not None else []):
            self.column_order = [c for c in self.column_order if c != name]
            self.column_visible.pop(name, None)
        self._refresh_listbox()
        self._changed()

    # ------------------------------------------------------------------
    # Core logic used by main.py
    # ------------------------------------------------------------------
    def apply_columns(self, df: pd.DataFrame) -> pd.DataFrame:
        if df is None or df.empty:
            return df

        # Remove duplicates first
        if self.remove_duplicates_var.get():
            col = self.duplicate_column_var.get()
            if col and col in df.columns:
                df = df.drop_duplicates(subset=[col], keep="first")

        # Apply formulas
        for f in self.formulas:
            name = f.get("name", "").strip()
            expr = f.get("expr", "").strip()
            if not name or not expr:
                continue
            try:
                df[name] = df.eval(expr, engine="python")
            except Exception:
                continue

        # Apply visibility + order
        visible_cols = [
            c for c in self.column_order if self.column_visible.get(c, True) and c in df.columns
        ]

        if visible_cols:
            df = df[visible_cols]

        return df

    # ------------------------------------------------------------------
    # Presets
    # ------------------------------------------------------------------
    def get_config(self) -> Dict[str, Any]:
        return {
            "order": self.column_order,
            "visible": self.column_visible,
            "formulas": self.formulas,
            "dedupe": {
                "enabled": self.remove_duplicates_var.get(),
                "column": self.duplicate_column_var.get(),
            },
        }

    def load_config(self, cfg: Dict[str, Any]):
        self.column_order = cfg.get("order", self.column_order)
        self.column_visible = cfg.get("visible", self.column_visible)
        self.formulas = [
            {"name": (f.get("name", "") if isinstance(f, dict) else ""), "expr": (f.get("expr", "") if isinstance(f, dict) else "")}
            for f in cfg.get("formulas", [])
            if isinstance(f, dict)
        ]

        dedupe = cfg.get("dedupe", {})
        self.remove_duplicates_var.set(dedupe.get("enabled", False))
        self.duplicate_column_var.set(dedupe.get("column", ""))

        for f in self.formulas:
            n = f.get("name", "")
            if n and n not in self.column_order:
                self.column_order.append(n)
                self.column_visible[n] = True

        self._refresh_formula_listbox()
        self._refresh_listbox()
        self._changed()

    def reset(self):
        if self.df is None:
            return
        self.column_order = list(self.df.columns)
        self.column_visible = {c: True for c in self.column_order}
        self.formulas = []
        self.formula_name_var.set("")
        self.formula_expr_var.set("")
        self.formula_column_picker_var.set("")
        self.remove_duplicates_var.set(False)
        self.duplicate_column_var.set("")
        self._refresh_formula_listbox()
        self._refresh_listbox()
        self._changed()

    def refresh_source_df(self, df: pd.DataFrame | None):
        self.df = df
        if df is None:
            return

        current_cols = list(df.columns)
        formula_cols = [f.get("name", "") for f in self.formulas if f.get("name", "")]

        for c in current_cols:
            if c not in self.column_order:
                self.column_order.append(c)
                self.column_visible[c] = True

        self.column_order = [c for c in self.column_order if c in current_cols or c in formula_cols]
        self.column_visible = {c: self.column_visible.get(c, True) for c in self.column_order}
        self.dup_col_cb["values"] = self.column_order
        self.formula_col_cb["values"] = current_cols
        if self.duplicate_column_var.get() not in self.column_order:
            self.duplicate_column_var.set("")

        self._refresh_formula_listbox()
        self._refresh_listbox()

    # ------------------------------------------------------------------
    def _changed(self):
        if callable(self.on_change):
            try:
                self.on_change()
            except Exception:
                pass
