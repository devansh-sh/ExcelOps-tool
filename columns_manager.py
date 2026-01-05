# columns_manager.py
import tkinter as tk
from tkinter import ttk
import pandas as pd
from typing import Optional, Dict, Any, List


class ColumnsManagerFrame(ttk.Frame):
    """
    Columns Manager:
      - Reorder columns
      - Hide / show columns (with visual indication)
      - Remove duplicates by column
    """

    def __init__(self, parent, on_change_callback=None, df: Optional[pd.DataFrame] = None):
        super().__init__(parent)
        self.on_change = on_change_callback
        self.df = df

        self.column_order: List[str] = list(df.columns) if df is not None else []
        self.column_visible: Dict[str, bool] = {c: True for c in self.column_order}

        self.remove_duplicates_var = tk.BooleanVar(value=False)
        self.duplicate_column_var = tk.StringVar()

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
        ttk.Button(btns, text="Show All", command=self._show_all).pack(side="left", padx=4)
        ttk.Button(btns, text="Apply Changes", command=self._apply).pack(side="right", padx=4)

        # ---------- Duplicates ----------
        dup = ttk.LabelFrame(self, text="Duplicates")
        dup.pack(fill="x", padx=6, pady=(0, 6))

        ttk.Checkbutton(
            dup,
            text="Remove duplicate rows based on column",
            variable=self.remove_duplicates_var,
            command=self._changed
        ).pack(anchor="w", padx=4, pady=2)

        self.dup_col_cb = ttk.Combobox(
            dup,
            textvariable=self.duplicate_column_var,
            state="readonly",
            values=self.column_order
        )
        self.dup_col_cb.pack(fill="x", padx=4, pady=2)
        self.dup_col_cb.bind("<<ComboboxSelected>>", lambda e: self._changed())

        self._refresh_listbox()

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

    def _apply(self):
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

        # Apply visibility + order
        visible_cols = [
            c for c in self.column_order
            if self.column_visible.get(c, True) and c in df.columns
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
            "dedupe": {
                "enabled": self.remove_duplicates_var.get(),
                "column": self.duplicate_column_var.get(),
            },
        }

    def load_config(self, cfg: Dict[str, Any]):
        self.column_order = cfg.get("order", self.column_order)
        self.column_visible = cfg.get("visible", self.column_visible)

        dedupe = cfg.get("dedupe", {})
        self.remove_duplicates_var.set(dedupe.get("enabled", False))
        self.duplicate_column_var.set(dedupe.get("column", ""))

        self._refresh_listbox()
        self._changed()

    def reset(self):
        if self.df is None:
            return
        self.column_order = list(self.df.columns)
        self.column_visible = {c: True for c in self.column_order}
        self.remove_duplicates_var.set(False)
        self.duplicate_column_var.set("")
        self._refresh_listbox()
        self._changed()

    def refresh_source_df(self, df: pd.DataFrame | None):
        self.df = df
        if df is None:
            return
        current_cols = list(df.columns)
        for c in current_cols:
            if c not in self.column_order:
                self.column_order.append(c)
                self.column_visible[c] = True
        self.column_order = [c for c in self.column_order if c in current_cols]
        self.column_visible = {c: self.column_visible.get(c, True) for c in self.column_order}
        self.dup_col_cb["values"] = self.column_order
        if self.duplicate_column_var.get() not in self.column_order:
            self.duplicate_column_var.set("")
        self._refresh_listbox()

    # ------------------------------------------------------------------
    def _changed(self):
        if callable(self.on_change):
            try:
                self.on_change()
            except Exception:
                pass
