# columns_manager.py
import tkinter as tk
from tkinter import ttk, messagebox
import pandas as pd
import re
from typing import Optional, Dict, Any, List


CALC_FUNCTIONS = [
    "ADD",
    "SUBTRACT",
    "MULTIPLY",
    "DIVIDE",
    "PERCENT",
    "ABS",
    "ROUND",
    "MIN",
    "MAX",
    "COALESCE",
    "IF",
    "AND",
    "OR",
]

BUILDER_FUNCTIONS = [
    "ADD", "SUBTRACT", "MULTIPLY", "DIVIDE", "PERCENT",
    "ABS", "ROUND", "MIN", "MAX", "COALESCE",
]


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

        self.calc_fn_var = tk.StringVar(value=BUILDER_FUNCTIONS[0])
        self.calc_col1_var = tk.StringVar()
        self.calc_col2_var = tk.StringVar()
        self.calc_constant_var = tk.StringVar()
        self.calc_decimals_var = tk.StringVar(value="2")

        self._build_ui()

    # ------------------------------------------------------------------
    # UI
    # ------------------------------------------------------------------
    def _build_ui(self):
        cols_frame = ttk.LabelFrame(self, text="Column Editor")
        cols_frame.pack(fill="both", expand=True, padx=6, pady=6)

        self.listbox = tk.Listbox(cols_frame, exportselection=False)
        self.listbox.pack(fill="both", expand=True, padx=4, pady=4)

        btns = ttk.Frame(cols_frame)
        btns.pack(fill="x", pady=4)

        ttk.Button(btns, text="↑ Move Up", command=self._move_up).pack(side="left", padx=4)
        ttk.Button(btns, text="↓ Move Down", command=self._move_down).pack(side="left", padx=4)
        ttk.Button(btns, text="👁 Show/Hide", command=self._toggle_visibility).pack(side="left", padx=4)
        ttk.Button(btns, text="✓ Select All", command=self._show_all).pack(side="left", padx=4)
        ttk.Button(btns, text="○ Unselect All", command=self._hide_all).pack(side="left", padx=4)
        ttk.Button(btns, text="Apply Changes", command=self._apply, style="Accent.TButton").pack(side="right", padx=4)

        dup = ttk.LabelFrame(self, text="Remove Duplicate Rows")
        dup.pack(fill="x", padx=6, pady=(0, 6))

        dup_row = ttk.Frame(dup)
        dup_row.pack(fill="x", padx=4, pady=4)
        ttk.Checkbutton(
            dup_row,
            text="Enable duplicate removal",
            variable=self.remove_duplicates_var,
            command=self._changed,
        ).pack(side="left")
        ttk.Label(dup_row, text="Column:").pack(side="left", padx=(16, 4))
        self.dup_col_cb = ttk.Combobox(
            dup_row,
            textvariable=self.duplicate_column_var,
            state="readonly",
            values=self.column_order,
            width=32,
        )
        self.dup_col_cb.pack(side="left", padx=(0, 6))
        self.dup_col_cb.bind("<<ComboboxSelected>>", lambda e: self._on_duplicate_column_selected())
        ttk.Label(
            dup,
            text="Select a column to keep the first row for each repeated value.",
            foreground="#4b5563",
        ).pack(anchor="w", padx=4, pady=(0, 4))

        calc = ttk.LabelFrame(self, text="Calculated Columns / Formula Builder")
        calc.pack(fill="both", padx=6, pady=(0, 6))

        ttk.Label(
            calc,
            text=(
                "Supported functions: ADD, SUBTRACT, MULTIPLY, DIVIDE, PERCENT, "
                "ABS, ROUND, MIN, MAX, COALESCE, IF, AND, OR"
            ),
        ).pack(anchor="w", padx=4, pady=(4, 2))

        ttk.Label(
            calc,
            text=(
                "Type formulas directly using column names. Examples: "
                "IF(A21 > B32, A21 * B3, 0)  •  "
                "IF(AND(Revenue >= 2500000, SellQty == 0), Revenue, 0)"
            ),
            foreground="#4b5563",
            wraplength=1050,
            justify="left",
        ).pack(anchor="w", padx=4, pady=(0, 4))

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

        builder = ttk.LabelFrame(calc, text="Easy Formula Builder")
        builder.pack(fill="x", padx=4, pady=(2, 6))

        row1 = ttk.Frame(builder)
        row1.pack(fill="x", padx=4, pady=2)
        ttk.Label(row1, text="Function").pack(side="left")
        self.fn_cb = ttk.Combobox(
            row1,
            textvariable=self.calc_fn_var,
            values=BUILDER_FUNCTIONS,
            state="readonly",
            width=14,
        )
        self.fn_cb.pack(side="left", padx=(6, 12))
        ttk.Label(row1, text="Column 1").pack(side="left")
        self.calc_col1_cb = ttk.Combobox(row1, textvariable=self.calc_col1_var, state="readonly", width=22)
        self.calc_col1_cb.pack(side="left", padx=(6, 12))
        ttk.Label(row1, text="Column 2").pack(side="left")
        self.calc_col2_cb = ttk.Combobox(row1, textvariable=self.calc_col2_var, state="readonly", width=22)
        self.calc_col2_cb.pack(side="left", padx=(6, 0))

        row2 = ttk.Frame(builder)
        row2.pack(fill="x", padx=4, pady=2)
        ttk.Label(row2, text="Constant (optional)").pack(side="left")
        ttk.Entry(row2, textvariable=self.calc_constant_var, width=16).pack(side="left", padx=(6, 12))
        ttk.Label(row2, text="Decimals (ROUND)").pack(side="left")
        ttk.Entry(row2, textvariable=self.calc_decimals_var, width=8).pack(side="left", padx=(6, 12))
        ttk.Button(row2, text="Build Formula", command=self._build_formula_from_builder).pack(side="left")

        calc_btns = ttk.Frame(calc)
        calc_btns.pack(fill="x", padx=4, pady=2)
        ttk.Button(calc_btns, text="Add / Update Formula", command=self._upsert_formula).pack(side="left", padx=4)
        ttk.Button(calc_btns, text="Validate Formula", command=self._validate_current_formula).pack(side="left", padx=4)
        ttk.Button(calc_btns, text="Delete Formula", command=self._delete_formula).pack(side="left", padx=4)

        self.formula_list = tk.Listbox(calc, height=5, exportselection=False)
        self.formula_list.pack(fill="x", padx=4, pady=(2, 4))
        self.formula_list.bind("<<ListboxSelect>>", lambda e: self._on_formula_selected())

        self._refresh_formula_column_choices(self.column_order)
        self._refresh_listbox()
        self._refresh_formula_listbox()

    # ------------------------------------------------------------------
    # Helpers
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

    def _refresh_formula_column_choices(self, columns: List[str]):
        self.formula_col_cb["values"] = columns
        self.calc_col1_cb["values"] = columns
        self.calc_col2_cb["values"] = columns

    def _selected_index(self) -> Optional[int]:
        sel = self.listbox.curselection()
        return sel[0] if sel else None

    def _on_duplicate_column_selected(self):
        if self.duplicate_column_var.get():
            self.remove_duplicates_var.set(True)
        self._changed()

    @staticmethod
    def _series_or_number(df: pd.DataFrame, token: str):
        raw = token.strip()
        if raw.startswith("`") and raw.endswith("`"):
            col_name = raw[1:-1]
            if col_name in df.columns:
                return df[col_name]
        if raw in df.columns:
            return df[raw]
        return pd.to_numeric(pd.Series([raw] * len(df)), errors="coerce")

    @staticmethod
    def _split_formula_args(text: str) -> List[str]:
        """Split function arguments while preserving nested calls and strings."""
        args, current = [], []
        depth = 0
        quote = None
        for char in text:
            if quote:
                current.append(char)
                if char == quote:
                    quote = None
                continue
            if char in {"'", '"'}:
                quote = char
                current.append(char)
            elif char == "(":
                depth += 1
                current.append(char)
            elif char == ")":
                depth -= 1
                if depth < 0:
                    raise ValueError("Formula contains an unmatched closing parenthesis.")
                current.append(char)
            elif char == "," and depth == 0:
                args.append("".join(current).strip())
                current = []
            else:
                current.append(char)
        if quote:
            raise ValueError("Formula contains an unterminated quoted value.")
        if depth != 0:
            raise ValueError("Formula contains unmatched parentheses.")
        args.append("".join(current).strip())
        return args

    @staticmethod
    def _outer_function(expr: str):
        match = re.match(r"^([A-Za-z_][A-Za-z0-9_]*)\s*\(", expr)
        if not match or not expr.endswith(")"):
            return None, None
        start = match.end() - 1
        depth = 0
        quote = None
        for index in range(start, len(expr)):
            char = expr[index]
            if quote:
                if char == quote:
                    quote = None
                continue
            if char in {"'", '"'}:
                quote = char
            elif char == "(":
                depth += 1
            elif char == ")":
                depth -= 1
                if depth == 0 and index != len(expr) - 1:
                    return None, None
        if depth != 0 or quote:
            raise ValueError("Formula contains unmatched parentheses or quotes.")
        return match.group(1).upper(), expr[start + 1:-1]

    @staticmethod
    def _normalise_excel_expression(expr: str) -> str:
        expr = expr.replace("<>", "!=")
        expr = re.sub(r"(?<![<>=!])=(?!=)", "==", expr)
        expr = re.sub(r"\bAND\b", "and", expr, flags=re.IGNORECASE)
        expr = re.sub(r"\bOR\b", "or", expr, flags=re.IGNORECASE)
        return expr

    @staticmethod
    def _normalise_formula_columns(df: pd.DataFrame, expr: str) -> str:
        """Resolve typed column names case-insensitively without changing strings."""
        column_map = {str(column).casefold(): str(column) for column in df.columns}

        def replace_segment(segment: str) -> str:
            def replace_identifier(match):
                typed = match.group(0)
                actual = column_map.get(typed.casefold())
                return f"`{actual}`" if actual is not None else typed

            return re.sub(r"\b[A-Za-z_][A-Za-z0-9_]*\b", replace_identifier, segment)

        protected_parts = re.split(r'(`[^`]+`|"(?:[^"\\]|\\.)*"|\'(?:[^\'\\]|\\.)*\')', expr)
        normalised_parts = []
        for part in protected_parts:
            if part.startswith("`") and part.endswith("`"):
                typed = part[1:-1]
                normalised_parts.append(f"`{column_map.get(typed.casefold(), typed)}`")
            elif part[:1] in {"'", '"'}:
                normalised_parts.append(part)
            else:
                normalised_parts.append(replace_segment(part))
        return "".join(normalised_parts)

    @staticmethod
    def _as_series(df: pd.DataFrame, value):
        if isinstance(value, pd.Series):
            return value.reindex(df.index)
        return pd.Series([value] * len(df), index=df.index)

    def _evaluate_basic_expression(self, df: pd.DataFrame, expr: str):
        raw = expr.strip()
        if not raw:
            raise ValueError("Formula expression cannot be empty.")
        if len(raw) >= 2 and raw[0] == raw[-1] and raw[0] in {"'", '"'}:
            return raw[1:-1]
        if raw.upper() == "TRUE":
            return True
        if raw.upper() == "FALSE":
            return False
        try:
            return float(raw) if "." in raw else int(raw)
        except ValueError:
            pass
        if raw.startswith("`") and raw.endswith("`") and raw[1:-1] in df.columns:
            return df[raw[1:-1]]
        if raw in df.columns:
            return df[raw]
        normalised = self._normalise_excel_expression(raw)
        normalised = self._normalise_formula_columns(df, normalised)
        return df.eval(normalised, engine="python")

    def _as_boolean_series(self, df: pd.DataFrame, value):
        series = self._as_series(df, value)
        if not pd.api.types.is_bool_dtype(series):
            series = series.fillna(False).astype(bool)
        return series.fillna(False)

    def _evaluate_formula_expr(self, df: pd.DataFrame, expr: str):
        expr = expr.strip()
        if expr.startswith("="):
            expr = expr[1:].strip()
        natural_if = re.match(r"^IF\s+(.+)$", expr, flags=re.IGNORECASE)
        if natural_if and not re.match(r"^IF\s*\(", expr, flags=re.IGNORECASE):
            expr = f"IF({natural_if.group(1)})"
        fn, args_raw = self._outer_function(expr)
        if fn is None or fn not in CALC_FUNCTIONS:
            return self._evaluate_basic_expression(df, expr)
        args = self._split_formula_args(args_raw)

        if fn == "IF":
            if len(args) not in (2, 3):
                raise ValueError("IF requires IF(condition, value_if_true, value_if_false). The false value is optional.")
            condition = self._as_boolean_series(df, self._evaluate_formula_expr(df, args[0]))
            true_value = self._as_series(df, self._evaluate_formula_expr(df, args[1]))
            false_value = self._evaluate_formula_expr(df, args[2]) if len(args) == 3 else pd.NA
            return true_value.where(condition, false_value)
        if fn in {"AND", "OR"}:
            if len(args) < 2:
                raise ValueError(f"{fn} requires at least two conditions.")
            result = self._as_boolean_series(df, self._evaluate_formula_expr(df, args[0]))
            for arg in args[1:]:
                condition = self._as_boolean_series(df, self._evaluate_formula_expr(df, arg))
                result = result & condition if fn == "AND" else result | condition
            return result

        if fn == "ADD" and len(args) == 2:
            return self._series_or_number(df, args[0]) + self._series_or_number(df, args[1])
        if fn == "SUBTRACT" and len(args) == 2:
            return self._series_or_number(df, args[0]) - self._series_or_number(df, args[1])
        if fn == "MULTIPLY" and len(args) == 2:
            return self._series_or_number(df, args[0]) * self._series_or_number(df, args[1])
        if fn == "DIVIDE" and len(args) == 2:
            b = self._series_or_number(df, args[1]).replace(0, pd.NA)
            return self._series_or_number(df, args[0]) / b
        if fn == "PERCENT" and len(args) == 2:
            b = self._series_or_number(df, args[1]).replace(0, pd.NA)
            return (self._series_or_number(df, args[0]) / b) * 100
        if fn == "ABS" and len(args) == 1:
            return self._series_or_number(df, args[0]).abs()
        if fn == "ROUND" and len(args) >= 1:
            n = 0
            if len(args) > 1:
                try:
                    n = int(float(args[1]))
                except Exception:
                    n = 0
            return self._series_or_number(df, args[0]).round(n)
        if fn == "MIN" and len(args) == 2:
            a = self._series_or_number(df, args[0])
            b = self._series_or_number(df, args[1])
            return pd.concat([a, b], axis=1).min(axis=1)
        if fn == "MAX" and len(args) == 2:
            a = self._series_or_number(df, args[0])
            b = self._series_or_number(df, args[1])
            return pd.concat([a, b], axis=1).max(axis=1)
        if fn == "COALESCE" and len(args) == 2:
            a = self._series_or_number(df, args[0])
            b = self._series_or_number(df, args[1])
            return a.fillna(b)

        raise ValueError(f"Invalid arguments for {fn}.")

    # ------------------------------------------------------------------
    # Column actions
    # ------------------------------------------------------------------
    def _move_up(self):
        idx = self._selected_index()
        if idx is None or idx == 0:
            return
        self.column_order[idx - 1], self.column_order[idx] = self.column_order[idx], self.column_order[idx - 1]
        self._refresh_listbox(idx - 1)
        self._changed()

    def _move_down(self):
        idx = self._selected_index()
        if idx is None or idx >= len(self.column_order) - 1:
            return
        self.column_order[idx + 1], self.column_order[idx] = self.column_order[idx], self.column_order[idx + 1]
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
        self.formula_expr_var.set(f"{current} {col_token}".strip() if current else col_token)

    def _build_formula_from_builder(self):
        fn = self.calc_fn_var.get().strip().upper()
        c1 = self.calc_col1_var.get().strip()
        c2 = self.calc_col2_var.get().strip()
        const = self.calc_constant_var.get().strip()
        decimals = self.calc_decimals_var.get().strip() or "2"

        def tok(v: str):
            return f"`{v}`" if " " in v else v

        formula = ""
        if fn in ("ADD", "SUBTRACT", "MULTIPLY", "DIVIDE", "PERCENT", "MIN", "MAX", "COALESCE"):
            a = tok(c1) if c1 else ""
            b = tok(c2) if c2 else const
            if a and b:
                formula = f"{fn}({a}, {b})"
        elif fn == "ABS" and c1:
            formula = f"ABS({tok(c1)})"
        elif fn == "ROUND" and c1:
            formula = f"ROUND({tok(c1)}, {decimals})"

        if not formula:
            messagebox.showwarning("Formula Builder", "Please select required columns/values for the chosen function.")
            return
        self.formula_expr_var.set(formula)

    def _formula_validation_df(self, formula_name: str) -> pd.DataFrame | None:
        if self.df is None or self.df.empty:
            return None
        sample = self.df.head(100).copy()
        for formula in self.formulas:
            name = formula.get("name", "").strip()
            expr = formula.get("expr", "").strip()
            if not name or not expr or name == formula_name:
                continue
            sample[name] = self._evaluate_formula_expr(sample, expr)
        return sample

    def _validate_current_formula(self, show_success: bool = True) -> bool:
        name = self.formula_name_var.get().strip()
        expr = self.formula_expr_var.get().strip()
        if not name or not expr:
            messagebox.showwarning("Formula", "Both Column Name and Formula are required.")
            return False
        sample = self._formula_validation_df(name)
        if sample is None:
            messagebox.showwarning("Formula", "Load data before validating a calculated column.")
            return False
        try:
            result = self._evaluate_formula_expr(sample, expr)
            result = self._as_series(sample, result)
            if len(result) != len(sample):
                raise ValueError("Formula result does not match the number of input rows.")
        except Exception as exc:
            messagebox.showerror(
                "Invalid Formula",
                f"ExcelOps could not evaluate this formula:\n\n{exc}\n\n"
                "Use column names (backticks for names with spaces) and functions such as "
                "IF, AND, and OR.",
            )
            return False
        if show_success:
            messagebox.showinfo("Formula Valid", "The formula is valid and ready to save in this workflow.")
        return True

    def _upsert_formula(self):
        name = self.formula_name_var.get().strip()
        expr = self.formula_expr_var.get().strip()
        if not name or not expr:
            messagebox.showwarning("Formula", "Both Column Name and Formula are required.")
            return
        if not self._validate_current_formula(show_success=False):
            return

        idx = next((i for i, f in enumerate(self.formulas) if f.get("name") == name), None)
        entry = {"name": name, "expr": expr}
        if idx is None:
            self.formulas.append(entry)
            if name not in self.column_order:
                self.column_order.append(name)
                self.column_visible[name] = True
        else:
            self.formulas[idx] = entry

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

        for f in self.formulas:
            name = f.get("name", "").strip()
            expr = f.get("expr", "").strip()
            if not name or not expr:
                continue
            try:
                df[name] = self._evaluate_formula_expr(df, expr)
            except Exception:
                continue

        if self.remove_duplicates_var.get():
            col = self.duplicate_column_var.get()
            if col and col in df.columns:
                df = df.drop_duplicates(subset=[col], keep="first")

        visible_cols = [c for c in self.column_order if self.column_visible.get(c, True) and c in df.columns]
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
            {
                "name": (f.get("name", "") if isinstance(f, dict) else ""),
                "expr": (f.get("expr", "") if isinstance(f, dict) else ""),
            }
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

        self._refresh_formula_column_choices(self.column_order)
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
        self.calc_col1_var.set("")
        self.calc_col2_var.set("")
        self.calc_constant_var.set("")
        self.calc_decimals_var.set("2")
        self.remove_duplicates_var.set(False)
        self.duplicate_column_var.set("")
        self._refresh_formula_column_choices(self.column_order)
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

        # Keep configured columns even when they are not present in the current
        # raw dataframe yet. This is required for presets/workflows where columns
        # are created later by pivot or VLOOKUP, then edited/calculated in this tab.
        self.column_visible = {c: self.column_visible.get(c, True) for c in self.column_order}
        self.dup_col_cb["values"] = self.column_order
        self._refresh_formula_column_choices(current_cols)
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
