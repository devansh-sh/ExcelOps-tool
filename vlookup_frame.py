import tkinter as tk
from tkinter import ttk


class VlookupFrame(ttk.Frame):
    def __init__(self, parent, on_vlookup, columns=None):
        super().__init__(parent)
        self.on_vlookup = on_vlookup
        self.columns = columns or []
        self.main_keys_var = tk.StringVar()
        self.lookup_keys_var = tk.StringVar()
        self.values_var = tk.StringVar()
        self.prefix_var = tk.StringVar()
        self.default_fill_var = tk.StringVar()
        self._build_ui()

    def _build_ui(self):
        pad = {"padx": 6, "pady": 6}
        ttk.Label(
            self,
            text=(
                "How to use VLOOKUP:\n"
                "1) Choose your main sheet and apply filters if needed.\n"
                "2) Fill in keys/value columns (comma-separated for multi-key).\n"
                "3) Click 'Run VLOOKUP' to select the lookup file.\n"
                "4) Preview updates automatically in the Preview tab."
            ),
            wraplength=760,
            justify="left"
        ).pack(anchor="w", **pad)

        form = ttk.LabelFrame(self, text="VLOOKUP Settings")
        form.pack(fill="x", **pad)
        form.columnconfigure(1, weight=1)

        ttk.Label(form, text="Main key(s)").grid(row=0, column=0, sticky="w", padx=6, pady=4)
        self.main_keys_cb = ttk.Combobox(
            form,
            textvariable=self.main_keys_var,
            values=self.columns,
            width=48
        )
        self.main_keys_cb.grid(row=0, column=1, sticky="ew", padx=6, pady=4)
        ttk.Label(form, text="(comma-separated)").grid(row=0, column=2, sticky="w", padx=6, pady=4)

        ttk.Label(form, text="Lookup key(s)").grid(row=1, column=0, sticky="w", padx=6, pady=4)
        ttk.Entry(form, textvariable=self.lookup_keys_var, width=50).grid(row=1, column=1, sticky="ew", padx=6, pady=4)
        ttk.Label(form, text="(comma-separated)").grid(row=1, column=2, sticky="w", padx=6, pady=4)

        ttk.Label(form, text="Lookup value columns").grid(row=2, column=0, sticky="w", padx=6, pady=4)
        ttk.Entry(form, textvariable=self.values_var, width=50).grid(row=2, column=1, sticky="ew", padx=6, pady=4)
        ttk.Label(form, text="(comma-separated)").grid(row=2, column=2, sticky="w", padx=6, pady=4)

        ttk.Label(form, text="Prefix (optional)").grid(row=3, column=0, sticky="w", padx=6, pady=4)
        ttk.Entry(form, textvariable=self.prefix_var, width=50).grid(row=3, column=1, sticky="ew", padx=6, pady=4)

        ttk.Label(form, text="Default fill (optional)").grid(row=4, column=0, sticky="w", padx=6, pady=4)
        ttk.Entry(form, textvariable=self.default_fill_var, width=50).grid(row=4, column=1, sticky="ew", padx=6, pady=4)

        actions = ttk.Frame(self)
        actions.pack(fill="x", **pad)

        ttk.Button(actions, text="Run VLOOKUP", command=self.on_vlookup).pack(side="left", padx=4)

    def get_config(self):
        return {
            "main_keys": self.main_keys_var.get(),
            "lookup_keys": self.lookup_keys_var.get(),
            "values": self.values_var.get(),
            "prefix": self.prefix_var.get(),
            "default_fill": self.default_fill_var.get(),
        }

    def load_config(self, cfg: dict):
        self.main_keys_var.set(cfg.get("main_keys", ""))
        self.lookup_keys_var.set(cfg.get("lookup_keys", ""))
        self.values_var.set(cfg.get("values", ""))
        self.prefix_var.set(cfg.get("prefix", ""))
        self.default_fill_var.set(cfg.get("default_fill", ""))

    def set_columns(self, columns):
        self.columns = columns or []
        if hasattr(self, "main_keys_cb"):
            self.main_keys_cb["values"] = self.columns
