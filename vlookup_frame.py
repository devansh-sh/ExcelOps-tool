import tkinter as tk
from tkinter import ttk


class VlookupFrame(ttk.Frame):
    def __init__(self, parent, on_vlookup, on_vlookup_multi):
        super().__init__(parent)
        self.on_vlookup = on_vlookup
        self.on_vlookup_multi = on_vlookup_multi
        self.mode_var = tk.StringVar(value="single")
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
                "2) Fill in keys and value columns below (saved in presets).\n"
                "3) Click 'Run VLOOKUP' to select the lookup file.\n"
                "4) Preview updates automatically in the Preview tab."
            ),
            wraplength=760,
            justify="left"
        ).pack(anchor="w", **pad)

        form = ttk.Frame(self)
        form.pack(fill="x", **pad)

        ttk.Label(form, text="Mode").grid(row=0, column=0, sticky="w")
        ttk.Radiobutton(form, text="Single Key", variable=self.mode_var, value="single").grid(row=0, column=1, sticky="w")
        ttk.Radiobutton(form, text="Multi-Key", variable=self.mode_var, value="multi").grid(row=0, column=2, sticky="w")

        ttk.Label(form, text="Main key(s)").grid(row=1, column=0, sticky="w")
        ttk.Entry(form, textvariable=self.main_keys_var, width=50).grid(row=1, column=1, columnspan=2, sticky="ew", pady=2)

        ttk.Label(form, text="Lookup key(s)").grid(row=2, column=0, sticky="w")
        ttk.Entry(form, textvariable=self.lookup_keys_var, width=50).grid(row=2, column=1, columnspan=2, sticky="ew", pady=2)

        ttk.Label(form, text="Lookup value columns").grid(row=3, column=0, sticky="w")
        ttk.Entry(form, textvariable=self.values_var, width=50).grid(row=3, column=1, columnspan=2, sticky="ew", pady=2)

        ttk.Label(form, text="Prefix (optional)").grid(row=4, column=0, sticky="w")
        ttk.Entry(form, textvariable=self.prefix_var, width=50).grid(row=4, column=1, columnspan=2, sticky="ew", pady=2)

        ttk.Label(form, text="Default fill (optional)").grid(row=5, column=0, sticky="w")
        ttk.Entry(form, textvariable=self.default_fill_var, width=50).grid(row=5, column=1, columnspan=2, sticky="ew", pady=2)

        actions = ttk.Frame(self)
        actions.pack(fill="x", **pad)

        ttk.Button(actions, text="Run VLOOKUP", command=self.on_vlookup).pack(side="left", padx=4)
        ttk.Button(actions, text="Run Multi-Key VLOOKUP", command=self.on_vlookup_multi).pack(side="left", padx=4)

    def get_config(self):
        return {
            "mode": self.mode_var.get(),
            "main_keys": self.main_keys_var.get(),
            "lookup_keys": self.lookup_keys_var.get(),
            "values": self.values_var.get(),
            "prefix": self.prefix_var.get(),
            "default_fill": self.default_fill_var.get(),
        }

    def load_config(self, cfg: dict):
        self.mode_var.set(cfg.get("mode", "single"))
        self.main_keys_var.set(cfg.get("main_keys", ""))
        self.lookup_keys_var.set(cfg.get("lookup_keys", ""))
        self.values_var.set(cfg.get("values", ""))
        self.prefix_var.set(cfg.get("prefix", ""))
        self.default_fill_var.set(cfg.get("default_fill", ""))
