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
        self.same_keys_var = tk.BooleanVar(value=True)
        self._build_ui()

    def _build_ui(self):
        pad = {"padx": 6, "pady": 6}
        ttk.Label(
            self,
            text=(
                "How to use VLOOKUP:\n"
                "1) Choose your main sheet and apply filters if needed.\n"
                "2) Pick main keys and value columns below.\n"
                "3) Click 'Run VLOOKUP' to select the lookup file.\n"
                "4) Preview updates automatically in the Preview tab."
            ),
            wraplength=760,
            justify="left"
        ).pack(anchor="w", **pad)

        form = ttk.LabelFrame(self, text="VLOOKUP Settings")
        form.pack(fill="x", **pad)
        form.columnconfigure(1, weight=1)
        form.columnconfigure(2, weight=1)

        ttk.Label(form, text="Main keys").grid(row=0, column=0, sticky="w", padx=6, pady=4)
        self.main_keys_lb = tk.Listbox(form, selectmode="multiple", height=6, exportselection=False)
        self.main_keys_lb.grid(row=1, column=0, sticky="nsew", padx=6, pady=4)

        ttk.Label(form, text="Lookup value columns").grid(row=0, column=1, sticky="w", padx=6, pady=4)
        self.values_lb = tk.Listbox(form, selectmode="multiple", height=6, exportselection=False)
        self.values_lb.grid(row=1, column=1, sticky="nsew", padx=6, pady=4)

        options = ttk.Frame(form)
        options.grid(row=1, column=2, sticky="nw", padx=6, pady=4)

        ttk.Checkbutton(
            options,
            text="Lookup keys match main keys",
            variable=self.same_keys_var,
            command=self._toggle_lookup_keys
        ).pack(anchor="w", pady=(0, 6))

        ttk.Label(options, text="Lookup key(s)").pack(anchor="w")
        self.lookup_keys_entry = ttk.Entry(options, textvariable=self.lookup_keys_var, width=28)
        self.lookup_keys_entry.pack(fill="x", pady=(2, 6))

        ttk.Label(options, text="Prefix (optional)").pack(anchor="w")
        ttk.Entry(options, textvariable=self.prefix_var, width=28).pack(fill="x", pady=(2, 6))

        ttk.Label(options, text="Default fill (optional)").pack(anchor="w")
        ttk.Entry(options, textvariable=self.default_fill_var, width=28).pack(fill="x", pady=(2, 6))

        actions = ttk.Frame(self)
        actions.pack(fill="x", **pad)

        ttk.Button(actions, text="Run VLOOKUP", command=self.on_vlookup).pack(side="left", padx=4)

    def get_config(self):
        main_keys = [self.main_keys_lb.get(i) for i in self.main_keys_lb.curselection()]
        values = [self.values_lb.get(i) for i in self.values_lb.curselection()]
        lookup_keys = "" if self.same_keys_var.get() else self.lookup_keys_var.get()
        return {
            "main_keys": ", ".join(main_keys),
            "lookup_keys": lookup_keys,
            "values": ", ".join(values),
            "prefix": self.prefix_var.get(),
            "default_fill": self.default_fill_var.get(),
        }

    def load_config(self, cfg: dict):
        self.main_keys_var.set(cfg.get("main_keys", ""))
        self.lookup_keys_var.set(cfg.get("lookup_keys", ""))
        self.values_var.set(cfg.get("values", ""))
        self.prefix_var.set(cfg.get("prefix", ""))
        self.default_fill_var.set(cfg.get("default_fill", ""))
        self.same_keys_var.set(cfg.get("lookup_keys", "") == "")
        self._apply_listbox_selections()
        self._toggle_lookup_keys()

    def set_columns(self, columns):
        self.columns = columns or []
        self._refresh_listbox(self.main_keys_lb, self.columns)
        self._refresh_listbox(self.values_lb, self.columns)
        self._apply_listbox_selections()

    def _refresh_listbox(self, listbox, values):
        listbox.delete(0, "end")
        for item in values:
            listbox.insert("end", item)

    def _apply_listbox_selections(self):
        main_keys = [c.strip() for c in self.main_keys_var.get().split(",") if c.strip()]
        values = [c.strip() for c in self.values_var.get().split(",") if c.strip()]
        self._select_values(self.main_keys_lb, main_keys)
        self._select_values(self.values_lb, values)

    def _select_values(self, listbox, values):
        listbox.selection_clear(0, "end")
        for i in range(listbox.size()):
            if listbox.get(i) in values:
                listbox.selection_set(i)

    def _toggle_lookup_keys(self):
        state = "disabled" if self.same_keys_var.get() else "normal"
        self.lookup_keys_entry.configure(state=state)
