import tkinter as tk
from tkinter import ttk


class VlookupFrame(ttk.Frame):
    def __init__(self, parent, on_vlookup, on_vlookup_multi):
        super().__init__(parent)
        self.on_vlookup = on_vlookup
        self.on_vlookup_multi = on_vlookup_multi
        self._build_ui()

    def _build_ui(self):
        pad = {"padx": 6, "pady": 6}
        ttk.Label(
            self,
            text=(
                "VLOOKUP: merge a lookup file into the current sheet.\n"
                "Use single-key for one matching column or multi-key for multiple keys."
            ),
            wraplength=760,
            justify="left"
        ).pack(anchor="w", **pad)

        actions = ttk.Frame(self)
        actions.pack(fill="x", **pad)

        ttk.Button(actions, text="Run VLOOKUP", command=self.on_vlookup).pack(side="left", padx=4)
        ttk.Button(actions, text="Run Multi-Key VLOOKUP", command=self.on_vlookup_multi).pack(side="left", padx=4)
