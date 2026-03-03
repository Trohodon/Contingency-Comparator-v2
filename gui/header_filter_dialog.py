# gui/header_filter_dialog.py

import tkinter as tk
from tkinter import ttk


class HeaderFilterDialog(tk.Toplevel):
    """
    Dialog to let the user choose which headers to filter out.
    The filtered headers are written to the main app's log via log_func.
    """

    def __init__(self, parent, headers, log_func):
        super().__init__(parent)
        self.title("Filter Columns from ViolationCTG Export")
        self.headers = headers
        self.log_func = log_func

        self.geometry("450x400")
        self.transient(parent)
        self.grab_set()

        ttk.Label(
            self,
            text=(
                "Select any columns you want to FILTER OUT.\n"
                "These selections will be logged so you can later\n"
                "hard-code them out of future exports."
            ),
            justify="left",
        ).pack(padx=10, pady=10, anchor="w")

        # Listbox for multi-select of headers
        frame = ttk.Frame(self)
        frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 10))

        self.listbox = tk.Listbox(frame, selectmode=tk.MULTIPLE)
        self.listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        scroll = ttk.Scrollbar(frame, orient="vertical", command=self.listbox.yview)
        scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.listbox.configure(yscrollcommand=scroll.set)

        for h in headers:
            self.listbox.insert(tk.END, str(h))

        btn_frame = ttk.Frame(self)
        btn_frame.pack(fill=tk.X, padx=10, pady=(0, 10))

        ttk.Button(
            btn_frame, text="Log filtered headers", command=self._log_filtered
        ).pack(side=tk.LEFT)

        ttk.Button(btn_frame, text="Close", command=self.destroy).pack(side=tk.RIGHT)

    def _log_filtered(self):
        indices = self.listbox.curselection()
        if not indices:
            self.log_func("No columns selected to filter out.")
            return

        filtered = [self.headers[i] for i in indices]

        self.log_func("\nUser-chosen filtered columns (to hide in future):")
        for h in filtered:
            self.log_func(f"  - {h}")
        self.log_func("End of filtered column list.\n")
