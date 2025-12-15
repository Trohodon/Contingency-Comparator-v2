# gui/app.py
from __future__ import annotations

import tkinter as tk
from tkinter import ttk

# Your existing tabs (keep these filenames/imports matching your project)
# If your modules/classes are named differently, adjust ONLY these imports.
from gui.case_processing_view import CaseProcessingView
from gui.compare_cases_view import CompareCasesView

# New Trends tab
from gui.trends_view import TrendsView


class App(tk.Tk):
    """
    Main window for the Contingency Comparison Tool.

    IMPORTANT:
    - App is the Tk root (so you can do App().mainloop()).
    - Each tab is a ttk.Frame subclass that receives (master=self.notebook).
    """

    def __init__(self):
        super().__init__()

        # ---- Window ----
        self.title("Contingency Comparison Tool")
        self.geometry("1250x760")
        self.minsize(1050, 650)

        # Use a modern theme if available
        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except Exception:
            pass

        # ---- Header ----
        header = ttk.Frame(self)
        header.pack(side="top", fill="x")

        title_row = ttk.Frame(header)
        title_row.pack(side="top", fill="x", padx=12, pady=(10, 2))

        self._title_label = ttk.Label(
            title_row,
            text="Contingency Comparison Tool",
            font=("Segoe UI", 16, "bold"),
        )
        self._title_label.pack(side="left")

        self._subtitle_label = ttk.Label(
            title_row,
            text="PowerWorld Results Export + Compare",
            font=("Segoe UI", 10),
        )
        self._subtitle_label.pack(side="left", padx=(12, 0))

        # ---- Notebook ----
        body = ttk.Frame(self)
        body.pack(side="top", fill="both", expand=True)

        self.notebook = ttk.Notebook(body)
        self.notebook.pack(fill="both", expand=True, padx=10, pady=10)

        # ---- Tabs ----
        # Each tab is responsible for its own UI and logic
        self.case_processing_tab = CaseProcessingView(self.notebook)
        self.compare_cases_tab = CompareCasesView(self.notebook)
        self.trends_tab = TrendsView(self.notebook)

        self.notebook.add(self.case_processing_tab, text="Case Processing")
        self.notebook.add(self.compare_cases_tab, text="Compare Cases")
        self.notebook.add(self.trends_tab, text="Trends")

        # Optional: default to Compare Cases tab
        # self.notebook.select(self.compare_cases_tab)

        # ---- Footer (optional) ----
        footer = ttk.Frame(self)
        footer.pack(side="bottom", fill="x")

        self._status_var = tk.StringVar(value="Ready")
        self._status_label = ttk.Label(
            footer,
            textvariable=self._status_var,
            anchor="w",
        )
        self._status_label.pack(side="left", fill="x", expand=True, padx=10, pady=6)

        # Quit shortcut
        self.bind("<Escape>", lambda _e: self.destroy())

    def set_status(self, text: str):
        """Tabs can call this if you wire it in later."""
        self._status_var.set(text)


if __name__ == "__main__":
    App().mainloop()