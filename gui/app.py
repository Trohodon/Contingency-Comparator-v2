# gui/app.py

import tkinter as tk
from tkinter import ttk

from gui.tab_case import CaseProcessingTab
from gui.tab_compare import CompareTab


class RibbonApp(tk.Tk):
    """
    Main application window for the PowerWorld Ribbon Tool.
    """

    def __init__(self):
        super().__init__()

        self.title("PowerWorld Ribbon Tool")
        self.geometry("1100x750")

        # Top-level notebook for "ribbon" tabs
        notebook = ttk.Notebook(self)
        notebook.pack(fill=tk.BOTH, expand=True)

        # Case Processing tab (tab 1)
        self.case_tab = CaseProcessingTab(notebook)
        notebook.add(self.case_tab, text="Case Processing")

        # Compare Cases tab (tab 2)
        self.compare_tab = CompareTab(notebook)
        notebook.add(self.compare_tab, text="Compare Cases")

        # Optional: hook their external_log_func to a common logger if you want.
        # For now, keep them separate. If later you add a global "Logs" tab,
        # you can wire it here.


def main():
    app = RibbonApp()
    app.mainloop()


if __name__ == "__main__":
    main()