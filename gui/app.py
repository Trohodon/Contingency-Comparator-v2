# gui/app.py

import tkinter as tk
from tkinter import ttk

from .tab_case import CaseProcessingTab
from .tab_compare import CompareTab
from .tab_settings import SettingsTab
from .tab_logs import LogsTab


class PwbExportApp(tk.Tk):
    def __init__(self):
        super().__init__()

        self.title("PowerWorld Ribbon Tool")
        self.geometry("1100x750")

        # Create ribbon (Notebook)
        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill=tk.BOTH, expand=True)

        # ---- Create tabs ----
        self.tab_case = CaseProcessingTab(self)
        self.tab_compare = CompareTab(self)
        self.tab_settings = SettingsTab(self)
        self.tab_logs = LogsTab(self)

        # ---- Add tabs to Notebook ----
        self.notebook.add(self.tab_case, text="Case Processing")
        self.notebook.add(self.tab_compare, text="Compare Cases")
        self.notebook.add(self.tab_settings, text="Settings")
        self.notebook.add(self.tab_logs, text="Logs")

        # Allow Case tab to log directly to Logs tab
        self.tab_case.external_log_func = self.tab_logs.write