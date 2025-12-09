# gui/tab_case.py

import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

import pandas as pd

from core.pwb_exporter import export_violation_ctg
from core.column_blacklist import (
    apply_blacklist,
    apply_row_filter,
    apply_limviolid_max_filter,
)

TARGET_PATTERNS = {
    "ACCA_LongTerm": "ACCA_LongTerm",
    "ACCA_P1,2,4,7": "ACCA_P1,2,4,7",
    "DCwACver_P1-7": "DCwACver_P1-7",
}


class CaseProcessingTab(ttk.Frame):
    def __init__(self, master):
        super().__init__(master)

        # Logs in this tab still shown locally
        self.local_log = None

        # External logger (Logs tab)
        self.external_log_func = None

        self.pwb_path = tk.StringVar(value="No .pwb file selected")
        self.folder_path = tk.StringVar(value="No folder selected")

        self.target_cases = {}

        self.max_filter_var = tk.BooleanVar(value=True)

        self._build_gui()

    def log(self, msg):
        """Logs to log window inside this tab AND to global Logs tab."""
        if self.local_log:
            self.local_log.insert(tk.END, msg + "\n")
            self.local_log.see(tk.END)

        if self.external_log_func:
            self.external_log_func(msg)

    # ---------------- GUI ---------------- #

    def _build_gui(self):
        # Top frame
        top = ttk.LabelFrame(self, text="Single case processing")
        top.pack(fill=tk.X, padx=10, pady=10)

        ttk.Label(top, text="Selected .pwb case:").grid(row=0, column=0, sticky="w")
        ttk.Label(top, textvariable=self.pwb_path, width=80).grid(
            row=1, column=0, columnspan=2, sticky="w"
        )

        browse_btn = ttk.Button(top, text="Browse .pwb…", command=self.browse_pwb)
        browse_btn.grid(row=1, column=2)

        run_btn = ttk.Button(
            top, text="Process selected .pwb", command=self.run_export_single
        )
        run_btn.grid(row=2, column=0, pady=6, sticky="w")

        # Folder area
        folder = ttk.LabelFrame(self, text="Folder processing")
        folder.pack(fill=tk.X, padx=10, pady=5)

        ttk.Label(folder, text="Selected folder:").grid(row=0, column=0, sticky="w")
        ttk.Label(folder, textvariable=self.folder_path, width=80).grid(
            row=1, column=0, columnspan=2, sticky="w"
        )

        ttk.Button(folder, text="Browse folder…", command=self.browse_folder).grid(
            row=1, column=2
        )

        ttk.Button(
            folder,
            text="Process ACCA/DC cases in folder",
            command=self.run_export_folder,
        ).grid(row=2, column=0, pady=5, sticky="w")

        # Tree for folder preview
        tree_frame = ttk.Frame(folder)
        tree_frame.grid(row=3, column=0, columnspan=3, sticky="nsew", pady=(8, 0))
        folder.rowconfigure(3, weight=1)
        folder.columnconfigure(0, weight=1)

        self.case_tree = ttk.Treeview(
            tree_frame, columns=("file", "type"), show="headings", height=8
        )
        self.case_tree.heading("file", text="File name")
        self.case_tree.heading("type", text="Type")
        self.case_tree.column("file", width=500)
        self.case_tree.column("type", width=180)
        self.case_tree.tag_configure("target", foreground="blue")

        tree_scroll = ttk.Scrollbar(
            tree_frame, orient="vertical", command=self.case_tree.yview
        )
        self.case_tree.configure(yscrollcommand=tree_scroll.set)
        self.case_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        tree_scroll.pack(side=tk.RIGHT, fill=tk.Y)

        # Filters
        filters = ttk.LabelFrame(self, text="Filters")
        filters.pack(fill=tk.X, padx=10, pady=5)

        ttk.Checkbutton(
            filters,
            text="Deduplicate LimViolID (max LimViolPct)",
            variable=self.max_filter_var,
        ).pack(anchor="w")

        # Log box
        log_frame = ttk.LabelFrame(self, text="Local Log")
        log_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        self.local_log = tk.Text(log_frame, height=10)
        self.local_log.pack(fill=tk.BOTH, expand=True)

    # ---------------- Functionality (SAME AS BEFORE) ---------------- #

    # everything below is identical to your working logic: browse_pwb, browse_folder,
    # scanning, row-filter, dedup filter, column blacklist, filtered CSV saving …