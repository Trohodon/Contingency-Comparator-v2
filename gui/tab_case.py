import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

from core.case_finder import scan_folder, TARGET_PATTERNS
from core.case_processor import process_case
from core.comparison_builder import build_workbook


class CaseProcessingTab(ttk.Frame):
    """
    GUI tab for:
      - Single case processing
      - Folder scan + processing of ACCA/DC cases
      - Multi-folder mode: each subfolder is a scenario to compare
    """

    def __init__(self, parent):
        super().__init__(parent)

        self.is_processing = False

        # -----------------------------
        # UI variables / settings
        # -----------------------------
        self.selected_root = tk.StringVar(value="")
        self.single_mode_var = tk.BooleanVar(value=False)       # if True, process a single folder (not multi)
        self.max_filter_var = tk.BooleanVar(value=True)         # dedup
        self.branch_mva_var = tk.BooleanVar(value=True)         # include Branch MVA
        self.bus_lv_var = tk.BooleanVar(value=False)            # include Bus Low Volts
        self.delete_original_var = tk.BooleanVar(value=True)    # delete unfiltered csv

        # internal storage for the scan results
        self.cases_found = []           # list of dicts: {"name","path","type","is_target"}
        self.target_cases = {}          # single-folder mode: label -> full path
        self.case_vars = {}             # checkbox vars for each found case in single-folder mode

        self._build_ui()

    # -----------------------------
    # UI layout
    # -----------------------------
    def _build_ui(self):
        root_frame = ttk.LabelFrame(self, text="Main Folder")
        root_frame.pack(fill=tk.X, padx=10, pady=10)

        ttk.Entry(root_frame, textvariable=self.selected_root, width=70).pack(
            side=tk.LEFT, padx=(10, 5), pady=8
        )
        ttk.Button(root_frame, text="Browse", command=self._browse_root).pack(
            side=tk.LEFT, padx=5, pady=8
        )
        ttk.Button(root_frame, text="Scan", command=self._scan_root).pack(
            side=tk.LEFT, padx=5, pady=8
        )

        mode_frame = ttk.Frame(self)
        mode_frame.pack(fill=tk.X, padx=10)

        ttk.Checkbutton(
            mode_frame,
            text="Single-folder mode (process cases in the selected folder only)",
            variable=self.single_mode_var,
            command=self._refresh_case_list,
        ).pack(side=tk.LEFT)

        # -----------------------------
        # Case list + controls
        # -----------------------------
        cases_frame = ttk.LabelFrame(self, text="Cases")
        cases_frame.pack(fill=tk.BOTH, expand=False, padx=10, pady=(10, 5))

        list_frame = ttk.Frame(cases_frame)
        list_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        self.case_tree = ttk.Treeview(
            list_frame,
            columns=("Type", "Path"),
            show="headings",
            height=8,
        )
        self.case_tree.heading("Type", text="Type")
        self.case_tree.heading("Path", text="Path")
        self.case_tree.column("Type", width=140, anchor=tk.W)
        self.case_tree.column("Path", width=700, anchor=tk.W)
        self.case_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        yscroll = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.case_tree.yview)
        yscroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.case_tree.configure(yscrollcommand=yscroll.set)

        action_frame = ttk.Frame(self)
        action_frame.pack(fill=tk.X, padx=10, pady=(0, 10))

        ttk.Button(
            action_frame,
            text="Process",
            command=self._process_selected,
        ).pack(side=tk.LEFT)

        # -----------------------------
        # Filters
        # -----------------------------
        filters_frame = ttk.LabelFrame(self, text="Filters")
        filters_frame.pack(fill=tk.X, padx=10, pady=(0, 10))

        ttk.Checkbutton(
            filters_frame,
            text="Expandable issue view (Excel dropdown: show max row, expand to see others)",
            variable=self.max_filter_var,
        ).grid(row=0, column=0, sticky="w", padx=10, pady=5)

        ttk.Checkbutton(
            filters_frame,
            text='Include "Branch MVA" rows',
            variable=self.branch_mva_var,
        ).grid(row=1, column=0, sticky="w", padx=10, pady=5)

        ttk.Checkbutton(
            filters_frame,
            text='Include "Bus Low Volts" rows',
            variable=self.bus_lv_var,
        ).grid(row=2, column=0, sticky="w", padx=10, pady=5)

        ttk.Checkbutton(
            filters_frame,
            text="Delete original (unfiltered) CSV after filtering",
            variable=self.delete_original_var,
        ).grid(row=3, column=0, sticky="w", padx=10, pady=5)

        # -----------------------------
        # Log
        # -----------------------------
        log_frame = ttk.LabelFrame(self, text="Case Processing Log")
        log_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 10))

        self.local_log = tk.Text(log_frame, height=14, wrap="word")
        self.local_log.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=10, pady=10)

        log_scroll = ttk.Scrollbar(log_frame, orient=tk.VERTICAL, command=self.local_log.yview)
        log_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.local_log.configure(yscrollcommand=log_scroll.set)

    # -----------------------------
    # Helpers
    # -----------------------------
    def log(self, msg: str):
        self.local_log.insert(tk.END, msg + "\n")
        self.local_log.see(tk.END)

    def _browse_root(self):
        folder = filedialog.askdirectory()
        if folder:
            self.selected_root.set(folder)

    def _scan_root(self):
        root = self.selected_root.get().strip()
        if not root or not os.path.isdir(root):
            messagebox.showerror("Invalid folder", "Please select a valid main folder.")
            return

        self.log(f"\nScanning folder:\n  {root}")
        cases, targets = scan_folder(root, log_func=self.log)

        self.cases_found = cases
        self.target_cases = targets
        self._refresh_case_list()

    def _refresh_case_list(self):
        # Clear tree
        for item in self.case_tree.get_children():
            self.case_tree.delete(item)

        # Single-folder mode: show the target cases found in this folder
        if self.single_mode_var.get():
            for label in TARGET_PATTERNS:
                path = self.target_cases.get(label, "")
                if path:
                    self.case_tree.insert("", "end", values=(label, path))
                else:
                    self.case_tree.insert("", "end", values=(label, "(missing)"))

        # Multi-folder mode: show subfolders that contain pwb files
        else:
            subfolders = []
            try:
                for name in sorted(os.listdir(self.selected_root.get().strip())):
                    fpath = os.path.join(self.selected_root.get().strip(), name)
                    if os.path.isdir(fpath):
                        subfolders.append(fpath)
            except Exception:
                subfolders = []

            for sf in subfolders:
                self.case_tree.insert("", "end", values=("ScenarioFolder", sf))

    def _process_selected(self):
        if self.is_processing:
            messagebox.showwarning(
                "Busy", "Processing is already running. Please wait for it to finish."
            )
            return

        root = self.selected_root.get().strip()
        if not root or not os.path.isdir(root):
            messagebox.showerror("Invalid folder", "Please select a valid main folder.")
            return

        self.is_processing = True
        try:
            if self.single_mode_var.get():
                self._process_single_folder(root)
            else:
                self._process_multi_folder(root)
        finally:
            self.is_processing = False

    # -----------------------------
    # Processing logic
    # -----------------------------
    def _process_single_folder(self, root: str):
        # categories based on checkboxes
        cats = set()
        if self.branch_mva_var.get():
            cats.add("Branch MVA")
        if self.bus_lv_var.get():
            cats.add("Bus Low Volts")

        # process each target case in this folder
        errors = []
        for label in TARGET_PATTERNS:
            pwb = self.target_cases.get(label)
            if not pwb:
                self.log(f"ERROR: Missing target case for [{label}] in folder.")
                errors.append(label)
                continue
            try:
                self.log(f"\nProcessing [{label}] case:\n  {pwb}")
                filtered_csv = process_case(
                    pwb,
                    dedup_enabled=self.max_filter_var.get(),
                    keep_categories=cats,
                    delete_original=self.delete_original_var.get(),
                    log_func=self.log,
                )
                if not filtered_csv:
                    raise RuntimeError("No filtered CSV was created.")
                self.log(f"Filtered CSV created:\n  {filtered_csv}")
            except Exception as e:
                msg = f"ERROR processing [{label}] case: {e}"
                self.log(msg)
                errors.append(msg)

        if errors:
            messagebox.showerror(
                "Processing completed with errors",
                "Some cases failed; see log for details.",
            )
        else:
            messagebox.showinfo(
                "Processing complete",
                "All cases processed successfully. See log for details.",
            )

    def _process_multi_folder(self, root: str):
        # categories based on checkboxes
        cats = set()
        if self.branch_mva_var.get():
            cats.add("Branch MVA")
        if self.bus_lv_var.get():
            cats.add("Bus Low Volts")

        self.log(f"\nMulti-folder mode: processing subfolders in:\n  {root}")
        subfolders = [
            f for f in sorted(os.listdir(root)) if os.path.isdir(os.path.join(root, f))
        ]

        folder_to_case_csvs = {}
        errors = []

        for sub in subfolders:
            sub_path = os.path.join(root, sub)
            self.log(f"\n--- Scenario folder: {sub} ---")
            cases, targets = scan_folder(sub_path, log_func=self.log)

            case_csvs = {}
            for label in TARGET_PATTERNS:
                pwb = targets.get(label)
                if not pwb:
                    self.log(f"  [{sub}] WARNING: Missing [{label}] case.")
                    continue
                try:
                    self.log(f"  [{sub}] Processing [{label}] case:\n    {pwb}")
                    filtered_csv = process_case(
                        pwb,
                        dedup_enabled=self.max_filter_var.get(),
                        keep_categories=cats,
                        delete_original=self.delete_original_var.get(),
                        log_func=self.log,
                    )
                    if not filtered_csv:
                        raise RuntimeError("No filtered CSV was created.")
                    case_csvs[label] = filtered_csv
                except Exception as e:
                    msg = f"  [{sub}] ERROR processing [{label}] case: {e}"
                    self.log(msg)
                    errors.append(msg)

            if case_csvs:
                folder_to_case_csvs[sub] = case_csvs
            else:
                self.log(f"  [{sub}] No filtered CSVs produced; no sheet will be made.")

        # Build the combined workbook in the root folder
        workbook_path = build_workbook(
            root,
            folder_to_case_csvs,
            group_details=self.max_filter_var.get(),
            include_branch_mva=self.branch_mva_var.get(),
            include_bus_lv=self.bus_lv_var.get(),
            log_func=self.log,
        )

        if workbook_path:
            self.log(f"\nCombined workbook created at:\n  {workbook_path}")
            if errors:
                messagebox.showerror(
                    "Multi-folder processing completed with errors",
                    f"Workbook created:\n{workbook_path}\n\n"
                    "Some cases failed; see log for details.",
                )
            else:
                messagebox.showinfo(
                    "Multi-folder processing complete",
                    f"Workbook created:\n{workbook_path}",
                )
        else:
            if errors:
                messagebox.showerror(
                    "Processing completed with errors",
                    "No combined workbook created. See log for details.",
                )
            else:
                messagebox.showwarning(
                    "Nothing processed",
                    "No valid subfolders / cases found to build a workbook.",
                )
