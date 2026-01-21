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

    def __init__(self, master):
        super().__init__(master)

        self.local_log = None
        self.external_log_func = None

        self.pwb_path = tk.StringVar(value="No .pwb file selected")
        self.folder_path = tk.StringVar(value="No folder selected")

        # For single-folder mode: label -> full path
        self.target_cases = {}

        # Filter options
        self.max_filter_var = tk.BooleanVar(value=True)        # dedup
        self.branch_mva_var = tk.BooleanVar(value=True)        # include Branch MVA
        self.bus_lv_var = tk.BooleanVar(value=False)           # include Bus Low Volts
        self.delete_original_var = tk.BooleanVar(value=False)  # delete unfiltered CSV

        # Flag to prevent re-entrancy / double-runs
        self._is_running = False

        self._build_gui()

    # ───────────── Logging helper ───────────── #

    def log(self, msg: str):
        if self.local_log is not None:
            self.local_log.insert(tk.END, msg + "\n")
            self.local_log.see(tk.END)

        if self.external_log_func:
            self.external_log_func(msg)

    # ───────────── GUI layout ───────────── #

    def _build_gui(self):
        # Top frame: single-case controls
        top = ttk.LabelFrame(self, text="Single case processing")
        top.pack(side=tk.TOP, fill=tk.X, padx=10, pady=10)

        ttk.Label(top, text="Selected .pwb case:").grid(row=0, column=0, sticky="w")
        ttk.Label(top, textvariable=self.pwb_path, width=80).grid(
            row=1, column=0, columnspan=2, sticky="w"
        )

        ttk.Button(top, text="Browse .pwb…", command=self.browse_pwb).grid(
            row=1, column=2, padx=(5, 0)
        )

        self.single_btn = ttk.Button(
            top,
            text="Process selected .pwb (export + filter)",
            command=self.run_export_single,
        )
        self.single_btn.grid(row=2, column=0, columnspan=3, pady=(8, 0), sticky="w")

        # Folder frame: folder selection + tree
        folder = ttk.LabelFrame(self, text="Folder processing (ACCA/DC cases)")
        folder.pack(side=tk.TOP, fill=tk.BOTH, expand=False, padx=10, pady=5)

        ttk.Label(folder, text="Selected folder:").grid(row=0, column=0, sticky="w")
        ttk.Label(folder, textvariable=self.folder_path, width=80).grid(
            row=1, column=0, columnspan=2, sticky="w"
        )

        ttk.Button(folder, text="Browse folder…", command=self.browse_folder).grid(
            row=1, column=2, padx=(5, 0)
        )

        self.process_folder_btn = ttk.Button(
            folder,
            text="Process ACCA/DC cases in folder / subfolders",
            command=self.run_export_folder,
        )
        self.process_folder_btn.grid(
            row=2, column=0, columnspan=3, pady=(8, 0), sticky="w"
        )

        # Tree view
        tree_frame = ttk.Frame(folder)
        tree_frame.grid(row=3, column=0, columnspan=3, sticky="nsew", pady=(8, 0))
        folder.rowconfigure(3, weight=1)
        folder.columnconfigure(0, weight=1)

        self.case_tree = ttk.Treeview(
            tree_frame,
            columns=("file", "type"),
            show="headings",
            height=8,
        )
        self.case_tree.heading("file", text="File name")
        self.case_tree.heading("type", text="Case type")
        self.case_tree.column("file", width=500, anchor="w")
        self.case_tree.column("type", width=180, anchor="w")
        self.case_tree.tag_configure("target", foreground="blue")

        tree_scroll = ttk.Scrollbar(
            tree_frame, orient="vertical", command=self.case_tree.yview
        )
        self.case_tree.configure(yscrollcommand=tree_scroll.set)
        self.case_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        tree_scroll.pack(side=tk.RIGHT, fill=tk.Y)

        # Filters frame
        filters = ttk.LabelFrame(self, text="Filters")
        filters.pack(side=tk.TOP, fill=tk.X, padx=10, pady=(4, 4))

        ttk.Checkbutton(
            filters,
            text="Deduplicate LimViolID (keep row(s) with max LimViolPct)",
            variable=self.max_filter_var,
        ).grid(row=0, column=0, sticky="w", padx=5, pady=2)

        ttk.Checkbutton(
            filters,
            text='Include "Branch MVA" rows',
            variable=self.branch_mva_var,
        ).grid(row=1, column=0, sticky="w", padx=5, pady=2)

        ttk.Checkbutton(
            filters,
            text='Include "Bus Low Volts" rows',
            variable=self.bus_lv_var,
        ).grid(row=2, column=0, sticky="w", padx=5, pady=2)

        ttk.Checkbutton(
            filters,
            text="Delete original (unfiltered) CSV after filtering",
            variable=self.delete_original_var,
        ).grid(row=3, column=0, sticky="w", padx=5, pady=(4, 2))

        # Local log box
        log_frame = ttk.LabelFrame(self, text="Case Processing Log")
        log_frame.pack(side=tk.TOP, fill=tk.BOTH, expand=True, padx=10, pady=10)

        self.local_log = tk.Text(log_frame, wrap="word", height=10)
        self.local_log.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        log_scroll = ttk.Scrollbar(
            log_frame, orient="vertical", command=self.local_log.yview
        )
        self.local_log.configure(yscrollcommand=log_scroll.set)
        log_scroll.pack(side=tk.RIGHT, fill=tk.Y)

    # ───────────── Helpers ───────────── #

    def _get_row_filter_categories(self):
        """Return set of LimViolCat values to keep."""
        cats = set()
        if self.branch_mva_var.get():
            cats.add("Branch MVA")
        if self.bus_lv_var.get():
            cats.add("Bus Low Volts")
        return cats

    def _get_output_workbook_path(self, root: str) -> str:
        """
        Naming rules (requested):
          - Branch only: {main folder}_BranchMVA_CTG_Comparison.xlsx
          - Bus only:    {main folder}_BusLowVolts_CTG_Comparison.xlsx
          - Both:        {main folder}_CombinedCTG_Comparison.xlsx
        """
        main_folder_name = os.path.basename(os.path.normpath(root))

        branch = self.branch_mva_var.get()
        bus = self.bus_lv_var.get()

        if branch and bus:
            filename = f"{main_folder_name}_CombinedCTG_Comparison.xlsx"
        elif branch and not bus:
            filename = f"{main_folder_name}_BranchMVA_CTG_Comparison.xlsx"
        elif bus and not branch:
            filename = f"{main_folder_name}_BusLowVolts_CTG_Comparison.xlsx"
        else:
            # Edge case: user unchecked both categories
            filename = f"{main_folder_name}_CTG_Comparison.xlsx"

        return os.path.join(root, filename)

    def _set_running(self, running: bool):
        """Enable/disable buttons while a run is in progress."""
        self._is_running = running
        state = "disabled" if running else "normal"
        self.single_btn.configure(state=state)
        self.process_folder_btn.configure(state=state)
        self.update_idletasks()
        self.update()

    # ───────────── Single-case callbacks ───────────── #

    def browse_pwb(self):
        path = filedialog.askopenfilename(
            title="Select PowerWorld case (.pwb)",
            filetypes=[("PowerWorld case", "*.pwb"), ("All files", "*.*")],
        )
        if path:
            self.pwb_path.set(path)
            self.log(f"Selected case: {path}")

    def run_export_single(self):
        if self._is_running:
            messagebox.showinfo(
                "Busy", "Processing is already running. Please wait for it to finish."
            )
            return

        pwb = self.pwb_path.get()
        if not pwb.lower().endswith(".pwb") or not os.path.exists(pwb):
            messagebox.showwarning(
                "No case selected", "Please select a valid .pwb file."
            )
            return

        cats = self._get_row_filter_categories()
        self.log("\n=== Processing single case ===")
        if not cats:
            self.log(
                "WARNING: No LimViolCat categories selected. Row filter will be skipped."
            )

        self._set_running(True)
        try:
            self.update_idletasks()
            self.update()

            filtered_csv = process_case(
                pwb,
                dedup_enabled=self.max_filter_var.get(),
                keep_categories=cats,
                delete_original=self.delete_original_var.get(),
                log_func=self.log,
            )
        except Exception as e:
            self.log(f"ERROR: {e}")
            messagebox.showerror("Error", str(e))
        else:
            if filtered_csv:
                messagebox.showinfo(
                    "Done", f"Processing complete.\nFiltered CSV:\n{filtered_csv}"
                )
            else:
                messagebox.showwarning(
                    "Done", "Processing finished, but no filtered CSV was created."
                )
        finally:
            self._set_running(False)

    # ───────────── Folder callbacks ───────────── #

    def browse_folder(self):
        folder = filedialog.askdirectory(title="Select folder containing .pwb cases")
        if not folder:
            return

        self.folder_path.set(folder)
        self._scan_and_display_folder(folder)

    def _scan_and_display_folder(self, folder: str):
        """
        Preview .pwb files directly in this folder.
        If there are no .pwb files but there ARE subfolders, show those instead.
        """
        self.case_tree.delete(*self.case_tree.get_children())
        self.target_cases = {}

        cases, target_cases = scan_folder(folder, self.log)
        self.target_cases = target_cases

        if cases:
            for info in cases:
                tag = "target" if info["is_target"] else ""
                self.case_tree.insert(
                    "",
                    "end",
                    values=(info["filename"], info["type"]),
                    tags=(tag,) if tag else (),
                )
            return

        subdirs = sorted(
            d for d in os.listdir(folder)
            if os.path.isdir(os.path.join(folder, d))
        )

        if not subdirs:
            self.log("No .pwb files or subfolders found in this folder.")
            return

        self.log(
            "No .pwb files directly in this folder; showing subfolders as scenarios."
        )

        for d in subdirs:
            self.case_tree.insert(
                "",
                "end",
                values=(d, "Scenario subfolder"),
            )

    def run_export_folder(self):
        if self._is_running:
            messagebox.showinfo(
                "Busy", "Processing is already running. Please wait for it to finish."
            )
            return

        root = self.folder_path.get()
        if not os.path.isdir(root):
            messagebox.showwarning(
                "No folder selected", "Please select a valid folder."
            )
            return

        cats = self._get_row_filter_categories()
        if not cats:
            self.log(
                "WARNING: No LimViolCat categories selected. Row filter will be skipped."
            )

        subdirs = sorted(
            d for d in os.listdir(root)
            if os.path.isdir(os.path.join(root, d))
        )

        self._set_running(True)
        try:
            if subdirs:
                self._run_export_multi_folder(root, subdirs, cats)
            else:
                _, target_cases = scan_folder(root, self.log)
                self.target_cases = target_cases
                self._run_export_single_folder(root, cats)
        finally:
            self._set_running(False)

    # ---------- Single-folder mode ---------- #

    def _run_export_single_folder(self, folder: str, cats):
        if not self.target_cases:
            messagebox.showwarning(
                "No target cases found",
                "No ACCA_LongTerm / ACCA_P1,2,4,7 / DCwACver_P1-7 cases detected.",
            )
            return

        self.log("\n=== Batch processing ACCA/DC cases in folder ===")

        errors = []
        for label in TARGET_PATTERNS:
            self.update_idletasks()
            self.update()

            pwb_path = self.target_cases.get(label)
            if not pwb_path:
                self.log(f"Skipping type [{label}] (not found).")
                continue

            self.log(f"\n--- Processing [{label}] case ---")
            self.log(f"Case path: {pwb_path}")
            try:
                filtered_csv = process_case(
                    pwb_path,
                    dedup_enabled=self.max_filter_var.get(),
                    keep_categories=cats,
                    delete_original=self.delete_original_var.get(),
                    log_func=self.log,
                )
                if not filtered_csv:
                    raise RuntimeError("No filtered CSV was created.")
            except Exception as e:
                msg = f"ERROR processing [{label}] case: {e}"
                self.log(msg)
                errors.append(msg)

        if errors:
            messagebox.showerror(
                "Batch processing completed with errors",
                "Some cases failed. Check the log window for details.",
            )
        else:
            messagebox.showinfo(
                "Batch processing complete",
                "All detected ACCA/DC cases in the folder have been processed.",
            )

    # ---------- Multi-folder mode ---------- #

    def _run_export_multi_folder(self, root: str, subdirs, cats):
        self.log(
            "\n=== Multi-folder mode: each subfolder is a case set to compare ==="
        )
        self.log(f"Root folder: {root}")
        self.log(f"Subfolders found: {', '.join(subdirs)}")

        folder_to_case_csvs = {}
        errors = []

        for sub in subdirs:
            self.update_idletasks()
            self.update()

            scenario_folder = os.path.join(root, sub)
            self.log(f"\n=== Processing scenario folder: {sub} ===")

            _, target_cases = scan_folder(scenario_folder, self.log)
            if not target_cases:
                self.log(f"  [{sub}] No ACCA/DC cases found; skipping.")
                continue

            case_csvs = {}

            for label in TARGET_PATTERNS:
                self.update_idletasks()
                self.update()

                pwb_path = target_cases.get(label)
                if not pwb_path:
                    self.log(f"  [{sub}] Skipping type [{label}] (not found).")
                    continue

                self.log(f"\n  [{sub}] --- Processing [{label}] case ---")
                self.log(f"  Case path: {pwb_path}")
                try:
                    filtered_csv = process_case(
                        pwb_path,
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

        # Build the combined workbook in the root folder with the requested naming rule
        output_workbook_path = self._get_output_workbook_path(root)
        workbook_path = build_workbook(root, folder_to_case_csvs, self.log, workbook_path=output_workbook_path)

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