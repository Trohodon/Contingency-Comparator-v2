# gui/tab_case.py

import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

from core.case_finder import scan_folder, TARGET_PATTERNS
from core.case_processor import process_case


class CaseProcessingTab(ttk.Frame):
    """
    GUI tab for:
      - Single case processing
      - Folder scan + processing of the 3 ACCA/DC cases
    """

    def __init__(self, master):
        super().__init__(master)

        self.local_log = None
        self.external_log_func = None

        self.pwb_path = tk.StringVar(value="No .pwb file selected")
        self.folder_path = tk.StringVar(value="No folder selected")

        self.target_cases = {}

        # Filter options
        self.max_filter_var = tk.BooleanVar(value=True)      # existing
        self.branch_mva_var = tk.BooleanVar(value=True)      # NEW – default ON
        self.bus_lv_var = tk.BooleanVar(value=False)         # NEW – default OFF
        self.delete_original_var = tk.BooleanVar(value=False)  # NEW – default OFF

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

        ttk.Button(
            top,
            text="Process selected .pwb (export + filter)",
            command=self.run_export_single,
        ).grid(row=2, column=0, columnspan=3, pady=(8, 0), sticky="w")

        # Folder frame
        folder = ttk.LabelFrame(self, text="Folder processing (3 ACCA/DC cases)")
        folder.pack(side=tk.TOP, fill=tk.BOTH, expand=False, padx=10, pady=5)

        ttk.Label(folder, text="Selected folder:").grid(row=0, column=0, sticky="w")
        ttk.Label(folder, textvariable=self.folder_path, width=80).grid(
            row=1, column=0, columnspan=2, sticky="w"
        )

        ttk.Button(folder, text="Browse folder…", command=self.browse_folder).grid(
            row=1, column=2, padx=(5, 0)
        )

        ttk.Button(
            folder,
            text="Process 3 ACCA/DC cases in folder",
            command=self.run_export_folder,
        ).grid(row=2, column=0, columnspan=3, pady=(8, 0), sticky="w")

        # Tree view
        tree_frame = ttk.Frame(folder)
        tree_frame.grid(row=3, column=0, columnspan=3, sticky="nsew", pady=(8, 0))
        folder.rowconfigure(3, weight=1)
        folder.columnconfigure(0, weight=1)

        self.case_tree = ttk.Treeview(
            tree_frame, columns=("file", "type"), show="headings", height=8
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
        cats = set()
        if self.branch_mva_var.get():
            cats.add("Branch MVA")
        if self.bus_lv_var.get():
            cats.add("Bus Low Volts")
        return cats

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

        try:
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
            return

        if filtered_csv:
            messagebox.showinfo(
                "Done", f"Processing complete.\nFiltered CSV:\n{filtered_csv}"
            )
        else:
            messagebox.showwarning(
                "Done", "Processing finished, but no filtered CSV was created."
            )

    # ───────────── Folder callbacks ───────────── #

    def browse_folder(self):
        folder = filedialog.askdirectory(title="Select folder containing .pwb cases")
        if not folder:
            return

        self.folder_path.set(folder)
        self._scan_and_display_folder(folder)

    def _scan_and_display_folder(self, folder: str):
        self.case_tree.delete(*self.case_tree.get_children())
        self.target_cases = {}

        cases, target_cases = scan_folder(folder, self.log)
        self.target_cases = target_cases

        for info in cases:
            tag = "target" if info["is_target"] else ""
            self.case_tree.insert(
                "",
                "end",
                values=(info["filename"], info["type"]),
                tags=(tag,) if tag else (),
            )

    def run_export_folder(self):
        folder = self.folder_path.get()
        if not os.path.isdir(folder):
            messagebox.showwarning(
                "No folder selected", "Please select a valid folder."
            )
            return

        if not self.target_cases:
            messagebox.showwarning(
                "No target cases found",
                "No ACCA_LongTerm / ACCA_P1,2,4,7 / DCwACver_P1-7 cases detected.",
            )
            return

        cats = self._get_row_filter_categories()
        self.log("\n=== Batch processing ACCA/DC cases in folder ===")
        if not cats:
            self.log(
                "WARNING: No LimViolCat categories selected. Row filter will be skipped."
            )

        errors = []
        for label in TARGET_PATTERNS:
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