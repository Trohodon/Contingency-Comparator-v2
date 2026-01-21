import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

from core.case_finder import find_case_files, TARGET_PATTERNS
from core.case_processor import process_case
from core.comparison_builder import build_workbook


class CaseTab(ttk.Frame):
    """
    Case tab:
      - Single folder: pick a folder with 1 .pwb, export, filter, output filtered CSV
      - Multi-folder: pick a root folder with multiple case folders, process all, build combined workbook
    """

    def __init__(self, parent):
        super().__init__(parent)

        # Mode selection
        self.mode_var = tk.StringVar(value="single")

        # Folder paths
        self.single_folder_var = tk.StringVar(value="")
        self.multi_root_var = tk.StringVar(value="")

        # Filter options
        self.expandable_var = tk.BooleanVar(value=True)     # Excel +/- dropdown view
        self.branch_mva_var = tk.BooleanVar(value=True)      # include Branch MVA
        self.bus_lv_var = tk.BooleanVar(value=False)         # include Bus Low Volts
        self.delete_original_var = tk.BooleanVar(value=False)  # delete unfiltered CSV

        # UI
        self._build_ui()

    # ---------------------------
    # UI
    # ---------------------------
    def _build_ui(self):
        top = ttk.Frame(self)
        top.pack(side=tk.TOP, fill=tk.X, padx=10, pady=10)

        # Mode selection
        mode_frame = ttk.LabelFrame(top, text="Mode")
        mode_frame.pack(side=tk.TOP, fill=tk.X)

        ttk.Radiobutton(
            mode_frame, text="Single Folder (one case)", variable=self.mode_var, value="single",
            command=self._on_mode_change
        ).grid(row=0, column=0, sticky="w", padx=5, pady=4)

        ttk.Radiobutton(
            mode_frame, text="Multi Folder (root with multiple cases)", variable=self.mode_var, value="multi",
            command=self._on_mode_change
        ).grid(row=1, column=0, sticky="w", padx=5, pady=4)

        # Single folder picker
        single = ttk.LabelFrame(top, text="Single Folder")
        single.pack(side=tk.TOP, fill=tk.X, pady=(8, 0))

        ttk.Entry(single, textvariable=self.single_folder_var).grid(row=0, column=0, sticky="we", padx=5, pady=5)
        ttk.Button(single, text="Browse...", command=self._browse_single).grid(row=0, column=1, padx=5, pady=5)
        single.columnconfigure(0, weight=1)

        # Multi root picker
        multi = ttk.LabelFrame(top, text="Multi Folder Root")
        multi.pack(side=tk.TOP, fill=tk.X, pady=(8, 0))

        ttk.Entry(multi, textvariable=self.multi_root_var).grid(row=0, column=0, sticky="we", padx=5, pady=5)
        ttk.Button(multi, text="Browse...", command=self._browse_multi).grid(row=0, column=1, padx=5, pady=5)
        multi.columnconfigure(0, weight=1)

        # Filters frame
        filters = ttk.LabelFrame(self, text="Filters")
        filters.pack(side=tk.TOP, fill=tk.X, padx=10, pady=(4, 4))

        ttk.Checkbutton(
            filters,
            text="Expandable issue view (Excel +/- dropdown: show max row, expand to see others)",
            variable=self.expandable_var,
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
        ).grid(row=3, column=0, sticky="w", padx=5, pady=2)

        # Log box
        log_frame = ttk.LabelFrame(self, text="Case Processing Log")
        log_frame.pack(side=tk.TOP, fill=tk.BOTH, expand=True, padx=10, pady=(6, 10))

        self.log_text = tk.Text(log_frame, height=18, wrap="word")
        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        yscroll = ttk.Scrollbar(log_frame, command=self.log_text.yview)
        yscroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.log_text.configure(yscrollcommand=yscroll.set)

        # Run button
        run_frame = ttk.Frame(self)
        run_frame.pack(side=tk.BOTTOM, fill=tk.X, padx=10, pady=(0, 10))

        ttk.Button(run_frame, text="Run", command=self._run).pack(side=tk.RIGHT)

        # Apply initial mode state
        self._on_mode_change()

    def _on_mode_change(self):
        # Nothing to disable currently; placeholder if you want to gray out irrelevant pickers
        pass

    def _browse_single(self):
        folder = filedialog.askdirectory(title="Select folder containing one .pwb case")
        if folder:
            self.single_folder_var.set(folder)

    def _browse_multi(self):
        folder = filedialog.askdirectory(title="Select root folder containing case folders")
        if folder:
            self.multi_root_var.set(folder)

    # ---------------------------
    # Logging
    # ---------------------------
    def log(self, msg: str):
        self.log_text.insert(tk.END, msg + "\n")
        self.log_text.see(tk.END)
        self.update_idletasks()

    # ---------------------------
    # Processing
    # ---------------------------
    def _get_row_filter_categories(self):
        keep = []
        if self.branch_mva_var.get():
            keep.append("Branch MVA")
        if self.bus_lv_var.get():
            keep.append("Bus Low Volts")
        return keep

    def _run(self):
        mode = self.mode_var.get()

        self.log_text.delete("1.0", tk.END)

        keep_categories = self._get_row_filter_categories()
        if not keep_categories:
            messagebox.showwarning("No filters selected", "Please select at least one row category to include.")
            return

        if mode == "single":
            folder = self.single_folder_var.get().strip()
            if not folder or not os.path.isdir(folder):
                messagebox.showerror("Invalid folder", "Please select a valid folder for Single Folder mode.")
                return

            self._run_single_folder(folder, keep_categories)

        else:
            root = self.multi_root_var.get().strip()
            if not root or not os.path.isdir(root):
                messagebox.showerror("Invalid folder", "Please select a valid root folder for Multi Folder mode.")
                return

            self._run_multi_folder(root, keep_categories)

    def _run_single_folder(self, folder: str, keep_categories):
        self.log("Scanning for .pwb files...")
        case_files = find_case_files(folder)
        if not case_files:
            messagebox.showerror("No cases found", "No .pwb files were found in this folder.")
            return

        if len(case_files) > 1:
            self.log(f"Found {len(case_files)} cases. Processing all...")

        for pwb_path in case_files:
            self.log(f"Processing: {pwb_path}")
            ok = process_case(
                pwb_path=pwb_path,
                dedup_enabled=(not self.expandable_var.get()),
                keep_categories=keep_categories,
                delete_original=self.delete_original_var.get(),
                log_func=self.log,
            )
            if ok:
                self.log("Done.\n")
            else:
                self.log("Failed.\n")

        messagebox.showinfo("Complete", "Single folder processing complete.")

    def _run_multi_folder(self, root: str, keep_categories):
        self.log("Scanning for case folders under root...")
        folders = [os.path.join(root, d) for d in os.listdir(root) if os.path.isdir(os.path.join(root, d))]
        if not folders:
            messagebox.showerror("No folders", "No subfolders found under the selected root.")
            return

        folder_to_case_csvs = {}  # {folder_name: {case_type_label: csv_path}}

        for folder in folders:
            folder_name = os.path.basename(folder)
            self.log(f"\n=== Folder: {folder_name} ===")

            case_files = find_case_files(folder)
            if not case_files:
                self.log("No .pwb found. Skipping.")
                continue

            if len(case_files) > 1:
                self.log(f"Found {len(case_files)} .pwb files. Processing all and using matching patterns.")

            # Process each pwb and keep the filtered csv paths keyed by case type
            case_map = {}

            for pwb_path in case_files:
                basename = os.path.basename(pwb_path)

                # Identify which case type this belongs to
                matched_label = None
                for label, patterns in TARGET_PATTERNS.items():
                    if any(pat.lower() in basename.lower() for pat in patterns):
                        matched_label = label
                        break

                if matched_label is None:
                    self.log(f"Skipping (not a target case type): {basename}")
                    continue

                self.log(f"Processing {matched_label}: {basename}")

                ok, filtered_csv_path = process_case(
                    pwb_path=pwb_path,
                    dedup_enabled=(not self.expandable_var.get()),
                    keep_categories=keep_categories,
                    delete_original=self.delete_original_var.get(),
                    log_func=self.log,
                    return_filtered_csv_path=True,
                )

                if ok and filtered_csv_path:
                    case_map[matched_label] = filtered_csv_path

            if case_map:
                folder_to_case_csvs[folder_name] = case_map

        if not folder_to_case_csvs:
            messagebox.showerror("No outputs", "No valid cases were processed. Workbook not created.")
            return

        self.log("\nBuilding combined workbook...")

        workbook_path = build_workbook(
            root,
            folder_to_case_csvs,
            include_branch_mva=self.branch_mva_var.get(),
            include_bus_low_volts=self.bus_lv_var.get(),
            group_details=self.expandable_var.get(),
            log_func=self.log,
        )

        if workbook_path:
            messagebox.showinfo("Complete", f"Workbook created:\n{workbook_path}")
        else:
            messagebox.showerror("Failed", "Workbook build failed. Check the log for details.")