# gui/app.py

import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

import pandas as pd

from core.pwb_exporter import export_violation_ctg
from core.column_blacklist import apply_blacklist, apply_row_filter


# Map human-friendly labels to filename substrings we look for
TARGET_PATTERNS = {
    "ACCA_LongTerm": "ACCA_LongTerm",
    "ACCA_P1,2,4,7": "ACCA_P1,2,4,7",
    "DCwACver_P1-7": "DCwACver_P1-7",
}


class PwbExportApp(tk.Tk):
    def __init__(self):
        super().__init__()

        self.title("PowerWorld Contingency Violations Export (ViolationCTG)")
        self.geometry("1050x650")

        self.pwb_path = tk.StringVar(value="No .pwb file selected")
        self.folder_path = tk.StringVar(value="No folder selected")

        # For folder mode: label -> full path
        self.target_cases = {}

        self._build_gui()

    # ───────────── GUI LAYOUT ───────────── #

    def _build_gui(self):
        # Top frame: single-case controls
        top = ttk.LabelFrame(self, text="Single case processing")
        top.pack(side=tk.TOP, fill=tk.X, padx=10, pady=10)

        ttk.Label(top, text="Selected .pwb case:").grid(row=0, column=0, sticky="w")
        ttk.Label(top, textvariable=self.pwb_path, width=80).grid(
            row=1, column=0, columnspan=2, sticky="w"
        )

        browse_btn = ttk.Button(top, text="Browse .pwb…", command=self.browse_pwb)
        browse_btn.grid(row=1, column=2, padx=(5, 0), sticky="e")

        run_btn = ttk.Button(
            top,
            text="Process selected .pwb (export + filter)",
            command=self.run_export_single,
        )
        run_btn.grid(row=2, column=0, columnspan=3, pady=(10, 0), sticky="w")

        # Folder frame: folder selection + tree
        folder_frame = ttk.LabelFrame(self, text="Folder processing (3 ACCA/DC cases)")
        folder_frame.pack(side=tk.TOP, fill=tk.BOTH, expand=False, padx=10, pady=5)

        ttk.Label(folder_frame, text="Selected folder:").grid(
            row=0, column=0, sticky="w"
        )
        ttk.Label(folder_frame, textvariable=self.folder_path, width=80).grid(
            row=1, column=0, columnspan=2, sticky="w"
        )

        browse_folder_btn = ttk.Button(
            folder_frame, text="Browse folder…", command=self.browse_folder
        )
        browse_folder_btn.grid(row=1, column=2, padx=(5, 0), sticky="e")

        process_folder_btn = ttk.Button(
            folder_frame,
            text="Process 3 ACCA/DC cases in folder",
            command=self.run_export_folder,
        )
        process_folder_btn.grid(row=2, column=0, columnspan=3, pady=(8, 0), sticky="w")

        # Tree showing .pwb files in selected folder
        tree_frame = ttk.Frame(folder_frame)
        tree_frame.grid(row=3, column=0, columnspan=3, sticky="nsew", pady=(8, 0))
        folder_frame.rowconfigure(3, weight=1)
        folder_frame.columnconfigure(0, weight=1)

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

        tree_scroll = ttk.Scrollbar(
            tree_frame, orient="vertical", command=self.case_tree.yview
        )
        self.case_tree.configure(yscrollcommand=tree_scroll.set)

        self.case_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        tree_scroll.pack(side=tk.RIGHT, fill=tk.Y)

        # Tag for the 3 important cases so they stand out
        self.case_tree.tag_configure("target", foreground="blue")

        # Log / output area
        log_frame = ttk.Frame(self)
        log_frame.pack(side=tk.TOP, fill=tk.BOTH, expand=True, padx=10, pady=10)

        ttk.Label(log_frame, text="Log:").pack(anchor="w")

        self.log_text = tk.Text(log_frame, wrap="word", height=18)
        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        scroll = ttk.Scrollbar(
            log_frame, orient="vertical", command=self.log_text.yview
        )
        scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.log_text.configure(yscrollcommand=scroll.set)

    def log(self, msg: str):
        """Append a line to the GUI log."""
        self.log_text.insert(tk.END, msg + "\n")
        self.log_text.see(tk.END)
        self.update_idletasks()

    # ───────────── CALLBACKS: SINGLE CASE ───────────── #

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

        try:
            self._process_single_case(pwb)
        except Exception as e:
            self.log(f"ERROR: {e}")
            messagebox.showerror("Error", str(e))
            return

        messagebox.showinfo("Done", f"Processing complete for:\n{pwb}")

    # ───────────── CALLBACKS: FOLDER MODE ───────────── #

    def browse_folder(self):
        folder = filedialog.askdirectory(title="Select folder containing .pwb cases")
        if not folder:
            return

        self.folder_path.set(folder)
        self._scan_folder(folder)

    def _scan_folder(self, folder: str):
        """Populate the tree with .pwb files and mark the 3 important ones."""
        self.case_tree.delete(*self.case_tree.get_children())
        self.target_cases = {}

        self.log(f"\nScanning folder for .pwb cases:\n{folder}")

        pwb_files = [f for f in os.listdir(folder) if f.lower().endswith(".pwb")]
        if not pwb_files:
            self.log("No .pwb files found in folder.")
            return

        for fname in sorted(pwb_files):
            fpath = os.path.join(folder, fname)

            ctype = "Other"
            tag = ""
            for label, pattern in TARGET_PATTERNS.items():
                if pattern in fname:
                    ctype = label
                    tag = "target"
                    # If there are duplicates for a type, keep the first and log it
                    if label not in self.target_cases:
                        self.target_cases[label] = fpath
                    else:
                        self.log(
                            f"WARNING: Multiple cases found for type '{label}'. "
                            f"Using first: {self.target_cases[label]}"
                        )
                    break

            self.case_tree.insert(
                "", "end", values=(fname, ctype), tags=(tag,) if tag else ()
            )

        self.log("Folder scan complete.")
        for label in TARGET_PATTERNS:
            if label in self.target_cases:
                self.log(f"  Found target case [{label}]: {self.target_cases[label]}")
            else:
                self.log(f"  WARNING: No case found for type [{label}] in this folder.")

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

        self.log("\n=== Batch processing 3 ACCA/DC cases in folder ===")

        errors = []
        for label in TARGET_PATTERNS:
            pwb_path = self.target_cases.get(label)
            if not pwb_path:
                self.log(f"Skipping type [{label}] (not found).")
                continue

            self.log(f"\n--- Processing [{label}] case ---")
            self.log(f"Case path: {pwb_path}")
            try:
                self._process_single_case(pwb_path)
            except Exception as e:
                err_msg = f"ERROR processing [{label}] case: {e}"
                self.log(err_msg)
                errors.append(err_msg)

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

    # ───────────── CORE PROCESSING HELPERS ───────────── #

    @staticmethod
    def _make_filtered_path(original_csv: str) -> str:
        base, ext = os.path.splitext(original_csv)
        if not ext:
            ext = ".csv"
        return f"{base}_Filtered{ext}"

    def _process_single_case(self, pwb_path: str):
        """
        For a single .pwb:
        - Export ViolationCTG to CSV via SimAuto
        - Read CSV, apply row filter & column blacklist
        - Write filtered CSV
        """
        self.log("\nConnecting to PowerWorld and exporting ViolationCTG...")
        csv_out = export_violation_ctg(pwb_path, self.log)

        self.log(f"Exported CSV path: {csv_out}")

        # Now post-process CSV as our "temporary Excel sheet"
        self._post_process_csv(csv_out)

    # ───────────── CSV / HEADER HANDLING ───────────── #

    def _post_process_csv(self, csv_path: str):
        """
        After export, read the CSV and:
        - Skip row 1 (the single 'ViolationCTG' cell)
        - Use row 2 as headers
        - Treat row 3+ as data
        - FIRST apply row filter (e.g., LimViolCat == 'Branch MVA')
        - THEN apply column blacklist
        - Save a new filtered CSV
        """
        self.log("\nReading CSV to detect headers...")
        try:
            # Skip the first row because it only has "ViolationCTG" in one column.
            # After skiprows=1:
            #   raw.iloc[0] -> original row 2 (the real header row)
            #   raw.iloc[1:] -> original rows 3+ (data)
            raw = pd.read_csv(csv_path, header=None, skiprows=1)

            if raw.shape[0] < 1:
                self.log("Not enough rows in CSV to extract headers (need at least 1).")
                return

            # First row in 'raw' is now the header row
            header_row = list(raw.iloc[0])
            self.log(f"Detected {len(header_row)} headers from row 2.")

            if raw.shape[0] > 1:
                # Data rows are index >= 1
                data = raw.iloc[1:].copy()
                data.columns = header_row

                # 1) Apply row filter FIRST (uses LimViolCat before we drop it)
                self.log("\nApplying row filter (e.g., only keep LimViolCat == 'Branch MVA')...")
                filtered_data, removed_rows = apply_row_filter(data, self.log)
                self.log(f"Rows removed by row filter: {removed_rows}")

                # 2) Apply column blacklist
                self.log("\nApplying column blacklist...")
                filtered_data, removed_cols = apply_blacklist(filtered_data)

                if removed_cols:
                    self.log("Columns removed by blacklist:")
                    for c in removed_cols:
                        self.log(f"  - {c}")
                else:
                    self.log("No columns matched blacklist; no columns removed.")

                # Save filtered CSV
                filtered_csv = self._make_filtered_path(csv_path)
                filtered_data.to_csv(filtered_csv, index=False)
                self.log(f"Filtered CSV saved to:\n  {filtered_csv}")

                # Preview first few rows of filtered data
                self.log("\nPreview of first few filtered data rows:")
                preview = filtered_data.head(10).to_string(index=False)
                self.log(preview)
            else:
                self.log("No data rows found after header row; nothing to filter.")

        except Exception as e:
            self.log(f"(Could not read CSV for header inspection: {e})")