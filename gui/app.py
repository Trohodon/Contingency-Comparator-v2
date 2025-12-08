# gui/app.py

import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

import pandas as pd

from .header_filter_dialog import HeaderFilterDialog
from core.pwb_exporter import export_violation_ctg
from core.column_blacklist import apply_blacklist, apply_row_filter


class PwbExportApp(tk.Tk):
    def __init__(self):
        super().__init__()

        self.title("PowerWorld Contingency Violations Export (ViolationCTG)")
        self.geometry("900x550")

        self.pwb_path = tk.StringVar(value="No .pwb file selected")
        self.csv_path = None

        self._build_gui()

    # ───────────── GUI LAYOUT ───────────── #

    def _build_gui(self):
        top = ttk.Frame(self)
        top.pack(side=tk.TOP, fill=tk.X, padx=10, pady=10)

        ttk.Label(top, text="Selected .pwb case:").grid(row=0, column=0, sticky="w")
        ttk.Label(top, textvariable=self.pwb_path, width=80).grid(
            row=1, column=0, columnspan=2, sticky="w"
        )

        browse_btn = ttk.Button(top, text="Browse…", command=self.browse_pwb)
        browse_btn.grid(row=1, column=2, padx=(5, 0), sticky="e")

        run_btn = ttk.Button(
            top,
            text="Export existing contingency violations (ViolationCTG)",
            command=self.run_export,
        )
        run_btn.grid(row=2, column=0, columnspan=3, pady=(10, 0), sticky="w")

        ttk.Separator(self, orient="horizontal").pack(fill=tk.X, padx=10, pady=5)

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

    # ───────────── CALLBACKS ───────────── #

    def browse_pwb(self):
        path = filedialog.askopenfilename(
            title="Select PowerWorld case (.pwb)",
            filetypes=[("PowerWorld case", "*.pwb"), ("All files", "*.*")],
        )
        if path:
            self.pwb_path.set(path)
            self.csv_path = None
            self.log(f"Selected case: {path}")

    def run_export(self):
        pwb = self.pwb_path.get()
        if not pwb.lower().endswith(".pwb") or not os.path.exists(pwb):
            messagebox.showwarning("No case selected", "Please select a valid .pwb file.")
            return

        try:
            # Call into core logic to export ViolationCTG to CSV
            csv_out = export_violation_ctg(pwb, self.log)
            self.csv_path = csv_out
        except Exception as e:
            self.log(f"ERROR: {e}")
            messagebox.showerror("Error", str(e))
            return

        # Treat the CSV as our "temporary Excel sheet"
        if self.csv_path and os.path.exists(self.csv_path):
            self._post_process_csv(self.csv_path)
        else:
            self.log("WARNING: CSV file does not exist after export.")

        messagebox.showinfo("Done", f"ViolationCTG exported to:\n{self.csv_path}")

    # ───────────── HELPERS ───────────── #

    @staticmethod
    def _make_filtered_path(original_csv: str) -> str:
        base, ext = os.path.splitext(original_csv)
        if not ext:
            ext = ".csv"
        return f"{base}_Filtered{ext}"

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
        - Show header filter dialog and log filtered headers.
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
            self.log("Header names:")
            for h in header_row:
                self.log(f"  - {h}")

            if raw.shape[0] > 1:
                # Data rows are index >= 1
                data = raw.iloc[1:].copy()
                data.columns = header_row

                # 1) Apply row filter FIRST (uses LimViolCat before we drop it)
                self.log("\nApplying row filter (e.g., LimViolCat == 'Branch MVA')...")
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

            # Still open dialog so you can manually explore headers and
            # decide which ones to add to the blacklist file.
            self.log(
                "\nOpening header filter dialog so you can choose which\n"
                "columns should be filtered out in future versions..."
            )
            HeaderFilterDialog(self, header_row, self.log)

        except Exception as e:
            self.log(f"(Could not read CSV for header inspection: {e})")