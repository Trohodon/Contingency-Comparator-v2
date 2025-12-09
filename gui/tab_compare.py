# gui/tab_compare.py

import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

from core import comparator


class CompareTab(ttk.Frame):
    """
    Tab for comparing two scenario sheets inside a Combined_ViolationCTG_Comparison.xlsx
    (or any workbook with the same formatted layout).
    """

    def __init__(self, master):
        super().__init__(master)

        self.workbook_path = tk.StringVar(value="No comparison workbook selected")
        self.base_sheet_var = tk.StringVar()
        self.new_sheet_var = tk.StringVar()

        # Case type filters
        self.acca_long_var = tk.BooleanVar(value=True)
        self.acca_var = tk.BooleanVar(value=True)
        self.dcwac_var = tk.BooleanVar(value=True)

        # Threshold & mode
        self.pct_threshold_var = tk.StringVar(value="0.0")
        self.mode_var = tk.StringVar(value="all")  # all, worse, better

        self._is_running = False

        self.local_log = None
        self.external_log_func = None  # wired from app.py if desired

        self._sheet_list = []

        self._build_gui()

    # ---------- Logging ---------- #

    def log(self, msg: str):
        if self.local_log is not None:
            self.local_log.insert(tk.END, msg + "\n")
            self.local_log.see(tk.END)
        if self.external_log_func:
            self.external_log_func(msg)

    def _set_running(self, running: bool):
        self._is_running = running
        state = "disabled" if running else "normal"
        self.browse_btn.configure(state=state)
        self.run_btn.configure(state=state)
        self.update_idletasks()
        self.update()

    # ---------- GUI ---------- #

    def _build_gui(self):
        top = ttk.LabelFrame(self, text="Comparison workbook")
        top.pack(side=tk.TOP, fill=tk.X, padx=10, pady=10)

        ttk.Label(top, text="Selected workbook:").grid(row=0, column=0, sticky="w")
        ttk.Label(top, textvariable=self.workbook_path, width=80).grid(
            row=1, column=0, columnspan=2, sticky="w"
        )

        self.browse_btn = ttk.Button(
            top, text="Browse .xlsx…", command=self.browse_workbook
        )
        self.browse_btn.grid(row=1, column=2, padx=(5, 0))

        # Scenario selection
        sc_frame = ttk.LabelFrame(self, text="Scenario selection")
        sc_frame.pack(side=tk.TOP, fill=tk.X, padx=10, pady=5)

        ttk.Label(sc_frame, text="Base scenario sheet:").grid(
            row=0, column=0, sticky="w", padx=5, pady=2
        )
        self.base_combo = ttk.Combobox(
            sc_frame, textvariable=self.base_sheet_var, state="readonly", width=30
        )
        self.base_combo.grid(row=0, column=1, sticky="w", padx=5, pady=2)

        ttk.Label(sc_frame, text="New scenario sheet:").grid(
            row=1, column=0, sticky="w", padx=5, pady=2
        )
        self.new_combo = ttk.Combobox(
            sc_frame, textvariable=self.new_sheet_var, state="readonly", width=30
        )
        self.new_combo.grid(row=1, column=1, sticky="w", padx=5, pady=2)

        # Options
        opt_frame = ttk.LabelFrame(self, text="Options")
        opt_frame.pack(side=tk.TOP, fill=tk.X, padx=10, pady=5)

        # Case types
        case_frame = ttk.Frame(opt_frame)
        case_frame.grid(row=0, column=0, sticky="w", padx=5, pady=2)

        ttk.Label(case_frame, text="Include case types:").grid(
            row=0, column=0, sticky="w"
        )

        ttk.Checkbutton(
            case_frame, text="ACCA LongTerm", variable=self.acca_long_var
        ).grid(row=1, column=0, sticky="w", padx=10)

        ttk.Checkbutton(case_frame, text="ACCA", variable=self.acca_var).grid(
            row=1, column=1, sticky="w", padx=10
        )

        ttk.Checkbutton(
            case_frame, text="DCwAC", variable=self.dcwac_var
        ).grid(row=1, column=2, sticky="w", padx=10)

        # Threshold + mode
        thresh_frame = ttk.Frame(opt_frame)
        thresh_frame.grid(row=1, column=0, sticky="w", padx=5, pady=(8, 2))

        ttk.Label(thresh_frame, text="Δ Percent Loading threshold (|Δ%| ≥):").grid(
            row=0, column=0, sticky="w"
        )
        ttk.Entry(thresh_frame, textvariable=self.pct_threshold_var, width=8).grid(
            row=0, column=1, sticky="w", padx=(3, 10)
        )
        ttk.Label(thresh_frame, text="%").grid(row=0, column=2, sticky="w")

        mode_frame = ttk.Frame(opt_frame)
        mode_frame.grid(row=2, column=0, sticky="w", padx=5, pady=(8, 2))

        ttk.Label(mode_frame, text="Show rows where:").grid(
            row=0, column=0, sticky="w"
        )
        ttk.Radiobutton(
            mode_frame, text="All differences", value="all", variable=self.mode_var
        ).grid(row=1, column=0, sticky="w", padx=10)

        ttk.Radiobutton(
            mode_frame,
            text="New scenario is WORSE (higher %)",
            value="worse",
            variable=self.mode_var,
        ).grid(row=1, column=1, sticky="w", padx=10)

        ttk.Radiobutton(
            mode_frame,
            text="New scenario is BETTER (lower %)",
            value="better",
            variable=self.mode_var,
        ).grid(row=1, column=2, sticky="w", padx=10)

        # Run button
        self.run_btn = ttk.Button(
            self, text="Run comparison", command=self.run_comparison
        )
        self.run_btn.pack(side=tk.TOP, anchor="w", padx=10, pady=(10, 0))

        # Log
        log_frame = ttk.LabelFrame(self, text="Compare Log")
        log_frame.pack(side=tk.TOP, fill=tk.BOTH, expand=True, padx=10, pady=10)

        self.local_log = tk.Text(log_frame, wrap="word", height=10)
        self.local_log.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        log_scroll = ttk.Scrollbar(
            log_frame, orient="vertical", command=self.local_log.yview
        )
        self.local_log.configure(yscrollcommand=log_scroll.set)
        log_scroll.pack(side=tk.RIGHT, fill=tk.Y)

    # ---------- Callbacks ---------- #

    def browse_workbook(self):
        """
        Let the user pick ANY comparison workbook (.xlsx), old or new.
        """
        path = filedialog.askopenfilename(
            title="Select comparison workbook (.xlsx)",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
        )
        if not path:
            return

        self.workbook_path.set(path)
        self.log(f"Selected workbook: {path}")

        # Load sheets
        try:
            self._sheet_list = comparator.list_sheets(path)
        except Exception as e:
            self.log(f"ERROR reading workbook sheets: {e}")
            messagebox.showerror("Error", str(e))
            return

        if not self._sheet_list:
            messagebox.showwarning("No sheets", "No sheets found in this workbook.")
            return

        self.base_combo["values"] = self._sheet_list
        self.new_combo["values"] = self._sheet_list

        # Auto-select first for base, second for new if available
        self.base_sheet_var.set(self._sheet_list[0])
        if len(self._sheet_list) > 1:
            self.new_sheet_var.set(self._sheet_list[1])
        else:
            self.new_sheet_var.set(self._sheet_list[0])

    def _get_case_type_filter(self):
        case_types = []
        if self.acca_long_var.get():
            case_types.append("ACCA_LongTerm")
        if self.acca_var.get():
            case_types.append("ACCA_P1,2,4,7")
        if self.dcwac_var.get():
            case_types.append("DCwACver_P1-7")
        return case_types

    def run_comparison(self):
        if self._is_running:
            messagebox.showinfo(
                "Busy", "A comparison is already running. Please wait for it to finish."
            )
            return

        path = self.workbook_path.get()
        if not path.lower().endswith(".xlsx") or not os.path.isfile(path):
            messagebox.showwarning(
                "No workbook selected", "Please select a valid .xlsx workbook."
            )
            return

        base_sheet = self.base_sheet_var.get()
        new_sheet = self.new_sheet_var.get()
        if not base_sheet or not new_sheet:
            messagebox.showwarning(
                "No sheets selected",
                "Please select both base and new scenario sheets.",
            )
            return

        if base_sheet == new_sheet:
            if not messagebox.askyesno(
                "Same sheet",
                "Base and New sheets are the same. Compare sheet against itself?",
            ):
                return

        # Threshold parsing
        try:
            thresh = float(self.pct_threshold_var.get())
            if thresh < 0:
                thresh = -thresh
        except ValueError:
            messagebox.showwarning(
                "Invalid threshold", "Threshold must be a number (e.g., 2 or 2.5)."
            )
            return

        case_types = self._get_case_type_filter()
        if not case_types:
            if not messagebox.askyesno(
                "No case types selected",
                "No case types selected; comparison will consider ALL case types "
                "found in the sheets.\n\nContinue?",
            ):
                return
            case_types = None  # means no explicit filter

        mode = self.mode_var.get()

        self.log(
            f"\nRunning comparison:\n  Base sheet: {base_sheet}\n  New sheet:  {new_sheet}"
        )

        self._set_running(True)
        try:
            # Let UI breathe
            self.update_idletasks()
            self.update()

            result_path = comparator.compare_scenarios(
                workbook_path=path,
                base_sheet=base_sheet,
                new_sheet=new_sheet,
                case_types_to_include=case_types,
                pct_threshold=thresh,
                mode=mode,
                log_func=self.log,
            )
        except Exception as e:
            self.log(f"ERROR during comparison: {e}")
            messagebox.showerror("Error", str(e))
        else:
            if result_path:
                messagebox.showinfo(
                    "Comparison complete",
                    f"Comparison sheet has been added to:\n{result_path}",
                )
        finally:
            self._set_running(False)