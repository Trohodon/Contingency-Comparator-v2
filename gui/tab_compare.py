# gui/tab_compare.py

import os
import math
import threading
from typing import Optional, List, Tuple

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

from core.comparator import (
    list_sheets,
    build_case_type_comparison,   # still used for live view
    build_batch_comparison_workbook,
)


class CompareTab(ttk.Frame):
    """
    Split-screen-style comparison tab.

    - Open any Combined_ViolationCTG_Comparison.xlsx (or compatible workbook)
    - Choose left/right sheets
    - Set a percent-loading threshold (default 80%)
    - Live view: for each case type (ACCA LongTerm, ACCA, DCwAC) show rows sorted
      highest-to-lowest by loading, with:
         Left %, Right %, Δ% (or 'Only in left/right' when unmatched)
    - Build queue:
         * "Add to queue": add current Left vs Right pair
         * "Delete selected": remove pair from queue
         * "Clear all": remove all queued pairs
         * "Build queued workbook": write a new .xlsx containing one sheet per pair

    NEW:
      - "Expandable issue view (+/-) in batch workbook": groups rows per Resulting Issue.

    IMPORTANT (UI responsiveness):
      - Batch workbook build runs in a background thread so the UI won't go "Not Responding".
      - Logging is thread-safe and marshaled to the Tk UI thread when needed.
    """

    CASE_TYPE_TABS = [
        ("ACCA LongTerm", "ACCA_LongTerm"),
        ("ACCA", "ACCA_P1,2,4,7"),
        ("DCwAC", "DCwACver_P1-7"),
    ]

    def __init__(self, master):
        super().__init__(master)

        self.workbook_path = tk.StringVar(value="No workbook loaded")
        self.left_sheet_var = tk.StringVar()
        self.right_sheet_var = tk.StringVar()

        # Percent loading threshold
        self.threshold_var = tk.StringVar(value="80")

        # NEW: expandable issue view for batch build workbook
        self.expandable_batch_var = tk.BooleanVar(value=True)

        self._sheets: List[str] = []
        self._is_running = False

        self.local_log: Optional[tk.Text] = None
        self.external_log_func = None

        # One Treeview per canonical case type
        self._trees: dict[str, ttk.Treeview] = {}

        # Queue of (left_sheet, right_sheet) pairs
        self._queue: List[Tuple[str, str]] = []
        self._queue_listbox: Optional[tk.Listbox] = None

        self._build_gui()

    # ---------------- Thread-safe UI helpers ---------------- #

    def _ui(self, func, *args, **kwargs):
        """Run func on Tk UI thread."""
        self.after(0, lambda: func(*args, **kwargs))

    def _set_cursor_busy(self, busy: bool):
        """Show a busy cursor while doing long work."""
        try:
            cursor = "watch" if busy else ""
            self.winfo_toplevel().configure(cursor=cursor)
        except Exception:
            pass

    # ---------------- Logging helpers ---------------- #

    def _log_ui(self, msg: str):
        if self.local_log is not None:
            self.local_log.insert(tk.END, msg + "\n")
            self.local_log.see(tk.END)
        if self.external_log_func:
            try:
                self.external_log_func(msg)
            except Exception:
                # external log should never crash the UI
                pass

    def log(self, msg: str):
        # If we're on a background thread, marshal to UI thread.
        if threading.current_thread() is threading.main_thread():
            self._log_ui(msg)
        else:
            self._ui(self._log_ui, msg)

    def _set_running(self, running: bool):
        self._is_running = running
        state = "disabled" if running else "normal"
        self.open_btn.configure(state=state)
        self.compare_btn.configure(state=state)
        self.add_btn.configure(state=state)
        self.build_btn.configure(state=state)
        self.delete_btn.configure(state=state)
        self.clear_all_btn.configure(state=state)

        self._set_cursor_busy(running)

        # Let Tk paint immediately
        self.update_idletasks()

    # ---------------- GUI layout ---------------- #

    def _build_gui(self):
        # Top: workbook selection
        wb_frame = ttk.Frame(self)
        wb_frame.pack(side=tk.TOP, fill=tk.X, padx=10, pady=10)

        self.open_btn = ttk.Button(
            wb_frame, text="Open Excel Workbook", command=self.browse_workbook
        )
        self.open_btn.grid(row=0, column=0, sticky="w", padx=(0, 8))

        ttk.Label(wb_frame, text="Loaded:").grid(row=0, column=1, sticky="w")
        ttk.Label(wb_frame, textvariable=self.workbook_path, width=60).grid(
            row=0, column=2, sticky="w"
        )

        ttk.Label(wb_frame, text="Percent loading threshold:").grid(
            row=0, column=3, sticky="e", padx=(10, 2)
        )
        ttk.Entry(wb_frame, textvariable=self.threshold_var, width=6).grid(
            row=0, column=4, sticky="w"
        )

        # NEW checkbox
        ttk.Checkbutton(
            wb_frame,
            text="Expandable issue view (+/-) in batch workbook",
            variable=self.expandable_batch_var,
        ).grid(row=1, column=2, sticky="w", pady=(6, 0))

        wb_frame.columnconfigure(2, weight=1)

        # Comparison controls + build queue
        cmp_frame = ttk.LabelFrame(self, text="Comparison")
        cmp_frame.pack(side=tk.TOP, fill=tk.X, padx=10, pady=(0, 8))

        ttk.Label(cmp_frame, text="Left sheet:").grid(
            row=0, column=0, sticky="w", padx=5, pady=2
        )
        self.left_combo = ttk.Combobox(
            cmp_frame, textvariable=self.left_sheet_var, state="readonly", width=30
        )
        self.left_combo.grid(row=0, column=1, sticky="w", padx=5, pady=2)

        ttk.Label(cmp_frame, text="Right sheet:").grid(
            row=0, column=2, sticky="w", padx=5, pady=2
        )
        self.right_combo = ttk.Combobox(
            cmp_frame, textvariable=self.right_sheet_var, state="readonly", width=30
        )
        self.right_combo.grid(row=0, column=3, sticky="w", padx=5, pady=2)

        self.add_btn = ttk.Button(cmp_frame, text="Add to queue", command=self.add_to_queue)
        self.add_btn.grid(row=0, column=4, sticky="w", padx=(10, 5), pady=2)

        self.compare_btn = ttk.Button(cmp_frame, text="Compare", command=self.run_comparison)
        self.compare_btn.grid(row=0, column=5, sticky="w", padx=(5, 5), pady=2)

        cmp_frame.columnconfigure(1, weight=1)
        cmp_frame.columnconfigure(3, weight=1)

        ttk.Label(cmp_frame, text="Queued comparisons:").grid(
            row=1, column=0, sticky="nw", padx=5, pady=(4, 4)
        )

        queue_frame = ttk.Frame(cmp_frame)
        queue_frame.grid(row=1, column=1, columnspan=3, sticky="nsew", pady=(4, 4))

        self._queue_listbox = tk.Listbox(queue_frame, height=4)
        self._queue_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        q_scroll = ttk.Scrollbar(queue_frame, orient="vertical", command=self._queue_listbox.yview)
        self._queue_listbox.configure(yscrollcommand=q_scroll.set)
        q_scroll.pack(side=tk.RIGHT, fill=tk.Y)

        self.delete_btn = ttk.Button(
            cmp_frame, text="Delete selected", command=self.delete_selected_queue_item
        )
        self.delete_btn.grid(row=1, column=4, sticky="nw", padx=(10, 5), pady=(4, 4))

        self.clear_all_btn = ttk.Button(cmp_frame, text="Clear all", command=self.clear_all_queue)
        self.clear_all_btn.grid(row=1, column=5, sticky="nw", padx=(5, 5), pady=(4, 4))

        self.build_btn = ttk.Button(
            cmp_frame, text="Build queued workbook", command=self.build_queued_workbook
        )
        self.build_btn.grid(row=2, column=5, sticky="nw", padx=(5, 5), pady=(4, 6))

        cmp_frame.rowconfigure(1, weight=1)
        cmp_frame.columnconfigure(1, weight=1)
        cmp_frame.columnconfigure(3, weight=1)

        # Notebook for live view
        nb = ttk.Notebook(self)
        nb.pack(side=tk.TOP, fill=tk.BOTH, expand=True, padx=10, pady=(0, 8))

        for label, canonical in self.CASE_TYPE_TABS:
            frame = ttk.Frame(nb)
            nb.add(frame, text=label)

            tree = ttk.Treeview(
                frame,
                columns=("cont", "issue", "left", "right", "delta"),
                show="headings",
            )
            self._trees[canonical] = tree

            tree.heading("cont", text="Contingency")
            tree.heading("issue", text="Resulting issue")
            tree.heading("left", text="Left %")
            tree.heading("right", text="Right %")
            tree.heading("delta", text="Δ% (Right - Left) / Status")

            tree.column("cont", width=420, anchor="w")
            tree.column("issue", width=420, anchor="w")
            tree.column("left", width=80, anchor="e")
            tree.column("right", width=80, anchor="e")
            tree.column("delta", width=160, anchor="e")

            vs = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
            tree.configure(yscrollcommand=vs.set)

            tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
            vs.pack(side=tk.RIGHT, fill=tk.Y)

        log_frame = ttk.LabelFrame(self, text="Compare Log")
        log_frame.pack(side=tk.TOP, fill=tk.BOTH, expand=False, padx=10, pady=(0, 10))

        self.local_log = tk.Text(log_frame, wrap="word", height=7)
        self.local_log.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        log_scroll = ttk.Scrollbar(log_frame, orient="vertical", command=self.local_log.yview)
        self.local_log.configure(yscrollcommand=log_scroll.set)
        log_scroll.pack(side=tk.RIGHT, fill=tk.Y)

    # ---------------- Queue helpers ---------------- #

    def add_to_queue(self):
        wb = self.workbook_path.get()
        if not wb.lower().endswith(".xlsx") or not os.path.isfile(wb):
            messagebox.showwarning("No workbook", "Please load a valid .xlsx workbook first.")
            return

        left_sheet = self.left_sheet_var.get()
        right_sheet = self.right_sheet_var.get()
        if not left_sheet or not right_sheet:
            messagebox.showwarning("No sheets selected", "Please select both left and right sheets.")
            return

        pair = (left_sheet, right_sheet)
        self._queue.append(pair)

        display = f"{left_sheet}  vs  {right_sheet}"
        self._queue_listbox.insert(tk.END, display)
        self.log(f"Added to queue: {display}")

    def delete_selected_queue_item(self):
        if not self._queue_listbox:
            return
        sel = list(self._queue_listbox.curselection())
        if not sel:
            return
        for idx in reversed(sel):
            self._queue_listbox.delete(idx)
            if 0 <= idx < len(self._queue):
                removed = self._queue.pop(idx)
                self.log(f"Removed from queue: {removed[0]} vs {removed[1]}")

    def clear_all_queue(self):
        if not self._queue:
            self.log("Queue is already empty.")
            return

        if not messagebox.askyesno("Clear queue", "Clear ALL queued comparisons?"):
            return

        count = len(self._queue)
        self._queue.clear()

        if self._queue_listbox:
            self._queue_listbox.delete(0, tk.END)

        self.log(f"Cleared queue ({count} item{'s' if count != 1 else ''}).")

    def build_queued_workbook(self):
        if self._is_running:
            messagebox.showinfo("Busy", "Another operation is running. Please wait.")
            return

        if not self._queue:
            messagebox.showinfo("Empty queue", "No comparisons in the build queue.")
            return

        wb = self.workbook_path.get()
        if not wb.lower().endswith(".xlsx") or not os.path.isfile(wb):
            messagebox.showwarning("No workbook", "Please load a valid .xlsx workbook first.")
            return

        try:
            thr_raw = self.threshold_var.get().strip()
            threshold = float(thr_raw) if thr_raw else 0.0
            if threshold < 0:
                threshold = 0.0
        except ValueError:
            messagebox.showwarning("Invalid threshold", "Percent loading threshold must be a number (e.g. 80).")
            return

        initial_dir = os.path.dirname(wb) if os.path.dirname(wb) else "."
        save_path = filedialog.asksaveasfilename(
            title="Save batch comparison workbook",
            defaultextension=".xlsx",
            initialdir=initial_dir,
            initialfile="Batch_Comparison.xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
        )
        if not save_path:
            return

        # Snapshot queue so changes during build don't affect the worker
        pairs_snapshot = list(self._queue)
        expandable = self.expandable_batch_var.get()

        self._set_running(True)
        self.log("Building queued workbook...")

        def worker():
            ok = False
            err_msg = None
            try:
                build_batch_comparison_workbook(
                    src_workbook=wb,
                    pairs=pairs_snapshot,
                    threshold=threshold,
                    output_path=save_path,
                    log_func=self.log,  # thread-safe
                    expandable_issue_view=expandable,
                )
                ok = True
            except Exception as e:
                err_msg = str(e)

            def finish_on_ui():
                self._set_running(False)
                if ok:
                    self.log(f"Batch comparison workbook created at:\n{save_path}")
                    messagebox.showinfo("Batch workbook created", f"Batch comparison workbook created at:\n{save_path}")
                else:
                    messagebox.showerror("Error", f"Failed to build workbook:\n{err_msg}")
                    self.log(f"ERROR building batch workbook: {err_msg}")

            self._ui(finish_on_ui)

        threading.Thread(target=worker, daemon=True).start()

    # ---------------- Main compare callbacks ---------------- #

    def browse_workbook(self):
        path = filedialog.askopenfilename(
            title="Select comparison workbook (.xlsx)",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
        )
        if not path:
            return

        if not os.path.isfile(path):
            messagebox.showerror("Error", f"File not found:\n{path}")
            return

        self.workbook_path.set(path)
        self.log(f"Loaded workbook: {path}")

        try:
            self._sheets = list_sheets(path)
        except Exception as e:
            self.log(f"ERROR reading sheet names: {e}")
            messagebox.showerror("Error", str(e))
            return

        if not self._sheets:
            messagebox.showwarning("No sheets", "Workbook has no sheets.")
            return

        self.left_combo["values"] = self._sheets
        self.right_combo["values"] = self._sheets

        self.left_sheet_var.set(self._sheets[0])
        if len(self._sheets) > 1:
            self.right_sheet_var.set(self._sheets[1])
        else:
            self.right_sheet_var.set(self._sheets[0])

    def run_comparison(self):
        if self._is_running:
            messagebox.showinfo("Busy", "A comparison is already running. Please wait.")
            return

        wb = self.workbook_path.get()
        if not wb.lower().endswith(".xlsx") or not os.path.isfile(wb):
            messagebox.showwarning("No workbook", "Please load a valid .xlsx workbook first.")
            return

        left_sheet = self.left_sheet_var.get()
        right_sheet = self.right_sheet_var.get()
        if not left_sheet or not right_sheet:
            messagebox.showwarning("No sheets selected", "Please select both left and right sheets.")
            return

        try:
            thr_raw = self.threshold_var.get().strip()
            threshold = float(thr_raw) if thr_raw else 0.0
            if threshold < 0:
                threshold = 0.0
        except ValueError:
            messagebox.showwarning("Invalid threshold", "Percent loading threshold must be a number (e.g. 80).")
            return

        self.log(
            f"\nComparing sheets:\n"
            f"  Left:  {left_sheet}\n"
            f"  Right: {right_sheet}\n"
            f"  Threshold: {threshold:.2f}% (rows below this on BOTH sides are hidden)"
        )

        self._set_running(True)
        try:
            for label, canonical in self.CASE_TYPE_TABS:
                self.update_idletasks()
                self._compare_one_case_type(
                    wb,
                    left_sheet,
                    right_sheet,
                    canonical,
                    label,
                    threshold,
                )
        finally:
            self._set_running(False)

    # ---------------- Internal helpers ---------------- #

    def _compare_one_case_type(
        self,
        workbook_path: str,
        left_sheet: str,
        right_sheet: str,
        case_type_canonical: str,
        display_label: str,
        threshold: float,
    ):
        tree = self._trees.get(case_type_canonical)
        if tree is None:
            return

        tree.delete(*tree.get_children())

        try:
            df = build_case_type_comparison(
                workbook_path,
                base_sheet=left_sheet,
                new_sheet=right_sheet,
                case_type=case_type_canonical,
                max_rows=None,
                log_func=self.log,
            )
        except Exception as e:
            msg = f"ERROR comparing {display_label}: {e}"
            self.log(msg)
            tree.insert("", "end", values=(msg, "", "", "", ""))
            return

        if df.empty:
            msg = f"No contingencies for {display_label} in either sheet."
            self.log(f"  {msg}")
            tree.insert("", "end", values=(msg, "", "", "", ""))
            return

        self.log(f"  {display_label}: raw rows={len(df)}")

        def is_nan(x) -> bool:
            return isinstance(x, float) and math.isnan(x)

        kept_count = 0

        for _, row in df.iterrows():
            cont = str(row.get("Contingency", "") or "")
            issue = str(row.get("ResultingIssue", "") or "")

            left_pct = row.get("LeftPct", math.nan)
            right_pct = row.get("RightPct", math.nan)
            delta_pct = row.get("DeltaPct", math.nan)

            values = []
            if not is_nan(left_pct):
                values.append(left_pct)
            if not is_nan(right_pct):
                values.append(right_pct)

            if not values:
                continue

            if max(values) < threshold:
                continue

            if is_nan(left_pct) and not is_nan(right_pct):
                delta_text = "Only in right"
            elif not is_nan(left_pct) and is_nan(right_pct):
                delta_text = "Only in left"
            elif is_nan(left_pct) and is_nan(right_pct):
                delta_text = ""
            else:
                try:
                    delta_text = f"{float(delta_pct):.2f}"
                except Exception:
                    delta_text = str(delta_pct)

            def fmt_pct(x):
                if is_nan(x):
                    return ""
                try:
                    return f"{float(x):.2f}"
                except Exception:
                    return str(x)

            tree.insert(
                "",
                "end",
                values=(cont, issue, fmt_pct(left_pct), fmt_pct(right_pct), delta_text),
            )
            kept_count += 1

        self.log(f"  {display_label}: shown rows={kept_count}")
        if kept_count == 0:
            tree.insert("", "end", values=(f"No rows >= {threshold:.2f}%", "", "", "", ""))
