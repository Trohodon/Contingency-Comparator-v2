# gui/tab_compare.py

import os
import math
from typing import Optional, List, Tuple

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

from core.comparator import (
    list_sheets,
    build_case_type_comparison,   # still used for live view
    CASE_TYPES_CANONICAL,
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

    # ---------------- Logging helpers ---------------- #

    def log(self, msg: str):
        if self.local_log is not None:
            self.local_log.insert(tk.END, msg + "\n")
            self.local_log.see(tk.END)
        if self.external_log_func:
            self.external_log_func(msg)

    def _set_running(self, running: bool):
        self._is_running = running
        state = "disabled" if running else "normal"
        self.open_btn.configure(state=state)
        self.compare_btn.configure(state=state)
        self.add_btn.configure(state=state)
        self.build_btn.configure(state=state)
        self.delete_btn.configure(state=state)

        # NEW: clear-all button also follows running state
        self.clear_all_btn.configure(state=state)

        self.update_idletasks()
        self.update()

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
        ttk.Entry(
            wb_frame, textvariable=self.threshold_var, width=6
        ).grid(row=0, column=4, sticky="w")

        wb_frame.columnconfigure(2, weight=1)

        # Comparison controls + build queue
        cmp_frame = ttk.LabelFrame(self, text="Comparison")
        cmp_frame.pack(side=tk.TOP, fill=tk.X, padx=10, pady=(0, 8))

        # Row 0: sheet selection + add/compare buttons
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

        # Add to queue (to the left of Compare)
        self.add_btn = ttk.Button(
            cmp_frame, text="Add to queue", command=self.add_to_queue
        )
        self.add_btn.grid(row=0, column=4, sticky="w", padx=(10, 5), pady=2)

        self.compare_btn = ttk.Button(
            cmp_frame, text="Compare", command=self.run_comparison
        )
        self.compare_btn.grid(row=0, column=5, sticky="w", padx=(5, 5), pady=2)

        cmp_frame.columnconfigure(1, weight=1)
        cmp_frame.columnconfigure(3, weight=1)

        # Row 1: queue list + delete/build buttons
        ttk.Label(cmp_frame, text="Queued comparisons:").grid(
            row=1, column=0, sticky="nw", padx=5, pady=(4, 4)
        )

        queue_frame = ttk.Frame(cmp_frame)
        queue_frame.grid(row=1, column=1, columnspan=3, sticky="nsew", pady=(4, 4))

        self._queue_listbox = tk.Listbox(queue_frame, height=4)
        self._queue_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        q_scroll = ttk.Scrollbar(
            queue_frame, orient="vertical", command=self._queue_listbox.yview
        )
        self._queue_listbox.configure(yscrollcommand=q_scroll.set)
        q_scroll.pack(side=tk.RIGHT, fill=tk.Y)

        # Buttons on the right of queue list
        self.delete_btn = ttk.Button(
            cmp_frame, text="Delete selected", command=self.delete_selected_queue_item
        )
        self.delete_btn.grid(row=1, column=4, sticky="nw", padx=(10, 5), pady=(4, 4))

        # NEW: Clear all button (clears entire queue)
        self.clear_all_btn = ttk.Button(
            cmp_frame, text="Clear all", command=self.clear_all_queue
        )
        self.clear_all_btn.grid(row=1, column=5, sticky="nw", padx=(5, 5), pady=(4, 4))

        # Build button moved to row 2 to make room
        self.build_btn = ttk.Button(
            cmp_frame,
            text="Build queued workbook",
            command=self.build_queued_workbook,
        )
        self.build_btn.grid(row=2, column=5, sticky="nw", padx=(5, 5), pady=(4, 6))

        cmp_frame.rowconfigure(1, weight=1)
        cmp_frame.columnconfigure(1, weight=1)
        cmp_frame.columnconfigure(3, weight=1)

        # Notebook for ACCA LongTerm / ACCA / DCwAC (live view)
        nb = ttk.Notebook(self)
        nb.pack(side=tk.TOP, fill=tk.BOTH, expand=True, padx=10, pady=(0, 8))

        for label, canonical in self.CASE_TYPE_TABS:
            frame = ttk.Frame(nb)
            nb.add(frame, text=label)

            tree = ttk.Treeview(
                frame,
                columns=("cont", "issue", "left", "right", "delta"),
                show=("tree", "headings"),
            )
            self._trees[canonical] = tree
            # Enable tree hierarchy (expand/collapse) so we can show
            # "other contingencies" under the worst row per Resulting Issue.
            tree.column("#0", width=24, stretch=False, anchor="w")
            tree.heading("#0", text="")

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

        # Log area at bottom
        log_frame = ttk.LabelFrame(self, text="Compare Log")
        log_frame.pack(side=tk.TOP, fill=tk.BOTH, expand=False, padx=10, pady=(0, 10))

        self.local_log = tk.Text(log_frame, wrap="word", height=7)
        self.local_log.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        log_scroll = ttk.Scrollbar(
            log_frame, orient="vertical", command=self.local_log.yview
        )
        self.local_log.configure(yscrollcommand=log_scroll.set)
        log_scroll.pack(side=tk.RIGHT, fill=tk.Y)

    # ---------------- Queue helpers ---------------- #

    def add_to_queue(self):
        wb = self.workbook_path.get()
        if not wb.lower().endswith(".xlsx") or not os.path.isfile(wb):
            messagebox.showwarning(
                "No workbook", "Please load a valid .xlsx workbook first."
            )
            return

        left_sheet = self.left_sheet_var.get()
        right_sheet = self.right_sheet_var.get()
        if not left_sheet or not right_sheet:
            messagebox.showwarning(
                "No sheets selected", "Please select both left and right sheets."
            )
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
        # delete from end to start so indices stay valid
        for idx in reversed(sel):
            self._queue_listbox.delete(idx)
            if 0 <= idx < len(self._queue):
                removed = self._queue.pop(idx)
                self.log(f"Removed from queue: {removed[0]} vs {removed[1]}")

    def clear_all_queue(self):
        """
        Clear the entire comparison queue (and the listbox).
        """
        if not self._queue:
            self.log("Queue is already empty.")
            return

        # Optional confirmation - remove these 3 lines if you don't want a popup.
        if not messagebox.askyesno("Clear queue", "Clear ALL queued comparisons?"):
            return

        count = len(self._queue)
        self._queue.clear()

        if self._queue_listbox:
            self._queue_listbox.delete(0, tk.END)

        self.log(f"Cleared queue ({count} item{'s' if count != 1 else ''}).")

    def build_queued_workbook(self):
        if not self._queue:
            messagebox.showinfo("Empty queue", "No comparisons in the build queue.")
            return

        wb = self.workbook_path.get()
        if not wb.lower().endswith(".xlsx") or not os.path.isfile(wb):
            messagebox.showwarning(
                "No workbook", "Please load a valid .xlsx workbook first."
            )
            return

        # Threshold
        try:
            thr_raw = self.threshold_var.get().strip()
            threshold = float(thr_raw) if thr_raw else 0.0
            if threshold < 0:
                threshold = 0.0
        except ValueError:
            messagebox.showwarning(
                "Invalid threshold",
                "Percent loading threshold must be a number (e.g. 80).",
            )
            return

        # Default save folder = folder of source workbook
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

        try:
            self._set_running(True)
            build_batch_comparison_workbook(
                src_workbook=wb,
                pairs=self._queue,
                threshold=threshold,
                output_path=save_path,
                log_func=self.log,
            )
        except Exception as e:
            messagebox.showerror("Error", f"Failed to build workbook:\n{e}")
            self.log(f"ERROR building batch workbook: {e}")
        finally:
            self._set_running(False)

        messagebox.showinfo(
            "Batch workbook created",
            f"Batch comparison workbook created at:\n{save_path}",
        )

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

        # Get sheet names
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

        # Default selection: first as left, second as right (if exists)
        self.left_sheet_var.set(self._sheets[0])
        if len(self._sheets) > 1:
            self.right_sheet_var.set(self._sheets[1])
        else:
            self.right_sheet_var.set(self._sheets[0])

    def run_comparison(self):
        if self._is_running:
            messagebox.showinfo(
                "Busy", "A comparison is already running. Please wait."
            )
            return

        wb = self.workbook_path.get()
        if not wb.lower().endswith(".xlsx") or not os.path.isfile(wb):
            messagebox.showwarning(
                "No workbook", "Please load a valid .xlsx workbook first."
            )
            return

        left_sheet = self.left_sheet_var.get()
        right_sheet = self.right_sheet_var.get()
        if not left_sheet or not right_sheet:
            messagebox.showwarning(
                "No sheets selected", "Please select both left and right sheets."
            )
            return

        # Parse percent threshold
        try:
            thr_raw = self.threshold_var.get().strip()
            threshold = float(thr_raw) if thr_raw else 0.0
            if threshold < 0:
                threshold = 0.0
        except ValueError:
            messagebox.showwarning(
                "Invalid threshold",
                "Percent loading threshold must be a number (e.g. 80).",
            )
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
                self.update()

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
        """
        Build comparison DF for one case type and push it into that tab's Treeview.

        NEW (v2):
          - We still compute the full (CTGLabel, LimViolID) outer-merge just like before.
          - Then we GROUP rows by ResultingIssue (LimViolID):
              * Parent row = the "worst" contingency for that issue
                (prefers highest RightPct, falls back to LeftPct).
              * Child rows = the remaining contingencies for that same issue.
          - Treeview supports expand/collapse via show=("tree","headings").

        Threshold behavior:
          - If a group's *worst* loading (max of Left/Right within the issue)
            is < threshold, the whole group is hidden.
          - Once a group is shown, children are shown even if they are below
            the threshold (because they are the "other contingencies" people asked for).

        Δ% column:
          - If both sides present: numeric Right - Left (2 decimals).
          - Only left present: 'Only in left'
          - Only right present: 'Only in right'
        """
        tree = self._trees.get(case_type_canonical)
        if tree is None:
            return

        # Clear any existing rows
        tree.delete(*tree.get_children())

        try:
            # max_rows=None -> show all; sorting is handled inside comparator
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
            tree.insert("", "end", text="", values=(msg, "", "", "", ""))
            return

        if df.empty:
            msg = f"No contingencies for {display_label} in either sheet."
            self.log(f"  {msg}")
            tree.insert("", "end", text="", values=(msg, "", "", "", ""))
            return

        self.log(f"  {display_label}: raw rows={len(df)}")

        def is_nan(x) -> bool:
            return isinstance(x, float) and math.isnan(x)

        def fmt_pct(x):
            if is_nan(x):
                return ""
            try:
                return f"{float(x):.2f}"
            except Exception:
                return str(x)

        def delta_text(left_pct, right_pct, delta_pct):
            if is_nan(left_pct) and not is_nan(right_pct):
                return "Only in right"
            if not is_nan(left_pct) and is_nan(right_pct):
                return "Only in left"
            if is_nan(left_pct) and is_nan(right_pct):
                return ""
            try:
                return f"{float(delta_pct):.2f}"
            except Exception:
                return str(delta_pct)

        def row_max_lr(lp, rp):
            vals = []
            if not is_nan(lp):
                vals.append(float(lp))
            if not is_nan(rp):
                vals.append(float(rp))
            return max(vals) if vals else float("nan")

        # ----------------------------
        # Group rows by ResultingIssue
        # ----------------------------
        kept_groups = 0
        kept_rows_total = 0

        if "ResultingIssue" not in df.columns:
            # Fallback: old behavior (should never happen)
            for _, row in df.iterrows():
                cont = str(row.get("Contingency", "") or "")
                issue = str(row.get("ResultingIssue", "") or "")

                left_pct = row.get("LeftPct", math.nan)
                right_pct = row.get("RightPct", math.nan)
                d_pct = row.get("DeltaPct", math.nan)

                if row_max_lr(left_pct, right_pct) < threshold:
                    continue

                tree.insert(
                    "",
                    "end",
                    text="",
                    values=(cont, issue, fmt_pct(left_pct), fmt_pct(right_pct), delta_text(left_pct, right_pct, d_pct)),
                )
                kept_rows_total += 1

            self.log(f"  {display_label}: shown rows={kept_rows_total}")
            if kept_rows_total == 0:
                tree.insert("", "end", text="", values=(f"No rows >= {threshold:.2f}%", "", "", "", ""))
            return

        # Build a stable parent ordering: sort groups by worst loading desc
        groups = []
        for issue, g in df.groupby("ResultingIssue", dropna=False):
            # compute group's worst loading (for threshold + ordering)
            g = g.copy()

            g["_RowMax"] = g.apply(
                lambda r: row_max_lr(r.get("LeftPct", math.nan), r.get("RightPct", math.nan)),
                axis=1,
            )
            group_worst = g["_RowMax"].max() if len(g) else float("nan")
            groups.append((issue, g, group_worst))

        # Sort groups by group_worst desc, NaN last
        groups.sort(key=lambda t: (-t[2] if not math.isnan(t[2]) else float("inf")))

        for issue, g, group_worst in groups:
            if math.isnan(group_worst) or group_worst < threshold:
                continue  # hide entire group

            # Pick the "parent" row:
            # Prefer highest RightPct; if all NaN, use LeftPct.
            if g["RightPct"].notna().any():
                parent_idx = g["RightPct"].astype(float).idxmax()
                sort_series = g["RightPct"].astype(float)
            else:
                parent_idx = g["LeftPct"].astype(float).idxmax()
                sort_series = g["LeftPct"].astype(float)

            parent_row = g.loc[parent_idx]

            # Children = all others, sorted high-to-low by same series used above
            child_df = g.drop(index=[parent_idx]).copy()
            # Ensure we don't crash on non-numeric values
            try:
                child_df["_Sort"] = child_df.apply(
                    lambda r: float(r.get("RightPct")) if (g["RightPct"].notna().any() and not is_nan(r.get("RightPct", math.nan)))
                    else (float(r.get("LeftPct")) if not is_nan(r.get("LeftPct", math.nan)) else float("-inf")),
                    axis=1,
                )
                child_df = child_df.sort_values(by="_Sort", ascending=False).drop(columns=["_Sort"], errors="ignore")
            except Exception:
                pass

            # Parent display values
            p_cont = str(parent_row.get("Contingency", "") or "")
            p_issue = str(parent_row.get("ResultingIssue", "") or "")

            p_left = parent_row.get("LeftPct", math.nan)
            p_right = parent_row.get("RightPct", math.nan)
            p_delta = parent_row.get("DeltaPct", math.nan)

            parent_iid = tree.insert(
                "",
                "end",
                text="",
                values=(p_cont, p_issue, fmt_pct(p_left), fmt_pct(p_right), delta_text(p_left, p_right, p_delta)),
                open=False,  # start collapsed; user can expand
            )
            kept_groups += 1
            kept_rows_total += 1

            # Child rows (other contingencies for same issue)
            for _, crow in child_df.iterrows():
                c_cont = str(crow.get("Contingency", "") or "")
                # keep issue blank on children to visually "drop in" under parent
                c_issue = ""

                c_left = crow.get("LeftPct", math.nan)
                c_right = crow.get("RightPct", math.nan)
                c_delta = crow.get("DeltaPct", math.nan)

                tree.insert(
                    parent_iid,
                    "end",
                    text="",
                    values=(c_cont, c_issue, fmt_pct(c_left), fmt_pct(c_right), delta_text(c_left, c_right, c_delta)),
                )
                kept_rows_total += 1

        self.log(f"  {display_label}: shown groups={kept_groups}, total rows (parents+children)={kept_rows_total}")
        if kept_groups == 0:
            tree.insert("", "end", text="", values=(f"No rows >= {threshold:.2f}%", "", "", "", ""))
