# gui/trends_view.py
from __future__ import annotations

import os
import threading
import queue
from dataclasses import dataclass
from typing import Dict, Optional, List, Tuple

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk
from matplotlib.figure import Figure

from core.trends import build_trends, CATEGORY_LABELS, TrendResult


@dataclass
class _BuildRequest:
    workbook_path: str
    category_key: str
    min_percent: float
    top_n: Optional[int]


class TrendsView(ttk.Frame):
    def __init__(self, master, log_callback=None):
        super().__init__(master)
        self._log_callback = log_callback

        self.workbook_path: Optional[str] = None
        self._trend_result: Optional[TrendResult] = None
        self._colors: Dict[str, str] = {}  # issue_key -> color (matplotlib color string)

        self._build_thread: Optional[threading.Thread] = None
        self._q: "queue.Queue[Tuple[str, object]]" = queue.Queue()

        self._selected_category = tk.StringVar(value="ACCA")  # category_key
        self._min_percent_var = tk.StringVar(value="80")
        self._top_n_var = tk.StringVar(value="")  # blank = all
        self._search_var = tk.StringVar(value="")

        self._init_ui()
        self._poll_queue()

    # ---------------- UI ----------------
    def _init_ui(self):
        # Top bar
        top = ttk.Frame(self)
        top.pack(fill="x", padx=10, pady=8)

        ttk.Button(top, text="Open Combined Workbook", command=self._open_workbook).pack(side="left")
        self._loaded_label = ttk.Label(top, text="Loaded: No workbook loaded")
        self._loaded_label.pack(side="left", padx=(10, 0))

        # Controls row
        ctrl = ttk.Frame(self)
        ctrl.pack(fill="x", padx=10, pady=(0, 8))

        ttk.Label(ctrl, text="Min % filter:").pack(side="left")
        ttk.Entry(ctrl, width=8, textvariable=self._min_percent_var).pack(side="left", padx=(6, 14))

        ttk.Label(ctrl, text="Top N:").pack(side="left")
        ttk.Entry(ctrl, width=8, textvariable=self._top_n_var).pack(side="left", padx=(6, 14))
        ttk.Label(ctrl, text="(blank = all)").pack(side="left", padx=(0, 14))

        ttk.Button(ctrl, text="Build Trend Data", command=self._build_clicked).pack(side="left")

        # Category buttons (same idea as Compare Cases)
        cat = ttk.Frame(self)
        cat.pack(fill="x", padx=10, pady=(0, 8))

        ttk.Label(cat, text="Category:").pack(side="left")
        for key, label in CATEGORY_LABELS.items():
            ttk.Radiobutton(
                cat,
                text=label,
                value=key,
                variable=self._selected_category,
                command=self._category_changed,
            ).pack(side="left", padx=8)

        # Main split
        main = ttk.PanedWindow(self, orient="horizontal")
        main.pack(fill="both", expand=True, padx=10, pady=8)

        left = ttk.Frame(main)
        right = ttk.Frame(main)
        main.add(left, weight=2)
        main.add(right, weight=3)

        # Left: search + issue list
        search_row = ttk.Frame(left)
        search_row.pack(fill="x", pady=(0, 6))
        ttk.Label(search_row, text="Search:").pack(side="left")
        e = ttk.Entry(search_row, textvariable=self._search_var)
        e.pack(side="left", fill="x", expand=True, padx=(6, 0))
        e.bind("<KeyRelease>", lambda _e: self._refresh_issue_list())

        self._issue_tree = ttk.Treeview(
            left,
            columns=("max", "count"),
            show="headings",
            selectmode="extended",
            height=18,
        )
        self._issue_tree.heading("max", text="Max %")
        self._issue_tree.heading("count", text="# Cases")
        self._issue_tree["displaycolumns"] = ("max", "count")

        # We also want the issue string visible; easiest is to use #0 "tree" column
        self._issue_tree.configure(show="tree headings")
        self._issue_tree.heading("#0", text="Resulting Issue")
        self._issue_tree.column("#0", width=380, anchor="w")
        self._issue_tree.column("max", width=70, anchor="e")
        self._issue_tree.column("count", width=70, anchor="e")

        yscroll = ttk.Scrollbar(left, orient="vertical", command=self._issue_tree.yview)
        self._issue_tree.configure(yscrollcommand=yscroll.set)

        self._issue_tree.pack(side="left", fill="both", expand=True)
        yscroll.pack(side="left", fill="y")

        btn_row = ttk.Frame(left)
        btn_row.pack(fill="x", pady=6)
        ttk.Button(btn_row, text="Plot Selected", command=self._plot_selected).pack(side="left")
        ttk.Button(btn_row, text="Clear Plot", command=self._clear_plot).pack(side="left", padx=8)

        # Right: plot
        self._fig = Figure(figsize=(6, 4), dpi=100)
        self._ax = self._fig.add_subplot(111)
        self._ax.set_title("Percent Loading Trend")
        self._ax.set_xlabel("Case / Sheet")
        self._ax.set_ylabel("Percent Loading")

        self._canvas = FigureCanvasTkAgg(self._fig, master=right)
        self._canvas.get_tk_widget().pack(fill="both", expand=True)

        toolbar = NavigationToolbar2Tk(self._canvas, right)
        toolbar.update()

        # Log
        log_frame = ttk.LabelFrame(self, text="Trends Log")
        log_frame.pack(fill="both", expand=False, padx=10, pady=(0, 10))

        self._log = tk.Text(log_frame, height=6, wrap="none")
        self._log.pack(fill="both", expand=True)

    # ---------------- Logging ----------------
    def _log_line(self, s: str):
        self._log.insert("end", s + "\n")
        self._log.see("end")
        if self._log_callback:
            try:
                self._log_callback(s)
            except Exception:
                pass

    # ---------------- Workbook ----------------
    def _open_workbook(self):
        path = filedialog.askopenfilename(
            title="Select combined violation workbook",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
        )
        if not path:
            return
        self.workbook_path = path
        self._loaded_label.config(text=f"Loaded: {os.path.basename(path)}")
        self._log_line(f"Loaded workbook: {path}")
        # Reset state
        self._trend_result = None
        self._colors.clear()
        self._issue_tree.delete(*self._issue_tree.get_children())
        self._clear_plot()

    # ---------------- Build data (threaded) ----------------
    def _build_clicked(self):
        if not self.workbook_path:
            messagebox.showwarning("No workbook", "Please open a combined workbook first.")
            return

        # parse controls
        try:
            min_pct = float(self._min_percent_var.get().strip() or "80")
        except Exception:
            messagebox.showwarning("Invalid min %", "Min % must be a number (ex: 80).")
            return

        top_n_txt = self._top_n_var.get().strip()
        top_n = None
        if top_n_txt:
            try:
                top_n = int(top_n_txt)
            except Exception:
                messagebox.showwarning("Invalid Top N", "Top N must be blank or an integer.")
                return

        req = _BuildRequest(
            workbook_path=self.workbook_path,
            category_key=self._selected_category.get(),
            min_percent=min_pct,
            top_n=top_n,
        )

        if self._build_thread and self._build_thread.is_alive():
            messagebox.showinfo("Working", "Already building trend data. Please wait.")
            return

        self._log_line(f"Building trend data: {CATEGORY_LABELS[req.category_key]} (min {req.min_percent}%) ...")
        self._build_thread = threading.Thread(target=self._build_worker, args=(req,), daemon=True)
        self._build_thread.start()

    def _build_worker(self, req: _BuildRequest):
        try:
            res = build_trends(req.workbook_path, req.category_key, min_percent=req.min_percent)

            # Apply Top N (by max %)
            if req.top_n is not None:
                items = list(res.issues.values())
                items.sort(key=lambda it: (it.max_value if it.max_value is not None else -1), reverse=True)
                keep = {it.issue_key for it in items[: req.top_n]}
                res.issues = {k: v for k, v in res.issues.items() if k in keep}

            self._q.put(("result", res))
        except Exception as e:
            self._q.put(("error", e))

    def _poll_queue(self):
        try:
            while True:
                msg, payload = self._q.get_nowait()
                if msg == "result":
                    self._trend_result = payload  # type: ignore
                    self._on_new_trend_result()
                elif msg == "error":
                    self._log_line(f"ERROR: {payload}")
                    messagebox.showerror("Trend build failed", str(payload))
        except queue.Empty:
            pass
        self.after(100, self._poll_queue)

    # ---------------- Category switching ----------------
    def _category_changed(self):
        # Switching category should rebuild (or prompt). We’ll just clear and wait for user to click build.
        self._trend_result = None
        self._colors.clear()
        self._issue_tree.delete(*self._issue_tree.get_children())
        self._clear_plot()
        self._log_line(f"Selected category: {CATEGORY_LABELS[self._selected_category.get()]} (click Build Trend Data)")

    # ---------------- Results -> list + colors ----------------
    def _on_new_trend_result(self):
        assert self._trend_result is not None
        self._log_line(f"Found {len(self._trend_result.issues)} issues across {len(self._trend_result.sheet_order)} sheets.")
        self._assign_colors()
        self._refresh_issue_list()
        # Auto-plot top 5 by default (nice demo)
        self._auto_plot_top(5)

    def _assign_colors(self):
        """
        Stable colors per issue key, based on sorted order.
        Uses matplotlib default color cycle.
        """
        if not self._trend_result:
            return
        keys = sorted(self._trend_result.issues.keys())
        # Matplotlib default cycle
        cycle = self._ax._get_lines.prop_cycler
        colors = []
        # Pull enough colors
        for _ in range(max(1, len(keys))):
            colors.append(next(cycle)["color"])
        # Map (repeat if needed)
        self._colors = {}
        for i, k in enumerate(keys):
            self._colors[k] = colors[i % len(colors)]

    def _refresh_issue_list(self):
        self._issue_tree.delete(*self._issue_tree.get_children())
        if not self._trend_result:
            return

        q = self._search_var.get().strip().upper()

        items = list(self._trend_result.issues.values())
        # Sort by max descending
        items.sort(key=lambda it: (it.max_value if it.max_value is not None else -1), reverse=True)

        for it in items:
            if q and (q not in it.issue_display.upper() and q not in it.issue_key):
                continue
            mx = it.max_value
            mx_txt = f"{mx:.2f}" if mx is not None else ""
            cnt_txt = str(it.count_present)
            self._issue_tree.insert("", "end", iid=it.issue_key, text=it.issue_display, values=(mx_txt, cnt_txt))

    # ---------------- Plotting ----------------
    def _clear_plot(self):
        self._ax.clear()
        self._ax.set_title("Percent Loading Trend")
        self._ax.set_xlabel("Case / Sheet")
        self._ax.set_ylabel("Percent Loading")
        self._canvas.draw_idle()

    def _auto_plot_top(self, n: int):
        if not self._trend_result:
            return
        items = list(self._trend_result.issues.values())
        items.sort(key=lambda it: (it.max_value if it.max_value is not None else -1), reverse=True)
        top = items[:n]
        if not top:
            return
        # select them in tree
        self._issue_tree.selection_set([it.issue_key for it in top if it.issue_key in self._issue_tree.get_children("")])
        self._plot_issues([it.issue_key for it in top])

    def _plot_selected(self):
        sel = list(self._issue_tree.selection())
        if not sel:
            messagebox.showinfo("Select issues", "Select one or more Resulting Issues to plot.")
            return
        self._plot_issues(sel)

    def _plot_issues(self, issue_keys: List[str]):
        if not self._trend_result:
            return

        sheet_order = self._trend_result.sheet_order
        x = list(range(len(sheet_order)))
        xlabels = sheet_order

        self._ax.clear()
        self._ax.set_title(f"Percent Loading Trend — {CATEGORY_LABELS[self._trend_result.category]}")
        self._ax.set_xlabel("Case / Sheet")
        self._ax.set_ylabel("Percent Loading")

        # To keep labels readable, show fewer ticks if lots of sheets
        if len(xlabels) <= 15:
            self._ax.set_xticks(x)
            self._ax.set_xticklabels(xlabels, rotation=30, ha="right")
        else:
            # Show every ~N
            step = max(1, len(xlabels) // 10)
            ticks = list(range(0, len(xlabels), step))
            self._ax.set_xticks(ticks)
            self._ax.set_xticklabels([xlabels[i] for i in ticks], rotation=30, ha="right")

        for k in issue_keys:
            it = self._trend_result.issues.get(k)
            if not it:
                continue
            y = []
            for s in sheet_order:
                y.append(it.per_sheet.get(s, None))

            # Convert None -> gap (matplotlib will break line if we use NaN)
            y2 = [float("nan") if v is None else float(v) for v in y]
            color = self._colors.get(k, None)
            label = it.issue_display
            self._ax.plot(x, y2, marker="o", label=label, color=color)

        self._ax.legend(loc="best", fontsize=8)
        self._ax.grid(True, which="both", axis="y", linestyle="--", alpha=0.4)
        self._fig.tight_layout()
        self._canvas.draw_idle()