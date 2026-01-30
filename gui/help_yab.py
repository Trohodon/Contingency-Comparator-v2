# gui/help_view.py

import tkinter as tk
from tkinter import ttk, messagebox

from core.help_search import rank_topics, probe
from core.menu_one_runner import launch_menu_one_detached


class HelpTab(ttk.Frame):
    """
    Professional Help / Documentation tab.

    Left = topic navigation + search
    Right = styled content

    Search behavior:
      - Type in search box -> ranks topics by relevance (title + content)
      - Press Enter -> opens top match
      - Secret trigger: "Menu One" opens Menu One easter egg (single-instance)
    """

    # Palette
    NAVY = "#0B2F5B"
    NAVY_2 = "#103A6B"
    CARD_BG = "#FFFFFF"
    CALLOUT_BG = "#EAF2FF"
    CODE_BG = "#F2F2F2"
    MUTED = "#5C6773"
    TEXT = "#1F2A44"
    DIVIDER = "#D6DFEA"
    HILITE_BG = "#FFF2A8"

    def __init__(self, master):
        super().__init__(master)
        self._current_topic = "Overview"
        self._topics_master = []
        self._menu_launching = False  # prevents spam in this GUI process
        self._build_gui()

    # ---------------- GUI ---------------- #

    def _build_gui(self):
        outer = ttk.Frame(self)
        outer.pack(fill=tk.BOTH, expand=True, padx=12, pady=12)
        outer.columnconfigure(0, weight=0)
        outer.columnconfigure(1, weight=1)
        outer.rowconfigure(0, weight=1)

        # LEFT: nav "card"
        nav_card = tk.Frame(outer, bg=self.CARD_BG, bd=1, relief="solid")
        nav_card.grid(row=0, column=0, sticky="nsw", padx=(0, 12))
        nav_card.grid_propagate(False)
        nav_card.configure(width=260)

        nav_header = tk.Frame(nav_card, bg=self.NAVY)
        nav_header.pack(fill=tk.X)

        tk.Label(
            nav_header,
            text="Help Topics",
            bg=self.NAVY,
            fg="white",
            font=("Segoe UI", 11, "bold"),
            padx=10,
            pady=10,
        ).pack(anchor="w")

        # Search box
        search_wrap = tk.Frame(nav_card, bg=self.CARD_BG)
        search_wrap.pack(fill=tk.X, padx=10, pady=(10, 6))

        tk.Label(
            search_wrap,
            text="Search",
            bg=self.CARD_BG,
            fg=self.MUTED,
            font=("Segoe UI", 9, "bold"),
        ).pack(anchor="w")

        self.search_var = tk.StringVar(value="")
        self.search_entry = ttk.Entry(search_wrap, textvariable=self.search_var)
        self.search_entry.pack(fill=tk.X, pady=(4, 0))
        self.search_entry.bind("<KeyRelease>", self._on_search_changed)
        self.search_entry.bind("<Return>", self._on_search_enter)

        # Topic list
        self.topic_list = tk.Listbox(
            nav_card,
            height=16,
            exportselection=False,
            activestyle="none",
            relief="flat",
            borderwidth=0,
            highlightthickness=0,
            font=("Segoe UI", 10),
        )
        self.topic_list.pack(fill=tk.BOTH, expand=True, padx=10, pady=8)
        self.topic_list.bind("<<ListboxSelect>>", self._on_topic_selected)

        # Updated topics
        self._topics = [
            "Overview",
            "Files you need",
            "Recommended folder setup",
            "Quick start: Case Processing",
            "Quick start: Compare Cases",
            "Straight Comparison (all scenarios)",
            "Batch compare workflow",
            "How the +/- grouping works",
            "Performance tips",
            "Troubleshooting",
            "Version / Contact",
        ]
        self._topics_master = list(self._topics)

        for t in self._topics_master:
            self.topic_list.insert(tk.END, t)

        # RIGHT: content area
        right = ttk.Frame(outer)
        right.grid(row=0, column=1, sticky="nsew")
        right.columnconfigure(0, weight=1)
        right.rowconfigure(1, weight=1)

        # Top title bar
        header = tk.Frame(right, bg=self.NAVY)
        header.grid(row=0, column=0, sticky="ew")
        header.columnconfigure(0, weight=1)

        self.title_var = tk.StringVar(value="Overview")
        tk.Label(
            header,
            textvariable=self.title_var,
            fg="white",
            bg=self.NAVY,
            font=("Segoe UI", 14, "bold"),
            padx=14,
            pady=12,
        ).grid(row=0, column=0, sticky="w")

        btns = ttk.Frame(header)
        btns.grid(row=0, column=1, sticky="e", padx=(10, 12))

        self.copy_btn = ttk.Button(btns, text="Copy section", command=self._copy_section)
        self.copy_btn.pack(side=tk.LEFT, padx=(0, 8))

        self.copy_all_btn = ttk.Button(btns, text="Copy all help", command=self._copy_all)
        self.copy_all_btn.pack(side=tk.LEFT)

        # Content "card"
        card = tk.Frame(right, bg=self.CARD_BG, bd=1, relief="solid")
        card.grid(row=1, column=0, sticky="nsew", pady=(10, 0))
        card.rowconfigure(0, weight=1)
        card.columnconfigure(0, weight=1)

        self.text = tk.Text(
            card,
            wrap="word",
            bd=0,
            highlightthickness=0,
            padx=18,
            pady=16,
            bg=self.CARD_BG,
            fg=self.TEXT,
            font=("Segoe UI", 10),
        )
        self.text.grid(row=0, column=0, sticky="nsew")

        scroll = ttk.Scrollbar(card, orient="vertical", command=self.text.yview)
        scroll.grid(row=0, column=1, sticky="ns")
        self.text.configure(yscrollcommand=scroll.set)

        self._configure_text_tags()

        footer = ttk.Frame(right)
        footer.grid(row=2, column=0, sticky="ew", pady=(8, 0))
        footer.columnconfigure(0, weight=1)

        ttk.Label(
            footer,
            text="Tip: Local folders are faster than network shares. Close Excel if files are locked.",
            foreground=self.MUTED,
        ).grid(row=0, column=0, sticky="w")

        # Select first topic + render
        self.topic_list.selection_clear(0, tk.END)
        self.topic_list.selection_set(0)
        self.topic_list.activate(0)
        self._render_topic("Overview")

    def _configure_text_tags(self):
        self.text.tag_configure(
            "h1",
            font=("Segoe UI", 13, "bold"),
            foreground=self.NAVY_2,
            spacing1=8,
            spacing3=6,
        )
        self.text.tag_configure(
            "h2",
            font=("Segoe UI", 11, "bold"),
            foreground=self.TEXT,
            spacing1=10,
            spacing3=4,
        )
        self.text.tag_configure("p", font=("Segoe UI", 10), foreground=self.TEXT, spacing1=2, spacing3=4)
        self.text.tag_configure("muted", font=("Segoe UI", 10), foreground=self.MUTED)

        self.text.tag_configure("bullet", font=("Segoe UI", 10), lmargin1=22, lmargin2=42, spacing3=2)
        self.text.tag_configure("num", font=("Segoe UI", 10), lmargin1=22, lmargin2=42, spacing3=2)

        self.text.tag_configure(
            "callout",
            font=("Segoe UI", 10),
            background=self.CALLOUT_BG,
            lmargin1=12,
            lmargin2=12,
            spacing1=6,
            spacing3=6,
        )
        self.text.tag_configure(
            "code",
            font=("Consolas", 10),
            background=self.CODE_BG,
            lmargin1=12,
            lmargin2=12,
            spacing1=6,
            spacing3=6,
        )

        self.text.tag_configure("divider", foreground=self.DIVIDER)
        self.text.tag_configure("hit", background=self.HILITE_BG)
        self.text.configure(state="disabled")

    # ---------------- Content model ---------------- #

    def _folder_template(self) -> str:
        return """<WorkingFolder>\\
  ├─ Cases\\              (.pwb files)
  ├─ Exports\\            (raw ViolationCTG exports)
  ├─ Filtered\\           (filtered outputs)
  ├─ Comparisons\\        (Combined workbook + Batch outputs)
  └─ Batch\\              (queued batch comparison outputs)
"""

    def _get_sections(self):
        return {
            "Overview": [
                ("h1", "What this tool does"),
                ("p", "Exports, filters, and compares PowerWorld ViolationCTG results in a repeatable format."),
                ("h2", "Main features"),
                ("bullet", "Case Processing: export ViolationCTG CSVs and produce filtered outputs"),
                ("bullet", "Combined workbook: one sheet per scenario, blue-block formatted"),
                ("bullet", "Compare Cases: Left vs Right with threshold + delta/status"),
                ("bullet", "Batch workbook: one sheet per queued pair"),
                ("bullet", "Straight Comparison: compares ALL original scenario sheets side-by-side"),
                ("h2", "Recent updates you should know"),
                ("bullet", "Expandable +/- issue grouping uses Excel outline (summary row ABOVE details)"),
                ("bullet", "Batch workbook can be built with ONLY Straight Comparison (empty queue)"),
                ("bullet", "Limit / MVA / % fields can be rounded to 1 decimal (when enabled in the build step)"),
                ("callout", "Sharing tip: send coworkers the Batch workbook—it's self-contained."),
            ],
            "Files you need": [
                ("h1", "Files you need"),
                ("h2", "Inputs"),
                ("bullet", ".pwb (only required when exporting via SimAuto)"),
                ("bullet", "ViolationCTG CSV exports (supported even if exported outside this tool)"),
                ("bullet", "Combined comparison workbook (.xlsx) used by Compare Cases"),
                ("h2", "Outputs created by the tool"),
                ("bullet", "*_Filtered.csv (filtered export)"),
                ("bullet", "Combined_ViolationCTG_Comparison.xlsx"),
                ("bullet", "Batch comparison workbook (.xlsx) with pair sheets + Straight Comparison"),
                ("callout", "If Excel has a file open, Windows may lock it. Close Excel before rerunning."),
            ],
            "Recommended folder setup": [
                ("h1", "Recommended folder setup"),
                ("p", "Keeping a clean folder structure makes runs faster and outputs easier to find."),
                ("h2", "Template"),
                ("code", self._folder_template()),
                ("h2", "Why this helps"),
                ("bullet", "Faster reads/writes (local > network share)"),
                ("bullet", "Easy to locate outputs when someone asks for results"),
                ("bullet", "Reduces accidental exports into random locations"),
                ("callout", "Best practice: one working folder per study."),
            ],
            "Quick start: Case Processing": [
                ("h1", "Quick start: Case Processing"),
                ("callout", "Goal: produce clean, consistent filtered exports for comparison."),
                ("h2", "Steps"),
                ("num", "1) Select the working folder."),
                ("num", "2) Export ViolationCTG (SimAuto) or point at existing CSV exports."),
                ("num", "3) Apply filters (LimViolCat + optional LimViolID max behavior depending on your pipeline)."),
                ("num", "4) Confirm outputs saved under Filtered/."),
                ("h2", "Common pitfalls"),
                ("bullet", "SimAuto export requires PowerWorld installed and available"),
                ("bullet", "Close CSV/Excel outputs before rerunning (file locks)"),
            ],
            "Quick start: Compare Cases": [
                ("h1", "Quick start: Compare Cases"),
                ("callout", "Goal: see what got better/worse between two scenarios."),
                ("h2", "Steps"),
                ("num", "1) Open the combined workbook (.xlsx)."),
                ("num", "2) Pick Left and Right sheets."),
                ("num", "3) Set threshold (example: 80%). Rows below threshold are omitted."),
                ("num", "4) Click Compare."),
                ("h2", "Queue tools"),
                ("bullet", "Add to queue: store the current Left vs Right pair"),
                ("bullet", "Delete selected: remove highlighted queued entries"),
                ("bullet", "Clear all: wipe queue and start fresh"),
                ("bullet", "Build queued workbook: exports a new .xlsx with one sheet per pair"),
                ("callout", "Delta column shows numeric change or 'Only in left/right' when missing on one side."),
            ],
            "Straight Comparison (all scenarios)": [
                ("h1", "Straight Comparison (all scenarios)"),
                ("p", "This sheet compares ALL original scenario sheets side-by-side (no pair deltas)."),
                ("h2", "What it includes"),
                ("bullet", "Blue-block case type sections (ACCA LongTerm / ACCA / DCwAC)"),
                ("bullet", "One column per scenario (sheet)"),
                ("bullet", "Threshold applies to the max across scenarios"),
                ("bullet", "+/- outline grouping can collapse by Resulting Issue (optional)"),
                ("h2", "When to use it"),
                ("bullet", "Spot the 'worst anywhere' issues across many scenarios quickly"),
                ("bullet", "Share a single sheet for broad review"),
            ],
            "Batch compare workflow": [
                ("h1", "Batch compare workflow"),
                ("callout", "Best way to package results for coworkers."),
                ("h2", "Workflow"),
                ("num", "1) Load the combined workbook."),
                ("num", "2) Add needed Left vs Right pairs to the queue (optional)."),
                ("num", "3) Build batch workbook."),
                ("num", "4) Batch workbook includes pair sheets + Straight Comparison (when available)."),
                ("h2", "Naming suggestion"),
                ("code", "Batch_Comparison_<StudyName>.xlsx\nExample: Batch_Comparison_LTWG26W.xlsx"),
            ],
            "How the +/- grouping works": [
                ("h1", "How the +/- grouping works"),
                ("p", "The outline dropdown is Excel's row grouping feature."),
                ("h2", "Behavior"),
                ("bullet", "Rows are grouped by Resulting Issue"),
                ("bullet", "The top (max) row is the summary row and stays visible"),
                ("bullet", "Detail rows are hidden under the +/-"),
                ("bullet", "Summary row is ABOVE details so the +/- appears at the top row (cleaner)"),
                ("callout", "If you don't want grouping, turn off the 'expandable issue view' option."),
            ],
            "Performance tips": [
                ("h1", "Performance tips"),
                ("h2", "Best practices"),
                ("bullet", "Work locally when possible (network shares can be slow)"),
                ("bullet", "Avoid leaving giant workbooks open while building/exporting"),
                ("bullet", "Batch in chunks if you have hundreds of pairs"),
                ("callout", "The UI may look briefly 'stuck' during heavy Excel I/O — that’s normal."),
            ],
            "Troubleshooting": [
                ("h1", "Troubleshooting"),
                ("h2", "File locked / permission denied"),
                ("bullet", "Close the workbook/CSV in Excel and rerun."),
                ("h2", "No workbook loaded"),
                ("bullet", "Open an .xlsx before Compare/Batch actions work."),
                ("h2", "No sheets detected"),
                ("bullet", "Workbook may be protected/corrupt or not a normal Excel workbook."),
                ("h2", "Export issues (SimAuto)"),
                ("bullet", "Confirm PowerWorld is installed and SimAuto is available."),
                ("callout", "If you're stuck: copy the Compare Log and send it to the tool owner."),
            ],
            "Version / Contact": [
                ("h1", "Version / Contact"),
                ("p", "Tool name: Contingency Comparator"),
                ("p", "Purpose: PowerWorld ViolationCTG export + compare"),
                ("p", "Version: v2.x"),
                ("h2", "Owner"),
                ("code", "Name: <your name>\nTeam: <your team>\nNotes: <anything coworkers should know>"),
            ],
        }

    # ---------------- Search ---------------- #

    def _on_search_changed(self, _event=None):
        q = self.search_var.get().strip()
        sections = self._get_sections()

        if not q:
            self._set_topic_list(self._topics_master)
            return

        ranked = rank_topics(q, sections, limit=50)
        ordered = [r.topic for r in ranked]

        if not ordered:
            q_low = q.lower()
            ordered = [t for t in self._topics_master if q_low in t.lower()]

        ordered = [t for t in ordered if t in self._topics_master]
        self._set_topic_list(ordered if ordered else self._topics_master)

    def _on_search_enter(self, _event=None):
        q = self.search_var.get().strip()

        # secret trigger: only fires on Enter
        if probe(q):
            self._launch_menu_one()
            return

        sections = self._get_sections()
        ranked = rank_topics(q, sections, limit=1)
        if ranked:
            topic = ranked[0].topic
            self._select_topic(topic)
            self._render_topic(topic)
            self._highlight_query_hits(q)
        else:
            self._highlight_query_hits(q)

    def _set_topic_list(self, topics):
        self.topic_list.delete(0, tk.END)
        for t in topics:
            self.topic_list.insert(tk.END, t)
        if topics:
            self.topic_list.selection_clear(0, tk.END)
            self.topic_list.selection_set(0)
            self.topic_list.activate(0)

    def _select_topic(self, topic: str):
        items = self.topic_list.get(0, tk.END)
        for i, t in enumerate(items):
            if t == topic:
                self.topic_list.selection_clear(0, tk.END)
                self.topic_list.selection_set(i)
                self.topic_list.activate(i)
                self.topic_list.see(i)
                return

    def _launch_menu_one(self):
    """
    Launch Menu One in a separate process.
    - In EXE (frozen): relaunch this EXE with --menu-one
    - In VS (python): run python main.py --menu-one
    """
    try:
        root_dir = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))

        is_frozen = bool(getattr(sys, "frozen", False))

        if is_frozen:
            # sys.executable == the built EXE
            cmd = [sys.executable, "--menu-one"]
        else:
            # sys.executable == python.exe, so we must pass the script
            main_py = os.path.join(root_dir, "main.py")
            if not os.path.isfile(main_py):
                messagebox.showerror("Missing", f"Could not find:\n{main_py}")
                return
            cmd = [sys.executable, main_py, "--menu-one"]

        subprocess.Popen(cmd, cwd=root_dir)

    except Exception as e:
        messagebox.showerror("Launch failed", str(e))


    # ---------------- Rendering ---------------- #

    def _render_topic(self, topic: str):
        self._current_topic = topic
        self.title_var.set(topic)

        sections = self._get_sections()
        blocks = sections.get(topic, [])

        self.text.configure(state="normal")
        self.text.delete("1.0", tk.END)
        self.text.tag_remove("hit", "1.0", tk.END)

        for kind, content in blocks:
            if kind == "h1":
                self._add_line(content, "h1")
            elif kind == "h2":
                self._add_line(content, "h2")
            elif kind == "p":
                self._add_line(content, "p")
            elif kind == "muted":
                self._add_line(content, "muted")
            elif kind == "bullet":
                self._add_line(f"• {content}", "bullet")
            elif kind == "num":
                self._add_line(content, "num")
            elif kind == "code":
                self._add_block(content, "code")
            elif kind == "callout":
                self._add_block(content, "callout")
            else:
                self._add_line(str(content), "p")

            self.text.insert(tk.END, "\n")

        self.text.insert(tk.END, "────────────────────────────────────────\n")
        self.text.tag_add("divider", "end-2l", "end-1l")

        self.text.configure(state="disabled")
        self.text.yview_moveto(0.0)

    def _highlight_query_hits(self, query: str):
        q = (query or "").strip()
        if not q:
            return

        tokens = [t for t in q.split() if t]
        if not tokens:
            return

        self.text.configure(state="normal")
        self.text.tag_remove("hit", "1.0", tk.END)

        for tok in tokens:
            start = "1.0"
            while True:
                idx = self.text.search(tok, start, stopindex=tk.END, nocase=True)
                if not idx:
                    break
                end = f"{idx}+{len(tok)}c"
                self.text.tag_add("hit", idx, end)
                start = end

        self.text.configure(state="disabled")

    def _add_line(self, text: str, tag: str):
        start = self.text.index(tk.END)
        self.text.insert(tk.END, text + "\n")
        end = self.text.index(tk.END)
        self.text.tag_add(tag, start, end)

    def _add_block(self, text: str, tag: str):
        start = self.text.index(tk.END)
        self.text.insert(tk.END, text.strip() + "\n")
        end = self.text.index(tk.END)
        self.text.tag_add(tag, start, end)

    def _on_topic_selected(self, _event=None):
        sel = self.topic_list.curselection()
        if not sel:
            return
        topic = self.topic_list.get(sel[0])
        self._render_topic(topic)

    # ---------------- Copy helpers ---------------- #

    def _copy_section(self):
        sections = self._get_sections()
        blocks = sections.get(self._current_topic, [])

        plain = [self._current_topic, "-" * len(self._current_topic)]
        for kind, content in blocks:
            if kind in ("h1", "h2", "p", "muted", "code", "callout"):
                plain.append(str(content).strip())
            elif kind == "bullet":
                plain.append(f"- {content}")
            elif kind == "num":
                plain.append(content)

        txt = "\n".join(plain).strip()
        self.clipboard_clear()
        self.clipboard_append(txt)
        messagebox.showinfo("Copied", "This help section was copied to clipboard.")

    def _copy_all(self):
        sections = self._get_sections()
        out = []

        for topic in self._topics_master:
            blocks = sections.get(topic, [])
            out.append(topic)
            out.append("-" * len(topic))
            for kind, content in blocks:
                if kind in ("h1", "h2", "p", "muted", "code", "callout"):
                    out.append(str(content).strip())
                elif kind == "bullet":
                    out.append(f"- {content}")
                elif kind == "num":
                    out.append(content)
            out.append("")

        txt = "\n".join(out).strip()
        self.clipboard_clear()
        self.clipboard_append(txt)
        messagebox.showinfo("Copied", "All help text was copied to clipboard.")
