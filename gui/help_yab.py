# gui/help_view.py

import tkinter as tk
from tkinter import ttk, messagebox


class HelpTab(ttk.Frame):
    """
    Professional Help / Documentation tab.

    Left = topic navigation
    Right = styled content

    Styled with tk.Text tags:
      - Title styling
      - Section headers
      - Paragraphs
      - Bullets / numbered steps
      - Callouts
      - Code/folder blocks (monospace)
    """

    # Palette
    NAVY = "#0B2F5B"
    NAVY_2 = "#103A6B"
    LIGHT_BG = "#F4F7FB"
    CARD_BG = "#FFFFFF"
    CALLOUT_BG = "#EAF2FF"
    CODE_BG = "#F2F2F2"
    MUTED = "#5C6773"
    TEXT = "#1F2A44"
    DIVIDER = "#D6DFEA"

    def __init__(self, master):
        super().__init__(master)
        self._current_topic = "Overview"
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
        nav_card.configure(width=220)

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
        self.topic_list.pack(fill=tk.BOTH, expand=True, padx=8, pady=8)

        self._topics = [
            "Overview",
            "Files you need",
            "Recommended folder setup",
            "Quick start: Case Processing",
            "Quick start: Compare Cases",
            "Batch compare workflow",
            "Performance tips",
            "Troubleshooting",
            "Version / Contact",
        ]
        for t in self._topics:
            self.topic_list.insert(tk.END, t)

        self.topic_list.bind("<<ListboxSelect>>", self._on_topic_selected)

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

        # Footer hint
        footer = ttk.Frame(right)
        footer.grid(row=2, column=0, sticky="ew", pady=(8, 0))
        footer.columnconfigure(0, weight=1)

        ttk.Label(
            footer,
            text="Tip: Keep your study in one working folder. Local drives are faster than network shares.",
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
        self.text.tag_configure(
            "p",
            font=("Segoe UI", 10),
            foreground=self.TEXT,
            spacing1=2,
            spacing3=4,
        )
        self.text.tag_configure("muted", font=("Segoe UI", 10), foreground=self.MUTED)

        # bullets + numbered steps
        self.text.tag_configure("bullet", font=("Segoe UI", 10), lmargin1=22, lmargin2=42, spacing3=2)
        self.text.tag_configure("num", font=("Segoe UI", 10), lmargin1=22, lmargin2=42, spacing3=2)

        # callout + code blocks
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

        # make read-only but selectable
        self.text.configure(state="disabled")

    # ---------------- Content model ---------------- #

    def _folder_template(self) -> str:
        # Use triple quotes to avoid raw-string "\" issues
        return """<WorkingFolder>\\
  ├─ Cases\\              (your .pwb files)
  ├─ Exports\\            (raw ViolationCTG exports)
  ├─ Filtered\\           (filtered outputs)
  ├─ Comparisons\\        (combined / compare workbooks)
  └─ Batch\\              (queued batch comparison outputs)
"""

    def _get_sections(self):
        return {
            "Overview": [
                ("h1", "What this tool does"),
                ("p", "This tool helps you export, filter, and compare PowerWorld ViolationCTG results in a repeatable, shareable format."),
                ("h2", "Main features"),
                ("bullet", "Export ViolationCTG to CSV via SimAuto (when available)"),
                ("bullet", "Filter rows (LimViolCat) and remove unwanted columns (blacklist)"),
                ("bullet", "Optional LimViolID max filter (keeps highest LimViolPct per LimViolID)"),
                ("bullet", "Compare two scenarios (Left vs Right) with a percent threshold"),
                ("bullet", "Queue multiple comparisons and build a batch workbook for sharing"),
                ("h2", "Recommended workflow"),
                ("num", "1) Run Case Processing to create clean, filtered exports."),
                ("num", "2) Build or choose a combined workbook that contains sheets for each scenario."),
                ("num", "3) Use Compare Cases for live review and batch workbook export."),
                ("callout", "Sharing tip: Send coworkers the batch comparison workbook — it is self-contained."),
            ],

            "Files you need": [
                ("h1", "Files you need"),
                ("h2", "Inputs"),
                ("bullet", ".pwb (PowerWorld case) — required only if exporting via SimAuto"),
                ("bullet", "ViolationCTG .csv — supported if exported outside this tool"),
                ("bullet", "Combined comparison .xlsx — used for Compare Cases tab"),
                ("h2", "Outputs created by the tool"),
                ("bullet", "*_Filtered.csv (filtered export)"),
                ("bullet", "Batch comparison workbook (.xlsx) with one sheet per queued pair"),
                ("callout", "If Excel has a workbook/CSV open, Windows may lock it. Close Excel before rerunning."),
            ],

            "Recommended folder setup": [
                ("h1", "Recommended folder setup"),
                ("p", "Keeping a clean folder structure makes runs faster and outputs easier to find."),
                ("h2", "Template"),
                ("code", self._folder_template()),
                ("h2", "Why this helps"),
                ("bullet", "Faster reads/writes (local > network share)"),
                ("bullet", "Easy to find outputs when someone asks for results"),
                ("bullet", "Reduces accidental exports into random locations"),
                ("callout", "Best practice: One working folder per study (LTWG, ACCA, etc.)."),
            ],

            "Quick start: Case Processing": [
                ("h1", "Quick start: Case Processing"),
                ("callout", "Goal: Produce clean filtered exports so comparisons are consistent."),
                ("h2", "Steps"),
                ("num", "1) Select the case folder / working folder (as your tab requires)."),
                ("num", "2) Export ViolationCTG (SimAuto)."),
                ("num", "3) Apply filters (LimViolCat + optional LimViolID max filter)."),
                ("num", "4) Confirm output files saved (Filtered outputs)."),
                ("h2", "Common pitfalls"),
                ("bullet", "PowerWorld + SimAuto must be available on the machine running exports"),
                ("bullet", "Close CSV/Excel outputs before rerunning to avoid file locks"),
            ],

            "Quick start: Compare Cases": [
                ("h1", "Quick start: Compare Cases"),
                ("callout", "Goal: See what got better/worse between two scenarios."),
                ("h2", "Steps"),
                ("num", "1) Open the combined comparison workbook (.xlsx)."),
                ("num", "2) Pick Left and Right sheets."),
                ("num", "3) Set threshold (default 80%). Rows below threshold are hidden."),
                ("num", "4) Click Compare."),
                ("h2", "Queue tools"),
                ("bullet", "Add to queue: store the current Left vs Right pair"),
                ("bullet", "Delete selected: remove highlighted queued entries"),
                ("bullet", "Clear all: wipe queue and start fresh"),
                ("bullet", "Build queued workbook: exports a new .xlsx with one sheet per queued pair"),
            ],

            "Batch compare workflow": [
                ("h1", "Batch compare workflow"),
                ("callout", "This is the best way to package results for coworkers."),
                ("h2", "Workflow"),
                ("num", "1) Load the combined workbook."),
                ("num", "2) Add all needed Left vs Right pairs to the queue."),
                ("num", "3) Build queued workbook and save it into your working folder."),
                ("h2", "Naming suggestion"),
                ("code", "Batch_Comparison_<StudyName>.xlsx\nExample: Batch_Comparison_LTWG26W.xlsx"),
            ],

            "Performance tips": [
                ("h1", "Performance tips"),
                ("h2", "Best practices"),
                ("bullet", "Work locally when possible (network shares can be slower)"),
                ("bullet", "Keep workbook + outputs in the same folder"),
                ("bullet", "Avoid leaving giant workbooks open in Excel while running"),
                ("bullet", "Batch in chunks if you have hundreds of sheet pairs"),
                ("callout", "UI may look briefly “stuck” during heavy Excel I/O — that’s normal for large workbooks."),
            ],

            "Troubleshooting": [
                ("h1", "Troubleshooting"),
                ("h2", "File locked / permission denied"),
                ("bullet", "Close the workbook/CSV in Excel and rerun."),
                ("h2", "No workbook loaded"),
                ("bullet", "You must open an .xlsx before Compare/Batch actions work."),
                ("h2", "No sheets detected"),
                ("bullet", "Workbook may be protected/corrupt or not a normal Excel workbook."),
                ("h2", "Export issues (SimAuto)"),
                ("bullet", "Confirm PowerWorld is installed and SimAuto is available."),
                ("callout", "If you’re stuck: copy the Compare Log and send it to the tool owner."),
            ],

            "Version / Contact": [
                ("h1", "Version / Contact"),
                ("p", "Tool name: Contingency Comparison Tool"),
                ("p", "Purpose: PowerWorld Results Export + Compare"),
                ("p", "Version: v1.0"),
                ("h2", "Owner"),
                ("code", "Name: <your name>\nTeam: <your team>\nNotes: <anything coworkers should know>"),
            ],
        }

    # ---------------- Rendering ---------------- #

    def _render_topic(self, topic: str):
        self._current_topic = topic
        self.title_var.set(topic)

        sections = self._get_sections()
        blocks = sections.get(topic, [])

        self.text.configure(state="normal")
        self.text.delete("1.0", tk.END)

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

        # subtle divider line at end
        self.text.insert(tk.END, "────────────────────────────────────────\n")
        self.text.tag_add("divider", "end-2l", "end-1l")

        self.text.configure(state="disabled")
        self.text.yview_moveto(0.0)

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

        for topic in self._topics:
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