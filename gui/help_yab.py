# gui/help_view.py

import tkinter as tk
from tkinter import ttk, messagebox


class HelpTab(ttk.Frame):
    """
    Professional Help / Documentation tab.

    Uses a tk.Text widget with tags for:
      - Title styling
      - Section headers
      - Bullets
      - Callouts (highlight boxes)
      - Code/folder blocks (monospace)
    """

    NAVY = "#0B2F5B"
    LIGHT_BG = "#F4F7FB"
    CALLOUT_BG = "#EAF2FF"
    CODE_BG = "#F2F2F2"
    MUTED = "#5C6773"

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

        # Left nav
        nav = ttk.Frame(outer)
        nav.grid(row=0, column=0, sticky="nsw", padx=(0, 12))
        ttk.Label(nav, text="Help Topics", font=("Segoe UI", 10, "bold")).pack(anchor="w", pady=(0, 6))

        self.topic_list = tk.Listbox(
            nav,
            height=14,
            exportselection=False,
            activestyle="none",
            relief="solid",
            borderwidth=1,
        )
        self.topic_list.pack(fill=tk.Y, expand=False)

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

        # Right content area
        right = ttk.Frame(outer)
        right.grid(row=0, column=1, sticky="nsew")
        right.columnconfigure(0, weight=1)
        right.rowconfigure(1, weight=1)

        header = ttk.Frame(right)
        header.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        header.columnconfigure(0, weight=1)

        self.title_var = tk.StringVar(value="Overview")
        title_lbl = tk.Label(
            header,
            textvariable=self.title_var,
            fg="white",
            bg=self.NAVY,
            font=("Segoe UI", 14, "bold"),
            padx=12,
            pady=10,
        )
        title_lbl.grid(row=0, column=0, sticky="ew")

        btns = ttk.Frame(header)
        btns.grid(row=0, column=1, sticky="e", padx=(10, 0))

        self.copy_btn = ttk.Button(btns, text="Copy section", command=self._copy_section)
        self.copy_btn.pack(side=tk.LEFT, padx=(0, 6))

        self.copy_all_btn = ttk.Button(btns, text="Copy all help", command=self._copy_all)
        self.copy_all_btn.pack(side=tk.LEFT)

        # Content card
        card = tk.Frame(right, bg=self.LIGHT_BG, bd=1, relief="solid")
        card.grid(row=1, column=0, sticky="nsew")
        card.rowconfigure(0, weight=1)
        card.columnconfigure(0, weight=1)

        self.text = tk.Text(
            card,
            wrap="word",
            bd=0,
            highlightthickness=0,
            padx=14,
            pady=12,
            bg=self.LIGHT_BG,
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
            text="Tip: Keep your study in one working folder. Local drives are faster than network shares.",
            foreground=self.MUTED,
        ).grid(row=0, column=0, sticky="w")

        self.topic_list.selection_set(0)
        self._render_topic("Overview")

    def _configure_text_tags(self):
        self.text.tag_configure("h1", font=("Segoe UI", 13, "bold"), foreground=self.NAVY, spacing1=6, spacing3=6)
        self.text.tag_configure("h2", font=("Segoe UI", 11, "bold"), foreground="#1F2A44", spacing1=8, spacing3=4)
        self.text.tag_configure("p", font=("Segoe UI", 10), foreground="#1F2A44", spacing1=2, spacing3=2)
        self.text.tag_configure("muted", font=("Segoe UI", 10), foreground=self.MUTED)

        self.text.tag_configure("bullet", font=("Segoe UI", 10), lmargin1=18, lmargin2=34)
        self.text.tag_configure("num", font=("Segoe UI", 10), lmargin1=18, lmargin2=34)

        self.text.tag_configure(
            "callout",
            font=("Segoe UI", 10),
            background=self.CALLOUT_BG,
            lmargin1=10,
            lmargin2=10,
            spacing1=4,
            spacing3=4,
        )

        self.text.tag_configure(
            "code",
            font=("Consolas", 10),
            background=self.CODE_BG,
            lmargin1=12,
            lmargin2=12,
            spacing1=3,
            spacing3=3,
        )

        self.text.configure(state="disabled")

    # ---------------- Content model ---------------- #

    def _folder_template(self) -> str:
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
                ("bullet", "Queue multiple comparisons and build a new batch workbook"),
                ("callout", "If you’re sharing with others: share the batch workbook output — it’s self-contained."),
            ],

            "Recommended folder setup": [
                ("h1", "Recommended folder setup"),
                ("p", "Keeping a clean folder structure makes runs faster and outputs easier to find."),
                ("h2", "Template"),
                ("code", self._folder_template()),
                ("h2", "Why this helps"),
                ("bullet", "Faster reads/writes (especially if working locally)"),
                ("bullet", "Easy to locate outputs when someone asks “where did the batch file go?”"),
                ("bullet", "Less chance of saving exports into random folders"),
            ],
        }

    # ---------------- Rendering ---------------- #

    def _render_topic(self, topic: str):
        self._current_topic = topic
        self.title_var.set(topic)

        blocks = self._get_sections().get(topic, [])

        self.text.configure(state="normal")
        self.text.delete("1.0", tk.END)

        for kind, content in blocks:
            if kind == "h1":
                self._add_line(content, "h1")
            elif kind == "h2":
                self._add_line(content, "h2")
            elif kind == "p":
                self._add_line(content, "p")
            elif kind == "bullet":
                self._add_line(f"• {content}", "bullet")
            elif kind == "code":
                self._add_block(content, "code")

            self.text.insert(tk.END, "\n")

        self.text.configure(state="disabled")

    def _add_line(self, text: str, tag: str):
        start = self.text.index(tk.END)
        self.text.insert(tk.END, text + "\n")
        end = self.text.index(tk.END)
        self.text.tag_add(tag, start, end)

    def _add_block(self, text: str, tag: str):
        start = self.text.index(tk.END)
        self.text.insert(tk.END, text.rstrip() + "\n")
        end = self.text.index(tk.END)
        self.text.tag_add(tag, start, end)

    def _on_topic_selected(self, _event=None):
        sel = self.topic_list.curselection()
        if sel:
            self._render_topic(self.topic_list.get(sel[0]))

    # ---------------- Copy helpers ---------------- #

    def _copy_section(self):
        self.clipboard_clear()
        self.clipboard_append(self.text.get("1.0", tk.END))
        messagebox.showinfo("Copied", "This help section was copied to clipboard.")

    def _copy_all(self):
        self.clipboard_clear()
        for topic in self._topics:
            self._render_topic(topic)
            self.clipboard_append(self.text.get("1.0", tk.END) + "\n\n")
        messagebox.showinfo("Copied", "All help text was copied to clipboard.")