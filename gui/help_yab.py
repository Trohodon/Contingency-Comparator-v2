# gui/help_view.py

import os
import tkinter as tk
from tkinter import ttk, messagebox


class HelpTab(ttk.Frame):
    """
    Help / Documentation tab for the Contingency Comparison Tool.

    Goals:
      - Explain what the tool does
      - Describe required inputs (PWB / CSV / XLSX)
      - Recommend folder layout for speed and organization
      - Provide quick-start steps and troubleshooting tips
    """

    def __init__(self, master):
        super().__init__(master)
        self._build_gui()

    def _build_gui(self):
        # Main layout: left "topics" + right content
        outer = ttk.Frame(self)
        outer.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        outer.columnconfigure(0, weight=0)
        outer.columnconfigure(1, weight=1)
        outer.rowconfigure(0, weight=1)

        # Left panel: topic list
        left = ttk.Frame(outer)
        left.grid(row=0, column=0, sticky="nsw", padx=(0, 10))
        ttk.Label(left, text="Help Topics").pack(anchor="w", pady=(0, 6))

        self.topic_list = tk.Listbox(left, height=12, exportselection=False)
        self.topic_list.pack(fill=tk.Y, expand=False)

        topics = [
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
        for t in topics:
            self.topic_list.insert(tk.END, t)

        self.topic_list.bind("<<ListboxSelect>>", self._on_topic_selected)

        # Right panel: content + actions
        right = ttk.Frame(outer)
        right.grid(row=0, column=1, sticky="nsew")
        right.rowconfigure(1, weight=1)
        right.columnconfigure(0, weight=1)

        header = ttk.Frame(right)
        header.grid(row=0, column=0, sticky="ew", pady=(0, 6))
        header.columnconfigure(0, weight=1)

        self.title_var = tk.StringVar(value="Overview")
        ttk.Label(
            header, textvariable=self.title_var, font=("Segoe UI", 12, "bold")
        ).grid(row=0, column=0, sticky="w")

        btns = ttk.Frame(header)
        btns.grid(row=0, column=1, sticky="e")

        self.copy_btn = ttk.Button(btns, text="Copy section", command=self._copy_section)
        self.copy_btn.pack(side=tk.LEFT, padx=(0, 6))

        self.copy_all_btn = ttk.Button(btns, text="Copy all help", command=self._copy_all)
        self.copy_all_btn.pack(side=tk.LEFT)

        # Text content area
        text_frame = ttk.Frame(right)
        text_frame.grid(row=1, column=0, sticky="nsew")
        text_frame.rowconfigure(0, weight=1)
        text_frame.columnconfigure(0, weight=1)

        self.text = tk.Text(text_frame, wrap="word")
        self.text.grid(row=0, column=0, sticky="nsew")

        scroll = ttk.Scrollbar(text_frame, orient="vertical", command=self.text.yview)
        scroll.grid(row=0, column=1, sticky="ns")
        self.text.configure(yscrollcommand=scroll.set)

        # Footer quick actions
        footer = ttk.Frame(right)
        footer.grid(row=2, column=0, sticky="ew", pady=(6, 0))
        footer.columnconfigure(0, weight=1)

        ttk.Label(
            footer,
            text="Tip: Keep your .pwb and generated .csv/.xlsx in the same working folder for easy exports.",
        ).grid(row=0, column=0, sticky="w")

        # Load initial section
        self._set_section("Overview")

        # Select first topic visually
        self.topic_list.selection_clear(0, tk.END)
        self.topic_list.selection_set(0)
        self.topic_list.activate(0)

    # ---------------- Content ---------------- #

    def _get_sections(self):
        folder_template = (
            "Recommended folder template:\n"
            "  <WorkingFolder>\\\n"
            "    ├─ Cases\\              (your .pwb files)\n"
            "    ├─ Exports\\            (exported ViolationCTG csv)\n"
            "    ├─ Filtered\\           (filtered csv outputs)\n"
            "    ├─ Comparisons\\        (Combined comparison xlsx outputs)\n"
            "    └─ Batch\\              (Batch_Comparison.xlsx outputs)\n"
        )

        return {
            "Overview": (
                "What this tool does\n"
                "- Exports PowerWorld ViolationCTG results to CSV (via SimAuto)\n"
                "- Filters CSV rows/columns to match Dominion-style workflows\n"
                "- Builds comparison workbooks (Left vs Right sheets)\n"
                "- Supports batching multiple comparisons into one new workbook\n"
                "\n"
                "Main tabs\n"
                "1) Case Processing\n"
                "   - Export + filter a case (or multiple cases depending on your setup)\n"
                "2) Compare Cases\n"
                "   - Compare any two sheets in a combined workbook\n"
                "   - Add multiple sheet pairs to a queue\n"
                "   - Build a batch workbook where each queued comparison becomes a sheet\n"
            ),
            "Files you need": (
                "Inputs you may use\n"
                "- .pwb files (PowerWorld cases) if you are exporting ViolationCTG\n"
                "- Exported ViolationCTG .csv (if you already exported outside the tool)\n"
                "- Combined_ViolationCTG_Comparison.xlsx (or compatible workbook) for comparing sheets\n"
                "\n"
                "Outputs the tool creates\n"
                "- *_Filtered.csv (filtered export)\n"
                "- Comparison workbook(s) in .xlsx format\n"
                "  - Single comparisons (if you do that flow)\n"
                "  - Batch workbook (queued comparisons)\n"
            ),
            "Recommended folder setup": (
                "Best practice folder layout\n"
                "- Keep everything for a study in ONE working folder.\n"
                "- Put your source workbook and any outputs in the SAME folder when possible.\n"
                "- Avoid working directly out of OneDrive/network shares if you notice slowness.\n"
                "\n"
                + folder_template +
                "\n"
                "Why this helps\n"
                "- Easier file picking\n"
                "- Faster reads/writes\n"
                "- No hunting for outputs\n"
            ),
            "Quick start: Case Processing": (
                "Case Processing workflow\n"
                "1) Choose your .pwb file (or whichever flow your Case Processing tab supports)\n"
                "2) Run export\n"
                "3) Apply filters (LimViolCat categories, LimViolID max filter if enabled)\n"
                "4) Confirm the *_Filtered.csv output\n"
                "\n"
                "Notes\n"
                "- If PowerWorld/SimAuto isn’t available, export must be run on a machine with it installed.\n"
                "- If a CSV is open in Excel, Windows can lock it. Close the CSV before rerunning.\n"
            ),
            "Quick start: Compare Cases": (
                "Compare Cases workflow\n"
                "1) Click 'Open Excel Workbook' and select your comparison workbook (.xlsx)\n"
                "2) Pick Left sheet and Right sheet\n"
                "3) Set threshold (example: 80 means hide rows < 80% on BOTH sides)\n"
                "4) Click Compare to view ACCA LongTerm / ACCA / DCwAC results\n"
                "\n"
                "Queue actions\n"
                "- Add to queue: stores the current Left vs Right pair\n"
                "- Delete selected: removes highlighted pair(s)\n"
                "- Clear all: wipes the entire queue so you can start fresh (if enabled)\n"
                "- Build queued workbook: creates a new .xlsx with one sheet per queued pair\n"
            ),
            "Batch compare workflow": (
                "Batch workflow (recommended)\n"
                "1) Load the combined workbook\n"
                "2) Add multiple Left vs Right sheet pairs to the queue\n"
                "3) Build queued workbook\n"
                "4) Share the new batch workbook (it’s self-contained)\n"
                "\n"
                "Naming suggestion\n"
                "- Batch_Comparison_<StudyName>.xlsx\n"
                "- Example: Batch_Comparison_LTWG26W.xlsx\n"
            ),
            "Performance tips": (
                "Performance tips\n"
                "- Keep working folders local (C:\\ or a local drive) when possible\n"
                "- Close Excel workbooks that the tool needs to read/write\n"
                "- Avoid extremely long sheet names (Excel has limits; the tool sanitizes names)\n"
                "- Don’t stack thousands of queued comparisons in one go (do batches)\n"
                "\n"
                "Common speed win\n"
                "- Put the source workbook and batch output in the same folder (less browsing + cleaner outputs)\n"
            ),
            "Troubleshooting": (
                "Troubleshooting\n"
                "\n"
                "1) 'File is open / permission denied'\n"
                "- Close the workbook/CSV in Excel and rerun.\n"
                "\n"
                "2) 'No workbook loaded'\n"
                "- You must open an .xlsx before you can compare or batch.\n"
                "\n"
                "3) 'Sheets list is empty'\n"
                "- Workbook may be protected, corrupt, or not a normal Excel workbook.\n"
                "\n"
                "4) Export issues (SimAuto)\n"
                "- Confirm PowerWorld is installed.\n"
                "- Confirm your environment can access SimAuto.\n"
                "\n"
                "If stuck\n"
                "- Copy the Compare Log and send it to the tool owner.\n"
            ),
            "Version / Contact": (
                "Version / Contact\n"
                "- Tool name: Contingency Comparison Tool\n"
                "- Purpose: PowerWorld Results Export + Compare\n"
                "- Version: v1.0\n"
                "\n"
                "Owner\n"
                "- (Put your name here)\n"
                "\n"
                "Suggested notes\n"
                "- If you distribute an .exe, keep a short README in the same folder.\n"
            ),
        }

    def _set_section(self, name: str):
        sections = self._get_sections()
        content = sections.get(name, "")
        self.title_var.set(name)
        self.text.configure(state="normal")
        self.text.delete("1.0", tk.END)
        self.text.insert("1.0", content)
        self.text.configure(state="normal")  # keep selectable/copyable

    def _on_topic_selected(self, _event=None):
        sel = self.topic_list.curselection()
        if not sel:
            return
        topic = self.topic_list.get(sel[0])
        self._set_section(topic)

    # ---------------- Copy helpers ---------------- #

    def _copy_section(self):
        txt = self.text.get("1.0", tk.END).strip()
        if not txt:
            return
        self.clipboard_clear()
        self.clipboard_append(txt)
        messagebox.showinfo("Copied", "This help section was copied to clipboard.")

    def _copy_all(self):
        sections = self._get_sections()
        big = []
        for k, v in sections.items():
            big.append(f"{k}\n" + ("-" * len(k)))
            big.append(v.strip())
            big.append("")  # blank line
        all_text = "\n".join(big).strip()

        self.clipboard_clear()
        self.clipboard_append(all_text)
        messagebox.showinfo("Copied", "All help text was copied to clipboard.")