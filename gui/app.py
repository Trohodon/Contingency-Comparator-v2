# gui/app.py
import tkinter as tk
from tkinter import ttk

from gui.tab_case import CaseProcessingTab
from gui.tab_compare import CompareTab
from gui.help_view import HelpTab  # NEW


APP_TITLE = "Contingency Comparison Tool"
APP_SUBTITLE = "PowerWorld Results Export + Compare"
APP_VERSION = "v1.0"


class App(ttk.Frame):
    def __init__(self, master):
        super().__init__(master)
        self.master = master

        self._status_var = tk.StringVar(value="Ready")

        self._configure_window()
        self._configure_style()

        self._build_header()
        self._build_body()
        self._build_status_bar()

    # ---------------- Window + Style ---------------- #

    def _configure_window(self):
        self.master.title(APP_TITLE)
        self.master.minsize(1100, 700)

        # If you want it to open centered-ish:
        try:
            w, h = 1200, 760
            sx = (self.master.winfo_screenwidth() - w) // 2
            sy = (self.master.winfo_screenheight() - h) // 3
            self.master.geometry(f"{w}x{h}+{sx}+{sy}")
        except Exception:
            pass

    def _configure_style(self):
        style = ttk.Style()

        # Use a native-ish theme when possible
        # On Windows, "vista" or "clam" usually looks clean.
        try:
            style.theme_use("vista")
        except Exception:
            try:
                style.theme_use("clam")
            except Exception:
                pass

        # Global font + padding (ttk is limited, but this helps a ton)
        base_font = ("Segoe UI", 10)

        style.configure(".", font=base_font)
        style.configure("TButton", padding=(10, 6))
        style.configure("TLabel", padding=(0, 0))
        style.configure("TEntry", padding=(6, 4))
        style.configure("TCombobox", padding=(6, 4))
        style.configure("TNotebook", padding=(0, 0))
        style.configure("TNotebook.Tab", padding=(12, 8))

        # “Card” style (LabelFrame) look
        style.configure("Card.TLabelframe", padding=(10, 8))
        style.configure("Card.TLabelframe.Label", font=("Segoe UI", 10, "bold"))

        # Status bar label style
        style.configure("Status.TLabel", font=("Segoe UI", 9))

    # ---------------- Layout ---------------- #

    def _build_header(self):
        header = tk.Frame(self.master, bg="#0B2F5B")  # deep navy
        header.pack(side=tk.TOP, fill=tk.X)

        title = tk.Label(
            header,
            text=APP_TITLE,
            bg="#0B2F5B",
            fg="white",
            font=("Segoe UI", 16, "bold"),
            padx=14,
            pady=10,
        )
        title.pack(side=tk.LEFT)

        subtitle = tk.Label(
            header,
            text=APP_SUBTITLE,
            bg="#0B2F5B",
            fg="#DCE7F5",
            font=("Segoe UI", 10),
            padx=10,
        )
        subtitle.pack(side=tk.LEFT)

        version = tk.Label(
            header,
            text=APP_VERSION,
            bg="#0B2F5B",
            fg="#DCE7F5",
            font=("Segoe UI", 10, "bold"),
            padx=14,
        )
        version.pack(side=tk.RIGHT)

    def _build_body(self):
        # Outer body frame with consistent margins
        body = ttk.Frame(self.master, padding=(12, 12))
        body.pack(side=tk.TOP, fill=tk.BOTH, expand=True)

        # Notebook with tabs
        self.notebook = ttk.Notebook(body)
        self.notebook.pack(side=tk.TOP, fill=tk.BOTH, expand=True)

        # --- Tab 1: Case Processing ---
        self.tab_case = ttk.Frame(self.notebook, padding=(8, 8))
        self.notebook.add(self.tab_case, text="Case Processing")

        # --- Tab 2: Compare Cases ---
        self.tab_compare = ttk.Frame(self.notebook, padding=(8, 8))
        self.notebook.add(self.tab_compare, text="Compare Cases")

        # --- Tab 3: Help ---
        self.tab_help = ttk.Frame(self.notebook, padding=(8, 8))
        self.notebook.add(self.tab_help, text="Help")

        # Mount your existing tab classes inside these frames
        self.case_processing_view = CaseProcessingTab(self.tab_case)
        self.case_processing_view.pack(fill=tk.BOTH, expand=True)

        self.compare_view = CompareTab(self.tab_compare)
        self.compare_view.pack(fill=tk.BOTH, expand=True)

        self.help_view = HelpTab(self.tab_help)  # NEW
        self.help_view.pack(fill=tk.BOTH, expand=True)

        # Optional: If your tab classes support external logging / status hooks
        # you can wire them up here without breaking anything.
        self._try_wire_hooks()

    def _build_status_bar(self):
        bar = ttk.Frame(self.master, padding=(10, 6))
        bar.pack(side=tk.BOTTOM, fill=tk.X)

        ttk.Separator(self.master, orient="horizontal").pack(side=tk.BOTTOM, fill=tk.X)

        status_label = ttk.Label(bar, textvariable=self._status_var, style="Status.TLabel")
        status_label.pack(side=tk.LEFT)

        # Right side “hint”
        hint = ttk.Label(
            bar,
            text="Tip: Add comparisons to the queue → Build queued workbook",
            style="Status.TLabel",
        )
        hint.pack(side=tk.RIGHT)

    # ---------------- Hooks ---------------- #

    def set_status(self, text: str):
        """Tabs can call this to update the status bar."""
        self._status_var.set(text)
        self.master.update_idletasks()

    def _try_wire_hooks(self):
        """
        Non-breaking optional wiring:
        - If CompareTab has `external_log_func`, we could set it.
        - If tabs want status updates, they can call self.master_app.set_status(...)
        """
        # Give tabs a reference to the app if they want to call set_status()
        for view in (self.case_processing_view, self.compare_view, self.help_view):
            try:
                view.master_app = self
            except Exception:
                pass


def run():
    root = tk.Tk()
    app = App(root)
    # app is a ttk.Frame mounted by itself, but we already packed everything via root frames
    root.mainloop()


if __name__ == "__main__":
    run()