# main.py
from __future__ import annotations

import os
import sys
import tkinter as tk
from tkinter import ttk


def resource_path(relative_path: str) -> str:
    try:
        base_path = sys._MEIPASS  # type: ignore[attr-defined]
    except Exception:
        base_path = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_path, relative_path)


def _set_windows_appusermodelid(app_id: str = "Dominion.DCC.ContingencyComparator"):
    if os.name != "nt":
        return
    try:
        import ctypes
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(app_id)
    except Exception:
        pass


def _set_app_icon(root: tk.Tk):
    ico_path = resource_path(os.path.join("assets", "app.ico"))
    if os.path.exists(ico_path):
        try:
            root.iconbitmap(ico_path)
        except Exception:
            pass


def _close_pyinstaller_splash():
    try:
        import pyi_splash  # type: ignore
        pyi_splash.close()
    except Exception:
        pass


def _show_loading_window(root: tk.Tk) -> tk.Toplevel:
    """
    Internal loading window (shows after Python starts).
    For onefile extraction time, use PyInstaller --splash.
    """
    root.withdraw()

    win = tk.Toplevel(root)
    win.title("Loading…")
    win.resizable(False, False)

    try:
        win.attributes("-topmost", True)
    except Exception:
        pass

    outer = tk.Frame(win, bg="white", bd=1, relief="solid")
    outer.pack(fill="both", expand=True)
    outer.grid_columnconfigure(0, weight=1)

    tk.Label(
        outer,
        text="Contingency Comparator",
        bg="white",
        fg="#0B2F5B",
        font=("Segoe UI", 14, "bold"),
        padx=18,
        pady=14,
    ).grid(row=0, column=0, sticky="ew")

    tk.Label(
        outer,
        text="Starting up…",
        bg="white",
        fg="#5C6773",
        font=("Segoe UI", 10),
        padx=18,
        pady=0,
    ).grid(row=1, column=0, sticky="ew")

    bar = ttk.Progressbar(outer, mode="indeterminate")
    bar.grid(row=2, column=0, sticky="ew", padx=18, pady=(14, 12))
    bar.start(12)

    # Use your existing splash.png as the image inside the loading window
    try:
        png_path = resource_path(os.path.join("assets", "splash.png"))
        if os.path.exists(png_path):
            img = tk.PhotoImage(file=png_path)
            win._loading_img = img  # type: ignore[attr-defined]
            tk.Label(outer, image=img, bg="white").grid(row=3, column=0, pady=(0, 14))
    except Exception:
        pass

    win.update_idletasks()
    w = win.winfo_width()
    h = win.winfo_height()
    x = (win.winfo_screenwidth() // 2) - (w // 2)
    y = (win.winfo_screenheight() // 2) - (h // 2)
    win.geometry(f"{w}x{h}+{x}+{y}")

    return win


def main():
    _set_windows_appusermodelid()

    # Easter egg mode
    if "--menu-one" in sys.argv:
        _close_pyinstaller_splash()
        from core.menu_one_runner import maybe_run_menu_one_from_argv
        maybe_run_menu_one_from_argv()
        return

    root = tk.Tk()
    _set_app_icon(root)
    _close_pyinstaller_splash()

    loading = _show_loading_window(root)

    try:
        from gui.app import App
        app = App(root)
        app.pack(fill="both", expand=True)
    finally:
        try:
            loading.destroy()
        except Exception:
            pass
        root.deiconify()

    root.mainloop()


if __name__ == "__main__":
    main()
