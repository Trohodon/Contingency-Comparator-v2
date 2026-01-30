# main.py
# Entry point for the Contingency Comparison Tool (DCC)
# - Supports PyInstaller onefile/onedir
# - Supports PyInstaller --splash (closes it once Tk is ready)
# - Shows an internal Tk "Loading..." screen while GUI constructs
# - Fixes taskbar icon consistency on Windows via AppUserModelID
# - Supports hidden tool: app.exe --menu-one

from __future__ import annotations

import os
import sys
import tkinter as tk
from tkinter import ttk


def resource_path(relative_path: str) -> str:
    """
    Return an absolute path to a resource.
    Works for:
      - normal python runs
      - PyInstaller --onefile / --onedir runs
    """
    try:
        base_path = sys._MEIPASS  # type: ignore[attr-defined]
    except Exception:
        base_path = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_path, relative_path)


def _set_windows_appusermodelid(app_id: str = "Dominion.DCC.ContingencyComparator"):
    """
    Helps Windows consistently use the correct taskbar icon for this process.
    Safe no-op on non-Windows.
    """
    if os.name != "nt":
        return
    try:
        import ctypes
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(app_id)
    except Exception:
        pass


def _set_app_icon(root: tk.Tk):
    """
    Set window icon. Taskbar icon primarily comes from EXE icon (--icon),
    but this still helps in some environments.
    """
    ico_path = resource_path(os.path.join("assets", "app.ico"))
    if os.path.exists(ico_path):
        try:
            root.iconbitmap(ico_path)
        except Exception:
            pass


def _close_pyinstaller_splash():
    """
    If built with PyInstaller --splash, close it once Tk is up.
    Safe no-op for normal python runs.
    """
    try:
        import pyi_splash  # type: ignore
        pyi_splash.close()
    except Exception:
        pass


def _show_loading_window(root: tk.Tk) -> tk.Toplevel:
    """
    Small internal loading window so users immediately see something
    while the main App() is building.
    This does NOT cover PyInstaller onefile extraction time (use --splash for that).
    """
    # Hide the main root until ready
    root.withdraw()

    win = tk.Toplevel(root)
    win.title("Loading…")
    win.resizable(False, False)

    # Make it feel like a splash: no maximize/minimize, centered, stays on top briefly
    try:
        win.attributes("-topmost", True)
    except Exception:
        pass

    # Basic "card" style
    outer = tk.Frame(win, bg="white", bd=1, relief="solid")
    outer.pack(fill="both", expand=True)

    outer.grid_columnconfigure(0, weight=1)

    title = tk.Label(
        outer,
        text="Contingency Comparator",
        bg="white",
        fg="#0B2F5B",
        font=("Segoe UI", 14, "bold"),
        padx=18,
        pady=14,
    )
    title.grid(row=0, column=0, sticky="ew")

    msg = tk.Label(
        outer,
        text="Starting up…",
        bg="white",
        fg="#5C6773",
        font=("Segoe UI", 10),
        padx=18,
        pady=0,
    )
    msg.grid(row=1, column=0, sticky="ew")

    bar = ttk.Progressbar(outer, mode="indeterminate")
    bar.grid(row=2, column=0, sticky="ew", padx=18, pady=(14, 18))
    bar.start(12)

    # Optional image (PNG). If missing, no big deal.
    # Put an image at: assets/loading.png
    try:
        png_path = resource_path(os.path.join("assets", "loading.png"))
        if os.path.exists(png_path):
            # PhotoImage supports PNG on modern Tk builds
            img = tk.PhotoImage(file=png_path)
            # Keep a reference so it doesn't get GC'd
            win._loading_img = img  # type: ignore[attr-defined]
            img_lbl = tk.Label(outer, image=img, bg="white")
            img_lbl.grid(row=3, column=0, pady=(0, 14))
    except Exception:
        pass

    # Center the loading window
    win.update_idletasks()
    w = win.winfo_width()
    h = win.winfo_height()
    x = (win.winfo_screenwidth() // 2) - (w // 2)
    y = (win.winfo_screenheight() // 2) - (h // 2)
    win.geometry(f"{w}x{h}+{x}+{y}")

    return win


def main():
    # Fix taskbar identity/icon behavior on Windows early
    _set_windows_appusermodelid()

    # Easter egg mode: app.exe --menu-one
    if "--menu-one" in sys.argv:
        _close_pyinstaller_splash()
        # Import only when needed (keeps normal startup lighter)
        from core.menu_one_runner import maybe_run_menu_one_from_argv
        maybe_run_menu_one_from_argv()
        return

    # Create Tk ASAP (then close PyInstaller splash)
    root = tk.Tk()
    _set_app_icon(root)
    _close_pyinstaller_splash()

    # Internal loading window while we import/build the main UI
    loading = _show_loading_window(root)

    try:
        # Import heavy GUI only after loading appears
        from gui.app import App

        app = App(root)
        app.pack(fill="both", expand=True)

    finally:
        # Remove loading UI and show real main window
        try:
            loading.destroy()
        except Exception:
            pass
        root.deiconify()

    root.mainloop()


if __name__ == "__main__":
    main()
