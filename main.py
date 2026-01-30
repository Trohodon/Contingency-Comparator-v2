# main.py
# Entry point for the Contingency Comparison Tool (DCC)
# - Sets Windows app icon (titlebar/taskbar) from assets/app.ico
# - Shows PyInstaller splash (if built with --splash) and closes it when ready
# - Handles normal runs AND PyInstaller onefile/onedir runs via resource_path()
# - Handles Menu One easter egg via "--menu-one" flag (no Tk)

import os
import sys
import tkinter as tk

from gui.app import App
from core.menu_one_runner import maybe_run_menu_one_from_argv


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


def _set_app_icon(root: tk.Tk):
    """
    Set the top-left window icon + (often) taskbar icon on Windows.
    """
    ico_path = resource_path(os.path.join("assets", "app.ico"))
    if os.path.exists(ico_path):
        try:
            root.iconbitmap(ico_path)
        except Exception:
            pass


def _close_pyinstaller_splash():
    """
    If built with PyInstaller --splash, close it.
    Safe no-op for normal python runs.
    """
    try:
        import pyi_splash  # type: ignore
        pyi_splash.close()
    except Exception:
        pass


def main():
    # --- Easter egg entrypoint (NO TK) ---
    # If this EXE was launched as: app.exe --menu-one
    # run the pygame program in this process and exit.
    if "--menu-one" in sys.argv:
        _close_pyinstaller_splash()
        if maybe_run_menu_one_from_argv():
            return
        # If flag was present but it didn't run for some reason, just exit silently.
        return

    # --- Normal GUI entrypoint ---
    root = tk.Tk()

    _set_app_icon(root)
    _close_pyinstaller_splash()

    app = App(root)
    app.pack(fill="both", expand=True)

    root.mainloop()


if __name__ == "__main__":
    main()
