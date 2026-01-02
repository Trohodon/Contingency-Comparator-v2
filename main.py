# main.py
# Entry point for the Contingency Comparison Tool (DCC)
# - Sets Windows app icon (titlebar/taskbar) from assets/app.ico
# - Shows PyInstaller splash (if built with --splash) and closes it when Tk is ready
# - Handles normal runs AND PyInstaller onefile/onedir runs via resource_path()

import os
import sys
import tkinter as tk

from gui.app import App


def resource_path(relative_path: str) -> str:
    """
    Return an absolute path to a resource.
    Works for:
      - normal python runs
      - PyInstaller --onefile / --onedir runs
    """
    try:
        # PyInstaller extraction/temporary folder
        base_path = sys._MEIPASS  # type: ignore[attr-defined]
    except Exception:
        # Folder containing this file
        base_path = os.path.dirname(os.path.abspath(__file__))

    return os.path.join(base_path, relative_path)


def _set_app_icon(root: tk.Tk):
    """
    Set the top-left window icon + (often) taskbar icon on Windows.
    Notes:
      - .ico is best on Windows.
      - Taskbar behavior can vary when running from python vs packaged exe.
      - For packaged exe, ALSO pass --icon=... to PyInstaller (you already are).
    """
    ico_path = resource_path(os.path.join("assets", "app.ico"))
    if os.path.exists(ico_path):
        try:
            root.iconbitmap(ico_path)
        except Exception:
            # Some environments may fail iconbitmap; ignore gracefully
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


def main():
    root = tk.Tk()

    # Set icon ASAP (top-left + often taskbar)
    _set_app_icon(root)

    # If we launched with PyInstaller splash, close it now that Tk exists
    _close_pyinstaller_splash()

    # Build and run app
    app = App(root)
    app.pack(fill="both", expand=True)

    root.mainloop()


if __name__ == "__main__":
    main()