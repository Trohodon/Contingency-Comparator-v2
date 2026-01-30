# main.py
import os
import sys
import tkinter as tk

from gui.app import App
from core.menu_one_runner import maybe_run_menu_one_from_argv


def resource_path(relative_path: str) -> str:
    try:
        base_path = sys._MEIPASS  # type: ignore[attr-defined]
    except Exception:
        base_path = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_path, relative_path)


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


def main():
    # If launched as: app.exe --menu-one
    # run pygame and exit without starting Tk.
    if "--menu-one" in sys.argv:
        _close_pyinstaller_splash()
        if maybe_run_menu_one_from_argv():
            return
        return

    root = tk.Tk()
    _set_app_icon(root)
    _close_pyinstaller_splash()

    app = App(root)
    app.pack(fill="both", expand=True)
    root.mainloop()


if __name__ == "__main__":
    main()
