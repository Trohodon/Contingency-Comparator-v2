from __future__ import annotations

import os
import sys
import tkinter as tk


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


def _pyi_splash_update(text: str):
    # Optional. If update_text isn't supported, it just does nothing.
    try:
        import pyi_splash  # type: ignore
        try:
            pyi_splash.update_text(text)
        except Exception:
            pass
    except Exception:
        pass


def _close_pyinstaller_splash():
    try:
        import pyi_splash  # type: ignore
        pyi_splash.close()
    except Exception:
        pass


def main():
    _set_windows_appusermodelid()

    if "--menu-one" in sys.argv:
        _close_pyinstaller_splash()
        from core.menu_one_runner import maybe_run_menu_one_from_argv
        maybe_run_menu_one_from_argv()
        return

    _pyi_splash_update("Starting...")

    root = tk.Tk()
    _set_app_icon(root)

    # Hide the window until the UI is fully constructed
    root.withdraw()

    _pyi_splash_update("Loading UI...")

    from gui.app import App
    app = App(root)
    app.pack(fill="both", expand=True)

    # Now that the real UI exists, close splash and show the window
    _close_pyinstaller_splash()
    root.deiconify()

    root.mainloop()


if __name__ == "__main__":
    main()
