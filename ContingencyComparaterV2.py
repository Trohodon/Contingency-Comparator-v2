from __future__ import annotations

import os
import sys
import tkinter as tk


def resource_path(relative_path: str) -> str:
    """
    Works for both dev runs and PyInstaller onefile/onedir.
    """
    try:
        base_path = sys._MEIPASS  # type: ignore[attr-defined]
    except Exception:
        base_path = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_path, relative_path)


def _set_windows_appusermodelid(app_id: str = "Dominion.DCC.ContingencyComparator") -> None:
    """
    Helps Windows group the app correctly and can impact taskbar identity/icon behavior.
    MUST be called before creating any windows.
    """
    if os.name != "nt":
        return
    try:
        import ctypes
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(app_id)
    except Exception:
        pass


def _pyi_splash_update(text: str) -> None:
    """
    Optional: updates the PyInstaller splash text (if supported).
    Safe no-op if splash module isn't present.
    """
    try:
        import pyi_splash  # type: ignore
        try:
            pyi_splash.update_text(text)
        except Exception:
            pass
    except Exception:
        pass


def _close_pyinstaller_splash() -> None:
    """
    Close the PyInstaller splash screen (if present).
    """
    try:
        import pyi_splash  # type: ignore
        pyi_splash.close()
    except Exception:
        pass


def _set_app_icon(root: tk.Tk) -> None:
    """
    Sets the Tk window icon.
    - iconbitmap(.ico) is the primary Windows method
    - iconphoto(.png) can help with Alt-Tab / some window managers
    """
    # 1) ICO (Windows)
    ico_path = resource_path(os.path.join("assets", "app.ico"))
    if os.path.exists(ico_path):
        try:
            root.iconbitmap(ico_path)
        except Exception:
            pass

    # 2) PNG fallback (optional but recommended)
    png_path = resource_path(os.path.join("assets", "app_256.png"))
    if os.path.exists(png_path):
        try:
            img = tk.PhotoImage(file=png_path)
            # keep a reference so Tk doesn't garbage-collect it
            root._app_icon_img = img  # type: ignore[attr-defined]
            # True = apply to all future toplevels too
            root.iconphoto(True, img)
        except Exception:
            pass


def main() -> None:
    _set_windows_appusermodelid()

    if "--menu-one" in sys.argv:
        _close_pyinstaller_splash()
        from core.menu_one_runner import maybe_run_menu_one_from_argv
        maybe_run_menu_one_from_argv()
        return

    _pyi_splash_update("Starting...")

    root = tk.Tk()
    _set_app_icon(root)

    # Hide the root window while the UI constructs
    root.withdraw()

    _pyi_splash_update("Loading UI...")

    from gui.app import App
    app = App(root)
    app.pack(fill="both", expand=True)

    # Now that the UI exists, close splash and show the window
    _close_pyinstaller_splash()

    # Ensure icon is applied again after UI creation (sometimes helps)
    _set_app_icon(root)

    root.deiconify()
    root.mainloop()


if __name__ == "__main__":
    main()
