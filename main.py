import os
import sys
import tkinter as tk
import ctypes

from gui.app import App


def resource_path(relative_path: str) -> str:
    """
    Works for:
      - normal runs
      - PyInstaller --onefile (sys._MEIPASS)
      - PyInstaller --onedir
    """
    try:
        base_path = sys._MEIPASS  # type: ignore[attr-defined]
    except Exception:
        base_path = os.path.dirname(os.path.abspath(__file__))

    return os.path.join(base_path, relative_path)


def set_windows_app_user_model_id(app_id: str):
    """
    Forces Windows to use THIS app identity for taskbar grouping + icon.
    Call BEFORE tk.Tk().
    """
    try:
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(app_id)
    except Exception:
        pass


def main():
    # 1) Taskbar identity (helps Windows stop using python default)
    set_windows_app_user_model_id("DCC.ContingencyComparatorV2")

    # 2) Build root
    root = tk.Tk()

    # 3) Set icon for title bar + taskbar button
    ico_path = resource_path(os.path.join("assets", "app.ico"))
    if os.path.exists(ico_path):
        try:
            root.iconbitmap(ico_path)
        except Exception:
            pass

    # 4) (Optional but often helps) also set iconphoto from PNG
    png_path = resource_path(os.path.join("assets", "app_256.png"))
    if os.path.exists(png_path):
        try:
            img = tk.PhotoImage(file=png_path)
            root.iconphoto(True, img)  # True = apply to all windows
        except Exception:
            pass

    app = App(root)
    root.mainloop()


if __name__ == "__main__":
    main()