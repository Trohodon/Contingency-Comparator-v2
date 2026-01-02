# main.py
import os
import sys
import tkinter as tk

from gui.app import App


def resource_path(relative_path: str) -> str:
    """
    Handles normal runs AND PyInstaller --onefile runs
    """
    try:
        base_path = sys._MEIPASS  # PyInstaller temp folder
    except Exception:
        base_path = os.path.dirname(os.path.abspath(__file__))

    return os.path.join(base_path, relative_path)


def main():
    root = tk.Tk()

    # ---- App icon (top-left + taskbar) ----
    ico_path = resource_path("assets/app.ico")
    if os.path.exists(ico_path):
        try:
            root.iconbitmap(ico_path)
        except Exception:
            pass  # if Windows can't apply it for some reason, don't crash

    App(root)  # App handles packing/frames internally in your version
    root.mainloop()


if __name__ == "__main__":
    main()