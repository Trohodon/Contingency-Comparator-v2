# main.py
import tkinter as tk
from gui.app import App

if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()