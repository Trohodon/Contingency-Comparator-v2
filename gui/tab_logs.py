import tkinter as tk
from tkinter import ttk

class LogsTab(ttk.Frame):
    def __init__(self, master):
        super().__init__(master)
        
        self.text = tk.Text(self)
        self.text.pack(fill=tk.BOTH, expand=True)

    def write(self, msg):
        self.text.insert(tk.END, msg + "\n")
        self.text.see(tk.END)