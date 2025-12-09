from tkinter import ttk

class SettingsTab(ttk.Frame):
    def __init__(self, master):
        super().__init__(master)
        ttk.Label(self, text="Settings tab (future features)").pack(pady=20)