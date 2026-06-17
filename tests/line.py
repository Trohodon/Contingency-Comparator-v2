import tkinter as tk
from tkinter import filedialog, messagebox
import win32com.client


def count_vba_lines():
    excel_file = filedialog.askopenfilename(
        title="Select Excel File",
        filetypes=[
            ("Excel Macro Files", "*.xlsm *.xls"),
            ("All Files", "*.*")
        ]
    )

    if not excel_file:
        return

    excel = None
    wb = None

    try:
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False

        wb = excel.Workbooks.Open(excel_file)

        results = []
        total_lines = 0

        for vb_component in wb.VBProject.VBComponents:
            lines = vb_component.CodeModule.CountOfLines
            results.append((vb_component.Name, lines))
            total_lines += lines

        results.sort(key=lambda x: x[1], reverse=True)

        output.delete("1.0", tk.END)

        for name, lines in results:
            output.insert(tk.END, f"{name:<35} {lines} lines\n")

        output.insert(tk.END, "\n" + "-" * 50 + "\n")
        output.insert(tk.END, f"TOTAL VBA LINES: {total_lines}\n")

    except Exception as e:
        messagebox.showerror("Error", str(e))

    finally:
        if wb:
            wb.Close(False)
        if excel:
            excel.Quit()


root = tk.Tk()
root.title("Excel VBA Line Counter")
root.geometry("600x450")

button = tk.Button(root, text="Select Excel File", command=count_vba_lines, width=25)
button.pack(pady=15)

output = tk.Text(root, width=70, height=22)
output.pack(padx=10, pady=10)

root.mainloop()
