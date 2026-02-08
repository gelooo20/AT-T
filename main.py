import win32com.client as win32
import tkinter as tk
from tkinter import filedialog, messagebox
import csv

def update_status(text):
    status_label.config(text=text)
    status_label.update_idletasks()

def run_process(report_path, csv_path):
    if not report_path or not csv_path:
        messagebox.showwarning(
            "Missing File",
            "Please select both the report workbook and the CSV file."
        )
        return

    excel = win32.Dispatch("Excel.Application")
    excel.Visible = True
    excel.DisplayAlerts = False
    excel.AskToUpdateLinks = False
    excel.EnableEvents = False
    excel.ScreenUpdating = False

    try:
        update_status("Opening report workbook...")
        wb = excel.Workbooks.Open(report_path, UpdateLinks=0)

        update_status("Accessing Raw sheet...")
        ws = wb.Worksheets("Raw")

        update_status("Clearing Raw sheet data...")
        if ws.UsedRange.Cells.Count > 1:
            ws.UsedRange.ClearContents()

        update_status("Reading CSV file...")
        with open(csv_path, newline='', encoding='utf-8-sig') as f:
            data = list(csv.reader(f))

        if not data:
            raise Exception("CSV file is empty.")

        rows = len(data)
        cols = len(data[0])

        update_status("Writing data to Raw sheet...")
        ws.Range(
            ws.Cells(1, 1),
            ws.Cells(rows, cols)
        ).Value = data

        update_status("Refreshing pivot tables...")
        wb.RefreshAll()
        excel.CalculateUntilAsyncQueriesDone()

        update_status("Saving report...")
        wb.Save()

        update_status("Process completed successfully.")
        messagebox.showinfo(
            "Success",
            "Raw data loaded and pivot tables refreshed."
        )

    except Exception as e:
        update_status("Error occurred.")
        messagebox.showerror("Error", str(e))

    finally:
        excel.ScreenUpdating = True
        excel.EnableEvents = True
        excel.DisplayAlerts = True

def browse(entry, filetypes):
    path = filedialog.askopenfilename(filetypes=filetypes)
    if path:
        entry.delete(0, tk.END)
        entry.insert(0, path)

# ---------------- CLEAN UI ---------------- #

root = tk.Tk()
root.title("AT&T Tool")
root.geometry("520x240")
root.resizable(False, False)

tk.Label(root, text="Report Workbook").pack(pady=(10, 0))
report_entry = tk.Entry(root, width=70)
report_entry.pack()
tk.Button(
    root,
    text="Browse",
    command=lambda: browse(
        report_entry,
        [("Excel Files", "*.xlsx *.xlsm *.xls")]
    )
).pack()

tk.Label(root, text="Raw CSV File").pack(pady=(10, 0))
csv_entry = tk.Entry(root, width=70)
csv_entry.pack()
tk.Button(
    root,
    text="Browse",
    command=lambda: browse(
        csv_entry,
        [("CSV Files", "*.csv")]
    )
).pack()

tk.Button(
    root,
    text="Run Process",
    height=2,
    width=20,
    command=lambda: run_process(
        report_entry.get(),
        csv_entry.get()
    )
).pack(pady=10)

status_label = tk.Label(
    root,
    text="Waiting for files...",
    anchor="w"
)
status_label.pack(fill="x", padx=10, pady=(5, 0))

root.mainloop()
