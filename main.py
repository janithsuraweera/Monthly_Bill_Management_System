import tkinter as tk
from tkinter import filedialog, messagebox
from datetime import date
from openpyxl import Workbook, load_workbook
import os

# Excel file setup
FILE_NAME = "bill_data.xlsx"

# Check if the file exists, if not create it
if not os.path.exists(FILE_NAME):
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Bill Data"
    sheet.append(["Account Number", "Bill Amount (Rs)", "Date", "Slip Path"])
    workbook.save(FILE_NAME)

# GUI Application
root = tk.Tk()
root.title("Monthly Bill Management")

# Variables
account_var = tk.StringVar()
bill_var = tk.StringVar()
date_var = tk.StringVar(value=date.today().strftime('%Y-%m-%d'))
slip_path = tk.StringVar()

# Functions
def upload_slip():
    file_path = filedialog.askopenfilename(
        filetypes=[("Image Files", "*.png;*.jpg;*.jpeg"), ("All Files", "*.*")]
    )
    if file_path:
        slip_path.set(file_path)
        messagebox.showinfo("Upload Success", "Slip uploaded successfully!")

def submit_data():
    account = account_var.get()
    bill_amount = bill_var.get()
    bill_date = date_var.get()
    slip = slip_path.get()

    if not account or not bill_amount or not bill_date or not slip:
        messagebox.showwarning("Missing Data", "Please fill in all fields!")
        return

    # Save data to Excel
    workbook = load_workbook(FILE_NAME)
    sheet = workbook.active
    sheet.append([account, bill_amount, bill_date, slip])
    workbook.save(FILE_NAME)

    messagebox.showinfo("Success", "Data saved successfully!")
    account_var.set("")
    bill_var.set("")
    slip_path.set("")

# Layout
tk.Label(root, text="Account Number").grid(row=0, column=0, padx=10, pady=5, sticky="w")
account_entry = tk.Entry(root, textvariable=account_var)
account_entry.grid(row=0, column=1, padx=10, pady=5)

tk.Label(root, text="Bill Amount (Rs)").grid(row=1, column=0, padx=10, pady=5, sticky="w")
bill_entry = tk.Entry(root, textvariable=bill_var)
bill_entry.grid(row=1, column=1, padx=10, pady=5)

tk.Label(root, text="Date").grid(row=2, column=0, padx=10, pady=5, sticky="w")
date_entry = tk.Entry(root, textvariable=date_var, state="readonly")
date_entry.grid(row=2, column=1, padx=10, pady=5)

tk.Button(root, text="Upload Slip", command=upload_slip).grid(row=3, column=0, padx=10, pady=5)
tk.Label(root, textvariable=slip_path, fg="blue").grid(row=3, column=1, padx=10, pady=5, sticky="w")

tk.Button(root, text="Submit", command=submit_data, bg="green", fg="white").grid(
    row=4, column=0, columnspan=2, pady=10
)

# Run the application
root.mainloop()
