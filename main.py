import tkinter as tk
from tkinter import ttk, messagebox
from datetime import datetime
import openpyxl
from openpyxl import Workbook
import os
from tkinter import filedialog

# Excel File Setup
file_path = "Monthly_Bills.xlsx"

if not os.path.exists(file_path):
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Bills"
    sheet.append(["Account Number", "Month", "Bill Amount", "Paid Amount", "Remaining Balance", "Date", "Time"])
    workbook.save(file_path)

# Functions
def submit_data():
    account = account_var.get()
    bill_amount = bill_var.get()
    paid_amount = amount_paid_var.get()
    remaining_balance = remaining_balance_var.get()
    date = date_var.get()
    time = time_var.get()

    if not account or not bill_amount or not paid_amount:
        messagebox.showerror("Input Error", "All fields are required.")
        return

    # Excel Writing
    workbook = openpyxl.load_workbook(file_path)
    sheet = workbook.active
    sheet.append([account, datetime.now().strftime("%B %Y"), bill_amount, paid_amount, remaining_balance, date, time])
    workbook.save(file_path)
    messagebox.showinfo("Success", "Bill data saved successfully!")
    clear_fields()

def calculate_balance():
    try:
        bill = float(bill_var.get())
        paid = float(amount_paid_var.get())
        remaining = bill - paid
        remaining_balance_var.set(f"{remaining:.2f}")
    except ValueError:
        messagebox.showerror("Calculation Error", "Enter valid numeric values for bill and paid amounts.")

def clear_fields():
    account_var.set("")
    bill_var.set("")
    amount_paid_var.set("")
    remaining_balance_var.set("")
    date_var.set(datetime.now().strftime("%Y-%m-%d"))
    time_var.set(datetime.now().strftime("%H:%M:%S"))

def edit_excel():
    password = password_var.get()
    if password == "admin002$":
        os.system(f'start excel "{file_path}"')
    else:
        messagebox.showerror("Access Denied", "Invalid password.")

# Upload file functionality
def upload_file():
    filetypes = (("PDF files", "*.pdf"), ("JPEG files", "*.jpg;*.jpeg"), ("All files", "*.*"))
    file_path = filedialog.askopenfilename(filetypes=filetypes)
    if file_path:
        messagebox.showinfo("File Upload", f"File uploaded: {os.path.basename(file_path)}")

# Create button with icon
def create_button_with_icon(master, text, command, icon_path, row, col):
    button = tk.Button(master, text=text, command=command, font=("Arial", 12), bg="lightgreen")
    try:
        icon = tk.PhotoImage(file=icon_path)  # Load the icon image
        button.config(image=icon, compound="left")
        button.image = icon  # Keep a reference to the image to prevent garbage collection
    except Exception as e:
        print(f"Error loading icon: {e}")
    button.grid(row=row, column=col, padx=10, pady=5)
    return button

# GUI Setup
root = tk.Tk()
root.title("Monthly Bill Management System")
root.geometry("600x400")

# Variables
account_var = tk.StringVar()
bill_var = tk.StringVar()
amount_paid_var = tk.StringVar()
remaining_balance_var = tk.StringVar()
date_var = tk.StringVar()
time_var = tk.StringVar()
password_var = tk.StringVar()

# Initialize date and time
date_var.set(datetime.now().strftime("%Y-%m-%d"))
time_var.set(datetime.now().strftime("%H:%M:%S"))

# Layout
tk.Label(root, text="Account Number", font=("Arial", 12)).grid(row=0, column=0, padx=10, pady=5, sticky="w")
account_menu = ttk.Combobox(root, textvariable=account_var, values=[
    "4515215604 (Home Up)",
    "4510306709 (Home Down)",
    "4521245706 (Shop)",
    "4523059306 (Loku Land)",
    "4500310703 (Ahangama)"
], state="readonly")
account_menu.grid(row=0, column=1, padx=10, pady=5)

tk.Label(root, text="Total Bill Amount", font=("Arial", 12)).grid(row=1, column=0, padx=10, pady=5, sticky="w")
tk.Entry(root, textvariable=bill_var, font=("Arial", 12)).grid(row=1, column=1, padx=10, pady=5)

tk.Label(root, text="Amount Paid", font=("Arial", 12)).grid(row=2, column=0, padx=10, pady=5, sticky="w")
tk.Entry(root, textvariable=amount_paid_var, font=("Arial", 12)).grid(row=2, column=1, padx=10, pady=5)

tk.Label(root, text="Remaining Balance", font=("Arial", 12)).grid(row=3, column=0, padx=10, pady=5, sticky="w")
tk.Entry(root, textvariable=remaining_balance_var, font=("Arial", 12), state="readonly").grid(row=3, column=1, padx=10, pady=5)

tk.Label(root, text="Date", font=("Arial", 12)).grid(row=4, column=0, padx=10, pady=5, sticky="w")
tk.Entry(root, textvariable=date_var, font=("Arial", 12), state="readonly").grid(row=4, column=1, padx=10, pady=5)

tk.Label(root, text="Time", font=("Arial", 12)).grid(row=5, column=0, padx=10, pady=5, sticky="w")
tk.Entry(root, textvariable=time_var, font=("Arial", 12), state="readonly").grid(row=5, column=1, padx=10, pady=5)

tk.Button(root, text="Calculate Balance", command=calculate_balance, font=("Arial", 12), bg="lightblue").grid(row=6, column=0, padx=10, pady=20)
tk.Button(root, text="Submit", command=submit_data, font=("Arial", 12), bg="lightgreen").grid(row=6, column=1, padx=10, pady=20)

tk.Label(root, text="Admin Password", font=("Arial", 12)).grid(row=7, column=0, padx=10, pady=5, sticky="w")
tk.Entry(root, textvariable=password_var, font=("Arial", 12), show="*").grid(row=7, column=1, padx=10, pady=5)

# Create "Upload Bill" button with icon
upload_button = create_button_with_icon(root, "Upload Bill", upload_file, "upload_icon.png", 0, 2)

tk.Button(root, text="Edit Excel", command=edit_excel, font=("Arial", 12), bg="orange").grid(row=8, column=0, columnspan=2, pady=10)

# Run GUI
root.mainloop()
