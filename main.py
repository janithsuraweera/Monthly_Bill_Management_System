import tkinter as tk
from tkinter import ttk, messagebox
from datetime import datetime
import openpyxl
from openpyxl import Workbook
import os

# Excel File Setup
file_path = "Monthly_Bills.xlsx"

# Check if file exists, create it if not
if not os.path.exists(file_path):
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Home Up"  # Default to one account sheet
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

    # Load the workbook
    workbook = openpyxl.load_workbook(file_path)

    # Check if the sheet for the account exists, create if not
    if account not in workbook.sheetnames:
        workbook.create_sheet(account)
        sheet = workbook[account]
        sheet.append(["Account Number", "Month", "Bill Amount", "Paid Amount", "Remaining Balance", "Date", "Time"])
    else:
        sheet = workbook[account]

    # Append data to the selected account's sheet
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


def delete_sheet():
    password = password_var.get()
    if password == "admin002$":
        account = account_var.get()
        if account:
            workbook = openpyxl.load_workbook(file_path)
            if account in workbook.sheetnames:
                del workbook[account]
                workbook.save(file_path)
                messagebox.showinfo("Success", f"Account '{account}' sheet deleted successfully.")
            else:
                messagebox.showerror("Error", "Sheet not found for this account.")
        else:
            messagebox.showerror("Error", "Please select an account to delete.")
    else:
        messagebox.showerror("Access Denied", "Invalid password.")


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
tk.Entry(root, textvariable=remaining_balance_var, font=("Arial", 12), state="readonly").grid(row=3, column=1, padx=10,
                                                                                              pady=5)

tk.Label(root, text="Date", font=("Arial", 12)).grid(row=4, column=0, padx=10, pady=5, sticky="w")
tk.Entry(root, textvariable=date_var, font=("Arial", 12), state="readonly").grid(row=4, column=1, padx=10, pady=5)

tk.Label(root, text="Time", font=("Arial", 12)).grid(row=5, column=0, padx=10, pady=5, sticky="w")
tk.Entry(root, textvariable=time_var, font=("Arial", 12), state="readonly").grid(row=5, column=1, padx=10, pady=5)

tk.Button(root, text="Calculate Balance", command=calculate_balance, font=("Arial", 12), bg="lightblue").grid(row=6,
                                                                                                              column=0,
                                                                                                              padx=10,
                                                                                                              pady=20)
tk.Button(root, text="Submit", command=submit_data, font=("Arial", 12), bg="lightgreen").grid(row=6, column=1, padx=10,
                                                                                              pady=20)

tk.Label(root, text="Admin Password", font=("Arial", 12)).grid(row=7, column=0, padx=10, pady=5, sticky="w")
tk.Entry(root, textvariable=password_var, font=("Arial", 12), show="*").grid(row=7, column=1, padx=10, pady=5)
tk.Button(root, text=".", command=edit_excel, font=("Arial", 12), bg="orange").grid(row=8, column=0,
                                                                                             columnspan=2, pady=10)


# Run GUI
root.mainloop()
