import tkinter as tk
import csv
from datetime import datetime
from tkcalendar import Calendar, DateEntry
from tkinter import *

def update_value(amount_vars, cash_list, product_vars):
    for cash_amount in cash_list:
        try:
            # Get the values from the Entry widgets and convert to integers
            if amount_vars[cash_amount].get():
                value1 = int(amount_vars[cash_amount].get())
                value2 = float(cash_amount)
                # Calculate the product
                product = value1 * value2
                # Update the StringVar associated with the Label widget
                product_vars[cash_amount].set("$"+str("{:.2f}".format(product)))
            else:
                # Update the StringVar associated with the Label widget
                product_vars[cash_amount].set("-")
        except ValueError:
            # If the values are not valid integers, clear the result
            amount_vars[cash_amount].set("")

class BankDepositGUI:
    def __init__(self, master):
        self.master = master
        master.title("Bank Deposit")

        self.cash_list = ["100", "50", "20", "10", "5", "2", "1", "0.50", "0.20", "0.10", "0.05"]

        # Create widgets
        self.number_label = tk.Label(master, text="Number of bills/coins:")
        self.amount_entry = {}
        self.amount_var = {}
        for cash_amount in self.cash_list:
            self.amount_var[cash_amount] = StringVar()
            self.amount_entry[cash_amount] = tk.Entry(master, textvariable=self.amount_var[cash_amount], width=14)
        self.date_label = tk.Label(master, text="Select date:")
        self.date_entry = DateEntry(master, width=12, background='darkblue', foreground='white', borderwidth=2)
        self.deposit_button = tk.Button(master, text="Deposit", command=self.deposit)

        # Grid widgets
        self.date_label.grid(row=0, column=0, padx=5, pady=5, sticky='w')
        self.date_entry.grid(row=0, column=2, padx=5, pady=5)
        self.number_label.grid(row=1, column=0, padx=5, pady=5)

        self.product_var = {}

        row_idx = 1
        for cash_amount in self.cash_list:
            tk.Label(master, text=f"${'{:.2f}'.format(float(cash_amount))}").grid(row=row_idx, column=1, padx=5, pady=5)
            self.amount_entry[cash_amount].grid(row=row_idx, column=2, padx=5, pady=5)
            self.product_var[cash_amount] = StringVar(value = "-")
            tk.Label(master, textvariable=self.product_var[cash_amount]).grid(row=row_idx, column=3, padx=5, pady=5)
            self.amount_entry[cash_amount].bind('<KeyRelease>', lambda event: update_value(self.amount_var, self.cash_list, self.product_var))
            row_idx=row_idx+1
        self.deposit_button.grid(row=12, column=1, padx=5, pady=5)


    def deposit(self):
        # Get the amount and date from the user inputs
        amount = self.amount_split.get()
        date = self.date_entry.get()

        # Validate user inputs
        if not all(entry.get().isdigit() for entry in self.amount_entry.values()):
            # Show an error message if any of the amount fields is not a valid integer
            tk.messagebox.showerror("Error", "Please enter a valid integer in all amount fields")
            return
        total_amount = sum(int(entry.get()) for entry in self.amount_entry.values())

        # Create a list with the deposit data
        deposit_data = [datetime.now().strftime("%m/%d/%Y %H:%M:%S"), amount, date.strftime("%m/%d/%Y")]

        # Write the data to a CSV file
        with open('deposits.csv', 'a', newline='') as csvfile:
            writer = csv.writer(csvfile)
            writer.writerow(deposit_data)

        # Show a confirmation message with the total amount deposited
        tk.messagebox.showinfo("Deposit Successful", f"Total amount deposited: ${total_amount:.2f}")
        
        # Clear the amount fields
        for entry in self.amount_entry.values():
            entry.delete(0, tk.END)

        # Set the date to today
        self.date_entry.set_date(datetime.today().date())

root = tk.Tk()
bank_deposit_gui = BankDepositGUI(root)
root.mainloop()
