import tkinter as tk
import csv
from datetime import datetime, timedelta
from tkcalendar import Calendar, DateEntry
from tkinter import *

def update_value(amount_vars, cash_list, product_vars, message_text_var):
    total_amount=0
    for cash_amount in cash_list:
        try:
            # Get the values from the Entry widgets and convert to integers
            if amount_vars[cash_amount].get():
                value1 = int(amount_vars[cash_amount].get())
                value2 = float(cash_amount)
                # Calculate the product
                product = value1 * value2
                # Add it to the total_amount
                total_amount +=product
                # Update the StringVar associated with the Label widget
                product_vars[cash_amount].set("$"+str("{:.2f}".format(product)))
            else:
                # Update the StringVar associated with the Label widget
                product_vars[cash_amount].set("-")
        except ValueError:
            # If the values are not valid integers, clear the result
            amount_vars[cash_amount].set("")
    # Set the text to display total amount
    message_text_var.set(f"Total amount is ${total_amount}")

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
        self.date_entry = DateEntry(master, locale= "en_AU", width=12, background='darkblue', foreground='white', borderwidth=2)
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
            self.amount_entry[cash_amount].bind('<KeyRelease>', lambda event: update_value(self.amount_var, self.cash_list, self.product_var, self.message_text_var))
            row_idx=row_idx+1
        self.message_text_var = StringVar()
        self.message_text=tk.Label(master, textvariable=self.message_text_var).grid(row=row_idx, column=0, columnspan = 4, padx=5, pady=5)
        self.deposit_button.grid(row=row_idx+1, column=1, padx=5, pady=5)


    def deposit(self):
        # Get the date from the user inputs
        date = self.date_entry.get_date().strftime("%A, %m/%d/%Y")

        # Validate user inputs
        deposit_data = []
        total_amount = 0
        for cash_amount in self.cash_list:
            amount = round(float(cash_amount),2)
            n_amount = int(self.amount_entry[cash_amount].get()) if self.amount_entry[cash_amount].get() else 0
            each_amount = amount*n_amount
            total_amount += each_amount
            # Create a list with the deposit data
            if n_amount > 0:
                deposit_data.append([date, f"${cash_amount}", f"${each_amount}"]) if len(deposit_data)==0 else deposit_data.append(["", f"${cash_amount}", f"${each_amount}"])
        deposit_data.append(["Total","",f"${total_amount}"])

        # Write the data to a CSV file
        with open('deposits.csv', 'a', newline='') as csvfile:
            writer = csv.writer(csvfile)
            for deposit_row in deposit_data:
                writer.writerow(deposit_row) 

        # Show a confirmation message with the total amount deposited
        self.message_text_var.set(f"Deposit Successful, Total amount deposited: ${total_amount:.2f}")
        
        # Clear entry fields
        for cash_amount in self.cash_list:
            self.amount_var[cash_amount].set("")

        # Set the date to today
        self.date_entry.set_date(self.date_entry.get_date()+timedelta(days=1))
        # Update the product var values
        update_value(self.amount_var, self.cash_list, self.product_var, self.self.message_text_var)

root = tk.Tk()
bank_deposit_gui = BankDepositGUI(root)
root.mainloop()
