from asyncio.windows_events import NULL
import tkinter as tk
import csv
from datetime import datetime, timedelta
from tkcalendar import Calendar, DateEntry
from tkinter import *

def next_widget(event):
    event.widget
    if event.keysym == 'Up':
        event.widget.tk_focusPrev().focus()
    elif event.keysym == 'Down' or event.keysym == 'Return':
        event.widget.tk_focusNext().focus()
    return "break"

class BankDepositGUI:
    def __init__(self, master):
        self.master = master
        master.title("Bank Deposit")
        self.deposit_data = {}

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
        self.date_entry.bind("<<DateEntrySelected>>", self.get_entries)
        self.deposit_button = tk.Button(master, text="Deposit", command=self.deposit)
        self.complete_button = tk.Button(master, text="Complete", command=self.generate_report)

        # Grid widgets
        self.date_label.grid(row=0, column=0, padx=5, pady=5, sticky='w')
        self.date_entry.grid(row=0, column=2, padx=5, pady=5)
        self.number_label.grid(row=1, column=0, padx=5, pady=5)

        # Display product value
        self.product_var = {}
        row_idx = 1
        for cash_amount in self.cash_list:
            tk.Label(master, text=f"${'{:.2f}'.format(float(cash_amount))}").grid(row=row_idx, column=1, padx=5, pady=5)
            self.amount_entry[cash_amount].grid(row=row_idx, column=2, padx=5, pady=5)
            self.product_var[cash_amount] = StringVar(value = "-")
            tk.Label(master, textvariable=self.product_var[cash_amount]).grid(row=row_idx, column=3, padx=5, pady=5)
            self.amount_entry[cash_amount].bind('<KeyRelease>', lambda event: self.update_value(self.message_text_var))
            row_idx=row_idx+1
        self.message_text_var = StringVar()
        self.message_text=tk.Label(master, textvariable=self.message_text_var).grid(row=row_idx, column=0, columnspan = 4, padx=5, pady=5)
        self.deposit_button.grid(row=row_idx+1, column=1, padx=5, pady=5)
        self.complete_button.grid(row=row_idx+2, column=1, padx=5, pady=5)

        # Use arrow key to navigate the entry
        self.master.bind('<Up>', next_widget)
        self.master.bind('<Down>', next_widget)
        self.master.bind('<Return>', next_widget)
        self.master.bind("<plus>", self.deposit)

    def update_value(self, message_text_var):
        total_amount=0
        for cash_amount in self.cash_list:
            try:
                # Get the values from the Entry widgets and convert to integers
                if self.amount_var[cash_amount].get():
                    value1 = int(self.amount_var[cash_amount].get())
                    value2 = float(cash_amount)
                    # Calculate the product
                    product = value1 * value2
                    # Add it to the total_amount
                    total_amount += round(product, 2)
                    # Update the StringVar associated with the Label widget
                    self.product_var[cash_amount].set("$"+str("{:.2f}".format(product)))
                else:
                    # Update the StringVar associated with the Label widget
                    self.product_var[cash_amount].set("-")
            except ValueError:
                # If the values are not valid integers, clear the result
                self.amount_var[cash_amount].set("")
        # Set the text to display total amount
        if (message_text_var):
            message_text_var.set(f"Total amount is ${total_amount}")

    def get_entries(self, event):
        # Get the date from the user inputs
        date = self.date_entry.get_date()
        # Get entries for selected date if the entries exist
        try:
            date_entry = self.deposit_data[date]
            if date_entry:
                for cash_amount in self.cash_list:
                    n_amount = date_entry[cash_amount]["n_amount"]
                    each_amount = date_entry[cash_amount]["each_amount"]
                    if n_amount:
                        self.amount_var[cash_amount].set(str(n_amount))
                        self.product_var[cash_amount].set(str(each_amount))
                    else:
                        self.amount_var[cash_amount].set("")
                        self.product_var[cash_amount].set("-")
                # Show a message with the total amount deposited on that day
                self.message_text_var.set(f"Daily amount: ${date_entry['total_daily_amount']}")
        except:
            # Clear entry fields
            for cash_amount in self.cash_list:
                self.amount_var[cash_amount].set("")
                self.product_var[cash_amount].set("")
                self.message_text_var.set(f"")

    def deposit(self, event=None):
        # Get the date from the user inputs
        date = self.date_entry.get_date()

        if any([self.amount_entry[cash_amount].get() for cash_amount in self.cash_list]):
            # Save user inputs
            self.deposit_data[date] = {}
            total_daily_amount = 0
            for cash_amount in self.cash_list:
                amount = round(float(cash_amount),2)
                n_amount = int(self.amount_entry[cash_amount].get()) if self.amount_entry[cash_amount].get() else 0
                each_amount = amount*n_amount
                total_daily_amount += each_amount
                # Save the data for each date
                self.deposit_data[date][cash_amount]={}
                self.deposit_data[date][cash_amount]["n_amount"]=n_amount
                self.deposit_data[date][cash_amount]["each_amount"]=each_amount
            self.deposit_data[date]["total_daily_amount"]=total_daily_amount

            # Show a confirmation message with the total amount deposited
            self.message_text_var.set(f"Deposit Successful, Daily amount deposited: ${total_daily_amount:.2f}")
            
            # Clear entry fields
            for cash_amount in self.cash_list:
                self.amount_var[cash_amount].set("")

            # Reset the cursor to the top entry
            self.amount_entry[self.cash_list[0]].focus_set()

            # Set the date to the next day
            self.date_entry.set_date(self.date_entry.get_date()+timedelta(days=1))
            # Clear the product var values
            self.update_value(NULL) #Don't update the message var

    def generate_report(self, event=None):
        # Check if there is any entry
        if any([self.amount_entry[cash_amount].get() for cash_amount in self.cash_list]):
            self.deposit()

        # Get the deposit data and write it in an excel file
        report = []
        total_amount = 0
        for date in sorted(self.deposit_data.keys(), reverse = False):
            daily_report = []
            for cash_amount in self.cash_list:
                n_amount = self.deposit_data[date][cash_amount]["n_amount"]
                each_amount = self.deposit_data[date][cash_amount]["each_amount"]
                # Create a list with the deposit data
                if n_amount > 0:
                    daily_report.append([date.strftime("%A, %m/%d/%Y"), f"${cash_amount}", f"{n_amount}", f"${each_amount}"]) if len(daily_report)==0 else daily_report.append(["", f"${cash_amount}", f"{n_amount}",f"${each_amount}"])
            report.extend(daily_report)
            total_daily_amount = self.deposit_data[date]["total_daily_amount"]
            report.append(["Total","","",f"${total_daily_amount}"])
            total_amount += total_daily_amount

        # Write the data to a CSV file
        with open('deposits.csv', 'a', newline='') as csvfile:
            writer = csv.writer(csvfile)
            for deposit_row in report:
                writer.writerow(deposit_row) 

        # Show a confirmation message with the total amount deposited
        self.message_text_var.set(f"Report Generated Successful, Total amount deposited: ${total_amount:.2f}")
        
        # Clear entry fields
        for cash_amount in self.cash_list:
            self.amount_var[cash_amount].set("")

        # Reset the cursor to the top entry
        self.amount_entry[self.cash_list[0]].focus_set()

        # Clear the product var values
        self.update_value(NULL) #Don't update the message var

        # Disable the buttons
        self.deposit_button.config(state="disabled")
        self.complete_button.config(state="disabled")

root = tk.Tk()
bank_deposit_gui = BankDepositGUI(root)
root.mainloop()
