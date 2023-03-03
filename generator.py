import tkinter as tk
from tkinter import filedialog
import os, sys
import pandas as pd

from HB_CashDeposit import generate_report
from excel_pdf import generate_pdf

class Generator_gui:
    def __init__(self):
        self.root = tk.Tk()
        self.root.geometry("600x250")
        self.root.title("Cash Deposit Report Generator")
        # Create text to browse for the input .xlsx/.csv file
        self.frame_browse_directory = tk.Frame(self.root, borderwidth=25)
        self.frame_browse_directory.pack(fill=tk.X)
        self.label_browse_directory = tk.Label(
            self.frame_browse_directory,
            text="File:",
        )
        self.label_browse_directory.pack(side=tk.TOP, anchor=tk.NW)
        # Create a text window to show the chosen path
        self.text_browse_directory = tk.Text(
            self.frame_browse_directory,
            height=1,
            width=60,
            bg="light grey",
        )
        self.text_browse_directory.pack(side=tk.LEFT)
        self.text_browse_directory.insert(1.0, "Please choose a .xlsx or.csv file")
        self.text_browse_directory.config(state=tk.DISABLED)
        # Create button to open the browse dialog window
        self.button_browse_directory = tk.Button(
            self.frame_browse_directory,
            text="Browse",
            command=self.get_file_path_for_match_file,
        ).pack(side=tk.RIGHT)
        # # Create a text window to get the sheet name
        self.frame_sheet_name = tk.Frame(self.root, padx=25)
        self.frame_sheet_name.pack(fill=tk.X)
        self.label_sheet_name = tk.Label(
            self.frame_sheet_name,
            text="Sheet name:",
        )
        self.label_sheet_name.pack(side=tk.LEFT)
        self.text_sheet_name = tk.Text(self.frame_sheet_name, height=1, width=58)
        self.text_sheet_name.pack(side=tk.RIGHT)
        self.text_sheet_name.insert(1.0, "Sheet1")
        ## Create dropdown menu to select the number of courts
        self.frame_select_court_numer = tk.Frame(
            self.root,
            padx=25,
        )

        ## Create frame for the run scheduler button
        self.frame_run_generator = tk.Frame(
            self.root,
            borderwidth=25,
        )
        self.frame_run_generator.pack(side=tk.BOTTOM, fill=tk.X)
        # Create run status text
        self.label_run_status = tk.Label(self.frame_run_generator, text="")
        self.label_run_status.pack(side=tk.TOP)
        # Create button to run the scheduler
        self.button_run_scheduler = tk.Button(
            self.frame_run_generator,
            text="Run generator",
            state=tk.DISABLED,
            command=self.get_report,
        )
        self.button_run_scheduler.pack(side=tk.BOTTOM)

        # Defining custom protocol for the 'x' button
        self.root.protocol("WM_DELETE_WINDOW", self.gui_exit_function)
        # Run the mainloop
        self.root.mainloop()

    def get_file_path_for_match_file(self):
        root = tk.Tk()
        root.withdraw()  # use to hide tkinter window
        type = [("Excel files", ".xlsx .xls"), ("CSV files", ".csv")]
        message = "Please select a csv file with the daily cash deposit entry"
        self.file_path = filedialog.askopenfilename(
            parent=root,
            initialdir=os.getcwd(),
            title=message,
            filetypes=type,
        )
        if len(self.file_path) > 0:
            self.text_browse_directory.config(state=tk.NORMAL)
            self.text_browse_directory.delete(1.0, tk.END)
            self.text_browse_directory.insert(1.0, self.file_path)
            self.text_browse_directory.config(state=tk.DISABLED)
            # Enable run scheduler buttons
            self.button_run_scheduler.config(state=tk.NORMAL)
        else:
            self.text_browse_directory.delete(1.0, tk.END)
            self.text_browse_directory.insert(1.0, "No file chosen")
            self.file_path = ""

    def get_report(self):
        ## Process the data from the input file
        selected_sheet_name = self.text_sheet_name.get(1.0, "end-1c")
        #Strip the extension from the filepath
        filename = os.path.splitext(self.file_path)[0]
        # Generate the report
        try:
            output_filename = generate_report(filename, selected_sheet_name)
            generate_pdf(os.path.abspath(output_filename))
            # Update the run status text
            self.label_run_status.config(
                text=output_filename.split("/")[-1] +".xlsx and .pdf \ngenerated successfully", foreground="green"
            )
            self.button_run_scheduler.config(state=tk.DISABLED)
        except:
            # Update the run status text
            self.label_run_status.config(
                text="Run failed, please ensure correct file or sheet name.", foreground="red"
            )

    def gui_exit_function(self):
        self.root.destroy()
        sys.exit()


if __name__ == "__main__":
    generator_gui = Generator_gui()