##Save worksheet as pdf
from win32com import client
import os

def generate_pdf(filename):
    try:
        excel = client.Dispatch("Excel.Application")
        sheets = excel.Workbooks.Open(f"{filename}.xlsx")
        work_sheets = sheets.Worksheets[1]
        work_sheets.ExportAsFixedFormat(0, f"{filename}.pdf")
        sheets.Close(True)
        excel.Quit()
        del excel
    except:
        excel = client.Dispatch("Excel.Application")
        excel.Quit()
        del excel

if __name__ == "__main__":
    generate_pdf(os.getcwd() + "\Data\Bank 2023 - NewFormat - 27Jan23")
