# Cash_Deposit

This is a repository for an app to generate a report for a cash deposit.

## Getting started
- Install python and pip from here
- Open cmd prompt if you are on Windows
- Run `pip install -r requirements.txt`

### Directly through the python scripts
- Option 1:
    - Run  `generator.py`
    - Follow the GUI prompt
- Option 2:
    - Put the file in `Data` folder, e.g. `CashDeposit.xlsx`
    - Run `HB_CashDeposit.py`
    - The generated .csv file will be located in the same `Data` folder
    - Run `excel_pdf.py`
        - Make sure that the name of the input file corresponds to the generated .csv file generated in the previous step
### Through an .exe file
- Navigate to directory where the python file to be compiled is located, e.g., `path_to_local_repository/src`
- Run `pyinstaller -F filename.py` or `pyinstaller --onefile filename.py`, e.g., `pyinstaller -F generator.py`
    - The executable will be created under a folder called `dist`!
    - Note: If it crashes because of babel.numbers, run `pyinstaller -F --hidden-import "babel.numbers" src/cash_input_gui.py` instead from the main directory
- Run the app and follow the GUI prompt

