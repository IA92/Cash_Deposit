# Cash_Deposit
This is a repository for an app to generate a report for a cash deposit.

## Getting started
- Install python and pip from here
- Open cmd prompt if you are on Windows
- Run `pip install -r requirements.txt`

### Directly through the python scripts
- Option 1:
    - Adjust the `generated_excel_filename` in report.py
        - Optional: Add/omit a `Data` folder to make it neater/straightforward
    - Run  `cash_input_gui.py`
    - Follow the GUI prompt
### Through an .exe file
- Navigate to main project directory, e.g., `path_to_local_repository`
- Run `pyinstaller -F filename.py` or `pyinstaller --onefile filename.py`, e.g., `pyinstaller -F generator.py`
    - The executable will be created under a folder called `dist`!
    - Note: If it crashes because of babel.numbers, run `pyinstaller -F --hidden-import "babel.numbers" src/cash_input_gui.py` instead from the main directory
- Run the app and follow the GUI prompt