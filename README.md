# Query_output_in_excel_sheets_with_python

---

This Python script is designed to modify an existing Excel file (`Final.xlsx`) using the `openpyxl` library. It performs several tasks such as copying cell values, performing calculations, adding titles, and applying color formatting to specific cells in the Excel sheet named `df_data_details`.

## Prerequisites

Before running the script, make sure you have the following installed:

- Python (https://www.python.org/)
- `openpyxl` library (`pip install openpyxl`)

## Usage

1. Clone this repository or download the script to your local machine.

2. Make sure you have the `Final.xlsx` file in the same directory as the script. This is the Excel file that will be modified.

3. Run the Python script using your preferred Python environment. It will load the existing Excel file, apply modifications, and save the changes.

4. The modified Excel file will be saved with the same name (`Final.xlsx`) in the same directory.

5. The modified Excel file will be opened automatically for viewing.

## Modification Tasks

The script performs the following modification tasks:

- Copies the value from cell `E2` to cell `G2`.
- Calculates the value in cell `I2` using a formula.
- Adds a title "DATA VOL in MB" in cell `C1`.
- Applies color formatting (yellow and blue fills) to specific cells.

You can customize these modification tasks by editing the script to match your specific requirements.

## Author

- Your Name

## License

This project is licensed under the MIT License - see the [LICENSE.md](LICENSE.md) file for details.

---
