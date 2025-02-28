import csv
import os
import sys
from openpyxl import Workbook
from openpyxl.styles import Font, numbers


def convert_chase_csv_to_expense_report(input_file):
    print("Converting...")

    # Get the directory and base filename of the input file
    directory = os.path.dirname(input_file)
    base_filename = os.path.splitext(os.path.basename(input_file))[0]
    output_file = os.path.join(directory, f"{base_filename}_converted.xlsx").replace("\\", "/")

    # Create a new Excel workbook
    wb = Workbook()
    ws = wb.active
    ws.title = "Expenses"

    # Write headers
    headers = ["Expense", "Amount", "Date"]
    ws.append(headers)

    with open(input_file, mode="r", newline="", encoding="utf-8") as infile:
        reader = csv.reader(infile)
        header = next(reader)  # Read and discard the header row

        # Identify column indices (case-insensitive search)
        try:
            date_idx = header.index("Transaction Date")
            desc_idx = header.index("Description")
            type_idx = header.index("Type")
            amount_idx = header.index("Amount")
        except ValueError:
            print("Error: Input CSV must contain 'Transaction Date', 'Description', 'Type', and 'Amount' columns.")
            sys.exit(1)

        data_rows = []
        for row in reader:
            try:
                # Read values
                date = row[date_idx]
                expense = row[desc_idx]
                trans_type = row[type_idx]
                amount = abs(float(row[amount_idx]))  # Ensure positive values only

                is_sale = trans_type == "Sale"  # Identify sales for red formatting
                data_rows.append([expense, amount, date, is_sale])
            except (IndexError, ValueError):
                print(f"Skipping malformed row: {row}")

    # Write data to the worksheet
    for row_data in data_rows:
        expense, amount, date, is_sale = row_data
        ws.append([expense, amount, date])

    # Apply formatting
    for row in ws.iter_rows(min_row=2, min_col=2, max_col=2):
        for cell in row:
            cell.number_format = numbers.FORMAT_CURRENCY_USD_SIMPLE
            if isinstance(cell.value, (int, float)):
                row_index = cell.row - 2  # Adjust for 0-based index in data_rows
                if data_rows[row_index][3]:  # Check if it was a sale (expense)
                    cell.font = Font(color="FF0000")  # Red for sales (expenses)
                else:
                    cell.font = Font(color="008000")  # Green for payments (earnings)

    wb.save(output_file)
    print(f"Converted file saved to: {output_file}")


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python chase.py csvfilelocation.csv")
        sys.exit(1)

    input_csv = sys.argv[1]
    convert_chase_csv_to_expense_report(input_csv)
