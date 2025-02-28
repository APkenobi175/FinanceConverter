# FinanceConverter
A personal python script that converts files so I can easily add them to my finances spreadsheets

It takes a CSV file generated from America First Credit union and organizes the data into an excel spreadsheet

Usage:
```bash
python afcu.py input.csv
```
The new excel sheet will have 3 columns. What the debit/credit was, how much it was, and the date it was made.

If the transaction was income it will be green, if it was not income it will be red.