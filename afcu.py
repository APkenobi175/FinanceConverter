import pandas as pd
import argparse
import os

def convert_csv_to_expense_report(input_file):
    # Load the CSV file
    df = pd.read_csv(input_file)
    
    # Fill missing Debit and Credit values with 0
    df["Debit"] = df["Debit"].fillna(0)
    df["Credit"] = df["Credit"].fillna(0)
    
    # Compute the amount column (negative for credits)
    df["Amount"] = df["Debit"] - df["Credit"]
    
    # Select relevant columns and rename them
    df_transformed = df[["Description", "Amount", "Date"]].rename(columns={"Description": "Expense"})
    
    # Generate output file path in the same directory
    output_file = os.path.join(os.path.dirname(input_file), "converted_expenses.csv")
    
    # Save to a new CSV file
    df_transformed.to_csv(output_file, index=False)
    
    print(f"Converted file saved to: {output_file}")

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Convert a CSV file to an expense report.")
    parser.add_argument("input_file", help="Path to the input CSV file")
    args = parser.parse_args()
    
    convert_csv_to_expense_report(args.input_file)
