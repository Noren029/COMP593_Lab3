from sys import argv  # Needed for command-line argument
from os import path  # Needed to check for file
import os  # Needed for directory creation
from datetime import date  # Needed for ISO date format
import pandas as pd  # Needed for DataFrame operations

def main():
    sales_csv = get_sales_csv()
    orders_dir = create_orders_dir(sales_csv)
    process_sales_data(sales_csv, orders_dir)

# Get path of sales data CSV file from command line
def get_sales_csv():
    if len(argv) < 2:
        print("Please provide the path to the CSV file. Exiting...")
        exit(1)

def get_sales_csv():
    sales_csv = "C:\\Users\\sibug\\Desktop\\GitHub Desktop\\COMP593-Lab3\\sales_data.csv"
    if not path.isfile(sales_csv):
        print(f"The file '{sales_csv}' does not exist. Exiting...")
        exit(1)

    return sales_csv

# Create the directory to hold the individual order Excel sheets
def create_orders_dir(sales_csv):
    dirname = path.dirname(path.realpath(sales_csv))
    isodate = date.today().isoformat()
    orders_dir = path.join(dirname, f"Orders_{isodate}")

    if not path.isdir(orders_dir):
        os.makedirs(orders_dir)

    return orders_dir  # Fix: Return the directory path

# Split the sales data into individual orders and save to Excel sheets
def process_sales_data(sales_csv, orders_dir):
    # Load the sales data
    df = pd.read_csv(sales_csv)  # Fix: Use the argument, not a hardcoded filename

    # Insert TOTAL PRICE column
    df["TOTAL PRICE"] = df["ITEM QUANTITY"] * df["ITEM PRICE"]

    # Remove unnecessary columns
    df.drop(columns=['ADDRESS', 'CITY', 'POSTAL CODE', 'COUNTRY'], inplace=True, errors='ignore')

    # Process each order
    for order_id, order_df in df.groupby('ORDER ID'):
        order_df = order_df.copy()  # Fix: Prevent SettingWithCopyWarning

        # Remove ORDER ID column
        order_df.drop(columns=['ORDER ID'], inplace=True)

        # Sort by ITEM NUMBER
        order_df.sort_values(by='ITEM NUMBER', inplace=True)

        # Append GRAND TOTAL row
        grand_total = order_df['TOTAL PRICE'].sum()
        grand_total_df = pd.DataFrame({'ITEM PRICE': ['GRAND TOTAL:'], 'TOTAL PRICE': [grand_total]})

        # Merge DataFrame
        order_df = pd.concat([order_df, grand_total_df], ignore_index=True)

        # Define output path
        output_file = path.join(orders_dir, f"Order_{order_id}.xlsx")

        # Save order to Excel
        order_df.to_excel(output_file, index=False, sheet_name=f"Order_{order_id}")

        print(f"Order {order_id} saved to: {output_file}")

if __name__ == '__main__':
    main()

