from sys import argv  # needed for command line parameter
from os import path  # Needed to check for file
import os  # Needed for makedirs
from datetime import date  # Needed for ISO date format
import pandas as pd

def main():
    sales_csv = get_sales_csv()
    orders_dir = create_orders_dir(sales_csv)
    process_sales_data(sales_csv, orders_dir)

# Get path of sales data CSV file from the command line
def get_sales_csv():
    if len(argv) < 2:    # Req-1, Req-2 Check whether provide parameter is valid path of file
        print("Please provide the path to the CSV file. Exiting...")
        exit(0)
    if not path.isfile(argv[1]):
        print("The path does not point to a file. Exiting...")
        exit(0)
    return argv[1]

# Create the directory to hold the individual order Excel sheets
def create_orders_dir(sales_csv):
    # Get directory in which sales data CSV file reside
    dirname = path.dirname(path.realpath(sales_csv))
    # Determine the name and path of the directory to hold the order data files
    isodate = date.today().isoformat()
    # Create the order directory if it does not already exist
    orders_dir = path.join(dirname, f"Orders_{isodate}")
    if not path.isdir(orders_dir):
        os.makedirs(orders_dir)
    return

# Split the sales data into individual orders and save to Excel sheets
def process_sales_data(sales_csv, orders_dir):
    # Import the sales data from the CSV file into a DataFrame
    df = pd.read_csv('sales_data.csv')
    # Insert a new "TOTAL PRICE" column into the DataFrame
    df.insert(7, 'TOTAL PRICE', df['ITEM QUANTITY'] * df['ITEM PRICE'])
    # Remove columns from the DataFrame that are not needed
    df.drop(columns=['ADDRESS', 'CITY', 'STATE', 'POSTAL CODE', 'COUNTRY'], inplace=True)
    # Group the rows in the DataFrame by order ID
    for order_id, order_df in df.groupby('ORDER ID'):
    
    # For each order ID:
        # Remove the "ORDER ID" column
        order_df.drop(columns=['ORDER ID'], inplace=True)
        # Sort the items by item number
        order_df.sort_values(by='ITEM NUMBER', inplace=True)
        # Append a "GRAND TOTAL" row
        grand_total = order_df['TOTAL PRICE'].sum()
        order_df['TOTAL PRICE'] = order_df['TOTAL PRICE'].apply(lambda x: f'${x:.2f}')
        order_df['ITEM PRICE'] = order_df['ITEM PRICE'].apply(lambda x: f'${x:.2f}')
        grand_total_df = pd.DataFrame({'ITEM NUMBER': [''],'ITEM PRICE': ['GRAND TOTAL:'], 'ITEM QUANTITY': [''], 'TOTAL PRICE': [f'${grand_total:.2f}']})
        # Determine the file name and full path of the Excel sheet
        print(grand_total_df)
        order_df = pd.concat([order_df, grand_total_df], ignore_index=True)
        # Export the data to an Excel sheet
        order_df[0:10].to_excel('pandasout.xlsx')
        # TODO: Format the Excel sheet
    return

if __name__ == '__main__':
    main()