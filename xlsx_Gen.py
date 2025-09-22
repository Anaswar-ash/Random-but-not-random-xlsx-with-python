import pandas as pd
import numpy as np
import random
from datetime import datetime, timedelta

def create_large_excel_file(num_sales=5000, num_customers=500):
    """
    Generates a large, realistic 'sales_data.xlsx' file with two sheets:
    'Sales' and 'Customers'. The data includes intentional "messy" records
    to simulate real-world data cleaning challenges.

    Args:
        num_sales (int): The number of sales records to generate.
        num_customers (int): The number of unique customers to generate.
    """
    print(f"--- Generating a large dataset with {num_sales} sales and {num_customers} customers... ---")

    # --- 1. Generate Customer Data ---
    regions = ['North', 'South', 'East', 'West']
    customer_ids = range(1, num_customers + 1)
    customer_data = {
        'CustomerID': customer_ids,
        'Region': [random.choice(regions) for _ in customer_ids]
    }
    customer_df = pd.DataFrame(customer_data)
    print("...Customer data generated.")

    # --- 2. Generate Realistic Sales Data ---
    product_categories = ['Electronics', 'Books', 'Home Goods', 'Stationery', 'Apparel']
    
    # Define realistic price ranges for each category
    price_ranges = {
        'Electronics': (150, 5000),
        'Books': (15, 250),
        'Home Goods': (50, 3000),
        'Stationery': (5, 90),
        'Apparel': (20, 400)
    }

    sales_records = []
    start_date = datetime(2023, 1, 1)
    end_date = datetime(2025, 9, 22)
    date_range_days = (end_date - start_date).days

    for i in range(num_sales):
        order_id = 101 + i
        customer_id = random.choice(customer_ids)
        order_date = start_date + timedelta(days=random.randint(0, date_range_days))
        category = random.choice(product_categories)
        
        # Generate a sale amount based on the category's price range
        min_price, max_price = price_ranges[category]
        sale_amount = round(random.uniform(min_price, max_price), 2)
        
        sales_records.append([order_id, customer_id, order_date, category, sale_amount])

    sales_df = pd.DataFrame(sales_records, columns=['OrderID', 'CustomerID', 'OrderDate', 'ProductCategory', 'SaleAmount'])
    print("...Sales data generated.")

    # --- 3. Introduce "Messiness" to Simulate Real-World Data ---
    
    # a) Introduce missing CustomerIDs for ~2% of records
    num_missing = int(num_sales * 0.02)
    missing_indices = random.sample(range(num_sales), num_missing)
    sales_df.loc[missing_indices, 'CustomerID'] = np.nan
    print(f"...Introduced {num_missing} missing CustomerIDs.")

    # b) Introduce zero SaleAmounts for ~1% of records
    num_zero_sales = int(num_sales * 0.01)
    zero_indices = random.sample(range(num_sales), num_zero_sales)
    sales_df.loc[zero_indices, 'SaleAmount'] = 0
    print(f"...Introduced {num_zero_sales} zero-value sales.")

    # --- 4. Write DataFrames to Excel File ---
    file_name = "sales_data.xlsx"
    try:
        with pd.ExcelWriter(file_name, engine='openpyxl') as writer:
            sales_df.to_excel(writer, sheet_name='Sales', index=False)
            customer_df.to_excel(writer, sheet_name='Customers', index=False)
        print(f"\nSuccessfully created '{file_name}' with 'Sales' and 'Customers' sheets.")
        print("You can now run your 'daily_sales_analysis.py' script on this large dataset.")
    except Exception as e:
        print(f"\n[ERROR] An error occurred while creating the Excel file: {e}")
        print("Please ensure you have the 'openpyxl' and 'numpy' libraries installed (`pip install openpyxl numpy`).")

if __name__ == "__main__":
    create_large_excel_file()
