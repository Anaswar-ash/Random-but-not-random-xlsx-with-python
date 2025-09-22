# Random-but-not-random-xlsx-with-python
Python script for you that will programmatically generate a large sales_data.xlsx file. When you run it, it will create an Excel file with 5,000 realistic but randomly generated sales records and 500 unique customers.
Realistic Sales Data Generator for Python Analysis
Author: ash

📈 Project Overview
This project contains a Python script (xlsx_Gen.py) designed to generate a large and realistic, multi-sheet Excel file (sales_data.xlsx). Its primary purpose is to create a high-quality sample dataset for testing and demonstrating data cleaning and analysis workflows, such as those performed by a script like daily_sales_analysis.py.

The script doesn't just create random data; it simulates a real-world scenario by intentionally introducing common data quality issues, such as missing values and invalid entries.

Key Features
Large-Scale Data: Generates 5,000 sales records across 500 customers by default, providing a substantial dataset for performance testing.

Realistic Generation: Sales amounts are generated based on logical price ranges for different product categories.

Intentional Messiness: To provide a real-world challenge for data cleaning, the script automatically introduces:

Missing CustomerIDs in a portion of the sales records.

Zero-value SaleAmounts for some transactions.

Multi-Sheet Excel Output: Creates a single .xlsx file with two distinct sheets: Sales and Customers, perfect for practicing VLOOKUP-style merges.

🚀 How to Run This Project
1. Prerequisites
Python 3.6 or newer.

pip (Python's package installer).

2. Setup
A. Clone the Repository

git clone <your-repository-url>
cd <repository-folder-name>

B. Set Up a Virtual Environment (Recommended)

# Create the environment
python -m venv venv

# Activate on Windows
.\venv\Scripts\activate

# Activate on macOS/Linux
source venv/bin/activate

3. Install Dependencies
Install the required Python libraries with one command:

pip install -r requirements.txt

(Note: You will need to create a requirements.txt file containing pandas, numpy, and openpyxl)

4. Run the Generator Script
Execute the Python script to create the Excel file:

python xlsx_Gen.py

📊 Expected Output
After running the script, you will see a console output confirming the generation process, similar to this:

--- Generating a large dataset with 5000 sales and 500 customers... ---
...Customer data generated.
...Sales data generated.
...Introduced 100 missing CustomerIDs.
...Introduced 50 zero-value sales.

Successfully created 'sales_data.xlsx' with 'Sales' and 'Customers' sheets.
You can now run your 'daily_sales_analysis.py' script on this large dataset.
