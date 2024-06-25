# WeeklyPLUDataCleaning

This script is designed to automate the cleaning and processing of Weekly PLU (Price Look-Up) data stored in Excel files as extracted from the H&L data report application. It performs several operations on each input file found in a specified folder, including:

1. Category Mapping: Extracts and maps categories based on specific conditions from the 'Unit Cost\nEx TAX' column.
2. Data Filtering: Removes rows where 'Qty Sold' is NaN or 0.
3. Value Ranges: Applies predefined value ranges to create new columns for 'LUC Cost Range' and 'Sales Price Range' based on 'Value Sold\nInc TAX' and 'TAX on\nSales' columns respectively.
4. Column Renaming: Standardizes column names for consistency and clarity.
5. Output: Saves the processed data to individual Excel files named after the original input files in a specified output folder.

# Usage
## Prerequisites
* Python 3.x
* pandas library
* Microsoft Excel (for reading and writing Excel files)
## Setup
1. Clone this repository to your local machine.
2. Install the required dependencies using pip install -r requirements.txt.
3. Ensure your Weekly PLU data files are located in the ToClean folder as specified in the script.
## Running the Script
2. Navigate to the directory containing DataCleaner.py.
3. Run the script using python DataCleaner.py.
## Output
* Processed Excel files will be saved in the cleanedReports folder.
* Each output file will be named after the corresponding input file with .xlsx appended.
## Notes
* Make sure all input Excel files (*.xls or *.xlsx) are formatted consistently for best results.
* Adjust file paths and column names in the script as needed for different datasets.
