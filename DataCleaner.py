import pandas as pd
import os

# Specify the folder path where your data files are located
folder_path = r"C:\Users\cmclaren\Desktop\Coding Projects\WeeklyPLUDataCleaning\ToClean"

# Initialize an empty list to store processed DataFrames
processed_dfs = []

# Iterate over each file in the folder
for filename in os.listdir(folder_path):
    if filename.endswith('.xls') or filename.endswith('.xlsx'):
        file_path = os.path.join(folder_path, filename)

        # Read the Excel file into a DataFrame
        df = pd.read_excel(file_path)

        # Create an empty dictionary to store category mappings
        category_mapping = {}

        # Create category variable
        category = ''

        # Iterate over each row of the DataFrame
        for index, row in df.iterrows():
            # Convert the value to string before checking prefix
            if str(row['Unit Cost\nEx TAX']).startswith('PLU Group:'):
                # Extract the category from the first column
                category = str(row['Unit Cost\nEx TAX']).split(' - ')[1].split(',')[0]
                # Update the category mapping dictionary
                category_mapping[category] = category
            # Assign category directly to 'Category' column
            df.at[index, 'Category'] = category

        # Fill the remaining empty 'Category' cells with the last known category
        df['Category'] = df['Category'].ffill()

        # Move 'Category' column to column C
        category_column = df.pop('Category')
        df.insert(2, 'Category', category_column)

        # Filter out rows where 'Qty Sold' is NaN or 0
        df = df[(df['Qty Sold'].notna()) & (df['Qty Sold'] != 0)]

        # Function to apply the conditions
        def value_range(value):
            if pd.notna(value) and value >= 0:
                if value <= 20:
                    return "$0 - $20"
                elif value <= 40:
                    return "$20 - $40"
                elif value <= 60:
                    return "$40 - $60"
                elif value <= 80:
                    return "$60 - $80"
                elif value <= 100:
                    return "$80 - $100"
                elif value <= 150:
                    return "$100 - $150"
                elif value <= 200:
                    return "$150 - $200"
                elif value <= 300:
                    return "$200 - $300"
                elif value <= 500:
                    return "$300 - $500"
                else:
                    return "$500+"
            else:
                return "Not a valid amount"

        # Apply the function to create a new column 'LUC Cost Range'
        df['LUC Cost Range'] = df['Value Sold\nInc TAX'].apply(value_range)
        # Move 'LUC Cost Range' column to column E
        LUC_Cost_Range = df.pop('LUC Cost Range')
        df.insert(3, 'LUC Cost Range', LUC_Cost_Range)

        # Apply the function to create a new column 'Sales Price Range'
        df['Sales Price Range'] = df['TAX on\nSales'].apply(value_range)
        # Move 'Sales Price Range' column to column E
        Sales_Price_Range = df.pop('Sales Price Range')
        df.insert(5, 'Sales Price Range', Sales_Price_Range)

        # Rename the columns
        df = df.rename(columns={
            'Unit Cost\nEx TAX': 'PLU Group',
            'Avg Price\nInc TAX': 'Description',
            'Value Sold\nInc TAX': 'Unit Cost EX TAX',
            'TAX on\nSales': 'Avg Price inc TAX',
            'Exp. COS\n Ex-TAX': 'Qty Sold',
            'Profit\nEx-TAX': 'Value Sold',
            'Profit %\nEx-TAX': 'TAX on Sales',
            'PLU': 'EXP. COS Ex-TAX',
            'Description': 'Profit Ex- TAX',
            'Qty Sold': 'Profit % Ex-Tax'
        })

        # Append the processed DataFrame to the list
        processed_dfs.append(df)

        # Specify the output directory path where you want to save the final DataFrame
        output_path = r"C:\Users\cmclaren\Desktop\Coding Projects\WeeklyPLUDataCleaning\cleanedReports"

        # Extract the filename without extension
        output_filename = os.path.splitext(filename)[0]

        # Save the DataFrame to an Excel file with the original filename
        output_file_path = os.path.join(output_path, f"{output_filename}.xlsx")
        df.to_excel(output_file_path, index=False)

        print(f"DataFrame saved to {output_file_path}")
