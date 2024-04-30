import pandas as pd
df = pd.read_excel("rosetta melb plu week 4.xls")

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
df['Category'] = df['Category'].fillna(method='ffill')

print(df)

df.to_excel("test.xlsx", index=False)

print("DataFrame saved to test.xlsx")