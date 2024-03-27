import pandas as pd
import numpy as np

# Load the Excel data into a DataFrame
df = pd.read_excel("Invoice June to August.xlsx")
# Initialize an empty list to store the transformed data
new_data = []

# Iterate through each row in the original DataFrame
for index, row in df.iterrows():
    # Extract common information
    common_info = row[0:11]

    # Iterate through each group of item-related columns (Quantity, Item, Description, etc.)
    for i in range(11, len(row), 6):
        item_info = row[i:i + 6]  # Extract item-related columns
        if item_info.notnull().sum():
            new_row = common_info.tolist() + item_info.tolist()  # Combine common and item-related info
            new_data.append(new_row)
            print(i)
            print(new_row)
            print(new_data,'\n')

# Determine the maximum number of columns in the data
max_columns = max(len(row) for row in new_data)

# Fill any missing columns with empty strings
new_data = [row + [''] * (max_columns - len(row)) for row in new_data]

# Create a new DataFrame from the list of transformed data
new_df = pd.DataFrame(new_data, columns=[
    'INVOICE #', 'DATE', 'P.O. No.', 'TERMS', 'REP.', 'SHIP', 'VIA', 'S.O. NO.', 'ORDERED BY',
    'Subtotal', 'Total', 'Quantity', 'Item', 'Description', 'U/M', 'Price Each', 'Amount'
])

# Check if the column is a string, and if so, remove the '$' sign
if new_df['Total'].dtype == 'object':
    new_df['Total'] = new_df['Total'].str.replace(',', '',regex=True).str.replace('$', '',regex=True).astype(float)

if new_df['Subtotal'].dtype == 'object':
    new_df['Subtotal'] = new_df['Subtotal'].str.replace(',', '',regex=True).str.replace('$', '',regex=True).astype(float)

new_df = new_df.astype({'Price Each':'str'})
new_df = new_df.astype({'Amount':'str'})
# Remove commas and replace '$' before converting to float
new_df['Price Each'] = new_df['Price Each'].str.replace(',', '', regex=True).str.replace('$', '', regex=True).astype(float)
new_df['Amount'] = new_df['Amount'].str.replace(',', '', regex=True).str.replace('$', '', regex=True).astype(float)

new_df = new_df.astype({'Quantity':'int'})
new_df = new_df.astype({'S.O. NO.':'int'})

# Save the result to a new Excel file
#new_df.to_excel("formatted_excel_file.xlsx", index=False)


# Load the existing Excel file
existing_file = pd.ExcelFile('Invoice June to August.xlsx')

# Create a new ExcelWriter object and specify the sheet name
sheet_name = 'pivot longer'  # Replace with the desired sheet name
with pd.ExcelWriter('Invoice June to August.xlsx', mode='a', engine='openpyxl') as writer:
    new_df.to_excel(writer, sheet_name=sheet_name, index=False)
