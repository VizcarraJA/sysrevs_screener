import os
import pandas as pd

# Define the path to your Excel file
excel_file_path = 'merged_file_no_elimination.xlsx'

# Load the Excel file into a DataFrame
df = pd.read_excel(excel_file_path)

# Convert the "DOI" column values to lowercase
df['DOI'] = df['DOI'].str.lower()

# Drop duplicate rows based on the lowercase "DOI" column
df_no_duplicates = df.drop_duplicates(subset=['DOI'], keep='first')

# Restore the original "DOI" column values to their original case (optional)
df_no_duplicates['DOI'] = df_no_duplicates['DOI'].str.upper()

# Define the path for the new Excel file
output_excel_file = 'output_no_duplicates.xlsx'

# Check if the new file already exists and generate a unique name if needed
file_counter = 1
while True:
    if file_counter > 1:
        # Append a counter to the file name to make it unique
        output_excel_file = f'output_no_duplicates_{file_counter}.xlsx'
    if not os.path.isfile(output_excel_file):
        # Save the DataFrame with duplicates removed to the new Excel file
        df_no_duplicates.to_excel(output_excel_file, index=False)
        break
    file_counter += 1

# Optionally, you can print the shape of the original and new DataFrames to see how many duplicates were removed
print(f'Original DataFrame shape: {df.shape}')
print(f'New DataFrame shape (no duplicates): {df_no_duplicates.shape}')

# Confirm that the new Excel file has been saved successfully
print(f'Data with duplicates removed saved to {output_excel_file}')
