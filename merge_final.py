import pandas as pd

# Define the file paths for the two Excel files
'''You would need to define your local paths for your excel files here, separate files for each databse queried'''
excel_file1 = 'records_embase.xlsx'
excel_file2 = 'records_pubmed.xlsx'

# Load the data from the Excel files into DataFrames
df1 = pd.read_excel(excel_file1)
df2 = pd.read_excel(excel_file2)

# Filter rows where "Aggregate Include" is 1 in both DataFrames
df1 = df1[df1['Aggregate Include'] == 1]
df2 = df2[df2['Aggregate Include'] == 1]

# Find the common columns between the two filtered DataFrames and convert to a list
common_columns = list(set(df1.columns).intersection(df2.columns))

# Reorder columns in both DataFrames to match the common columns
df1 = df1[common_columns]
df2 = df2[common_columns]

# Concatenate the DataFrames
merged_df = pd.concat([df1, df2], ignore_index=True)

# Save the merged DataFrame to a new Excel file (elimination based on all columns)
merged_excel_file_all = 'merged_file_all_columns.xlsx'
merged_df.to_excel(merged_excel_file_all, index=False)

# Optionally, you can print the first few rows of the merged DataFrame (elimination based on all columns)
print(f'Merged data based on all columns saved to {merged_excel_file_all}')

# Duplicate elimination based on title column
# Remove duplicates based on the "Title" column only
merged_df_title_only = merged_df.drop_duplicates(subset=['Title'])

# Save the DataFrame with duplicate elimination based on title to a new Excel file
merged_excel_file_title_only = 'merged_file_title_only.xlsx'
merged_df_title_only.to_excel(merged_excel_file_title_only, index=False)

# Optionally, you can print the first few rows of the merged DataFrame (elimination based on title only)
print(f'Merged data based on title only saved to {merged_excel_file_title_only}')

# Save the original merged DataFrame without eliminating duplicates
merged_excel_file_no_elimination = 'merged_file_no_elimination.xlsx'
merged_df.to_excel(merged_excel_file_no_elimination, index=False)

# Optionally, you can print the first few rows of the original merged DataFrame (no elimination)
print(f'Original merged data without elimination saved to {merged_excel_file_no_elimination}')
