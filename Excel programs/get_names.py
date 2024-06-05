import pandas as pd

# Load the Excel file
file_path = 'path_to_your_excel_file.xlsx'
df = pd.read_excel(file_path)

# Assuming the names are in a column named 'Names'
unique_names = df['Names'].unique()

# Print unique names
for name in unique_names:
    print(name)
