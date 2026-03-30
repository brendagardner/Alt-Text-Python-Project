#combine excel files

import os
import pandas as pd
print("Combine xls and xlsx")
cwd = os.path.abspath('/Users/brendagardner/Library/CloudStorage/OneDrive-SaltLakeCommunityCollege/Shared Alt Text Files/1050 Alt Text Files/1050_Textbook_with_alt_text')
files = os.listdir(cwd)

output_dir = os.path.join(cwd, 'Combined')
os.makedirs(output_dir, exist_ok=True)

## get all sheets of a given file
all_dfs = []

for file in files:                         # loop through Excel files
    if file.endswith('.xls') or file.endswith('.xlsx'):
        full_path = os.path.join(cwd, file)  # build full path to the file
        excel_file = pd.ExcelFile(full_path)
        sheets = excel_file.sheet_names
        for sheet in sheets:               # loop through sheets inside an Excel file
            print(file, sheet)
            df = excel_file.parse(sheet_name=sheet)
            all_dfs.append(df)

if all_dfs:
    df_total = pd.concat(all_dfs, ignore_index=True)
else:
    df_total = pd.DataFrame()

print("Loaded, ENTER to combine:")
input()

output_path = os.path.join(output_dir, 'combined_file.xlsx')
df_total.to_excel(output_path, index=False)
print(f"Done. File saved to: {output_path}")
input()