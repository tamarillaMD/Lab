import os
import pandas as pd

folder_path = r"C:\Users\nshcher\Desktop\New folder"
renaming_table_path = r"C:\Users\nshcher\Desktop\renaming_table.xlsx"

# Load the renaming table into a DataFrame
renaming_df = pd.read_excel(renaming_table_path)

# Iterate over each row in the renaming table
for index, row in renaming_df.iterrows():
    old_name = row['VarFish ID']
    new_name = row['real Org ID']

    # Ensure the new name ends with '.xlsx' if it's an Excel file
    if not new_name.endswith(".xlsx"):
        new_name += ".xlsx"

    # Construct the full paths for old and new files
    old_file_path = os.path.join(folder_path, old_name)
    new_file_path = os.path.join(folder_path, new_name)

    # Check if the old file exists and rename it
    if os.path.exists(old_file_path):
        os.rename(old_file_path, new_file_path)
        print(f"Renamed '{old_name}' to '{new_name}'")
    else:
        print(f"File '{old_name}' not found in the folder")

print("Renaming process complete.")