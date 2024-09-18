import os
import shutil

# Define the paths for the source folders and the destination folder
source_folders = [
    r"C:\Users\nshcher\Desktop\test\Rome_Panel",
    r"C:\Users\nshcher\Desktop\test\TrueSight_Cardio_Panel_2016_2018",
    r"C:\Users\nshcher\Desktop\test\TrueSight_Cardio_Panel_2019",
    r"C:\Users\nshcher\Desktop\test\exomes",
    r"C:\Users\nshcher\Desktop\test\genomes"
]
destination_folder = r"C:\Users\nshcher\Desktop\test\allinone"

# Ensure the destination folder exists
if not os.path.exists(destination_folder):
    os.makedirs(destination_folder)

# Dictionary to store the chosen files
chosen_files = {}

# List to store the files that were not copied
not_copied_files = []

# Iterate through each folder
for folder_path in source_folders:
    folder_name = os.path.basename(folder_path).lower()  # Get the folder name
    
    # Iterate through the Excel files in each folder
    for file_name in os.listdir(folder_path):
        if file_name.endswith(".xlsx"):  # Only consider Excel files
            file_path = os.path.join(folder_path, file_name)
            
            # Check if the file already exists in the chosen_files dictionary
            if file_name in chosen_files:
                # If the current folder is "exomes" or "genomes", prioritize it
                if 'exomes' in folder_name or 'genomes' in folder_name:
                    # Log that we're replacing the previously chosen file
                    not_copied_files.append(chosen_files[file_name])
                    chosen_files[file_name] = file_path
                else:
                    # Log that this file is not being copied
                    not_copied_files.append(file_path)
            else:
                # If the file is not yet chosen, add it to the dictionary
                chosen_files[file_name] = file_path

# Copy the chosen files to the destination folder
for file_name, file_path in chosen_files.items():
    destination_path = os.path.join(destination_folder, file_name)
    shutil.copy2(file_path, destination_path)  # Copy file with metadata
    print(f"Copied {file_name} to {destination_folder}")

# Log the files that were not copied
if not_copied_files:
    print("\nThe following files were not copied due to prioritization:")
    for file_path in not_copied_files:
        print(f" - {file_path}")

print("\nFile copying complete.")