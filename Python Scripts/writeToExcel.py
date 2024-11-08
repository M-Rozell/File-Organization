import os
import pandas as pd

# Specify the path to the parent project directory
parent_folder_path = r"X:\017560-12 - JEFFCO 2022 AMP08 - MAINLINE\NoMP4\2023\02"

# Get a list of all folder names in the parent directory
folder_names = [name for name in os.listdir(parent_folder_path) if os.path.isdir(os.path.join(parent_folder_path, name))]

# Convert to DataFrame
df = pd.DataFrame(folder_names, columns=["Folder Name"])

# Set the output file path and name
output_path = os.path.join(parent_folder_path, "folder_list.xlsx")

# Save to Excel
df.to_excel(output_path, index=False)

print(f"Folder list saved to {output_path}")