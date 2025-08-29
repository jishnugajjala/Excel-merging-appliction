import pandas as pd
import os

# List of Excel files to merge
excel_files = ['file1.xlsx', 'file2.xlsx']

# Initialize an empty list to store DataFrames
data_frames = []

# Read each Excel file and append the DataFrame to the list
for file in excel_files:
    df = pd.read_excel(file)
    data_frames.append(df)

# Concatenate all DataFrames into a single DataFrame
merged_df = pd.concat(data_frames, ignore_index=True)

# Save the merged DataFrame to a new Excel file
merged_df.to_excel('merged_file.xlsx', index=False)

print("Excel files merged successfully.")

