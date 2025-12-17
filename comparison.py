import pandas as pd
import os # We'll use this library to check if the file exists

# --- 1. Define the List of Excel Files ---

# Replace these placeholders with the actual names and paths of your Excel files.
excel_files_to_process = [
    'file1.xlsx',
    'file2.xlsx'
]

# A dictionary to store ALL data from ALL files, categorized by their original file name
all_project_data = {}

# --- 2. Loop Through Each File Path ---

for file_path in excel_files_to_process:
    
    # Check if the file actually exists before trying to read it
    if not os.path.exists(file_path):
        print(f" Skipping file: '{file_path}'. File not found.")
        continue # Skip to the next file in the list
    
    print(f"\n=======================================================")
    print(f" **Starting Processing for File: {file_path}**")
    print(f"=======================================================")
    
    # --- 3. Read All Sheets from the Current File ---
    
    try:
        # Use sheet_name=None to read ALL sheets into a dictionary
        file_sheets = pd.read_excel(file_path, sheet_name=None)
        
        # Store the dictionary of sheets for this file in our main data structure
        # Key: The file path, Value: The dictionary of DataFrames (sheets)
        all_project_data[file_path] = file_sheets 

        # --- 4. Loop Through Sheets in the Current File ---
        
        for sheet_name, df in file_sheets.items():
            print(f"\n Reading Sheet: **{sheet_name}**")
            
            # --- START OF YOUR DATA PROCESSING LOGIC ---
            
            # Example: Count the number of rows and columns
            rows, columns = df.shape
            print(f" - Dimensions: {rows} rows and {columns} columns.")
            
            # # Example: Print the column names
            # print(f"      - Columns: {', '.join(df.columns)}")

            # # You can add functions here for merging, cleaning, analysis, etc.
            
            # # --- END OF YOUR DATA PROCESSING LOGIC ---
            
    except Exception as e:
        print(f"An error occurred while reading file '{file_path}': {e}")
        
print("\n--- Processing Complete ---")
print(f"Successfully loaded data from {len(all_project_data)} files.")