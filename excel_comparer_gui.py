import pandas as pd
import os
from functools import reduce
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import threading

# --- 1. CORE COMPARISON LOGIC ---

def run_comparison(file_paths_str, sheet_name, key_column):
    """
    Executes the Excel comparison logic using inputs from the GUI.
    """
    # --- A. INITIAL INPUT VALIDATION ---
    if not sheet_name or not key_column:
        return "ERROR: Configuration fields (Sheet Name and Key Column) must be filled."

    file_paths = [p.strip() for p in file_paths_str.split('\n') if p.strip()]
    
    if len(file_paths) < 2:
        return "ERROR: Please list at least two Excel file paths to compare."

    # --- B. DATA LOADING AND CORE VALIDATION ---
    target_dfs = [] 
    original_dfs = [] 
    
    for file_path in file_paths:
        if not os.path.exists(file_path):
            return f"ERROR: File not found: '{file_path}'."
        
        try:
            df = pd.read_excel(file_path, sheet_name=sheet_name)
            
            if key_column not in df.columns:
                return f"ERROR: Column '{key_column}' not found in sheet '{sheet_name}' of file '{file_path}'."
                
            original_dfs.append({'filename': file_path, 'df': df.copy()}) 
            
            base_name = os.path.basename(file_path).split('.')[0]
            df = df.rename(columns={
                col: f"{col}_{base_name}" 
                for col in df.columns if col != key_column
            })
            
            target_dfs.append(df)
            
        except ValueError:
            return f"ERROR: Sheet '{sheet_name}' not found in file '{file_path}'."
        except Exception as e:
            return f"ERROR reading file '{file_path}': {e}"
            
    # --- C. COMPARISON LOGIC ---
    # Inner Join to find the Intersection
    # 
    df_intersection = reduce(
        lambda left, right: pd.merge(left, right, on=key_column, how='inner'),
        target_dfs
    )
    
    intersection_keys = set(df_intersection[key_column].unique())
    
    # --- D. FILE SAVING ---
    output_message = []
    
    # Intersection Save
    output_intersection_file = 'Intersection_Results.xlsx'
    if os.path.exists(output_intersection_file):
        output_message.append(f"WARNING: '{output_intersection_file}' already exists. Skipped.")
    else:
        try:
            df_intersection.to_excel(output_intersection_file, index=False)
            output_message.append(f"SUCCESS: Saved '{output_intersection_file}'.")
        except Exception as e:
            output_message.append(f"ERROR saving intersection: {e}")

    # Individual Difference Save
    # 
    total_difference_count = 0
    for data in original_dfs:
        file_path = data['filename']
        df_original = data['df']
        
        is_in_intersection = df_original[key_column].isin(intersection_keys)
        df_difference = df_original[~is_in_intersection]
        
        base_name = os.path.basename(file_path).split('.')[0]
        output_diff_file = f"Difference_{base_name}.xlsx"
        total_difference_count += len(df_difference)

        if os.path.exists(output_diff_file):
            output_message.append(f"WARNING: '{output_diff_file}' already exists. Skipped.")
        else:
            try:
                df_difference.to_excel(output_diff_file, index=False)
                output_message.append(f"SUCCESS: Saved '{output_diff_file}'.")
            except Exception as e:
                output_message.append(f"ERROR saving difference for {base_name}: {e}")

    final_summary = f"\n*** COMPLETE ***\nTotal unique records: {total_difference_count}"
    return "\n".join(output_message) + final_summary

# --- 2. TKINTER GUI SETUP ---

class ExcelComparerApp:
    def __init__(self, master):
        self.master = master
        master.title("Excel Comparison Tool")
        self.timeout_id = None 
        self.create_widgets()

    def create_widgets(self):
        # --- Configuration Frame ---
        config_frame = tk.LabelFrame(self.master, text="Configuration", padx=10, pady=10)
        config_frame.pack(padx=10, pady=10, fill="x")
        
        # Target Sheet Name (Empty)
        tk.Label(config_frame, text="Target Sheet Name:").grid(row=0, column=0, sticky="w", padx=5, pady=2)
        self.sheet_name_entry = tk.Entry(config_frame, width=30)
        self.sheet_name_entry.grid(row=0, column=1, padx=5, pady=2)

        # Unique Key Column (Empty)
        tk.Label(config_frame, text="Unique Key Column:").grid(row=1, column=0, sticky="w", padx=5, pady=2)
        self.key_column_entry = tk.Entry(config_frame, width=30)
        self.key_column_entry.grid(row=1, column=1, padx=5, pady=2)

        # --- File Paths Frame (Empty) ---
        path_frame = tk.LabelFrame(self.master, text="Excel File Paths (One per line)", padx=10, pady=5)
        path_frame.pack(padx=10, pady=10, fill="both", expand=True)
        self.paths_text = scrolledtext.ScrolledText(path_frame, height=8, width=70)
        self.paths_text.pack(padx=5, pady=5, fill="both", expand=True)
        
        # --- Run Button ---
        self.run_button = tk.Button(self.master, text="Run Comparison", command=self.start_comparison_thread, bg="lightblue")
        self.run_button.pack(pady=10)

        # --- Output Frame ---
        output_frame = tk.LabelFrame(self.master, text="Output Log", padx=10, pady=5)
        output_frame.pack(padx=10, pady=10, fill="x")
        self.output_text = scrolledtext.ScrolledText(output_frame, height=10, width=70, state='disabled', bg='lightyellow')
        self.output_text.pack(padx=5, pady=5, fill="x")

    def log_output(self, message):
        self.output_text.config(state='normal')
        self.output_text.delete('1.0', tk.END)
        self.output_text.insert(tk.END, message)
        self.output_text.config(state='disabled')

    def display_timeout_message(self):
        self.log_output("Something Wrong! Check Your Inputs!!\n(Timeout reached: Check if files are open or inputs are misspelled.)")
        self.cleanup_ui()

    def cleanup_ui(self):
        if self.timeout_id:
            self.master.after_cancel(self.timeout_id)
        self.run_button.config(state='normal')

    def comparison_finished(self, result_message):
        self.log_output(result_message)
        self.cleanup_ui()

    def execute_comparison_threaded(self):
        file_paths = self.paths_text.get('1.0', tk.END).strip()
        sheet_name = self.sheet_name_entry.get().strip()
        key_column = self.key_column_entry.get().strip()
        
        result_message = run_comparison(file_paths, sheet_name, key_column)
        self.master.after(0, lambda: self.comparison_finished(result_message))

    def start_comparison_thread(self):
        self.log_output("Starting comparison...")
        if not self.sheet_name_entry.get().strip() or not self.key_column_entry.get().strip():
             self.log_output("ERROR: All configuration fields must be filled.")
             return
             
        self.run_button.config(state='disabled')
        self.timeout_id = self.master.after(10000, self.display_timeout_message)
        threading.Thread(target=self.execute_comparison_threaded, daemon=True).start()

if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelComparerApp(root)
    root.mainloop()