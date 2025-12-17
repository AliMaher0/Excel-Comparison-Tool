# 📊 Excel Comparison Tool

A lightweight, professional desktop application built with Python and Tkinter to identify intersections and differences across multiple Excel datasets.

##  Key Features
- **Multi-File Comparison:** Compare two or more Excel files simultaneously.
- **Intersection Analysis:** Identify records present in every single file.
- **Difference Tracking:** Generate individual reports for records unique to each specific file.
- **Fast & Responsive:** Built with multi-threading to ensure the UI never freezes during heavy data processing.
- **Portable:** Can be compiled into a single `.exe` for use without Python.

## 🧠 The Logic Behind the Tool

The tool uses Set Theory and Pandas Join operations to analyze your data:

### 1. Intersection Results (Inner Join)
Finds rows where the **Unique Key** exists in **all** provided files.


### 2. Difference Results (Left Anti-Join)
For each file, the tool identifies rows that are **not** part of the intersection.


## 🛠️ How to Use

1. **Launch the App:** Run the `.exe` or `python excel_comparer_gui.py`.
2. **Configure:**
   - Enter the **Target Sheet Name** (e.g., `Sheet1`).
   - Enter the **Unique Key Column** (e.g., `Write the column that is considered as a key`).
3. **Input Files:** Paste the full file paths of your Excel files (one per line).
4. **Run:** Click `Run Comparison`.
5. **Check Output:** Results are saved in the same directory as the program:
   - `Intersection_Results.xlsx`
   - `Difference_[FileName].xlsx`

## 📦 Installation for Developers

If you want to run the source code directly:

1. Clone the repository:
   ```bash
   git clone [https://github.com/YourUsername/RepositoryName.git](https://github.com/YourUsername/RepositoryName.git)
2. Install dependencies:
   ```bash
   pip install pandas openpyxl
3. Run the application:
   ```bash
   pip install pandas openpyxl   
   
