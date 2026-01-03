
import pandas as pd

file_path = 'd:/Workspace/Manage_TO/TO_Table.xlsx'
sheet_name = '2-본부'

try:
    df = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
    
    # Header Row 10 (Index 9)
    # Check Cols I(8), J(9), K(10)
    print("--- Headers Row 10 (Index 9) ---")
    headers = df.iloc[9, 8:12] # I to L
    print(headers)
    
    # Data Rows 20-30
    print("\n--- Data Rows 20-30 (Cols D, I, J, K) ---")
    subset = df.iloc[19:30, [3, 8, 9, 10]] # D, I, J, K
    print(subset)

except Exception as e:
    print(f"Error: {e}")
