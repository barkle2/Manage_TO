
import pandas as pd

file_path = 'd:/Workspace/Manage_TO/TO_Table.xlsx'
sheet_name = '2-본부'

try:
    # Read specific rows to check structure
    # Header row is index 9 (row 10). Data starts index 19 (row 20)?
    # Reading a larger chunk to be sure.
    df = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
    
    # Inspect Header at Row 10 (Index 9)
    print("--- Row 10 (Index 9): Grade Headers ---")
    print(df.iloc[9, 0:10]) # Cols A-J
    
    # Check Cols G(6), H(7) headers specifically?
    # User said G, H have Political data.
    print(f"Col G(6) Header: {df.iloc[9, 6]}")
    print(f"Col H(7) Header: {df.iloc[9, 7]}")
    
    # Inspect Data Rows 20-40 (Indices 19-39)
    # Check Cols: D(Dept), E(Type), G, H
    print("\n--- Rows 20-40 Data ---")
    subset = df.iloc[19:40, [3, 4, 6, 7]] # D, E, G, H
    print(subset)

except Exception as e:
    print(f"Error: {e}")
