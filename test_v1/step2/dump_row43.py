
import pandas as pd

file_path = 'd:/Workspace/Manage_TO/TO_Table.xlsx'
sheet_name = '2-본부'

try:
    df = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
    row = df.iloc[42] # Row 43
    print(f"--- Row 43 (Index 42) ---")
    print(row.tolist())
    
    # 3rd element (Name)
    name = row[3]
    print(f"Name (Col 3): '{name}'")
    print(f"Name repr: {repr(name)}")
    
    # 1st element (Type)
    # k=4 usually. k-3 = 1.
    val_type = row[1]
    print(f"Type (Col 1): '{val_type}'")
    print(f"Type repr: {repr(val_type)}")

except Exception as e:
    print(f"Error: {e}")
