
import pandas as pd

file_path = 'd:/Workspace/Manage_TO/TO_Table.xlsx'
sheet_name = '2-본부'

try:
    df = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
    idx = 43 # Row 44
    row = df.iloc[idx]
    
    print(f"--- Row 44 (Index 43) ---")
    
    # Check '기준정원' location
    k = -1
    for col_idx in range(10): 
        if str(row[col_idx]).strip() == '기준정원':
            k = col_idx
            break
            
    if k != -1:
        name = row[k-1]
        print(f"Name (Col {k-1}): '{name}'")
        print(f"Name repr: {repr(name)}")
        
        type_ = row[k-3]
        print(f"Type (Col {k-3}): '{type_}'")
        print(f"Type repr: {repr(type_)}")
    else:
        print("Could not find '기준정원'")

except Exception as e:
    print(f"Error: {e}")
