
import pandas as pd

file_path = 'd:/Workspace/Manage_TO/TO_Table.xlsx'
sheet_name = '2-본부'

try:
    df = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
    
    # Iterate all rows starting from 19
    print("Searching for Political Appointees (G > 0 or H > 0)...")
    for i in range(19, len(df)):
        row = df.iloc[i]
        dept_name = str(row[3]).strip()
        type_ = str(row[4]).strip() # Type
        
        # Check G (Index 6)
        try:
            val_G = float(row[6])
            if val_G > 0:
                print(f"Row {i} (Excel {i+1}): {dept_name} ({type_}) -> G={val_G}")
        except:
            pass
            
        # Check H (Index 7)
        try:
            val_H = float(row[7])
            if val_H > 0:
                print(f"Row {i} (Excel {i+1}): {dept_name} ({type_}) -> H={val_H}")
        except:
            pass
            
except Exception as e:
    print(f"Error: {e}")
