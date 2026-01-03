
import pandas as pd

file_path = 'd:/Workspace/Manage_TO/TO_Table.xlsx'
sheet_name = '2-본부'

try:
    df = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
    # Select rows around 145 (e.g., 140 to 150)
    # Python index 139 is row 140
    subset = df.iloc[139:155]
    
    # Save to text file
    subset.to_csv('d:/Workspace/Manage_TO/debug_rows.txt', index=True, header=False, sep='\t', encoding='utf-8')
    print("Dumped key rows to debug_rows.txt")
    
except Exception as e:
    print(f"Error: {e}")
