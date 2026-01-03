
import pandas as pd

file_path = 'd:/Workspace/Manage_TO/TO_Table.xlsx'
sheet_name = '2-본부'

try:
    df = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
    # Rows 20 to 35 (Indices 19 to 34)
    subset = df.iloc[19:35]
    
    # Save to text file
    subset.to_csv('d:/Workspace/Manage_TO/step1/debug_political.txt', index=True, header=False, sep='\t', encoding='utf-8')
    print("Dumped rows 20-35 to debug_political.txt")
    
except Exception as e:
    print(f"Error: {e}")
