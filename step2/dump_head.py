
import pandas as pd

file_path = 'd:/Workspace/Manage_TO/TO_Table.xlsx'
sheet_name = '2-본부'

try:
    df = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
    subset = df.iloc[15:21] # Rows 16 to 21
    subset.to_csv('d:/Workspace/Manage_TO/step1/debug_political_head.txt', index=True, header=False, sep='\t', encoding='utf-8')
    print("Dumped rows 15-21")
    
except Exception as e:
    print(f"Error: {e}")
