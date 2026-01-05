
import pandas as pd

file_path = 'd:/Workspace/Manage_TO/TO_Table.xlsx'
sheet_name = '2-본부'

try:
    df = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
    # SanAn HQ usually row 50 in csv, so row ~73 in Excel?
    # Let's subset 65-80
    subset = df.iloc[65:80]
    subset.to_csv('d:/Workspace/Manage_TO/step1/debug_sanan.txt', index=True, header=False, sep='\t', encoding='utf-8')
    print("Dumped rows 65-80")
    
except Exception as e:
    print(f"Error: {e}")
