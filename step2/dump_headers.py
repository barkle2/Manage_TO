
import pandas as pd

file_path = 'd:/Workspace/Manage_TO/TO_Table.xlsx'
sheet_name = '2-본부'

try:
    df = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
    headers = df.iloc[12, 12:16] # M, N, O, P
    headers.to_csv('d:/Workspace/Manage_TO/step2/headers_Senior.txt', index=True, header=False, sep='\t', encoding='utf-8')
    print("Dumped Headers M-P")
    
except Exception as e:
    print(f"Error: {e}")
