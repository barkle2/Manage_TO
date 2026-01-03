
import pandas as pd

file_path = 'd:/Workspace/Manage_TO/TO_Table.xlsx'
sheet_name = '2-본부'

try:
    df = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
    subset = df.iloc[190:200]
    subset.to_csv('d:/Workspace/Manage_TO/step2/debug_committee.txt', index=True, header=False, sep='\t', encoding='utf-8')
    print("Dumped rows 190-200")
    
except Exception as e:
    print(f"Error: {e}")
