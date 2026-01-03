
import pandas as pd

file_path = 'd:/Workspace/Manage_TO/TO_Table.xlsx'
sheet_name = '2-본부'

try:
    df = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
    print("Shape:", df.shape)
    print("First 5 rows:")
    print(df.head(5))
    
    print("\nRows 140 to 150:")
    print(df.iloc[139:150]) # python is 0-indexed, so 140th row is index 139
    
except Exception as e:
    print(f"Error reading excel: {e}")
