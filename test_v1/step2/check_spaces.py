
import pandas as pd
import sys
sys.stdout.reconfigure(encoding='utf-8')

file_path = 'd:/Workspace/Manage_TO/TO_list.csv'

try:
    df = pd.read_csv(file_path, encoding='utf-8-sig')
    print("--- TO List Sample Rows (General Service) ---")
    # General service usually starts around row 30
    subset = df.iloc[30:40]
    for idx, row in subset.iterrows():
        print(f"Row {idx}: Rank='{row['직급']}'")

except Exception as e:
    print(f"Error: {e}")
