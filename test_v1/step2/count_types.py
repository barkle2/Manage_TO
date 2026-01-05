
import pandas as pd
import sys

sys.stdout.reconfigure(encoding='utf-8')
file_path = 'd:/Workspace/Manage_TO/TO_list.csv'

try:
    df = pd.read_csv(file_path, encoding='utf-8-sig')
    print(f"Total Rows: {len(df)}")
    
    # Count by Type
    type_counts = df['공무원_종류'].value_counts()
    print("\n--- Counts by Type ---")
    print(type_counts)
    
    # Filter non-general
    non_general = df[df['공무원_종류'] != '일반직']
    print("\n--- Non-General Rows ---")
    print(non_general[['부서6', '직급', '공무원_종류']])
    
    # Check Senior
    senior = df[df['직급'].str.contains('고위공무원단')]
    print(f"\nTotal Senior (Header match): {len(senior)}")
    print(senior[['부서6', '직급', '공무원_종류']])

except Exception as e:
    print(f"Error: {e}")
