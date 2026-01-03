
import pandas as pd
import openpyxl
from openpyxl.utils import cell

file_path = 'd:/Workspace/Manage_TO/TO_Table.xlsx'
sheet_name = '2-본부'

def col2num(col_str):
    """Convert column string (e.g., 'Q') to 0-based index."""
    # openpyxl uses 1-based index
    return cell.column_index_from_string(col_str) - 1

try:
    df = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
    
    idx_Q = col2num('Q')
    idx_CB = col2num('CB')
    idx_CC = col2num('CC')
    idx_CI = col2num('CI')
    
    print(f"Index Q: {idx_Q}, Index CB: {idx_CB}")
    print(f"Index CC: {idx_CC}, Index CI: {idx_CI}")
    
    # Range 1: Q - CB. Headers from Row 13 (Index 12)
    # Using openpyxl logic for verification or just pandas iloc
    
    headers_1_raw = df.iloc[12, idx_Q : idx_CB+1]
    print("\n--- Headers Q-CB (Row 13 / Index 12) ---")
    print(headers_1_raw.tolist()) # Print only values to see pattern
    
    # Range 2: CC - CI. Headers from Row 10 (Index 9)
    headers_2_raw = df.iloc[9, idx_CC : idx_CI+1]
    print("\n--- Headers CC-CI (Row 10 / Index 9) ---")
    print(headers_2_raw.tolist())

    # Quick Count Check
    print("\n--- Quick Count Check (Rows 19+) ---")
    total_count = 0
    
    # Scan all rows for these columns
    for i in range(19, len(df)):
        row = df.iloc[i]
        dept = str(row[3]).strip()
        type_ = str(row[4]).strip()
        
        if type_ != '기준정원':
            continue
            
        # Sum Range 1
        s1 = 0
        try:
            s1 = row[idx_Q : idx_CB+1].sum()
        except: pass
        
        # Sum Range 2
        s2 = 0
        try:
            s2 = row[idx_CC : idx_CI+1].sum()
        except: pass
        
        row_min_sum = s1 + s2
        if row_min_sum > 0:
            print(f"Row {i} ({dept}): {row_min_sum}")
            total_count += row_min_sum
            
    print(f"\nTotal Remaining General Service Count: {total_count}")

except Exception as e:
    print(f"Error: {e}")
