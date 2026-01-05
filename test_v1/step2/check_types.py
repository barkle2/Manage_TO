
import pandas as pd
from openpyxl.utils import cell
import sys

sys.stdout.reconfigure(encoding='utf-8')

file_path = 'd:/Workspace/Manage_TO/TO_Table.xlsx'
sheet_name = '2-본부'

def col2num(col_str):
    return cell.column_index_from_string(col_str) - 1

try:
    df = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
    
    idx_Q = col2num('Q')
    idx_CB = col2num('CB')
    idx_CC = col2num('CC')
    idx_CI = col2num('CI')
    
    type_counts = {}
    
    for i in range(19, len(df)):
        row = df.iloc[i]
        
        # Find '기준정원'
        k = -1
        for col_idx in range(10): 
            if str(row[col_idx]).strip() == '기준정원':
                k = col_idx
                break
        
        if k == -1: continue
        
        if k-3 < 0: continue
        unit_type = str(row[k-3]).strip()
        
        # Calc Sum
        s = 0
        try: s += row[idx_Q : idx_CB+1].sum()
        except: pass
        try: s += row[idx_CC : idx_CI+1].sum()
        except: pass
        
        if s > 0:
            if unit_type not in type_counts:
                type_counts[unit_type] = 0
            type_counts[unit_type] += s
            
    print("\n--- Counts by Unit Type ---")
    total_leaf = 0
    for t, c in type_counts.items():
        print(f"[{t}]: {c}")
        if t in ['과', '팀', '담당관', '상황실']: # Potential Leaf types
            total_leaf += c
            
    print(f"\nTotal Leaf (Gwa/Team/Dang/Room): {total_leaf}")

except Exception as e:
    print(f"Error: {e}")
