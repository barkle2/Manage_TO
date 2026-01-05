
import pandas as pd

file_path = 'd:/Workspace/Manage_TO/TO_Table.xlsx'
sheet_name = '2-본부'

try:
    df = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
    
    # Header Row 13 (Index 12)
    # Check Cols M(12), N(13), O(14), P(15)
    print("--- Headers Row 13 (Index 12) ---")
    headers = df.iloc[12, 12:16] # M to P
    print(headers)
    
    print("\n--- Searching by Type [실, 국, 관, 단, 위원회] (Dynamic Col) ---")
    
    target_types = ['실', '국', '관', '단', '위원회', '본부']
    
    match_count = 0
    
    for i in range(19, len(df)):
        row = df.iloc[i]
        
        # Find '기준정원'
        k = -1
        for col_idx in range(10): # Check first 10 columns
            if str(row[col_idx]).strip() == '기준정원':
                k = col_idx
                break
        
        if k == -1:
            continue
            
        # Dept Name is k-1
        dept = str(row[k-1]).strip()
        
        # Unit Type is k-3 (based on observation for Sil/Guk/Gwan)
        # Check index validity
        if k-3 < 0:
            continue
            
        unit_type = str(row[k-3]).strip()
        
        # Check if type is one of the targets
        if unit_type in target_types:
            # Special logic for '본부' -> '산업안전보건본부' only
            if unit_type == '본부' and '산업안전보건본부' not in dept:
                continue
                
            # Check counts
            m = 0
            n = 0
            p = 0
            try: m = float(row[12])
            except: pass
            try: n = float(row[13])
            except: pass
            try: p = float(row[15])
            except: pass
            
            if m > 0 or n > 0 or p > 0:
                print(f"Row {i} [{unit_type}] {dept}: M={m}, N={n}, P={p}")
                match_count += 1
                
    print(f"\nTotal Matching Rows: {match_count}")

except Exception as e:
    print(f"Error: {e}")
