
import pandas as pd
import re
import unicodedata
from openpyxl.utils import cell

file_path = 'd:/Workspace/Manage_TO/TO_Table.xlsx'
sheet_name = '2-본부'
match_csv_path = 'd:/Workspace/Manage_TO/Match_Table1.csv'
output_csv = 'd:/Workspace/Manage_TO/TO_list.csv'

# Helper to normalize
def normalize(s):
    # Remove all whitespace characters
    s = str(s).replace(' ', '').replace('\u3000', '').strip()
    s = re.sub(r'\s+', '', s)
    return unicodedata.normalize('NFC', s)

# Read Excel
df = pd.read_excel(file_path, sheet_name=sheet_name, header=None)

# Read Match CSV to get hierarchy map
match_df = pd.read_csv(match_csv_path, encoding='utf-8-sig')

hierarchy_map = {}
hierarchy_map_d5 = {} # Fallback using Dept5
for idx, row in match_df.iterrows():
    raw_key = str(row['부서6'])
    key = normalize(raw_key)
    hierarchy_map[key] = row
    
    # Map d5 as well
    raw_key_d5 = str(row['부서5'])
    key_d5 = normalize(raw_key_d5)
    hierarchy_map_d5[key_d5] = row

# Extracted Rows
output_rows = []

# Headers for output
columns = list(match_df.columns) + ['공무원_종류', '직급', '정원수']

# ---------------------------------------------------------
# Political Appointees (Rows 19+)
# ---------------------------------------------------------
# Headers Row 10 (Index 9)
header_G = str(df.iloc[9, 6]).replace(' ', '').strip()
header_H = str(df.iloc[9, 7]).replace(' ', '').strip()
header_I = str(df.iloc[9, 8]).replace(' ', '').strip()
header_J = str(df.iloc[9, 9]).replace(' ', '').strip()
header_K = str(df.iloc[9, 10]).replace(' ', '').strip()

total_rows = len(df)

# ---------------------------------------------------------
# Senior Civil Service Helpers
# ---------------------------------------------------------
# Row 13 (Index 12) Headers
header_M_Ga = "고위공무원단 " + str(df.iloc[12, 12]).replace(' ', '').strip() # Ga
header_N_Na = "고위공무원단 " + str(df.iloc[12, 13]).replace(' ', '').strip() # Na
header_P_Term = "고위공무원단 " + str(df.iloc[12, 15]).replace(' ', '').strip() # Term Na

target_types_senior = ['실', '국', '관', '단', '위원회', '본부']
ga_grade_depts = ['대변인', '기획조정실', '고용정책실', '노동정책실', '산업안전보건정책실']

# ---------------------------------------------------------
# Remaining General Service Helpers (Cols Q-CI)
# ---------------------------------------------------------
def col2num(col_str): return cell.column_index_from_string(col_str) - 1
idx_Q = col2num('Q')
idx_CB = col2num('CB')
idx_CC = col2num('CC')
idx_CI = col2num('CI')

# Range 1: Q-CB (Row 13 / Index 12)
headers_1 = [str(x).replace('\n', ' ').strip() for x in df.iloc[12, idx_Q : idx_CB+1]]
# Range 2: CC-CI (Row 10 / Index 9)
headers_2 = [str(x).replace('\n', ' ').strip() for x in df.iloc[9, idx_CC : idx_CI+1]]

target_leaf_types = ['과', '팀', '팀(총액)', '담당관', '상황실'] # Added based on check_types.py

# =========================================================
# Main Loop (scan all rows once per section? No, let's keep separate blocks for clarity)
# =========================================================

print("--- Processing Political Appointees ---")
for i in range(19, total_rows): 
    row = df.iloc[i]
    
    # Find '기준정원'
    k = -1
    for col_idx in range(10): 
        if str(row[col_idx]).strip() == '기준정원':
            k = col_idx
            break
    if k == -1: continue # Only process rows with '기준정원'
    
    dept_name_raw = str(row[k-1])
    dept_name = normalize(dept_name_raw)
    
    # 1. Political Appointee Check (Cols G, H, I, J, K)
    # Filter: Below Row 32, ignore Sil/Guk aggregates EXCEPT SanAn HQ
    if i >= 31: 
        if '산업안전보건본부' not in dept_name:
             # Skip Political check for these rows?
             # But wait, Political check is usually for Minister/Vice.
             # SanAn HQ Head is row 433.
             # So we must allow it.
             # But we should not process Senior/General here.
             pass
    
    # Political Logic filters
    # Only process if we find counts in G-K
    # And apply specific logic for Minister/Vice Minister Role vs Office
    
    # Override Key for Minister/Vice Minister
    lookup_key = dept_name
    if dept_name == '장관실':
        lookup_key = '장관'
    elif dept_name == '차관실':
        lookup_key = '차관'
        
    counts_found = False
    
    # Check G
    val_G = row[6]
    if pd.notna(val_G) and val_G > 0:
        counts_found = True
        count = int(val_G)
        if lookup_key in hierarchy_map:
            h_row = hierarchy_map[lookup_key]
            for _ in range(count):
                new_row = h_row.to_dict()
                new_row['공무원_종류'] = '정무직'
                new_row['직급'] = header_G
                new_row['정원수'] = 1
                new_row['부서6'] = lookup_key # Ensure role name
                output_rows.append(new_row)

    # Check H
    val_H = row[7]
    if pd.notna(val_H) and val_H > 0:
        counts_found = True
        count = int(val_H)
        if lookup_key in hierarchy_map:
            h_row = hierarchy_map[lookup_key]
            for _ in range(count):
                new_row = h_row.to_dict()
                new_row['공무원_종류'] = '정무직'
                new_row['직급'] = header_H
                new_row['정원수'] = 1
                new_row['부서6'] = lookup_key
                output_rows.append(new_row)
                
    # Check I, J, K (Special Service)
    for col_idx, hdr in [(8, header_I), (9, header_J), (10, header_K)]:
        val = row[col_idx]
        if pd.notna(val) and val > 0:
            counts_found = True
            count = int(val)
            if lookup_key in hierarchy_map:
                h_row = hierarchy_map[lookup_key]
                for _ in range(count):
                    new_row = h_row.to_dict()
                    new_row['공무원_종류'] = '별정직'
                    new_row['직급'] = hdr
                    new_row['정원수'] = 1
                    output_rows.append(new_row)

print("--- Processing Senior Civil Service ---")
for i in range(19, total_rows):
    row = df.iloc[i]
    
    k = -1
    for col_idx in range(10): 
        if str(row[col_idx]).strip() == '기준정원':
            k = col_idx; break
    if k == -1: continue
    
    dept_name_raw = str(row[k-1])
    dept_name = normalize(dept_name_raw)
    
    if k-3 < 0: continue
    unit_type = str(row[k-3]).strip()
    
    if unit_type not in target_types_senior:
        # Fallback for Committee
        if '위원회' in dept_name: unit_type = '위원회'
        else: continue
            
    # Filter SanAn HQ (Na grade) exclusion
    if dept_name == '산업안전보건본부': continue
    if unit_type == '본부' and '산업안전보건본부' not in dept_name: continue
    
    lookup_key = dept_name
    
    # Committee Special Logic
    if '고용보험심사위원회' in dept_name:
        parent_key = '고용서비스정책관'
        if parent_key in hierarchy_map:
            h_row = hierarchy_map[parent_key]
            rank = header_P_Term
            val = float(row[15]) if pd.notna(row[15]) else 0 # P
            if val > 0:
                 new_row = h_row.to_dict()
                 new_row['공무원_종류'] = '일반직'
                 new_row['직급'] = header_P_Term
                 new_row['정원수'] = 1
                 new_row['부서6'] = dept_name
                 output_rows.append(new_row)
                 print(f"  -> Added Committee Term Na")
        continue

    # Determine Rank
    rank = header_N_Na
    if dept_name in ga_grade_depts: rank = header_M_Ga
    
    # Check M and N cols
    # Only if present
    count_m = float(row[12]) if pd.notna(row[12]) else 0
    count_n = float(row[13]) if pd.notna(row[13]) else 0
    
    total_sr = count_m + count_n
    if total_sr > 0:
        if lookup_key in hierarchy_map:
            h_row = hierarchy_map[lookup_key]
            new_row = h_row.to_dict()
            new_row['공무원_종류'] = '일반직'
            new_row['직급'] = rank
            new_row['정원수'] = 1
            output_rows.append(new_row)
            print(f"  -> Added {rank} for {dept_name}")
        else:
            # Fallback d5
            if lookup_key in hierarchy_map_d5:
                 h_row = hierarchy_map_d5[lookup_key]
                 new_row = h_row.to_dict()
                 new_row['공무원_종류'] = '일반직'
                 new_row['직급'] = rank
                 new_row['정원수'] = 1
                 new_row['부서6'] = dept_name
                 output_rows.append(new_row)
                 print(f"  -> Added {rank} for {dept_name} (Fallback)")
            else:
                 print(f"  -> SKIP Senior: {dept_name} not in map")


print("--- Processing Remaining General Service ---")
for i in range(19, total_rows):
    row = df.iloc[i]
    
    k = -1
    for col_idx in range(10): 
        if str(row[col_idx]).strip() == '기준정원':
            k = col_idx; break
    if k == -1: continue
    
    dept_name_raw = str(row[k-1])
    dept_name = normalize(dept_name_raw)

    if k-3 < 0: continue
    unit_type = str(row[k-3]).strip()
    if unit_type == 'nan': unit_type = ''
    
    # Filter Logic
    is_target = False
    if unit_type in target_leaf_types: is_target = True
    elif unit_type == '' or unit_type.lower() == 'nan': is_target = True
    
    if not is_target: continue
    
    lookup_key = dept_name
    # Manual Override for known issues (Garbled chars / Encoding)
    if '대변인실' in dept_name:
        lookup_key = '대변인실'
    elif i == 43: # Row 44 (Spokesperson's Office)
        lookup_key = '대변인실'
    
    # Hierarchy Lookup
    h_row = None
    if lookup_key in hierarchy_map:
        h_row = hierarchy_map[lookup_key]
    elif lookup_key in hierarchy_map_d5:
        h_row = hierarchy_map_d5[lookup_key]
    
    if h_row is None:
        print(f"  -> SKIP Row {i+1}: '{lookup_key}' (Type: {unit_type})")
        continue
        
    # Process Q-CB
    for c_idx in range(len(headers_1)):
        val = 0
        try: val = float(row[idx_Q + c_idx])
        except: pass
        if val > 0:
            count = int(val)
            rank_str = headers_1[c_idx]
            for _ in range(count):
                new_row = h_row.to_dict()
                new_row['공무원_종류'] = '일반직'
                new_row['직급'] = rank_str
                new_row['정원수'] = 1
                # Ensure d6 is set correctly if fallback was used
                if lookup_key not in hierarchy_map:
                    new_row['부서6'] = dept_name
                output_rows.append(new_row)

    # Process CC-CI
    for c_idx in range(len(headers_2)):
        val = 0
        try: val = float(row[idx_CC + c_idx])
        except: pass
        if val > 0:
            count = int(val)
            rank_str = headers_2[c_idx]
            for _ in range(count):
                new_row = h_row.to_dict()
                new_row['공무원_종류'] = '일반직'
                new_row['직급'] = rank_str
                new_row['정원수'] = 1
                if lookup_key not in hierarchy_map:
                    new_row['부서6'] = dept_name
                output_rows.append(new_row)

# Write output
out_df = pd.DataFrame(output_rows, columns=columns)
out_df.to_csv(output_csv, index=False, encoding='utf-8-sig')
print(f"Created {output_csv} with {len(output_rows)} rows.")
