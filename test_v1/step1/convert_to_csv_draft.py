
import pandas as pd
import os

# Files
input_csv = 'd:/Workspace/Manage_TO/remaining_depts.csv'
target_csv = 'd:/Workspace/Manage_TO/Match_Table1.csv'

# Read input
df = pd.read_csv(input_csv)

# Current State
current_d1 = "고용노동부"
current_d2 = "본부"
current_d3 = "차관"
current_d4 = "" # 실
current_d5 = "" # 국/관/단
current_d6 = "" # 과/팀

rows_to_append = []

# Logic:
# Iterate through each row in remaining_depts
# Update state based on Type
# Generate row string
# Handle the "Fill Down" logic for d4/d5/d6 if they are the leaf

for index, row in df.iterrows():
    dtype = str(row['Type']).strip()
    name = str(row['Name']).strip()
    
    # 1. Update Hierarchy State
    if dtype == '본부':
        current_d2 = name
        # Reset deeper levels
        current_d3 = "차관" # Default? Or copy name? 
        # For '산업안전보건본부', let's assume d3 is empty or same?
        # If I look at the csv file, for '본부', d3 is '차관'.
        # If d2 changes to '산업안전보건본부', '차관' matches? 
        # Usually SanAnHeadquarters has a Head (Bonbujang).
        # But for safety, if it's a 본부 line, we just update d2. We keep d3 as '차관' unless explicitly changed?
        # Actually, if d2 is '산업안전보건본부', d3 should probably be '산업안전보건본부장' or similar if it existed.
        # Since it's not in the list, we might just put name in d3 to be safe or keep '차관'.
        # Let's keep '차관' but if it looks wrong we fix later. 
        # BUT wait, the row itself is the entity.
        # So we write the row for '산업안전보건본부'.
        current_d3 = name # Temporary fill for this row?
        current_d4 = ""
        current_d5 = ""
        current_d6 = ""
        
    elif dtype == '실':
        current_d4 = name
        current_d5 = ""
        current_d6 = ""
        # If we just switched to a Sil, ensures we are in the right d2? 
        # Assuming sequential order, we are fine.
        
    elif dtype in ['관', '국', '단']:
        # If we are under a Sil, this is d5.
        # If we are NOT under a Sil (d4 is empty or previous d4 finished?), this might be d4?
        # But '고용지원정책관' follows '...팀'. 
        # If d4 is set to '고용정책실', and we see '관', it becomes child d5.
        # If '통합고용정책국' (Row 20), is it child of '고용정책실' or independent L4?
        # By standard, 'policy bureau' is under 'policy office'.
        # 'Integrated Employment Policy Bureau' usually independent L4?
        # Let's check Row 33 '노동정책실'.
        # If '통합고용정책국' was under '고용정책실', d4 would persist.
        # If it is independent d4:
        # Heuristic: If '국' follows '과' which was under 'Sil', it might be a new Bureau.
        # But usually 'Sil' contains 'Gwan'.
        # Let's simple assumption: '관'/'국'/'단' are always L5 (under d4).
        # EXCEPT if d4 is empty.
        # If d4 is empty, then '국' becomes d4.
        
        # However, looking at the list:
        # Row 2: 실
        # ...
        # Row 20: 국 -> If I keep d4='고용정책실', this becomes d5.
        # Row 33: 실 (New L4).
        # This seems safest. Assume nested.
        current_d5 = name
        current_d6 = ""
        
    elif dtype in ['과', '팀', '팀(총액)', '담당관']:
        current_d6 = name
        
    else:
        # Fallback or 'nan'
        # e.g. 'nan, 고용보험심사위원회'. 
        # Treated as L6? Or L5?
        # Usually committees are attached to L4 or L3.
        # Let's put in d6 for visibility.
        current_d6 = name

    # 2. Construct Row Values
    # We need: LDAP_CODE, d1, d2, d3, d4, d5, d6
    # LDAP_CODE is empty in current file
    ldap = ""
    
    # Fill logic:
    # If the row is for '실' (d4), then d5, d6 should be same as d4?
    # Based on '장관실, 장관실, 장관실': Yes.
    
    out_d1 = current_d1
    out_d2 = current_d2
    # For d3: if d2 is '산업안전보건본부', d3 might ideally be '산업안전보건본부'. 
    # If d2 is '본부', d3 is '차관'.
    if current_d2 == '산업안전보건본부':
        out_d3 = current_d2 # Fill with HQ name
    else:
        out_d3 = current_d3
        
    out_d4 = current_d4 if current_d4 else out_d3
    out_d5 = current_d5 if current_d5 else out_d4
    out_d6 = current_d6 if current_d6 else out_d5
    
    # Override logic:
    # If we are processing '실', current_d5/d6 are empty strings.
    # So out_d5 = out_d4.
    # If we are processing '국', current_d6 is empty.
    # So out_d6 = out_d5.
    
    row_str = f"{ldap},{out_d1},{out_d2},{out_d3},{out_d4},{out_d5},{out_d6}"
    rows_to_append.append(row_str)

# Append to file
with open(target_csv, 'a', encoding='cp949') as f: # CSV seems to be ANSI/CP949 or UTF-8?
    # Inspecting original file content...
    # The 'view_file' output showed Korean chars correctly. 
    # Python default 'open' might use cp1252 on Windows.
    # Let's check encoding of original file. 
    # I'll try utf-8 first.
    pass

# Actually, let's read the file first to detect encoding or just append with utf-8-sig
# Most modern Excel CSVs are cp949 in Korea or utf-8-sig.
# Match_Table1.csv content from view_file looked correct.
# I will use 'utf-8-sig' which is safe for Excel.
# But wait, if original is cp949, mixing encodings destroys it.
# Let's try to detect first.
