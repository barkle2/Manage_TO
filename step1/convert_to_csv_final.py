
import pandas as pd
import os

input_csv = 'd:/Workspace/Manage_TO/remaining_depts.csv'
target_csv = 'd:/Workspace/Manage_TO/Match_Table1.csv'

# Read input
try:
    df = pd.read_csv(input_csv)
except Exception as e:
    print(f"Error reading input: {e}")
    exit(1)

# Current State (Initialize with what we know is the "context" at row 145)
# Row 145 was '국제개발협력팀' -> under '국제협력관' -> under '기획조정실'.
# But the first item in remaining_depts is '실, 고용정책실'.
# So we can just start fresh d4/d5/d6 resetting.
# But d1, d2, d3 default values need to be correct.

current_d1 = "고용노동부"
current_d2 = "본부"
current_d3 = "차관"

current_d4 = "" 
current_d5 = "" 
current_d6 = ""

rows_to_append = []

# Logic
for index, row in df.iterrows():
    dtype = str(row['Type']).strip()
    name = str(row['Name']).strip()
    
    # 1. Update Hierarchy State
    if dtype == '본부':
        current_d2 = name
        current_d3 = "차관" 
        current_d4 = ""
        current_d5 = ""
        current_d6 = ""
        
    elif dtype == '실':
        current_d4 = name
        current_d5 = ""
        current_d6 = ""
        # IMPORTANT: User wants 'Sil' managed at d6 too. 
        # So we create a row immediately for this 'Sil' as a leaf.
        # This will be handled by the common output logic below if we treat this row as a valid record.
        # But we need to ensure d5, d6 are filled with name.
        current_d6 = name # Provisional for this row output?
        # But wait, next iteration will be children. 
        # If we set current_d6 = name here, it will be emitted.
        # BUT for subsequent children (e.g. 'Gwan' under 'Sil'), we need to know d4 is 'Sil', but d6 should reset?
        # No, current_d6 is reset in loop usually. using `current_d6` to store context is risky if it persists.
        # Let's separate "State" from "Row Output".
        
    elif dtype in ['관', '국', '단']:
        current_d5 = name
        current_d6 = ""
        # Similarly, emit row for this Bureau as leaf.
        current_d6 = name
        
    elif dtype in ['과', '팀', '팀(총액)', '담당관', '센', '센터']: 
        current_d6 = name
        
    else:
        if name != 'nan' and name != '':
             current_d6 = name

    # 2. Fill Logic for Output
    # We want to output a row for EVERY entry in the input csv.
    # The input csv contains:
    # 실, 고용정책실 -> Output row
    # 관, 노동시장정책관 -> Output row
    # 과, 고용정책총괄과 -> Output row
    
    # State maintenance for children:
    # If this row is '실', we set current_d4.
    # If this row is '관', we set current_d5.
    
    # Re-evaluating variables to prevent "sticky" d6 from previous row.
    # We used `current_d6` above. 
    # Logic adjustment:
    
    # Global state (persists across rows)
    # current_d1, current_d2, current_d3, current_d4, current_d5
    
    # Row specific:
    # row_d6
    
    row_d6 = ""
    
    if dtype == '본부':
        # User Logic Change:
        # Previously: d2=Name, d3=Cha-gwan
        # New Request: d2="본부", d3=Name (e.g. SanAn HQ)?
        # Let's verify file content (Line 73):
        # ,고용노동부,본부,산업안전보건본부,산업안전보건본부,산업안전보건본부,산업안전보건본부
        # So d2 is '본부'. d3 is '산업안전보건본부'.
        # And it seems d4, d5, d6 fill down with '산업안전보건본부' for the "Headquarters" row itself?
        # Yes, line 73 shows that.
        
        current_d2 = "본부" # Enforce '본부' for d2
        current_d3 = name  # Set d3 to the Branch Name (SanAn HQ)
        current_d4 = ""
        current_d5 = ""
        
        # We assume this row represents the entity self-row.
        # Check Line 73: ,고용노동부,본부,산업안전보건본부,산업안전보건본부,산업안전보건본부,산업안전보건본부
        # So we should emit a row.
        # But wait, original code skipped 'Bonbu' row emission?
        # Let's assume we DO emit if name is '산업안전보건본부'.
        if name == '산업안전보건본부':  
             row_d6 = name
        
    elif dtype == '실':
        current_d4 = name
        current_d5 = "" # Reset d5 when new Sil starts
        row_d6 = name # The row represents the Sil itself
        
    elif dtype in ['관', '국', '단']:
        current_d5 = name
        row_d6 = name # The row represents the Bureau itself
        
    elif dtype in ['과', '팀', '팀(총액)', '담당관', '센', '센터']: 
        row_d6 = name
        
    else:
        if name != 'nan' and name != '':
             row_d6 = name
             
    # Prepare output components
    out_d1 = current_d1
    out_d2 = current_d2
    
    # Logic for d3
    # If d3 was set by 'Bonbu' logic above, use it.
    # If not (e.g. we are in 'Goyong' Sil), current_d3 is likely '차관'.
    # We need to make sure we don't overwrite SanAn HQ with 'Cha-gwan' if we are in SanAn HQ context.
    # How do we know context? 
    # current_d3 holds it. 
    # Initial state was current_d3 = "차관".
    # If we hit 'SanAn HQ', we updated current_d3 = "SanAn HQ".
    # So valid for subsequent rows.
    
    out_d3 = current_d3
        
    # Logic for d4
    # If d4 is empty (e.g. SanAn HQ straight to Guk? No, usually Sil first? 
    # Or in SanAn HQ, d4 might be empty if we go straight to 'Policy Bureau'?)
    # Line 75: ...,산업안전보건정책실,... 
    # So d4 becomes 'SanAn Policy Office'.
    # Line 73 (HQ self row): d4=HQ name?
    # Yes: ,고용노동부,본부,산업안전보건본부,산업안전보건본부...
    # logic: if d4 empty, use d3.
    out_d4 = current_d4 if current_d4 else out_d3
    
    # Logic for d5
    out_d5 = current_d5 if current_d5 else out_d4
    
    # Logic for d6
    out_d6 = row_d6 if row_d6 else out_d5
    
    # Minister/Vice Minister logic?
    # User mentioned: "장관, 차관도 따로 있어야 할 것 같아서"
    # These are rows 2-5 in the user file.
    # They are not in 'remaining_depts.csv' (which starts at row 146).
    # So I don't need to generate them. I just need to match the pattern for new data.
    
    # Check "SanAn HQ" self-row generation logic again.
    # If dtype='본부' and name='산업안전보건본부':
    # d4 (empty) -> d3 (SanAn HQ).
    # d5 (empty) -> d4 (SanAn HQ).
    # row_d6 (SanAn HQ) -> d6 (SanAn HQ).
    # Result: ..., 본부, 산업안전보건본부, 산업안전보건본부, 산업안전보건본부, 산업안전보건본부.
    # Matches Line 73. CORRECT.

    ldap = ""
    row_str = f"{ldap},{out_d1},{out_d2},{out_d3},{out_d4},{out_d5},{out_d6}"
    rows_to_append.append(row_str)

# Append
try:
    with open(target_csv, 'a', encoding='utf-8') as f:
        # Ensure we start on new line?
        # If file ends with newline, good.
        # We rely on print to add newline or manual \n.
        # f.write("\n".join(rows_to_append)) -> adds between.
        # Need leading newline?
        # Let's check if file ends with newline?
        # We can just write \n prefix to first line if needed?
        # Safer: read last char?
        pass

    # Actually re-opening in 'r' to check last char is safer.
    need_newline = False
    with open(target_csv, 'rb') as f:
        f.seek(0, 2) # End
        if f.tell() > 0:
            f.seek(-1, 2)
            last = f.read(1)
            if last != b'\n':
                need_newline = True
                
    with open(target_csv, 'a', encoding='utf-8') as f:
        if need_newline:
            f.write('\n')
        for r in rows_to_append:
            f.write(r + '\n')
            
    print(f"Appended {len(rows_to_append)} rows to {target_csv}")
    
except Exception as e:
    print(f"Error writing: {e}")
