
import pandas as pd
import os

input_csv = 'd:/Workspace/Manage_TO/remaining_depts.csv'
target_csv = 'd:/Workspace/Manage_TO/Match_Table1_verify.csv'

# Read input
df = pd.read_csv(input_csv)

current_d1 = "고용노동부"
current_d2 = "본부"
current_d3 = "차관"

current_d4 = "" 
current_d5 = "" 
current_d6 = ""

rows_to_append = []

for index, row in df.iterrows():
    dtype = str(row['Type']).strip()
    name = str(row['Name']).strip()
    
    row_d6 = ""
    
    if dtype == '본부':
        current_d2 = "본부"
        current_d3 = name
        current_d4 = ""
        current_d5 = ""
        if name == '산업안전보건본부':  
             row_d6 = name
        
    elif dtype == '실':
        current_d4 = name
        current_d5 = ""
        row_d6 = name 
        
    elif dtype in ['관', '국', '단']:
        current_d5 = name
        row_d6 = name
        
    elif dtype in ['과', '팀', '팀(총액)', '담당관', '센', '센터']: 
        row_d6 = name
        
    else:
        if name != 'nan' and name != '':
             row_d6 = name
             
    out_d1 = current_d1
    out_d2 = current_d2
    out_d3 = current_d3
    out_d4 = current_d4 if current_d4 else out_d3
    out_d5 = current_d5 if current_d5 else out_d4
    out_d6 = row_d6 if row_d6 else out_d5
    
    ldap = ""
    row_str = f"{ldap},{out_d1},{out_d2},{out_d3},{out_d4},{out_d5},{out_d6}"
    rows_to_append.append(row_str)

with open(target_csv, 'w', encoding='utf-8') as f:
    # Just write rows
    for r in rows_to_append:
        f.write(r + '\n')
        
print(f"Generated {len(rows_to_append)} rows to {target_csv}")
