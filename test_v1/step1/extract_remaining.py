
import pandas as pd

file_path = 'd:/Workspace/Manage_TO/TO_Table.xlsx'
sheet_name = '2-본부'
output_path = 'd:/Workspace/Manage_TO/remaining_depts.csv'

try:
    df = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
    
    # Slice from index 144 (Row 145) to include the starting point context, 
    # but user said 145 is done. Row 145 in Excel is Index 144.
    # User worked "up to 145". So 1-145 done. 146+ needs doing.
    # Start from index 144 (Row 145) to check what it is (it was `평가정원` of prev dept).
    # Wait, in debug_rows.txt:
    # 145: 고용정책실 (Index 144? No wait. debug_rows.txt has explicit index column 139 for row 140 in excel?
    # No, df index 139 is row 140.
    # In debug_rows.txt:
    # Line 7 starts with `145`. That is the index provided by `subset.to_csv`.
    # Index 145 => Excel Row 146.
    # Content: `고용정책실`.
    # Previous index 144 => Excel Row 145. Content: `국제개발협력팀`.
    # So if user done 145, they finished `국제개발협력팀`.
    # Next is `고용정책실` (Index 145).
    
    subset = df.iloc[145:]
    
    # Extract Col 1 (Type) and Col 3 (Name)
    # Filter unique names (keep first occurrence)
    data = []
    seen_names = set()
    
    for idx, row in subset.iterrows():
        type_ = str(row[1]).strip()
        name = str(row[3]).strip()
        
        # We only care about rows that define a department.
        # Often duplicated for different headcount types.
        if name not in seen_names and name != 'nan' and name != '':
            data.append({'Type': type_, 'Name': name})
            seen_names.add(name)
            
    # Save to CSV
    out_df = pd.DataFrame(data)
    out_df.to_csv(output_path, index=False, encoding='utf-8-sig')
    print(f"Saved {len(data)} unique departments to {output_path}")
    
except Exception as e:
    print(f"Error: {e}")
