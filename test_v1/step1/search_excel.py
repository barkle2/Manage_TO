
import pandas as pd

file_path = 'd:/Workspace/Manage_TO/TO_Table.xlsx'
sheet_name = '2-본부'

try:
    df = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
    
    # Search for '운영지원과' in column 3 (D)
    target_names = ['운영지원과', '기획조정실', '정책기획관', '국제개발협력팀']
    
    print("Searching for targets...")
    for name in target_names:
        # Find index
        # We assume column 3 contains the names
        matches = df.index[df[3] == name].tolist()
        print(f"\n--- {name} found at indices: {matches} ---")
        for idx in matches:
            # Print the row (slice columns 0-10)
            print(df.iloc[idx, 0:10].to_string())
            
except Exception as e:
    print(f"Error: {e}")
