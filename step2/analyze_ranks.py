
import pandas as pd
import sys

# Set encoding for output
sys.stdout.reconfigure(encoding='utf-8')

file_path = 'd:/Workspace/Manage_TO/TO_list.csv'

try:
    df = pd.read_csv(file_path, encoding='utf-8-sig')
    
    # Filter for '일반직' (General Service)
    general_df = df[df['공무원_종류'] == '일반직']
    
    total_count = len(general_df)
    print(f"Total '일반직' Count: {total_count}")
    
    # Group by Rank
    rank_counts = general_df['직급'].value_counts()
    
    print("\n--- Rank Breakdown ---")
    print(rank_counts)
    
    # Check if we can approximate 629
    # Filter out '고위공무원단' (Senior)
    non_senior = general_df[~general_df['직급'].str.contains('고위공무원단')]
    print(f"\nNon-Senior Count: {len(non_senior)}")
    
except Exception as e:
    print(f"Error: {e}")
