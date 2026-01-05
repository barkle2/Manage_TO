import csv
import json
import os

csv_path = r'c:\Workspace\Manage_TO\TO_HEAD_Table.csv'
html_path = r'c:\Workspace\Manage_TO\TO_HEAD_Viewer.html'

data = []

# Try UTF-8 first
try:
    with open(csv_path, 'r', encoding='utf-8-sig', newline='') as f:
        reader = csv.reader(f)
        for row in reader:
            data.append(row)
except UnicodeDecodeError:
    # Fallback to cp949 (common for Korean Windows CSVs)
    print("UTF-8 decode failed, trying cp949")
    data = []
    with open(csv_path, 'r', encoding='cp949', newline='') as f:
        reader = csv.reader(f)
        for row in reader:
            data.append(row)

# Handle potential empty rows at the end if just commas
# User's view showed lines 801-945 were mostly empty.
# We can filter rows that are completely empty or just have empty strings?
# But maybe we keep them to be faithful. The viewer handles large data fine.

json_data = json.dumps(data, ensure_ascii=False)

html_content = f"""<!DOCTYPE html>
<html lang="ko">
<head>
    <meta charset="UTF-8">
    <title>TO_HEAD_Table Viewer</title>
    <style>
        body {{ font-family: 'Segoe UI', sans-serif; margin: 0; padding: 20px; background: #f9f9f9; }}
        h1 {{ margin-bottom: 20px; color: #333; }}
        .controls {{ margin-bottom: 15px; display: flex; gap: 10px; align-items: center; }}
        .controls button {{
            padding: 8px 16px; 
            border: 1px solid #ccc; 
            background: white; 
            cursor: pointer; 
            border-radius: 4px; 
            font-size: 14px;
        }}
        .controls button:hover {{ background: #f0f0f0; }}
        #table-container {{ 
            overflow: auto; 
            height: 85vh; 
            border: 1px solid #ccc; 
            background: white; 
            box-shadow: 0 4px 6px rgba(0,0,0,0.1); 
        }}
        table {{ 
            border-collapse: separate; 
            border-spacing: 0; 
            min-width: 100%; 
            font-size: 13px; 
        }}
        th, td {{ 
            border-bottom: 1px solid #e0e0e0; 
            border-right: 1px solid #e0e0e0; 
            padding: 6px 10px; 
            white-space: nowrap; 
        }}
        th {{ 
            background-color: #f5f5f5; 
            position: sticky; 
            top: 0; 
            z-index: 10; 
            font-weight: 600; 
            color: #444; 
            border-top: 1px solid #e0e0e0; 
        }}
        td {{ color: #333; }}
        tr:nth-child(even) td {{ background-color: #fafafa; }}
        tr:hover td {{ background-color: #e6f7ff; }}
    </style>
</head>
<body>
    <h1>📊 TO_HEAD_Table.csv Viewer</h1>
    <div class="controls">
         <button onclick="window.print()">Print / PDF</button>
         <div style="flex:1"></div>
         <small style="color:#888;">{len(data)} rows loaded</small>
    </div>
    <div id="table-container"></div>
    <script>
        const rows = {json_data};
        
        const container = document.getElementById('table-container');
        const table = document.createElement('table');
        const tbody = document.createElement('tbody');
        
        // Render headings? 
        // The file structure is complex, just render everything as a grid.
        
        rows.forEach((row, index) => {{
            const tr = document.createElement('tr');
            row.forEach(cell => {{
                // Use th for the first few rows if they look like headers?
                // Let's just use td to keep it simple, but style row index < 9
                const td = document.createElement('td');
                td.textContent = cell;
                
                if (index < 9) {{
                    td.style.backgroundColor = '#fff3cd';
                    td.style.fontWeight = 'bold';
                    td.style.textAlign = 'center';
                }}
                
                tr.appendChild(td);
            }});
            tbody.appendChild(tr);
        }});
        
        table.appendChild(tbody);
        container.appendChild(table);
    </script>
</body>
</html>
"""

with open(html_path, 'w', encoding='utf-8') as f:
    f.write(html_content)

print(f"Successfully created {{html_path}} with {{len(data)}} rows.")
