
import http.server
import socketserver
import json
import os
import shutil
import datetime

PORT = 8000
CSV_FILE = 'TO_list.csv'

class CustomHandler(http.server.SimpleHTTPRequestHandler):
    def do_POST(self):
        if self.path == '/save':
            try:
                # 1. Read Content-Length
                content_length = int(self.headers['Content-Length'])
                post_data = self.rfile.read(content_length)
                
                # 2. Parse JSON
                data = json.loads(post_data.decode('utf-8'))
                csv_content = data.get('csvContent')
                
                if not csv_content:
                    raise ValueError("No content provided")

                # 3. Generate Timestamped Filename
                # Format: TO_list_YYMMDD_HHMMSS.csv
                timestamp = datetime.datetime.now().strftime("%y%m%d_%H%M%S")
                new_filename = f"TO_list_{timestamp}.csv"
                
                # 4. Write new file (utf-8-sig for Excel compatibility)
                with open(new_filename, 'w', encoding='utf-8-sig') as f:
                    f.write(csv_content)
                
                print(f"Saved to {new_filename}")

                # 5. Send Response
                self.send_response(200)
                self.send_header('Content-type', 'application/json')
                self.end_headers()
                response = {
                    'status': 'success', 
                    'message': f'Saved successfully to {new_filename}',
                    'filename': new_filename
                }
                self.wfile.write(json.dumps(response).encode('utf-8'))
                
            except Exception as e:
                print(f"Error saving: {e}")
                self.send_response(500)
                self.send_header('Content-type', 'application/json')
                self.end_headers()
                response = {'status': 'error', 'message': str(e)}
                self.wfile.write(json.dumps(response).encode('utf-8'))
        else:
            self.send_error(404, "Not Found")

print(f"Starting Editor Server on port {PORT}...")
print(f"Access at http://localhost:{PORT}")

# Allow reuse address to avoid "Address already in use" if restarting quickly
socketserver.TCPServer.allow_reuse_address = True

with socketserver.TCPServer(("", PORT), CustomHandler) as httpd:
    try:
        httpd.serve_forever()
    except KeyboardInterrupt:
        pass
    print("Server stopped.")
