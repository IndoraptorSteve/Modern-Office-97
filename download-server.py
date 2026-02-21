#!/usr/bin/env python3
"""Download server for Office 97 installer"""
import http.server
import os

PORT = 8096
DIST_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'dist')

HTML = """<!DOCTYPE html>
<html>
<head>
<title>Microsoft Office 97 - Download</title>
<style>
body {
    background: #008080;
    font-family: 'MS Sans Serif', Tahoma, sans-serif;
    display: flex;
    justify-content: center;
    align-items: center;
    min-height: 100vh;
    margin: 0;
}
.window {
    background: #c0c0c0;
    border: 2px solid;
    border-color: #ffffff #808080 #808080 #ffffff;
    width: 520px;
    box-shadow: 1px 1px 0 #000;
}
.titlebar {
    background: linear-gradient(90deg, #000080, #1084d0);
    color: white;
    padding: 3px 5px;
    font-weight: bold;
    font-size: 12px;
}
.content {
    padding: 20px;
    text-align: center;
}
.logo { display: flex; gap: 8px; justify-content: center; margin-bottom: 12px; }
.app-icon {
    width: 40px; height: 40px;
    display: flex; align-items: center; justify-content: center;
    color: #fff; font-size: 20px; font-weight: bold;
    font-family: 'Times New Roman', serif;
}
.w { background: #2b579a; }
.x { background: #217346; }
.p { background: #d04423; }
.a { background: #a4373a; }
h2 { margin: 5px 0; font-size: 16px; }
p { font-size: 12px; color: #000; margin: 8px 0; }
.btn {
    background: #c0c0c0;
    border: 2px solid;
    border-color: #ffffff #808080 #808080 #ffffff;
    padding: 6px 24px;
    font-family: inherit;
    font-size: 12px;
    cursor: pointer;
    text-decoration: none;
    color: #000;
    display: inline-block;
    margin-top: 10px;
}
.btn:active { border-color: #808080 #ffffff #ffffff #808080; }
.apps { font-size: 11px; color: #444; margin-top: 12px; text-align: left; padding: 8px; background: #fff; border: 1px solid #808080; }
.apps b { color: #000; }
.info { font-size: 10px; color: #808080; margin-top: 10px; }
</style>
</head>
<body>
<div class="window">
    <div class="titlebar">Microsoft Office 97 Professional - Download</div>
    <div class="content">
        <div class="logo">
            <div class="app-icon w">W</div>
            <div class="app-icon x">X</div>
            <div class="app-icon p">P</div>
            <div class="app-icon a">A</div>
        </div>
        <h2>Microsoft Office 97 Professional Edition</h2>
        <p>Version 1.0.0 | Windows x64 | SIZE_PLACEHOLDER</p>
        <div class="apps">
            <b>Included Applications:</b><br>
            - Microsoft Word 97 (with Clippy!)<br>
            - Microsoft Excel 97<br>
            - Microsoft PowerPoint 97<br>
            - Microsoft Access 97<br><br>
            <b>Features:</b> Opens real .docx/.xlsx/.pptx files, Win95 UI, Clippy assistant
        </div>
        <a class="btn" href="/download">Download Setup</a>
        <div class="info">Run the installer to set up Office 97 on your computer.</div>
    </div>
</div>
</body>
</html>"""

class Handler(http.server.BaseHTTPRequestHandler):
    def do_GET(self):
        if self.path == '/':
            exe = self._find_exe()
            size = os.path.getsize(exe) if exe else 0
            html = HTML.replace('SIZE_PLACEHOLDER', f"{size/1024/1024:.0f} MB")
            self.send_response(200)
            self.send_header('Content-Type', 'text/html')
            self.end_headers()
            self.wfile.write(html.encode())
        elif self.path != '/':
            exe = self._find_exe()
            if exe:
                self.send_response(200)
                self.send_header('Content-Type', 'application/octet-stream')
                self.send_header('Content-Disposition', 'attachment; filename="Microsoft Office 97 Setup.exe"')
                self.send_header('Content-Length', str(os.path.getsize(exe)))
                self.end_headers()
                with open(exe, 'rb') as f:
                    while chunk := f.read(65536):
                        self.wfile.write(chunk)
            else:
                self.send_error(404)
        else:
            self.send_error(404)

    def _find_exe(self):
        for f in os.listdir(DIST_DIR):
            if f.endswith('.exe') and 'Setup' in f:
                return os.path.join(DIST_DIR, f)
        return None

    def log_message(self, fmt, *args):
        print(f"[{self.client_address[0]}] {args[0]}")

if __name__ == '__main__':
    server = http.server.HTTPServer(('0.0.0.0', PORT), Handler)
    print(f"Office 97 Download Server on http://0.0.0.0:{PORT}")
    server.serve_forever()
