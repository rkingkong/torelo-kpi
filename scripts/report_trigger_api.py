#!/usr/bin/env python3
"""
Torelo KPI — Report Runner API
Lightweight HTTP endpoint to trigger report generation from the web dashboard.

Usage:
    python3 report_trigger_api.py

Runs on port 5001. The dashboard calls POST /run-reports to trigger generation.

To run as a service, create: /etc/systemd/system/torelo-trigger.service
    [Unit]
    Description=Torelo Report Trigger API
    After=network.target

    [Service]
    ExecStart=/opt/torelo-kpi/venv/bin/python3 /opt/torelo-kpi/scripts/report_trigger_api.py
    WorkingDirectory=/opt/torelo-kpi/scripts
    Restart=always
    User=root

    [Install]
    WantedBy=multi-user.target

Then: systemctl enable torelo-trigger && systemctl start torelo-trigger
"""

import http.server
import json
import subprocess
import threading
import os
import time
from datetime import datetime

PORT = 5001
SCRIPTS_DIR = '/opt/torelo-kpi/scripts'
VENV_PYTHON = '/opt/torelo-kpi/venv/bin/python3'
WEB_DIR = '/opt/torelo-kpi/web'

# Track running state
run_state = {
    'running': False,
    'last_run': None,
    'last_status': None,
    'last_duration': None,
    'output': ''
}

def run_reports_background():
    """Run the daily reports script in background"""
    global run_state
    run_state['running'] = True
    run_state['output'] = ''
    start = time.time()
    
    try:
        result = subprocess.run(
            [VENV_PYTHON, os.path.join(SCRIPTS_DIR, 'run_daily_reports.py')],
            capture_output=True, text=True, timeout=600,  # 10 min timeout
            cwd=SCRIPTS_DIR,
            env={**os.environ, 'PYTHONIOENCODING': 'utf-8'}
        )
        
        run_state['output'] = result.stdout[-2000:] if result.stdout else ''  # Last 2000 chars
        run_state['last_status'] = 'success' if result.returncode == 0 else 'error'
        
        if result.returncode != 0 and result.stderr:
            run_state['output'] += '\n\nERRORS:\n' + result.stderr[-1000:]
            
    except subprocess.TimeoutExpired:
        run_state['last_status'] = 'timeout'
        run_state['output'] = 'Script exceeded 10 minute timeout'
    except Exception as e:
        run_state['last_status'] = 'error'
        run_state['output'] = str(e)
    
    run_state['running'] = False
    run_state['last_run'] = datetime.now().isoformat()
    run_state['last_duration'] = round(time.time() - start, 1)


class ReportHandler(http.server.BaseHTTPRequestHandler):
    def _cors(self):
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Access-Control-Allow-Methods', 'GET, POST, OPTIONS')
        self.send_header('Access-Control-Allow-Headers', 'Content-Type')
    
    def _json(self, code, data):
        self.send_response(code)
        self.send_header('Content-Type', 'application/json')
        self._cors()
        self.end_headers()
        self.wfile.write(json.dumps(data).encode())
    
    def do_OPTIONS(self):
        self.send_response(200)
        self._cors()
        self.end_headers()
    
    def do_GET(self):
        if self.path == '/status':
            self._json(200, run_state)
        else:
            self._json(404, {'error': 'Not found'})
    
    def do_POST(self):
        if self.path == '/run-reports':
            if run_state['running']:
                self._json(409, {'error': 'Reports are already running', **run_state})
                return
            
            # Start in background thread
            thread = threading.Thread(target=run_reports_background, daemon=True)
            thread.start()
            
            self._json(202, {
                'message': 'Report generation started',
                'status': 'running'
            })
        else:
            self._json(404, {'error': 'Not found'})
    
    def log_message(self, format, *args):
        print(f"[{datetime.now().strftime('%H:%M:%S')}] {args[0]}")


if __name__ == '__main__':
    server = http.server.HTTPServer(('0.0.0.0', PORT), ReportHandler)
    print(f"🚀 Report Trigger API running on port {PORT}")
    print(f"   POST /run-reports  — Start report generation")
    print(f"   GET  /status       — Check run status")
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\nShutting down.")
        server.server_close()