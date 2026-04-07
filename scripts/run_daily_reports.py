#!/usr/bin/env python3
"""
Torelo KPI — Daily Report Orchestrator
Runs all report scripts daily at 6:30 AM GT (via cron).

Improvements over original:
  - Retry logic: each script retries up to MAX_RETRIES times on failure
  - Email notifications: sends alert when any script fails
  - Per-script timing: tracks duration of each script run
  - Timeout protection: kills scripts that exceed SCRIPT_TIMEOUT seconds
  - Detailed status.json: includes errors, retries, timing per script
"""
import os
import sys
import subprocess
import time
import smtplib
import traceback
from datetime import datetime
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import shutil
import json

# ─── PATHS ───────────────────────────────────────────────────────────────────
SCRIPTS_DIR = '/opt/torelo-kpi/scripts'
REPORTS_DIR = '/opt/torelo-kpi/reports'
WEB_DIR = '/opt/torelo-kpi/web'

# ─── RETRY CONFIG ────────────────────────────────────────────────────────────
MAX_RETRIES = 2          # retry failed scripts up to this many times
RETRY_DELAY = 10         # seconds between retries
SCRIPT_TIMEOUT = 300     # 5 minutes max per script execution

# ─── EMAIL NOTIFICATIONS ─────────────────────────────────────────────────────
# Set ENABLE_EMAIL = True and fill in your SMTP details to receive alerts.
# Works with Gmail (use App Password), Outlook, or any SMTP server.
ENABLE_EMAIL = False
SMTP_CONFIG = {
    'host': 'smtp.gmail.com',       # or 'smtp.office365.com' for Outlook
    'port': 587,
    'user': '',                      # e.g. 'alerts@torelo.net'
    'password': '',                  # App-specific password (never your real pw)
    'from': 'KPI Reports <alerts@torelo.net>',
    'to': ['rkong@torelo.net'],      # who gets notified on failure
}

# ─── SCRIPTS TO RUN ──────────────────────────────────────────────────────────
SCRIPTS = [
    '00_master_stock_movements_complete.py',
    '01_inventory_report_generator.py',
    '02_cash_flow.py',
    '07_odoo_project_plan.py',
    '08_purchase_order_report.py',
    '09_daily_attendance_report.py',
]


# ═════════════════════════════════════════════════════════════════════════════
# CORE LOGIC
# ═════════════════════════════════════════════════════════════════════════════

def run_single_script(script_name):
    """
    Run one script with retry logic and timeout.
    Returns a result dict with status, timing, errors, and retry count.
    """
    result = {
        'script': script_name,
        'status': 'pending',
        'attempts': 0,
        'duration_seconds': 0,
        'output_preview': '',
        'error': None,
    }

    env = os.environ.copy()
    env['REPORTS_DIR'] = REPORTS_DIR

    for attempt in range(1, MAX_RETRIES + 1):
        result['attempts'] = attempt
        start = time.time()

        try:
            os.chdir(SCRIPTS_DIR)
            proc = subprocess.run(
                ['/opt/torelo-kpi/venv/bin/python', script_name],
                capture_output=True,
                text=True,
                check=True,
                env=env,
                timeout=SCRIPT_TIMEOUT,
            )
            elapsed = round(time.time() - start, 1)
            result['status'] = 'success'
            result['duration_seconds'] = elapsed
            result['output_preview'] = (proc.stdout or '')[:300]

            label = f"(attempt {attempt})" if attempt > 1 else ""
            print(f"  ✅ {script_name} completed in {elapsed}s {label}")
            return result

        except subprocess.TimeoutExpired:
            elapsed = round(time.time() - start, 1)
            result['duration_seconds'] = elapsed
            result['error'] = f"Timed out after {SCRIPT_TIMEOUT}s"
            print(f"  ⏱️  {script_name} timed out (attempt {attempt}/{MAX_RETRIES})")

        except subprocess.CalledProcessError as e:
            elapsed = round(time.time() - start, 1)
            result['duration_seconds'] = elapsed
            result['error'] = (e.stderr or str(e))[:500]
            print(f"  ❌ {script_name} failed (attempt {attempt}/{MAX_RETRIES})")
            print(f"     {result['error'][:200]}")

        except Exception as e:
            elapsed = round(time.time() - start, 1)
            result['duration_seconds'] = elapsed
            result['error'] = str(e)[:500]
            print(f"  ❌ {script_name} exception (attempt {attempt}/{MAX_RETRIES}): {e}")

        # If we still have retries left, wait before trying again
        if attempt < MAX_RETRIES:
            print(f"     ⏳ Retrying in {RETRY_DELAY}s...")
            time.sleep(RETRY_DELAY)

    # All retries exhausted
    result['status'] = 'failed'
    return result


def collect_report_files(today_reports, latest_dir):
    """
    Copy generated files from today's folder into web/latest/.
    Returns (report_files, attendance_files) lists.
    """
    os.makedirs(latest_dir, exist_ok=True)

    # Clear old files (keep status files)
    for f in os.listdir(latest_dir):
        if f not in ('status.json', 'attendance_status.json'):
            fp = os.path.join(latest_dir, f)
            if os.path.isfile(fp):
                os.remove(fp)

    report_files = []
    attendance_files = []

    # Copy from dated folder
    if os.path.isdir(today_reports):
        for root, _dirs, files in os.walk(today_reports):
            for f in files:
                src = os.path.join(root, f)
                dst = os.path.join(latest_dir, f)

                if f.endswith(('.xlsx', '.csv', '.json')):
                    shutil.copy2(src, dst)
                    if 'asistencia' in f.lower() or 'attendance' in f.lower():
                        attendance_files.append(f)
                    else:
                        report_files.append(f)
                    print(f"   ✓ Copied {f}")

    # Also grab any files written directly to REPORTS_DIR today
    for f in os.listdir(REPORTS_DIR):
        if f.endswith(('.xlsx', '.csv')) and f not in report_files and f not in attendance_files:
            src = os.path.join(REPORTS_DIR, f)
            if datetime.fromtimestamp(os.path.getmtime(src)).date() == datetime.now().date():
                dst = os.path.join(latest_dir, f)
                shutil.copy2(src, dst)
                if 'asistencia' in f.lower() or 'attendance' in f.lower():
                    attendance_files.append(f)
                else:
                    report_files.append(f)
                print(f"   ✓ Copied {f}")

    return report_files, attendance_files


def send_failure_email(failed_results, all_results):
    """
    Send an email summarizing which scripts failed and why.
    Silently skips if ENABLE_EMAIL is False or config is incomplete.
    """
    if not ENABLE_EMAIL:
        return
    if not SMTP_CONFIG.get('user') or not SMTP_CONFIG.get('password'):
        print("  ⚠️  Email enabled but SMTP credentials not configured — skipping notification.")
        return

    subject = f"🚨 KPI Reports Failed — {datetime.now().strftime('%Y-%m-%d %H:%M')}"

    # Build a plain-text body
    lines = [
        "One or more KPI report scripts failed today.",
        f"Date: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')} (Guatemala)",
        "",
        "─── FAILED SCRIPTS ───",
    ]
    for r in failed_results:
        lines.append(f"")
        lines.append(f"Script:   {r['script']}")
        lines.append(f"Attempts: {r['attempts']}")
        lines.append(f"Duration: {r['duration_seconds']}s")
        lines.append(f"Error:    {r.get('error', 'unknown')}")

    lines.append("")
    lines.append("─── ALL RESULTS ───")
    for r in all_results:
        icon = "✅" if r['status'] == 'success' else "❌"
        retry_note = f" (after {r['attempts']} attempts)" if r['attempts'] > 1 else ""
        lines.append(f"  {icon} {r['script']} — {r['duration_seconds']}s{retry_note}")

    lines.append("")
    lines.append("Server: kpi.torelo.net")
    lines.append("Check logs: /opt/torelo-kpi/reports/")

    body = "\n".join(lines)

    msg = MIMEMultipart()
    msg['From'] = SMTP_CONFIG['from']
    msg['To'] = ', '.join(SMTP_CONFIG['to'])
    msg['Subject'] = subject
    msg.attach(MIMEText(body, 'plain', 'utf-8'))

    try:
        with smtplib.SMTP(SMTP_CONFIG['host'], SMTP_CONFIG['port'], timeout=15) as server:
            server.ehlo()
            server.starttls()
            server.ehlo()
            server.login(SMTP_CONFIG['user'], SMTP_CONFIG['password'])
            server.send_message(msg)
        print("  📧 Failure notification email sent.")
    except Exception as e:
        print(f"  ⚠️  Could not send email: {e}")


# ═════════════════════════════════════════════════════════════════════════════
# MAIN
# ═════════════════════════════════════════════════════════════════════════════

def run_reports():
    run_start = time.time()

    print(f"\n{'=' * 60}")
    print(f"🚀 Torelo KPI Reports — {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"{'=' * 60}")
    print(f"📋 {len(SCRIPTS)} scripts | max {MAX_RETRIES} attempts each | {SCRIPT_TIMEOUT}s timeout")
    print(f"{'=' * 60}\n")

    # Create dated output folder
    date_folder = datetime.now().strftime('%Y-%m-%d')
    today_reports = os.path.join(REPORTS_DIR, date_folder)
    os.makedirs(today_reports, exist_ok=True)

    # ── Run each script ──
    all_results = []
    for script in SCRIPTS:
        print(f"\n🔄 {script}")
        result = run_single_script(script)
        all_results.append(result)

    # ── Collect files ──
    print(f"\n{'=' * 60}")
    print("📁 Organizing files for web access...")
    latest_dir = os.path.join(WEB_DIR, 'latest')
    report_files, attendance_files = collect_report_files(today_reports, latest_dir)

    # ── Build status.json ──
    successes = [r for r in all_results if r['status'] == 'success']
    failures = [r for r in all_results if r['status'] == 'failed']
    total_duration = round(time.time() - run_start, 1)

    status = {
        'last_update': datetime.now().isoformat(),
        'total_duration_seconds': total_duration,
        'success_count': len(successes),
        'failure_count': len(failures),
        'total_scripts': len(SCRIPTS),
        'failed_scripts': [r['script'] for r in failures],
        'files_generated': report_files,
        'attendance_files': attendance_files,
        'date_folder': date_folder,
        'script_results': [
            {
                'script': r['script'],
                'status': r['status'],
                'attempts': r['attempts'],
                'duration_seconds': r['duration_seconds'],
                'error': r.get('error'),
            }
            for r in all_results
        ],
        'summary': {
            'inventory_reports': len([f for f in report_files if 'inventory' in f.lower()]),
            'cash_flow_reports': len([f for f in report_files if 'cash' in f.lower()]),
            'project_reports': len([f for f in report_files if 'project' in f.lower()]),
            'stock_reports': len([f for f in report_files if 'stock' in f.lower()]),
            'attendance_reports': len(attendance_files),
            'total_files': len(report_files) + len(attendance_files),
        },
    }

    with open(os.path.join(latest_dir, 'status.json'), 'w') as f:
        json.dump(status, f, indent=2)

    # ── Send notification if anything failed ──
    if failures:
        send_failure_email(failures, all_results)

    # ── Print summary ──
    print(f"\n{'=' * 60}")
    print("✨ REPORT GENERATION COMPLETED!")
    print(f"{'=' * 60}")
    print(f"📊 Success: {len(successes)}/{len(SCRIPTS)} scripts")
    print(f"⏱️  Total time: {total_duration}s")
    print(f"📁 Files generated: {len(report_files) + len(attendance_files)}")
    print(f"   - Inventory/Stock: {status['summary']['inventory_reports'] + status['summary']['stock_reports']}")
    print(f"   - Cash Flow: {status['summary']['cash_flow_reports']}")
    print(f"   - Project Plans: {status['summary']['project_reports']}")
    print(f"   - Attendance: {status['summary']['attendance_reports']}")

    if failures:
        print(f"\n⚠️  FAILED ({len(failures)}):")
        for r in failures:
            print(f"   ❌ {r['script']} — {r.get('error', 'unknown')[:120]}")
        if ENABLE_EMAIL:
            print("   📧 Notification email sent.")
        else:
            print("   💡 Tip: set ENABLE_EMAIL = True to get notified of failures.")

    # Per-script breakdown
    print(f"\n📋 Per-script breakdown:")
    for r in all_results:
        icon = "✅" if r['status'] == 'success' else "❌"
        retry_note = f" ({r['attempts']} attempts)" if r['attempts'] > 1 else ""
        print(f"   {icon} {r['script']:45s} {r['duration_seconds']:6.1f}s{retry_note}")

    print(f"\n🌐 Reports available at: https://kpi.torelo.net")
    print(f"📂 Files saved in: {today_reports}")
    print(f"📊 Latest files in: {latest_dir}")
    print(f"{'=' * 60}\n")

    return len(failures) == 0


if __name__ == "__main__":
    success = run_reports()
    sys.exit(0 if success else 1)