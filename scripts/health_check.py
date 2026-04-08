#!/usr/bin/env python3
"""
Torelo KPI — Health Check Generator
Generates /web/latest/health.json with status of all scripts and reports.
Called at the end of run_daily_reports.py or independently via cron.

Output: health.json with last run times, script statuses, file ages, and alerts.
"""

import os
import sys
import json
import glob
import logging
from datetime import datetime, timedelta

# ── Config ──────────────────────────────────────────────────────────────────
BASE_DIR = os.environ.get('TORELO_KPI_DIR', '/opt/torelo-kpi')
REPORTS_DIR = os.path.join(BASE_DIR, 'reports')
WEB_DIR = os.path.join(BASE_DIR, 'web')
LATEST_DIR = os.path.join(WEB_DIR, 'latest')
SCRIPTS_DIR = os.path.join(BASE_DIR, 'scripts')
LOG_DIR = os.path.join(BASE_DIR, 'logs')

OUTPUT_FILE = os.path.join(LATEST_DIR, 'health.json')

# Scripts we expect to run daily
EXPECTED_SCRIPTS = {
    '00_master_stock_movements_complete.py': {
        'outputs': ['stock_movements_complete.csv', 'stock_movements_complete.xlsx'],
        'description': 'Stock movements from Odoo',
    },
    '01_inventory_report_generator.py': {
        'outputs': ['inventory_report.xlsx'],
        'description': 'Current inventory with demand',
    },
    '02_cash_flow.py': {
        'outputs': ['cash_flow.xlsx', 'cash_flow.json'],
        'description': 'Weekly cash flow projections',
    },
    '07_odoo_project_plan.py': {
        'outputs': ['project_plan.xlsx'],
        'description': 'Project plan with tasks',
    },
    '08_purchase_order_report.py': {
        'outputs': ['purchase_orders.xlsx', 'purchase_order_report.xlsx'],
        'description': 'PO tracking and status',
    },
    '09_daily_attendance_report.py': {
        'outputs': ['daily_attendance.xlsx', 'comprehensive_attendance.json'],
        'description': 'Employee attendance',
    },
}

# Max age in hours before a report is considered stale
STALE_THRESHOLD_HOURS = 26  # ~1 day + buffer

# ── Logging ─────────────────────────────────────────────────────────────────
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s [%(levelname)s] %(message)s',
    handlers=[
        logging.StreamHandler(sys.stdout),
    ]
)
logger = logging.getLogger('health_check')

# Also log to file if log dir exists
if os.path.isdir(LOG_DIR):
    fh = logging.FileHandler(os.path.join(LOG_DIR, 'health_check.log'))
    fh.setFormatter(logging.Formatter('%(asctime)s [%(levelname)s] %(message)s'))
    logger.addHandler(fh)


def get_file_age_hours(filepath):
    """Return age of file in hours, or None if file doesn't exist."""
    try:
        mtime = os.path.getmtime(filepath)
        age = (datetime.now().timestamp() - mtime) / 3600
        return round(age, 2)
    except (OSError, FileNotFoundError):
        return None


def check_script_status(script_name, config):
    """Check if a script's output files exist and are fresh."""
    result = {
        'script': script_name,
        'description': config['description'],
        'status': 'unknown',
        'outputs': [],
        'last_output': None,
        'age_hours': None,
        'alerts': [],
    }

    latest_mtime = None

    for output_file in config['outputs']:
        # Check in latest/ directory
        latest_path = os.path.join(LATEST_DIR, output_file)
        age = get_file_age_hours(latest_path)

        file_info = {
            'file': output_file,
            'exists': age is not None,
            'age_hours': age,
            'path': latest_path if age is not None else None,
            'size_kb': round(os.path.getsize(latest_path) / 1024, 1) if age is not None else None,
        }
        result['outputs'].append(file_info)

        if age is not None:
            mtime = os.path.getmtime(latest_path)
            if latest_mtime is None or mtime > latest_mtime:
                latest_mtime = mtime
                result['last_output'] = output_file
                result['age_hours'] = age

    # Determine status
    if result['age_hours'] is None:
        result['status'] = 'missing'
        result['alerts'].append(f'No output files found for {script_name}')
    elif result['age_hours'] > STALE_THRESHOLD_HOURS:
        result['status'] = 'stale'
        result['alerts'].append(f'Output is {result["age_hours"]:.1f}h old (threshold: {STALE_THRESHOLD_HOURS}h)')
    else:
        result['status'] = 'ok'

    # Check if any expected outputs are missing
    missing = [o for o in result['outputs'] if not o['exists']]
    if missing and result['status'] == 'ok':
        result['status'] = 'partial'
        result['alerts'].append(f'Missing: {", ".join(o["file"] for o in missing)}')

    return result


def find_today_report_folder():
    """Find today's report folder in /reports/YYYY-MM-DD/."""
    today = datetime.now().strftime('%Y-%m-%d')
    folder = os.path.join(REPORTS_DIR, today)
    if os.path.isdir(folder):
        files = os.listdir(folder)
        return {
            'date': today,
            'path': folder,
            'file_count': len(files),
            'files': files[:20],  # Cap at 20 for JSON size
        }
    return None


def get_recent_log_errors():
    """Scan recent log files for errors."""
    errors = []
    if not os.path.isdir(LOG_DIR):
        return errors

    log_files = sorted(glob.glob(os.path.join(LOG_DIR, '*.log')), key=os.path.getmtime, reverse=True)[:5]
    cutoff = datetime.now() - timedelta(hours=24)

    for lf in log_files:
        try:
            with open(lf, 'r', encoding='utf-8', errors='ignore') as f:
                for line in f:
                    if any(kw in line.upper() for kw in ['ERROR', 'CRITICAL', 'EXCEPTION', 'TRACEBACK']):
                        errors.append({
                            'file': os.path.basename(lf),
                            'message': line.strip()[:200],
                        })
                        if len(errors) >= 10:
                            return errors
        except Exception:
            pass

    return errors


def generate_health():
    """Generate the complete health check report."""
    logger.info('Starting health check...')

    now = datetime.now()

    # Check each script
    scripts = {}
    all_alerts = []
    ok_count = 0
    total_count = len(EXPECTED_SCRIPTS)

    for script_name, config in EXPECTED_SCRIPTS.items():
        result = check_script_status(script_name, config)
        scripts[script_name] = result
        all_alerts.extend(result['alerts'])
        if result['status'] == 'ok':
            ok_count += 1

    # Overall status
    if ok_count == total_count:
        overall = 'healthy'
    elif ok_count > 0:
        overall = 'degraded'
    else:
        overall = 'down'

    # Today's report folder
    today_folder = find_today_report_folder()

    # Recent errors
    recent_errors = get_recent_log_errors()
    if recent_errors:
        all_alerts.append(f'{len(recent_errors)} error(s) found in recent logs')

    # Build health object
    health = {
        'status': overall,
        'generated_at': now.isoformat(),
        'last_run': now.isoformat(),
        'timezone': 'America/Guatemala',
        'scripts': {
            'total': total_count,
            'ok': ok_count,
            'details': scripts,
        },
        'today_reports': today_folder,
        'alerts': all_alerts,
        'recent_errors': recent_errors[:5],
        'system': {
            'base_dir': BASE_DIR,
            'disk_free_mb': get_disk_free_mb(BASE_DIR),
        },
    }

    # Write output
    os.makedirs(os.path.dirname(OUTPUT_FILE), exist_ok=True)
    with open(OUTPUT_FILE, 'w', encoding='utf-8') as f:
        json.dump(health, f, indent=2, ensure_ascii=False, default=str)

    logger.info(f'Health check complete: {overall} ({ok_count}/{total_count} scripts OK)')
    logger.info(f'Output: {OUTPUT_FILE}')

    if all_alerts:
        for alert in all_alerts:
            logger.warning(f'ALERT: {alert}')

    return health


def get_disk_free_mb(path):
    """Get free disk space in MB."""
    try:
        stat = os.statvfs(path)
        return round((stat.f_bavail * stat.f_frsize) / (1024 * 1024), 0)
    except Exception:
        return None


if __name__ == '__main__':
    health = generate_health()

    # Exit code based on status
    if health['status'] == 'healthy':
        sys.exit(0)
    elif health['status'] == 'degraded':
        sys.exit(1)
    else:
        sys.exit(2)