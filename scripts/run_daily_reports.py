#!/usr/bin/env python3
import os
import sys
import subprocess
from datetime import datetime
import shutil
import json

# Setup paths
SCRIPTS_DIR = '/opt/torelo-kpi/scripts'
REPORTS_DIR = '/opt/torelo-kpi/reports'
WEB_DIR = '/opt/torelo-kpi/web'

def run_reports():
    print(f"\n{'='*60}")
    print(f"🚀 Starting Torelo KPI Reports - {datetime.now()}")
    print(f"{'='*60}\n")
    
    # Create dated folder
    date_folder = datetime.now().strftime('%Y-%m-%d')
    today_reports = os.path.join(REPORTS_DIR, date_folder)
    os.makedirs(today_reports, exist_ok=True)
    
    # Scripts to run daily - INCLUDING the new attendance report
    scripts = [
        '00_master_stock_movements_complete.py',
        '01_inventory_report_generator.py',
        '02_cash_flow.py',
        '07_odoo_project_plan.py',
        '08_purchase_order_report.py',
        '09_daily_attendance_report.py'  # Now runs DAILY for attendance tracking
    ]
    
    print(f"📋 Will run {len(scripts)} reports today")
    print("="*60)
    
    success_count = 0
    failed_scripts = []
    
    for script in scripts:
        print(f"\n🔄 Running {script}...")
        try:
            # Change to scripts directory to ensure relative imports work
            os.chdir(SCRIPTS_DIR)
            
            # Set environment variable for reports directory
            env = os.environ.copy()
            env['REPORTS_DIR'] = REPORTS_DIR
            
            result = subprocess.run(
                ['/opt/torelo-kpi/venv/bin/python', script],
                capture_output=True,
                text=True,
                check=True,
                env=env
            )
            print(f"✅ {script} completed successfully")
            if result.stdout:
                # Show first 200 chars of output
                output_preview = result.stdout[:200]
                if len(result.stdout) > 200:
                    output_preview += "..."
                print(f"   Output: {output_preview}")
            success_count += 1
            
        except subprocess.CalledProcessError as e:
            print(f"❌ {script} failed!")
            print(f"   Error: {e.stderr[:500]}")  # Show first 500 chars of error
            failed_scripts.append(script)
        except Exception as e:
            print(f"❌ {script} failed with exception: {str(e)}")
            failed_scripts.append(script)
    
    # Copy generated files to web directory
    print("\n" + "="*60)
    print("📁 Organizing files for web access...")
    
    # Create latest directory
    latest_dir = os.path.join(WEB_DIR, 'latest')
    os.makedirs(latest_dir, exist_ok=True)
    
    # Clear old files (but keep status files)
    for file in os.listdir(latest_dir):
        if file not in ['status.json', 'attendance_status.json']:
            file_path = os.path.join(latest_dir, file)
            if os.path.isfile(file_path):
                os.remove(file_path)
    
    # Find and copy all report files
    report_files = []
    attendance_files = []
    
    # Look in dated folder for today's reports
    for root, dirs, files in os.walk(today_reports):
        for file in files:
            if file.endswith(('.xlsx', '.csv')):
                src = os.path.join(root, file)
                dst = os.path.join(latest_dir, file)
                shutil.copy2(src, dst)
                
                # Categorize files
                if 'asistencia' in file.lower() or 'attendance' in file.lower():
                    attendance_files.append(file)
                else:
                    report_files.append(file)
                    
                print(f"   ✓ Copied {file}")
            
            # Also copy status files
            elif file.endswith('_status.json'):
                src = os.path.join(root, file)
                dst = os.path.join(latest_dir, file)
                shutil.copy2(src, dst)
                print(f"   ✓ Copied {file}")
    
    # Also check main reports directory for any files not in dated folder
    for file in os.listdir(REPORTS_DIR):
        if file.endswith(('.xlsx', '.csv')) and file not in report_files and file not in attendance_files:
            src = os.path.join(REPORTS_DIR, file)
            # Only copy if file was modified today
            if datetime.fromtimestamp(os.path.getmtime(src)).date() == datetime.now().date():
                dst = os.path.join(latest_dir, file)
                shutil.copy2(src, dst)
                
                if 'asistencia' in file.lower() or 'attendance' in file.lower():
                    attendance_files.append(file)
                else:
                    report_files.append(file)
                    
                print(f"   ✓ Copied {file}")
    
    # Create main status file
    status = {
        'last_update': datetime.now().isoformat(),
        'success_count': success_count,
        'total_scripts': len(scripts),
        'failed_scripts': failed_scripts,
        'files_generated': report_files,
        'attendance_files': attendance_files,
        'date_folder': date_folder,
        'summary': {
            'inventory_reports': len([f for f in report_files if 'inventory' in f.lower()]),
            'cash_flow_reports': len([f for f in report_files if 'cash' in f.lower()]),
            'project_reports': len([f for f in report_files if 'project' in f.lower()]),
            'attendance_reports': len(attendance_files),
            'total_files': len(report_files) + len(attendance_files)
        }
    }
    
    with open(os.path.join(latest_dir, 'status.json'), 'w') as f:
        json.dump(status, f, indent=2)
    
    # Print summary
    print("\n" + "="*60)
    print("✨ REPORT GENERATION COMPLETED!")
    print("="*60)
    print(f"📊 Success: {success_count}/{len(scripts)} scripts")
    print(f"📁 Files generated: {len(report_files) + len(attendance_files)}")
    print(f"   - Inventory/Stock: {status['summary']['inventory_reports']}")
    print(f"   - Cash Flow: {status['summary']['cash_flow_reports']}")
    print(f"   - Project Plans: {status['summary']['project_reports']}")
    print(f"   - Attendance: {status['summary']['attendance_reports']}")
    
    if failed_scripts:
        print(f"\n⚠️ Failed scripts: {', '.join(failed_scripts)}")
    
    print(f"\n🌐 Reports available at: http://18.222.39.18/")
    print(f"📂 Files saved in: {today_reports}")
    print(f"📊 Latest files in: {latest_dir}")
    print("="*60 + "\n")
    
    return success_count == len(scripts)

if __name__ == "__main__":
    success = run_reports()
    sys.exit(0 if success else 1)