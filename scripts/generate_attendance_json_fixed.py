#!/usr/bin/env python3
"""
Fixed script with proper Guatemala timezone handling
Save as: generate_attendance_json_fixed.py
"""

import sys
import os
sys.path.append('/opt/torelo-kpi/scripts')

from config import DATABASE_CONFIG, BASE_DIR, WEB_DIR
import pandas as pd
from sqlalchemy import create_engine, text
from datetime import datetime, timedelta
import pytz
import json

GUATEMALA_TZ = pytz.timezone('America/Guatemala')
UTC_TZ = pytz.UTC

TRACKED_DEPARTMENTS = [
    'Obra',
    'Seguridad Industrial', 
    'Operativo',
    'Bodega',
    'Administración',
    'Administration',
    'Mantenimiento'
]

def connect_to_db():
    """Create database connection"""
    db_params = DATABASE_CONFIG
    conn_str = f"postgresql://{db_params['user']}:{db_params['password']}@{db_params['host']}:{db_params['port']}/{db_params['dbname']}"
    return create_engine(conn_str)

print("📊 Generando JSON de asistencias con timezone correcto...")
engine = connect_to_db()

# Get data for last 30 days
end_date = datetime.now(GUATEMALA_TZ).date()
start_date = end_date - timedelta(days=30)

# Fetch all attendance - IMPORTANT: Times are stored in UTC in database
query = """
SELECT 
    a.employee_id,
    e.name as employee_name,
    COALESCE(d.name, 'Sin Departamento') as department_name,
    a.check_in AT TIME ZONE 'UTC' as check_in_utc,
    a.check_out AT TIME ZONE 'UTC' as check_out_utc,
    DATE(a.check_in AT TIME ZONE 'UTC' AT TIME ZONE 'America/Guatemala') as work_date
FROM hr_attendance a
JOIN hr_employee e ON e.id = a.employee_id
LEFT JOIN hr_department d ON e.department_id = d.id
WHERE DATE(a.check_in AT TIME ZONE 'UTC' AT TIME ZONE 'America/Guatemala') >= :start_date
    AND DATE(a.check_in AT TIME ZONE 'UTC' AT TIME ZONE 'America/Guatemala') <= :end_date
ORDER BY work_date DESC, a.employee_id, a.check_in
"""

df = pd.read_sql_query(
    text(query),
    engine,
    params={"start_date": start_date, "end_date": end_date}
)

print(f"Found {len(df)} attendance records")

# Convert UTC times to Guatemala timezone properly
if not df.empty:
    # Parse as UTC first
    df['check_in_utc'] = pd.to_datetime(df['check_in_utc'], utc=True)
    df['check_out_utc'] = pd.to_datetime(df['check_out_utc'], utc=True)
    
    # Convert to Guatemala timezone
    df['check_in_gt'] = df['check_in_utc'].dt.tz_convert(GUATEMALA_TZ)
    df['check_out_gt'] = df['check_out_utc'].dt.tz_convert(GUATEMALA_TZ)

# Get all employees
emp_query = """
SELECT DISTINCT
    e.id,
    e.name,
    COALESCE(d.name, 'Sin Departamento') as department,
    e.mobile_phone as phone
FROM hr_employee e
LEFT JOIN hr_department d ON e.department_id = d.id
WHERE e.active = TRUE
ORDER BY e.name
"""

employees_df = pd.read_sql_query(text(emp_query), engine)
print(f"Found {len(employees_df)} employees")

# Process data by date
attendance_by_date = {}

for date in df['work_date'].unique():
    date_str = str(date)
    date_records = df[df['work_date'] == date]
    
    employees_data = []
    for emp_id in date_records['employee_id'].unique():
        emp_records = date_records[date_records['employee_id'] == emp_id].sort_values('check_in_gt')
        
        # Get clock records in Guatemala time
        clock_records = []
        total_seconds = 0
        lunch_duration_seconds = 0
        
        for idx, record in emp_records.iterrows():
            # Add entry time in Guatemala timezone
            if pd.notna(record['check_in_gt']):
                clock_records.append({
                    'type': 'entry',
                    'time': record['check_in_gt'].strftime('%H:%M:%S')
                })
            
            # Add exit time in Guatemala timezone
            if pd.notna(record['check_out_gt']):
                clock_records.append({
                    'type': 'exit',
                    'time': record['check_out_gt'].strftime('%H:%M:%S')
                })
                
                # Calculate duration for this segment
                duration = (record['check_out_gt'] - record['check_in_gt']).total_seconds()
                total_seconds += duration
        
        # Calculate lunch breaks (gaps between records)
        if len(emp_records) > 1:
            for i in range(len(emp_records) - 1):
                current_out = emp_records.iloc[i]['check_out_gt']
                next_in = emp_records.iloc[i + 1]['check_in_gt']
                
                if pd.notna(current_out) and pd.notna(next_in):
                    gap_seconds = (next_in - current_out).total_seconds()
                    # Consider gaps between 20 minutes and 2 hours as lunch
                    if 1200 <= gap_seconds <= 7200:
                        lunch_duration_seconds += gap_seconds
        
        # Determine if still on site
        last_record = emp_records.iloc[-1]
        is_currently_inside = pd.isna(last_record['check_out_gt'])
        
        # If still inside, calculate hours until now
        if is_currently_inside and pd.notna(last_record['check_in_gt']):
            now_gt = datetime.now(GUATEMALA_TZ)
            duration_until_now = (now_gt - last_record['check_in_gt']).total_seconds()
            total_seconds = duration_until_now
        
        # Convert to hours
        total_hours = total_seconds / 3600
        lunch_hours = lunch_duration_seconds / 3600
        effective_hours = total_hours  # Already excludes lunch time
        
        # Get first entry and last exit
        first_entry = None
        last_exit = None
        
        if len(clock_records) > 0:
            first_entry = clock_records[0]['time'] if clock_records[0]['type'] == 'entry' else None
            
            # Find last exit
            for record in reversed(clock_records):
                if record['type'] == 'exit':
                    last_exit = record['time']
                    break
        
        employees_data.append({
            'employee_id': int(emp_id),
            'employee_name': emp_records.iloc[0]['employee_name'],
            'department': emp_records.iloc[0]['department_name'],
            'clock_records': clock_records,
            'total_hours': round(total_hours, 2),
            'lunch_duration': round(lunch_hours, 2),
            'effective_hours': round(effective_hours, 2),
            'is_currently_inside': bool(is_currently_inside),
            'first_entry': first_entry,
            'last_exit': last_exit
        })
    
    attendance_by_date[date_str] = {
        'date': date_str,
        'total_employees': len(employees_data),
        'currently_inside': sum(1 for e in employees_data if e['is_currently_inside']),
        'employees': employees_data
    }

# Build final structure
all_employees = []
for _, emp in employees_df.iterrows():
    all_employees.append({
        'id': int(emp['id']),
        'name': emp['name'],
        'department': emp['department'],
        'phone': emp['phone'] if pd.notna(emp['phone']) else ''
    })

# Get current time in Guatemala for the report
now_gt = datetime.now(GUATEMALA_TZ)

comprehensive_data = {
    'generated_at': now_gt.isoformat(),
    'generated_at_formatted': now_gt.strftime('%d/%m/%Y %H:%M:%S'),
    'timezone': 'America/Guatemala',
    'data_range': {
        'start': start_date.isoformat(),
        'end': end_date.isoformat()
    },
    'all_employees': all_employees,
    'attendance_by_date': attendance_by_date,
    'summary': {
        'total_employees': len(all_employees),
        'total_days': len(attendance_by_date),
        'generated_time_gt': now_gt.strftime('%H:%M:%S')
    }
}

# Save to multiple locations
locations = [
    '/opt/torelo-kpi/web/comprehensive_attendance.json',
    '/opt/torelo-kpi/web/latest/comprehensive_attendance.json',
    f'/opt/torelo-kpi/reports/{end_date}/comprehensive_attendance.json'
]

for location in locations:
    try:
        os.makedirs(os.path.dirname(location), exist_ok=True)
        with open(location, 'w', encoding='utf-8') as f:
            json.dump(comprehensive_data, f, indent=2, ensure_ascii=False, default=str)
        print(f"✅ Saved to: {location}")
    except Exception as e:
        print(f"❌ Failed to save to {location}: {e}")

# Print sample of today's data to verify times
print("\n📍 Verificación de horas de hoy (Guatemala Time):")
today_data = attendance_by_date.get(str(end_date), {})
if today_data:
    print(f"Total empleados hoy: {today_data['total_employees']}")
    print(f"Actualmente en obra: {today_data['currently_inside']}")
    
    # Show first 3 employees as sample
    for emp in today_data.get('employees', [])[:3]:
        print(f"\n👤 {emp['employee_name']}:")
        for record in emp.get('clock_records', [])[:4]:  # Show first 4 records
            emoji = "🟢" if record['type'] == 'entry' else "��"
            print(f"   {emoji} {record['type']}: {record['time']} GT")

print(f"\n✅ JSON generation complete at {now_gt.strftime('%H:%M:%S')} Guatemala Time!")
print(f"📊 Total employees: {len(all_employees)}")
print(f"📅 Total days with data: {len(attendance_by_date)}")

engine.dispose()
