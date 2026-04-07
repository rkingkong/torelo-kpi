#!/usr/bin/env python3
"""
Test script to generate comprehensive attendance JSON
"""

import sys
import os
sys.path.append('/opt/torelo-kpi/scripts')

# Import the main script's configuration
from config import DATABASE_CONFIG, BASE_DIR, WEB_DIR

import pandas as pd
from sqlalchemy import create_engine, text
from datetime import datetime, timedelta
import pytz
import json

GUATEMALA_TZ = pytz.timezone('America/Guatemala')
TRACKED_DEPARTMENTS = [
    'Obra',
    'Seguridad Industrial', 
    'Operativo',
    'Bodega',
    'Administración'
]

def connect_to_db():
    """Create database connection"""
    db_params = DATABASE_CONFIG
    conn_str = f"postgresql://{db_params['user']}:{db_params['password']}@{db_params['host']}:{db_params['port']}/{db_params['dbname']}"
    return create_engine(conn_str)

print("📊 Generando JSON de asistencias...")
engine = connect_to_db()

# Get data for last 30 days
end_date = datetime.now(GUATEMALA_TZ).date()
start_date = end_date - timedelta(days=30)

# Fetch all attendance
query = """
SELECT 
    a.employee_id,
    e.name as employee_name,
    COALESCE(d.name, 'Sin Departamento') as department_name,
    a.check_in,
    a.check_out,
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
        emp_records = date_records[date_records['employee_id'] == emp_id]
        
        # Get clock records
        clock_records = []
        for _, record in emp_records.iterrows():
            if pd.notna(record['check_in']):
                check_in_time = pd.to_datetime(record['check_in'])
                if check_in_time.tzinfo is None:
                    check_in_time = GUATEMALA_TZ.localize(check_in_time)
                else:
                    check_in_time = check_in_time.astimezone(GUATEMALA_TZ)
                    
                clock_records.append({
                    'type': 'entry',
                    'time': check_in_time.strftime('%H:%M:%S')
                })
            
            if pd.notna(record['check_out']):
                check_out_time = pd.to_datetime(record['check_out'])
                if check_out_time.tzinfo is None:
                    check_out_time = GUATEMALA_TZ.localize(check_out_time)
                else:
                    check_out_time = check_out_time.astimezone(GUATEMALA_TZ)
                    
                clock_records.append({
                    'type': 'exit',
                    'time': check_out_time.strftime('%H:%M:%S')
                })
        
        # Calculate hours
        total_hours = 0
        if len(emp_records) > 0:
            first_in = pd.to_datetime(emp_records.iloc[0]['check_in'])
            last_record = emp_records.iloc[-1]
            
            if pd.notna(last_record['check_out']):
                last_out = pd.to_datetime(last_record['check_out'])
                total_hours = (last_out - first_in).total_seconds() / 3600
            else:
                # Still inside
                total_hours = (datetime.now(GUATEMALA_TZ).replace(tzinfo=None) - first_in.replace(tzinfo=None)).total_seconds() / 3600
        
        employees_data.append({
            'employee_id': int(emp_id),
            'employee_name': emp_records.iloc[0]['employee_name'],
            'department': emp_records.iloc[0]['department_name'],
            'clock_records': clock_records,
            'total_hours': round(total_hours, 2),
            'lunch_duration': 0,
            'effective_hours': round(total_hours, 2),
            'is_currently_inside': pd.isna(emp_records.iloc[-1]['check_out']),
            'first_entry': clock_records[0]['time'] if clock_records else None,
            'last_exit': clock_records[-1]['time'] if len(clock_records) > 1 and clock_records[-1]['type'] == 'exit' else None
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

comprehensive_data = {
    'generated_at': datetime.now(GUATEMALA_TZ).isoformat(),
    'timezone': 'America/Guatemala',
    'data_range': {
        'start': start_date.isoformat(),
        'end': end_date.isoformat()
    },
    'all_employees': all_employees,
    'attendance_by_date': attendance_by_date,
    'summary': {
        'total_employees': len(all_employees),
        'total_days': len(attendance_by_date)
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

print("\n✅ JSON generation complete!")
print(f"📊 Total employees: {len(all_employees)}")
print(f"📅 Total days with data: {len(attendance_by_date)}")
print(f"📍 Today's attendance: {attendance_by_date.get(str(end_date), {}).get('total_employees', 0)} employees")

engine.dispose()
