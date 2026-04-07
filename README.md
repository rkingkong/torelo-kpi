# Torelo KPI Dashboard

Automated business intelligence portal pulling data from Odoo via PostgreSQL.

**Live:** https://kpi.torelo.net

## Architecture

```
/opt/torelo-kpi/
├── scripts/          ← Python report generators (this repo)
├── web/              ← HTML frontends (this repo)
│   └── latest/       ← Symlinked latest reports (generated, not in repo)
├── reports/          ← Generated Excel/CSV output (not in repo)
└── venv/             ← Python virtualenv (not in repo)
```

## Report Scripts

| Script | Output | Description |
|--------|--------|-------------|
| `00_master_stock_movements_complete.py` | CSV + XLSX | Stock movements from Odoo |
| `01_inventory_report_generator.py` | XLSX | Current inventory with demand |
| `02_cash_flow.py` | XLSX + JSON | Weekly cash flow projections |
| `07_odoo_project_plan.py` | XLSX | Project plan with task dependencies |
| `08_purchase_order_report.py` | XLSX | PO tracking and status |
| `09_daily_attendance_report.py` | XLSX + JSON | Employee attendance |

## Web Viewers

| File | Description |
|------|-------------|
| `index.html` | Main dashboard |
| `gantt-viewer.html` | Gantt chart for projects |
| `cashflow-viewer.html` | Cash flow explorer |
| `attendance-viewer.html` | Attendance explorer |
| `search.html` | Search interface |
| `chat-widget.html` | AI chat widget |

## Automation

- `run_daily_reports.py` — Orchestrator, runs all scripts daily at 6:30 AM GT
- `fetch_odoo_attachments.py` — Pulls Odoo file attachments
- Shell scripts in repo handle cron setup and logging

## Setup

```bash
# 1. Clone
git clone git@github.com:YOUR_USER/torelo-kpi.git /opt/torelo-kpi/repo

# 2. Create config
cp config.example.py config.py
# Edit config.py with your database credentials

# 3. Install dependencies
python3 -m venv /opt/torelo-kpi/venv
source /opt/torelo-kpi/venv/bin/activate
pip install -r requirements.txt

# 4. Set up cron
bash setup_auto_attachments.sh
```