# Torelo KPI Dashboard

Automated business intelligence portal pulling data from Odoo via PostgreSQL.

**Live:** https://kpi.torelo.net

---

## Architecture

```
/opt/torelo-kpi/
├── scripts/                        ← Python report generators
│   ├── 00_master_stock_movements_complete.py   → CSV + XLSX
│   ├── 01_inventory_report_generator.py        → XLSX
│   ├── 02_cash_flow.py                         → XLSX + JSON
│   ├── 07_odoo_project_plan.py                 → XLSX
│   ├── 08_purchase_order_report.py             → XLSX
│   ├── 09_daily_attendance_report.py           → XLSX + JSON
│   ├── generate_attendance_json_fixed.py       → JSON (attendance helper)
│   ├── run_daily_reports.py                    → Orchestrator (cron 6:30 AM GT)
│   ├── fetch_odoo_attachments.py               → Pulls Odoo file attachments
│   ├── health_check.py                         → Generates health.json status
│   └── test_generate_json.py                   → Test/debug script
│
├── web/                            ← HTML frontends (this repo)
│   ├── theme.css                   → Shared design tokens & components
│   ├── nav.js                      → Shared sidebar navigation
│   ├── index.html                  → Main dashboard
│   ├── gantt-viewer.html           → Gantt chart for projects
│   ├── cashflow-viewer.html        → Cash flow explorer
│   ├── attendance-viewer.html      → Real-time attendance
│   ├── stock-explorer.html         → Stock movements explorer
│   ├── po-viewer.html              → Purchase order viewer
│   ├── inventory-viewer.html       → Inventory dashboard with alerts
│   ├── search.html                 → Document search interface
│   ├── chat-widget.html            → AI-powered chat assistant
│   ├── autoindex.html              → Simple file listing fallback
│   └── latest/                     → Symlinked latest reports (generated, not in repo)
│
├── email-search/                   ← Email search microservice
│   └── email_search_tool.py        → Flask API for email searching
│
├── reports/                        ← Generated Excel/CSV output (not in repo)
│   └── YYYY-MM-DD/                 → Daily report folders
│
├── logs/                           ← Script execution logs (not in repo)
├── venv/                           ← Python virtualenv (not in repo)
├── config.py                       ← Database & API credentials (not in repo)
├── config.example.py               ← Template for config.py
└── requirements.txt                ← Python dependencies
```

## Report Scripts

| Script | Output | Description |
|--------|--------|-------------|
| `00_master_stock_movements_complete.py` | CSV + XLSX | All stock movements from Odoo |
| `01_inventory_report_generator.py` | XLSX | Current inventory levels with demand & reorder points |
| `02_cash_flow.py` | XLSX + JSON | Weekly cash flow projections and payment tracking |
| `07_odoo_project_plan.py` | XLSX | Project plan with task dependencies and PO amounts |
| `08_purchase_order_report.py` | XLSX | Purchase order tracking, status, and vendor analysis |
| `09_daily_attendance_report.py` | XLSX + JSON | Employee attendance, clock-in/out, hours worked |

## Web Viewers

| File | Nav ID | Description |
|------|--------|-------------|
| `index.html` | `dashboard` | Main dashboard with KPIs, report downloads, and archive |
| `gantt-viewer.html` | `gantt` | MS Project-style Gantt chart with dependencies |
| `cashflow-viewer.html` | `cashflow` | Weekly cash flow projections and liquidity analysis |
| `attendance-viewer.html` | `attendance` | Real-time personnel tracking with entry/exit logs |
| `stock-explorer.html` | `stock` | Interactive stock movement analysis with charts |
| `po-viewer.html` | `po` | Purchase order tracking with vendor and status filters |
| `inventory-viewer.html` | `inventory` | Inventory dashboard with stock level alerts |
| `search.html` | `search` | Full-text document search across Odoo attachments |
| `chat-widget.html` | — | AI chat overlay for natural-language document search |

## Shared Components

### `theme.css` — Design System
Single source of truth for all visual styling. Includes:
- CSS variables (colors, spacing, radii)
- Layout components (topbar, sidebar, content area)
- KPI cards, section cards, viewer cards
- Data tables, filters, pagination
- Bar charts, status pills, modals
- Responsive breakpoints and print styles

**Usage:** `<link rel="stylesheet" href="theme.css">`

### `nav.js` — Sidebar Navigation
Auto-injected sidebar with navigation, server status, and mobile hamburger.

**Usage:** `<script src="nav.js" data-active="stock"></script>`

Available `data-active` values: `dashboard`, `gantt`, `cashflow`, `attendance`, `stock`, `po`, `inventory`, `search`

### `health_check.py` — System Health
Generates `latest/health.json` with script statuses, file ages, alerts, and disk space. Called after `run_daily_reports.py` or via cron.

## Automation

- **`run_daily_reports.py`** — Orchestrator, runs all scripts daily at **6:30 AM (Guatemala time)**
- **`fetch_odoo_attachments.py`** — Pulls file attachments from Odoo
- **`health_check.py`** — Generates system health status JSON
- Shell scripts handle cron setup and logging

### Recommended cron entries:
```cron
# Daily reports at 6:30 AM Guatemala time
30 6 * * * /opt/torelo-kpi/venv/bin/python /opt/torelo-kpi/scripts/run_daily_reports.py >> /opt/torelo-kpi/logs/daily.log 2>&1

# Health check after reports (7:00 AM) and at noon
0 7,12 * * * /opt/torelo-kpi/venv/bin/python /opt/torelo-kpi/scripts/health_check.py >> /opt/torelo-kpi/logs/health.log 2>&1

# Odoo attachments every 2 hours during business hours
0 8,10,12,14,16,18 * * 1-5 /opt/torelo-kpi/venv/bin/python /opt/torelo-kpi/scripts/fetch_odoo_attachments.py >> /opt/torelo-kpi/logs/attachments.log 2>&1
```

## Setup

```bash
# 1. Clone
git clone git@github.com:YOUR_USER/torelo-kpi.git /opt/torelo-kpi

# 2. Create config
cp config.example.py config.py
# Edit config.py with your database credentials

# 3. Install dependencies
python3 -m venv /opt/torelo-kpi/venv
source /opt/torelo-kpi/venv/bin/activate
pip install -r requirements.txt

# 4. Create required directories
mkdir -p /opt/torelo-kpi/{logs,reports,web/latest}

# 5. Set up cron
bash setup_auto_attachments.sh
# Or manually add cron entries from above

# 6. Run first report generation
python scripts/run_daily_reports.py
python scripts/health_check.py
```

## Adding a New Viewer

1. Create `web/your-viewer.html` — use `<link rel="stylesheet" href="theme.css">` and `<script src="nav.js" data-active="your-id"></script>`
2. Add the nav entry in `web/nav.js` → `navItems` array
3. Add a viewer card on `index.html` if desired
4. The viewer should load data from `latest/your-data.xlsx` (or `.json`)

## Tech Stack

- **Backend:** Python 3, PostgreSQL (via Odoo), pandas, openpyxl, SQLAlchemy
- **Frontend:** Vanilla HTML/CSS/JS, DM Sans + JetBrains Mono fonts
- **Libraries:** PapaParse (CSV), SheetJS (XLSX), Chart.js (where used)
- **Server:** Nginx serving static files from `/opt/torelo-kpi/web/`
- **Data source:** Odoo ERP via direct PostgreSQL queries