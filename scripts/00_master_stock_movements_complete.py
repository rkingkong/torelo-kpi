"""
TORELO Material Movement Extraction Script

NOTE: Usage Purpose Names - Using Reference Table
- Script includes hardcoded mapping for usage purpose IDs to descriptions
- Based on custom_usage_purpose export data provided by user
"""
import sys
sys.path.append('/opt/torelo-kpi/scripts')
from config import DATABASE_CONFIG, BASE_DIR, WEB_DIR
db_params = DATABASE_CONFIG
import os
import pandas as pd
from sqlalchemy import create_engine, text
import logging
from datetime import datetime

# Usage Purpose Reference Table (from custom_usage_purpose export)
USAGE_PURPOSE_MAPPING = {
    1: "SOTANO 5B",
    2: "SOTANO 5A", 
    3: "SOTANO 4B",
    4: "SOTANO 4A",
    5: "SOTANO 3B",
    6: "SOTANO 3A",
    7: "SOTANO 2B",
    8: "SOTANO 2A",
    9: "SOTANO 1B",
    10: "SOTANO 1A",
    11: "N1",
    12: "N2",
    13: "N3",
    14: "N4",
    15: "N5",
    16: "N6",
    17: "N7",
    18: "N8",
    19: "N9",
    20: "N10",
    21: "N11",
    22: "N12",
    23: "N13",
    24: "N14",
    25: "N15",
    26: "N16",
    27: "N17",
    28: "N18",
    29: "N20",
    30: "N21",
    31: "N22",
    32: "n8",
    33: "sotano5",
    34: "NIVEL PB",
    35: "N0",
    36: "n9.",
    37: "n111",
    38: "N6",
    39: "N7, n8",
    40: "N7",
    41: "N113.",
    42: "nivel 10",
    43: "nivel 12",
    44: "n133",
    45: "NIVEL 13",
    46: "n147",
    47: "#4"
}

# Configure logging - FIXED: Removed emojis and configured UTF-8 encoding
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.StreamHandler(),
        logging.FileHandler('torelo_material_movements.log', encoding='utf-8')
    ]
)

# Database configuration
DATABASE_CONFIG = {
    "dbname": "master",
    "user": "rudy",
    "password": "Kaibil10!z",
    "host": "3.18.71.107",
    "port": "5432"
}

# Create date-based output directory
date_folder = datetime.now().strftime('%Y-%m-%d')
output_dir = os.path.join(BASE_DIR, date_folder)
os.makedirs(output_dir, exist_ok=True)

# Output file paths
OUTPUT_PATH = os.path.join(output_dir, "master_stock_movements.csv")
EXCEL_OUTPUT_PATH = os.path.join(output_dir, "Material_Movement_Report.xlsx")

# FIXED COMPREHENSIVE MATERIAL MOVEMENT QUERY - Removed problematic stock_production_lot table
query_material_movements = """
SELECT
    -- Core Movement Information
    sm.id AS movement_id,
    sm.name AS movement_reference,
    DATE(sm.date) AS movement_date,
    sm.date AS movement_datetime,
    sm.create_date AS created_datetime,
    
    -- Product Details
    COALESCE(pt.name->>'es_GT', pt.name->>'en_US', pt.name::text) AS product_name,
    pp.default_code AS product_code,
    pp.barcode AS product_barcode,
    pc.name AS product_category,
    pt.type AS product_type,
    
    -- Movement Quantities and Values
    sm.product_uom_qty AS quantity,
    sm.quantity_done AS quantity_confirmed,
    COALESCE(uom.name->>'es_GT', uom.name->>'en_US', uom.name::text) AS unit_of_measure,
    COALESCE(ip.value_float, 0) AS unit_cost,
    (sm.product_uom_qty * COALESCE(ip.value_float, 0)) AS total_value,
    sm.price_unit AS transaction_price,
    
    -- Location Information (From and To)
    sl_source.name AS source_location,
    sl_source.complete_name AS source_location_full_path,
    sl_source.usage AS source_location_type,
    sl_dest.name AS destination_location,
    sl_dest.complete_name AS destination_location_full_path,
    sl_dest.usage AS destination_location_type,
    
    -- Movement Classification
    CASE
        WHEN sm.location_dest_id = 8 THEN 'Stock In'
        WHEN sm.location_id = 8 THEN 'Stock Out'
        WHEN sl_source.usage = 'supplier' THEN 'Receipt from Supplier'
        WHEN sl_dest.usage = 'customer' THEN 'Delivery to Customer'
        WHEN sl_source.usage = 'production' OR sl_dest.usage = 'production' THEN 'Production Movement'
        WHEN sl_source.usage = 'inventory' OR sl_dest.usage = 'inventory' THEN 'Inventory Adjustment'
        ELSE 'Internal Transfer'
    END AS movement_type,
    
    -- Status and State
    sm.state AS movement_status,
    sm.procure_method AS procurement_method,
    
    -- Usage Purpose (IDs only due to permission restrictions)
    CAST(sm.usage_purpose_id AS text) AS usage_purpose_id,
    
    -- Document Origins and References
    sm.origin AS document_origin,
    sm.reference AS internal_reference,
    sp.name AS operation_reference,
    sp.origin AS picking_origin,
    sp.note AS movement_notes,
    
    -- Person Responsible
    create_partner.name AS created_by_person,
    create_user.login AS created_by_login,
    write_partner.name AS last_modified_by,
    write_user.login AS modified_by_login,
    
    -- Partner Information (Supplier/Customer)
    partner.name AS partner_name,
    CASE 
        WHEN sl_source.usage = 'supplier' THEN 'Supplier'
        WHEN sl_dest.usage = 'customer' THEN 'Customer'
        ELSE 'Internal'
    END AS partner_type,
    
    -- Purchase Order Information (What it came from)
    po.name AS purchase_order,
    po.date_order AS po_order_date,
    po.state AS po_status,
    po_vendor.name AS vendor_name,
    pol.name AS po_line_description,
    
    -- Sale Order Information (Where it's going)
    so.name AS sale_order,
    so.date_order AS so_order_date,
    so.state AS so_status,
    so_customer.name AS customer_name,
    sol.name AS so_line_description,
    
    -- Manufacturing Information (What it will be used for)
    mo.name AS manufacturing_order,
    mo.state AS mo_status,
    mo.date_planned_start AS mo_planned_start,
    mo.date_planned_finished AS mo_planned_finish,
    mo_product.name AS manufactured_product,
    
    -- Project Information (Purpose/Usage)
    COALESCE(mo_aa.name, so_aa.name) AS project_name,
    pt_task.name AS task_name,
    pt_task.date_deadline AS task_deadline,
    stage.name AS task_status,
    
    -- Lot/Serial Information (FIXED: Simplified without stock_production_lot table)
    sml.lot_id AS lot_serial_number,
    CAST(sml.lot_id AS text) AS lot_serial_name,
    
    -- Additional Context
    emp.name AS employee_name,
    wh.name AS warehouse,
    company.name AS company
    
FROM stock_move sm
LEFT JOIN product_product pp ON sm.product_id = pp.id
LEFT JOIN product_template pt ON pp.product_tmpl_id = pt.id
LEFT JOIN product_category pc ON pt.categ_id = pc.id

-- Location details
LEFT JOIN stock_location sl_source ON sm.location_id = sl_source.id
LEFT JOIN stock_location sl_dest ON sm.location_dest_id = sl_dest.id

-- UOM and Cost
LEFT JOIN uom_uom uom ON sm.product_uom = uom.id
LEFT JOIN ir_property ip ON ip.res_id = CONCAT('product.product,', pp.id) 
    AND ip.name = 'standard_price'

-- Stock move lines for lot tracking (FIXED: Removed problematic join to stock_production_lot)
LEFT JOIN stock_move_line sml ON sm.id = sml.move_id

-- Stock picking information
LEFT JOIN stock_picking sp ON sm.picking_id = sp.id
LEFT JOIN res_partner partner ON sp.partner_id = partner.id

-- User information (who did the transaction)
LEFT JOIN res_users create_user ON sm.create_uid = create_user.id
LEFT JOIN res_users write_user ON sm.write_uid = write_user.id
LEFT JOIN res_partner create_partner ON create_user.partner_id = create_partner.id
LEFT JOIN res_partner write_partner ON write_user.partner_id = write_partner.id

-- Employee information
LEFT JOIN hr_employee emp ON create_user.id = emp.user_id

-- Purchase Order connections (where it came from)
LEFT JOIN purchase_order_line pol ON sm.purchase_line_id = pol.id
LEFT JOIN purchase_order po ON pol.order_id = po.id
LEFT JOIN res_partner po_vendor ON po.partner_id = po_vendor.id

-- Sale Order connections (where it's going)
LEFT JOIN sale_order_line sol ON sm.sale_line_id = sol.id
LEFT JOIN sale_order so ON sol.order_id = so.id
LEFT JOIN res_partner so_customer ON so.partner_id = so_customer.id

-- Manufacturing Order connections (what it will be used for)
LEFT JOIN mrp_production mo ON sm.production_id = mo.id
LEFT JOIN product_product mo_product_product ON mo.product_id = mo_product_product.id
LEFT JOIN product_template mo_product ON mo_product_product.product_tmpl_id = mo_product.id

-- Project connections (purpose/usage)
LEFT JOIN account_analytic_account mo_aa ON mo.analytic_account_id = mo_aa.id
LEFT JOIN account_analytic_account so_aa ON so.analytic_account_id = so_aa.id

-- Project task information
LEFT JOIN project_task pt_task ON COALESCE(mo_aa.id, so_aa.id) = pt_task.analytic_account_id
LEFT JOIN project_task_type stage ON pt_task.stage_id = stage.id

-- Warehouse and company
LEFT JOIN stock_warehouse wh ON sl_source.warehouse_id = wh.id OR sl_dest.warehouse_id = wh.id
LEFT JOIN res_company company ON wh.company_id = company.id

WHERE
    sm.state = 'done'  -- Only completed movements
    AND sm.product_uom_qty != 0  -- Only actual movements
    AND (sm.location_id = 8 OR sm.location_dest_id = 8)  -- Main stock location movements

ORDER BY sm.date DESC, sm.id DESC;
"""

def create_database_connection():
    """Create SQLAlchemy database connection"""
    try:
        conn_str = f"postgresql://{DATABASE_CONFIG['user']}:{DATABASE_CONFIG['password']}@{DATABASE_CONFIG['host']}:{DATABASE_CONFIG['port']}/{DATABASE_CONFIG['dbname']}"
        engine = create_engine(conn_str)
        
        # Test connection
        with engine.connect() as conn:
            result = conn.execute(text("SELECT version();"))
            version = result.fetchone()[0]
            logging.info(f"Database connected: {version[:50]}...")
            
        return engine
    except Exception as e:
        logging.error(f"Database connection failed: {e}")
        raise

def fetch_material_movements(engine, query):
    """Fetch comprehensive material movement data"""
    try:
        logging.info("Extracting material movement data...")
        
        # Get record count first
        count_query = text("""
        SELECT COUNT(*) FROM stock_move sm
        WHERE sm.state = 'done' 
            AND sm.product_uom_qty != 0
            AND (sm.location_id = 8 OR sm.location_dest_id = 8);
        """)
        
        with engine.connect() as conn:
            result = conn.execute(count_query)
            expected_count = result.fetchone()[0]
            logging.info(f"Expected records: {expected_count:,}")
        
        # Execute main query
        start_time = datetime.now()
        df = pd.read_sql_query(query, engine)
        duration = (datetime.now() - start_time).total_seconds()
        
        logging.info(f"SUCCESS: Extracted {len(df):,} material movements in {duration:.1f} seconds")
        return df
        
    except Exception as e:
        logging.error(f"Error extracting data: {e}")
        raise

def map_usage_purpose(usage_purpose_id):
    """Map usage purpose ID to description using reference table"""
    if pd.isna(usage_purpose_id) or usage_purpose_id == '':
        return ''
    try:
        purpose_id = int(float(str(usage_purpose_id)))
        return USAGE_PURPOSE_MAPPING.get(purpose_id, f"Unknown Purpose ID: {purpose_id}")
    except (ValueError, TypeError):
        return ''

def clean_and_enhance_data(df):
    """Clean and enhance the material movement data"""
    try:
        logging.info("Processing material movement data...")
        
        # Convert dates
        date_columns = ['movement_date', 'movement_datetime', 'created_datetime', 
                       'po_order_date', 'so_order_date', 'mo_planned_start', 
                       'mo_planned_finish', 'task_deadline']
        
        for col in date_columns:
            if col in df.columns:
                df[col] = pd.to_datetime(df[col], errors='coerce')
        
        # Clean text fields
        text_columns = df.select_dtypes(include=['object']).columns
        df[text_columns] = df[text_columns].fillna('')
        
        # Clean numeric fields
        numeric_columns = ['quantity', 'quantity_confirmed', 'unit_cost', 'total_value', 'transaction_price']
        for col in numeric_columns:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
        
        # Map usage purpose IDs to descriptions
        logging.info("Mapping usage purpose IDs to descriptions...")
        df['usage_purpose'] = df['usage_purpose_id'].apply(map_usage_purpose)
        
        # Add helpful summary columns
        df['value_per_unit'] = df.apply(lambda row: row['total_value'] / row['quantity'] if row['quantity'] > 0 else 0, axis=1)
        df['has_project'] = df['project_name'] != ''
        df['has_purchase_order'] = df['purchase_order'] != ''
        df['has_manufacturing_order'] = df['manufacturing_order'] != ''
        df['has_sale_order'] = df['sale_order'] != ''
        df['has_usage_purpose'] = (df['usage_purpose'] != '') | (df['usage_purpose_id'] != '')
        
        # Create a comprehensive description
        df['transaction_summary'] = df.apply(create_transaction_summary, axis=1)
        
        logging.info("SUCCESS: Data processing completed")
        return df
        
    except Exception as e:
        logging.error(f"Error processing data: {e}")
        raise

def create_transaction_summary(row):
    """Create a comprehensive summary of each transaction"""
    summary_parts = []
    
    # Basic movement
    summary_parts.append(f"{row['quantity']:.2f} {row['unit_of_measure']} of {row['product_name']}")
    summary_parts.append(f"moved from {row['source_location']} to {row['destination_location']}")
    
    # Usage Purpose (mapped from reference table)
    if row['usage_purpose']:
        summary_parts.append(f"for purpose: {row['usage_purpose']}")
    elif row['usage_purpose_id']:
        summary_parts.append(f"usage purpose ID: {row['usage_purpose_id']}")
    
    # Movement notes from stock picking
    if row['movement_notes']:
        summary_parts.append(f"notes: {row['movement_notes']}")
    
    # Purpose/Origin
    if row['purchase_order']:
        summary_parts.append(f"via Purchase Order {row['purchase_order']}")
        if row['vendor_name']:
            summary_parts.append(f"from {row['vendor_name']}")
    
    if row['sale_order']:
        summary_parts.append(f"for Sale Order {row['sale_order']}")
        if row['customer_name']:
            summary_parts.append(f"to {row['customer_name']}")
    
    if row['manufacturing_order']:
        summary_parts.append(f"for Manufacturing Order {row['manufacturing_order']}")
        if row['manufactured_product']:
            summary_parts.append(f"to produce {row['manufactured_product']}")
    
    if row['project_name']:
        summary_parts.append(f"for project: {row['project_name']}")
        if row['task_name']:
            summary_parts.append(f"task: {row['task_name']}")
    
    # Value
    if row['total_value'] > 0:
        summary_parts.append(f"(Value: Q{row['total_value']:.2f})")
    
    # Person responsible
    if row['created_by_person']:
        summary_parts.append(f"by {row['created_by_person']}")
    
    return " ".join(summary_parts)

def generate_summary_statistics(df):
    """Generate summary statistics for the report"""
    try:
        logging.info("Generating summary statistics...")
        
        stats = {
            'total_movements': len(df),
            'total_value': df['total_value'].sum(),
            'stock_in_movements': len(df[df['movement_type'] == 'Stock In']),
            'stock_out_movements': len(df[df['movement_type'] == 'Stock Out']),
            'date_range': f"{df['movement_date'].min()} to {df['movement_date'].max()}",
            'unique_products': df['product_name'].nunique(),
            'unique_categories': df['product_category'].nunique(),
            'movements_with_projects': len(df[df['has_project']]),
            'movements_with_purchase_orders': len(df[df['has_purchase_order']]),
            'movements_with_manufacturing': len(df[df['has_manufacturing_order']]),
            'movements_with_sales': len(df[df['has_sale_order']]),
            'movements_with_usage_purpose': len(df[df['has_usage_purpose']]),
            'unique_usage_purposes': df['usage_purpose'].nunique() if df['usage_purpose'].nunique() > 0 else 0,
        }
        
        for key, value in stats.items():
            if isinstance(value, float):
                logging.info(f"  {key}: Q{value:,.2f}" if 'value' in key else f"  {key}: {value:,.2f}")
            else:
                logging.info(f"  {key}: {value:,}" if isinstance(value, int) else f"  {key}: {value}")
        
        return stats
        
    except Exception as e:
        logging.error(f"Error generating statistics: {e}")
        return {}

def clean_for_excel(df):
    """Clean DataFrame for Excel compatibility"""
    df_clean = df.copy()
    
    # Clean text columns to remove illegal Excel characters
    for col in df_clean.select_dtypes(include=['object']).columns:
        df_clean[col] = df_clean[col].astype(str).apply(lambda x: 
            ''.join(char for char in x if ord(char) >= 32 and ord(char) != 127) if x else x
        )
    
    return df_clean

def save_to_files(df, stats):
    """Save the material movement data to CSV and Excel files"""
    try:
        # Ensure directories exist
        os.makedirs(os.path.dirname(OUTPUT_PATH), exist_ok=True)
        os.makedirs(os.path.dirname(EXCEL_OUTPUT_PATH), exist_ok=True)
        
        # Save to CSV
        logging.info("Saving comprehensive CSV report...")
        df.to_csv(OUTPUT_PATH, index=False, encoding='utf-8-sig')
        file_size_mb = os.path.getsize(OUTPUT_PATH) / (1024 * 1024)
        logging.info(f"SUCCESS: CSV saved: {OUTPUT_PATH} ({file_size_mb:.2f} MB)")
        
        # Clean data for Excel compatibility
        logging.info("Cleaning data for Excel compatibility...")
        df_excel = clean_for_excel(df)
        
        # Save to Excel with multiple sheets
        logging.info("Creating Excel report with multiple sheets...")
        with pd.ExcelWriter(EXCEL_OUTPUT_PATH, engine='openpyxl') as writer:
            # Main data
            df_excel.to_excel(writer, sheet_name='Material Movements', index=False)
            
            # Summary by movement type
            summary_by_type = df_excel.groupby('movement_type').agg({
                'quantity': 'sum',
                'total_value': 'sum',
                'movement_id': 'count'
            }).rename(columns={'movement_id': 'count'})
            summary_by_type.to_excel(writer, sheet_name='Summary by Type')
            
            # Summary by category
            summary_by_category = df_excel.groupby('product_category').agg({
                'quantity': 'sum',
                'total_value': 'sum',
                'movement_id': 'count'
            }).rename(columns={'movement_id': 'count'}).sort_values('total_value', ascending=False)
            summary_by_category.to_excel(writer, sheet_name='Summary by Category')
            
            # Summary by usage purpose (if data exists)
            if df_excel['usage_purpose'].nunique() > 1:  # More than just empty values
                summary_by_purpose = df_excel[df_excel['usage_purpose'] != ''].groupby('usage_purpose').agg({
                    'quantity': 'sum',
                    'total_value': 'sum',
                    'movement_id': 'count'
                }).rename(columns={'movement_id': 'count'}).sort_values('total_value', ascending=False)
                summary_by_purpose.to_excel(writer, sheet_name='Summary by Purpose')
            
            # Monthly summary
            df_excel['month'] = df_excel['movement_date'].dt.to_period('M')
            monthly_summary = df_excel.groupby('month').agg({
                'quantity': 'sum',
                'total_value': 'sum',
                'movement_id': 'count'
            }).rename(columns={'movement_id': 'count'})
            monthly_summary.to_excel(writer, sheet_name='Monthly Summary')
            
            # Statistics summary
            stats_df = pd.DataFrame(list(stats.items()), columns=['Metric', 'Value'])
            stats_df.to_excel(writer, sheet_name='Statistics', index=False)
        
        logging.info(f"SUCCESS: Excel report saved: {EXCEL_OUTPUT_PATH}")
        
    except Exception as e:
        logging.error(f"Error saving files: {e}")
        raise

def main():
    """Main execution function"""
    logging.info("TORELO MATERIAL MOVEMENT EXTRACTION STARTED")
    logging.info("=" * 60)
    logging.info("Using built-in usage purpose reference table for ID mapping")
    start_time = datetime.now()
    
    try:
        # Connect to database
        engine = create_database_connection()
        
        try:
            # Extract data
            df = fetch_material_movements(engine, query_material_movements)
            
            # Process data
            df_processed = clean_and_enhance_data(df)
            
            # Generate statistics
            stats = generate_summary_statistics(df_processed)
            
            # Save results
            save_to_files(df_processed, stats)
            
            # Final summary
            duration = (datetime.now() - start_time).total_seconds()
            logging.info("=" * 60)
            logging.info(f"SUCCESS: EXTRACTION COMPLETED in {duration:.1f} seconds")
            logging.info(f"Total movements extracted: {len(df_processed):,}")
            logging.info(f"Total value: Q{df_processed['total_value'].sum():,.2f}")
            logging.info(f"Usage purposes mapped: {df_processed['usage_purpose'].nunique()} unique purposes")
            logging.info(f"Files saved:")
            logging.info(f"   CSV: {OUTPUT_PATH}")
            logging.info(f"   Excel: {EXCEL_OUTPUT_PATH}")
            logging.info("=" * 60)
            
        finally:
            engine.dispose()
            
    except Exception as e:
        logging.error(f"EXTRACTION FAILED: {e}")
        raise

if __name__ == "__main__":
    main()