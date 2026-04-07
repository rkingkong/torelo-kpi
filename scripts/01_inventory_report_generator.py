#!/usr/bin/env python3
"""
Complete Corrected Inventory Report
Shows: En Mano, Demanda MO, Total Consumido (quantity_done), Disponible Real
"""
import sys
sys.path.append('/opt/torelo-kpi/scripts')
from config import DATABASE_CONFIG, BASE_DIR, WEB_DIR
import psycopg2
import pandas as pd
from datetime import datetime
import os
import re

def clean_for_excel(text):
    """Remove illegal characters for Excel"""
    if pd.isna(text):
        return text
    # Convert to string
    text = str(text)
    # Remove illegal XML characters
    illegal_chars = [(0x00, 0x08), (0x0B, 0x0C), (0x0E, 0x1F)]
    for start, end in illegal_chars:
        for char in range(start, end + 1):
            text = text.replace(chr(char), '')
    # Replace special characters that might cause issues
    text = text.replace('\x00', '')  # Null character
    text = text.replace('\r\n', ' ')  # Line endings
    text = text.replace('\n', ' ')
    text = text.replace('\r', ' ')
    return text

def generate_inventory_report():
    """Generate the inventory report with correct consumed quantities"""
    
    # Create date-based output directory
    date_folder = datetime.now().strftime('%Y-%m-%d')
    output_dir = os.path.join(BASE_DIR, date_folder)
    os.makedirs(output_dir, exist_ok=True)
    
    # Connect to database
    print("Connecting to database...")
    conn = psycopg2.connect(**DATABASE_CONFIG)
    
    # Main query
    query = """
    WITH 
    -- Current stock on hand
    current_inventory AS (
        SELECT 
            sq.product_id,
            SUM(sq.quantity) AS quantity_on_hand
        FROM stock_quant sq
        INNER JOIN stock_location sl ON sq.location_id = sl.id
        WHERE sl.usage = 'internal'
        GROUP BY sq.product_id
    ),
    -- Total demand from ALL open MOs
    mo_demand AS (
        SELECT 
            sm.product_id,
            SUM(sm.product_uom_qty) AS total_demand_qty
        FROM stock_move sm
        INNER JOIN mrp_production mo ON sm.raw_material_production_id = mo.id
        WHERE mo.state NOT IN ('done', 'cancel')
            AND sm.state NOT IN ('done', 'cancel')
        GROUP BY sm.product_id
    ),
    -- Total consumed = Sum of quantity_done for moves not in done or cancel state
    consumed_quantities AS (
        SELECT 
            sm.product_id,
            SUM(sm.quantity_done) AS total_consumed_qty
        FROM stock_move sm
        WHERE sm.raw_material_production_id IS NOT NULL 
            AND sm.state NOT IN ('done', 'cancel')
            AND sm.quantity_done > 0
        GROUP BY sm.product_id
    )
    SELECT 
        COALESCE(pc.complete_name, 'Sin Categoría') AS "Categoría",
        COALESCE(
            pt.name->>'es_GT',
            pt.name->>'en_US',
            pt.name::text
        ) AS "Producto",
        COALESCE(
            uom.name->>'es_GT',
            uom.name->>'en_US',
            uom.name::text
        ) AS "UDM",
        ROUND(COALESCE(ci.quantity_on_hand, 0), 2) AS "En Mano",
        ROUND(COALESCE(md.total_demand_qty, 0), 2) AS "Demanda MO Abierta",
        ROUND(COALESCE(cq.total_consumed_qty, 0), 2) AS "Total Consumido",
        ROUND(COALESCE(ci.quantity_on_hand, 0) - COALESCE(cq.total_consumed_qty, 0), 2) AS "Disponible Real"
    FROM product_product pp
    INNER JOIN product_template pt ON pp.product_tmpl_id = pt.id
    LEFT JOIN product_category pc ON pt.categ_id = pc.id
    LEFT JOIN uom_uom uom ON pt.uom_id = uom.id
    LEFT JOIN current_inventory ci ON pp.id = ci.product_id
    LEFT JOIN mo_demand md ON pp.id = md.product_id
    LEFT JOIN consumed_quantities cq ON pp.id = cq.product_id
    WHERE pp.active = true 
        AND pt.active = true
        AND pt.type = 'product'
        AND (COALESCE(ci.quantity_on_hand, 0) != 0 
            OR COALESCE(md.total_demand_qty, 0) != 0 
            OR COALESCE(cq.total_consumed_qty, 0) != 0)
    ORDER BY pc.complete_name, pt.name
    """
    
    print("Running query...")
    df = pd.read_sql_query(query, conn)
    
    # Clean all text columns for Excel compatibility
    print("Cleaning data for Excel...")
    text_columns = df.select_dtypes(include=['object']).columns
    for col in text_columns:
        df[col] = df[col].apply(clean_for_excel)
    
    # Generate filename with timestamp
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    filename = f'Reporte_Inventario_{timestamp}.xlsx'
    filepath = os.path.join(output_dir, filename)
    
    # Save to Excel
    print("Saving to Excel...")
    df.to_excel(filepath, index=False)
    
    print(f"\n✓ Reporte guardado en: {filepath}")
    print(f"✓ Total de productos: {len(df)}")
    
    # Quick summary
    print("\nResumen:")
    print(f"Total En Mano: {df['En Mano'].sum():,.2f}")
    print(f"Total Consumido: {df['Total Consumido'].sum():,.2f}")
    print(f"Total Disponible Real: {df['Disponible Real'].sum():,.2f}")
    
    conn.close()
    return filepath

if __name__ == "__main__":
    generate_inventory_report()