import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import datetime
from collections import OrderedDict
import random
import io
import re
import numpy as np

# Page configuration
st.set_page_config(
    page_title="RGI MRP System Dashboard",
    page_icon="🏭",
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- Session State Initialization ---
if 'rm_stock' not in st.session_state:
    st.session_state.rm_stock = pd.DataFrame(columns=['RM Code', 'Quantity'])
if 'rm_po' not in st.session_state:
    st.session_state.rm_po = pd.DataFrame(columns=['RM Code', 'Quantity', 'Arrival Date'])
if 'fg_formulas' not in st.session_state:
    st.session_state.fg_formulas = pd.DataFrame(columns=['FG Code', 'RM Code', 'Quantity'])
if 'fg_analysis_order' not in st.session_state:
    st.session_state.fg_analysis_order = OrderedDict()
if 'fg_expected_capacity' not in st.session_state:
    st.session_state.fg_expected_capacity = {}
if 'calculation_margin' not in st.session_state:
    st.session_state.calculation_margin = 3
if 'fg_colors' not in st.session_state:
    st.session_state.fg_colors = {}
if 'analysis_completed' not in st.session_state:
    st.session_state.analysis_completed = False
if 'select_all_trigger' not in st.session_state:
    st.session_state.select_all_trigger = False
if 'multiselect_key' not in st.session_state:
    st.session_state.multiselect_key = 0

# NEW: Session states for replacement and dilution
if 'rm_replacement_rules' not in st.session_state:
    st.session_state.rm_replacement_rules = pd.DataFrame(columns=['Old RM Code', 'New RM Code'])
if 'rm_dilution_rules' not in st.session_state:
    st.session_state.rm_dilution_rules = pd.DataFrame(columns=['RM Code', 'Component RM Code', 'Percentage'])
if 'modified_fg_formulas' not in st.session_state:
    st.session_state.modified_fg_formulas = pd.DataFrame(columns=['FG Code', 'RM Code', 'Quantity'])
if 'formulas_modified' not in st.session_state:
    st.session_state.formulas_modified = False
if 'dilution_applied' not in st.session_state:
    st.session_state.dilution_applied = False

# App Title and Description
st.title("🏭 RGI MRP System Dashboard")
st.markdown("""
**Material Requirements Planning System for Production Planning**
- **Tab 1**: Upload and manage Raw Material stock & purchase orders
- **Tab 2**: Upload Finished Goods formulas and configure settings
- **Tab 3**: Apply RM code replacement rules
- **Tab 4**: Apply RM dilution breakdown rules
- **Tab 5**: Generate production planning with detailed analysis
""")

# Create 5 tabs
tab1, tab2, tab3, tab4, tab5 = st.tabs(["📦 Stock & PO Management", "🧪 FG Formulas & Settings", "🔄 RM Replacement", "💧 RM Dilution", "📊 Production Planning"])

# Function to generate or get color for FG code
def get_fg_color(fg_code):
    if fg_code not in st.session_state.fg_colors:
        random.seed(hash(fg_code) % 1000)
        colors_list = [
            '#1f77b4', '#ff7f0e', '#2ca02c', '#d62728', '#9467bd',
            '#8c564b', '#e377c2', '#7f7f7f', '#bcbd22', '#17becf',
            '#aec7e8', '#ffbb78', '#98df8a', '#ff9896', '#c5b0d5',
            '#c49c94', '#f7b6d2', '#c7c7c7', '#dbdb8d', '#9edae5'
        ]
        color = colors_list[len(st.session_state.fg_colors) % len(colors_list)]
        st.session_state.fg_colors[fg_code] = color
    return st.session_state.fg_colors[fg_code]

# Function to preserve 8-character codes
def preserve_8char_code(code):
    """Ensure codes are treated as 8-character strings with leading zeros"""
    if pd.isna(code):
        return ""
    
    # Convert to string first
    code_str = str(code)
    
    # Remove any whitespace
    code_str = code_str.strip()
    
    # Handle special case: if code looks like a number (e.g., 123 -> '00000123')
    if code_str.isdigit():
        # Ensure it's 8 characters with leading zeros
        return code_str.zfill(8)
    else:
        # For alphanumeric codes, ensure they're 8 characters
        # Pad with zeros on the right if shorter than 8
        if len(code_str) < 8:
            return code_str.ljust(8, '0')
        elif len(code_str) > 8:
            return code_str[:8]
        else:
            return code_str

# Function to apply RM replacement rules
def apply_rm_replacement():
    """Apply RM code replacement rules to FG formulas"""
    if st.session_state.rm_replacement_rules.empty or st.session_state.fg_formulas.empty:
        return st.session_state.fg_formulas.copy()
    
    # Create a mapping dictionary for faster lookups
    replacement_map = {}
    for _, row in st.session_state.rm_replacement_rules.iterrows():
        old_rm = preserve_8char_code(row['Old RM Code'])
        new_rm = preserve_8char_code(row['New RM Code'])
        replacement_map[old_rm] = new_rm
    
    # Apply replacements to FG formulas
    modified_formulas = st.session_state.fg_formulas.copy()
    
    # Replace RM codes
    modified_formulas['RM Code'] = modified_formulas['RM Code'].apply(
        lambda x: replacement_map.get(preserve_8char_code(x), preserve_8char_code(x))
    )
    
    # Group by FG Code and RM Code to sum quantities if same RM appears multiple times
    modified_formulas = modified_formulas.groupby(['FG Code', 'RM Code'])['Quantity'].sum().reset_index()
    
    return modified_formulas

# Function to apply dilution rules
def apply_dilution_rules(formulas_df):
    """Apply dilution rules to break down RM codes into components"""
    if st.session_state.rm_dilution_rules.empty:
        return formulas_df
    
    # Create a list to store the final diluted formulas
    diluted_formulas = []
    
    # Process each formula row
    for _, formula_row in formulas_df.iterrows():
        fg_code = formula_row['FG Code']
        rm_code = preserve_8char_code(formula_row['RM Code'])
        quantity = formula_row['Quantity']
        
        # Check if this RM has dilution rules
        dilution_components = st.session_state.rm_dilution_rules[
            st.session_state.rm_dilution_rules['RM Code'] == rm_code
        ]
        
        if dilution_components.empty:
            # No dilution for this RM, keep as is
            diluted_formulas.append({
                'FG Code': fg_code,
                'RM Code': rm_code,
                'Quantity': quantity
            })
        else:
            # Apply dilution: break down into components
            for _, dil_row in dilution_components.iterrows():
                component_rm = preserve_8char_code(dil_row['Component RM Code'])
                percentage = dil_row['Percentage']
                
                # Calculate component quantity
                component_qty = quantity * (percentage / 100.0)
                
                # Apply strict rule: if result is 0.0000, make it 0.0001
                if abs(component_qty) < 0.00005:  # Tolerance for 4 decimal places
                    component_qty = 0.0001
                
                # Round to 4 decimal places
                component_qty = round(component_qty, 4)
                
                diluted_formulas.append({
                    'FG Code': fg_code,
                    'RM Code': component_rm,
                    'Quantity': component_qty
                })
    
    # Create DataFrame and group by FG Code and RM Code to sum quantities
    diluted_df = pd.DataFrame(diluted_formulas)
    if not diluted_df.empty:
        diluted_df = diluted_df.groupby(['FG Code', 'RM Code'])['Quantity'].sum().reset_index()
    
    return diluted_df

# Function to generate HTML report (fallback if PDF fails)
def generate_html_report(results, shortage_details, prod_date, total_volume, ready_fgs, delayed_pos, po_status):
    html_content = f"""
    <!DOCTYPE html>
    <html>
    <head>
        <title>MRP Production Planning Summary Report</title>
        <style>
            body {{ font-family: Arial, sans-serif; margin: 40px; }}
            .header {{ text-align: center; margin-bottom: 30px; }}
            .title {{ font-size: 24px; font-weight: bold; color: #333; }}
            .subtitle {{ font-size: 14px; color: #666; margin-top: 10px; }}
            .section {{ margin: 20px 0; }}
            .section-title {{ font-size: 18px; font-weight: bold; color: #2c3e50; border-bottom: 2px solid #3498db; padding-bottom: 5px; margin-bottom: 15px; }}
            table {{ width: 100%; border-collapse: collapse; margin: 10px 0; }}
            th {{ background-color: #3498db; color: white; padding: 10px; text-align: left; }}
            td {{ padding: 8px; border: 1px solid #ddd; }}
            tr:nth-child(even) {{ background-color: #f2f2f2; }}
            .metric {{ display: inline-block; margin: 10px 20px 10px 0; padding: 10px; background-color: #ecf0f1; border-radius: 5px; }}
            .metric-label {{ font-weight: bold; color: #7f8c8d; }}
            .metric-value {{ font-size: 18px; color: #2c3e50; }}
            .status-ready {{ color: #27ae60; font-weight: bold; }}
            .status-shortage {{ color: #e74c3c; font-weight: bold; }}
            .footer {{ margin-top: 40px; text-align: center; color: #7f8c8d; font-size: 12px; border-top: 1px solid #ddd; padding-top: 20px; }}
            .company-footer {{ margin-top: 30px; text-align: center; color: #3498db; font-weight: bold; font-size: 14px; }}
        </style>
    </head>
    <body>
        <div class="header">
            <div class="title">MRP Production Planning Summary Report</div>
            <div class="subtitle">Generated on: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}</div>
            <div class="subtitle">Production Date: {prod_date.strftime('%d/%m/%Y')}</div>
        </div>
        
        <div class="section">
            <div class="section-title">Summary Metrics</div>
            <div class="metric">
                <div class="metric-label">Planned Production Date</div>
                <div class="metric-value">{prod_date.strftime('%d/%m/%Y')}</div>
            </div>
            <div class="metric">
                <div class="metric-label">Producible FG Types</div>
                <div class="metric-value">{len(ready_fgs)}</div>
            </div>
            <div class="metric">
                <div class="metric-label">Total Production Volume</div>
                <div class="metric-value">{total_volume:,.1f} Kg</div>
            </div>
            <div class="metric">
                <div class="metric-label">Delayed Purchase Orders</div>
                <div class="metric-value">{delayed_pos}</div>
            </div>
        </div>
    """
    
    # Production Capability List
    if results:
        html_content += """
        <div class="section">
            <div class="section-title">Production Capability List</div>
            <table>
                <tr>
                    <th>FG Code</th>
                    <th>Expected</th>
                    <th>Max (Kg)</th>
                    <th>Actual (Kg)</th>
                    <th>Status</th>
                    <th>Missing RM</th>
                    <th>Batches</th>
                </tr>
        """
        
        for item in results:
            status_class = "status-ready" if "✅" in item['Status'] else "status-shortage"
            html_content += f"""
                <tr>
                    <td>{item['FG']}</td>
                    <td>{item['Expected']}</td>
                    <td>{item['Max']}</td>
                    <td>{item['Actual']}</td>
                    <td class="{status_class}">{item['Status']}</td>
                    <td>{item['Missing']}</td>
                    <td>{item['Batches']}</td>
                </tr>
            """
        
        html_content += """
            </table>
        </div>
        """
    
    # Shortage Details
    shortage_exists = False
    for fg in shortage_details:
        if shortage_details[fg]:
            shortage_exists = True
            break
    
    if shortage_exists:
        html_content += """
        <div class="section">
            <div class="section-title">Shortage Details</div>
        """
        
        for fg in shortage_details:
            if shortage_details[fg]:
                html_content += f"""
                <div style="margin: 15px 0;">
                    <div style="font-weight: bold; color: #e74c3c;">FG Code: {fg}</div>
                    <ul style="margin: 5px 0 20px 20px;">
                """
                
                for item in shortage_details[fg]:
                    html_content += f"<li>{item}</li>")
                
                html_content += """
                    </ul>
                </div>
                """
        
        html_content += """
        </div>
        """
    
    # PO Delay Tracker
    if po_status is not None and not po_status.empty:
        html_content += """
        <div class="section">
            <div class="section-title">Purchase Order Delay Status</div>
            <table>
                <tr>
                    <th>RM Code</th>
                    <th>Quantity</th>
                    <th>Arrival Date</th>
                    <th>Status</th>
                </tr>
        """
        
        for _, row in po_status.iterrows():
            status_class = "status-shortage" if row['Status'] == 'Delayed' else ""
            html_content += f"""
                <tr>
                    <td>{row['RM Code']}</td>
                    <td>{row['Quantity']:,.4f} Kg</td>
                    <td>{row['Arrival Date'].strftime('%d/%m/%Y') if hasattr(row['Arrival Date'], 'strftime') else str(row['Arrival Date'])}</td>
                    <td class="{status_class}">{row['Status']}</td>
                </tr>
            """
        
        html_content += """
            </table>
        </div>
        """
    
    # Settings Information
    html_content += f"""
        <div class="section">
            <div class="section-title">System Settings</div>
            <div style="margin: 10px 0;">
                • Decimal Precision: {st.session_state.calculation_margin} places<br>
                • FIFO Order: {', '.join(st.session_state.fg_analysis_order.keys()) if st.session_state.fg_analysis_order else 'Not set'}<br>
                • Report Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}
            </div>
        </div>
        
        <div class="company-footer">
            RGI - Supply Chain Department
        </div>
        
        <div class="footer">
            Report generated by MRP Dashboard System<br>
            --- End of Report ---
        </div>
    </body>
    </html>
    """
    
    return html_content

# Function to generate PDF report
def generate_report(results, shortage_details, prod_date, total_volume, ready_fgs, delayed_pos, po_status):
    try:
        from reportlab.lib.pagesizes import A4
        from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
        from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
        from reportlab.lib import colors
        
        buffer = io.BytesIO()
        
        doc = SimpleDocTemplate(buffer, pagesize=A4, 
                              rightMargin=72, leftMargin=72,
                              topMargin=72, bottomMargin=72)
        
        elements = []
        styles = getSampleStyleSheet()
        
        title_style = ParagraphStyle(
            'CustomTitle',
            parent=styles['Heading1'],
            fontSize=18,
            spaceAfter=30,
            alignment=1
        )
        
        heading_style = ParagraphStyle(
            'CustomHeading',
            parent=styles['Heading2'],
            fontSize=14,
            spaceAfter=12,
            spaceBefore=12
        )
        
        normal_style = styles['Normal']
        
        elements.append(Paragraph("MRP Production Planning Summary Report", title_style))
        elements.append(Paragraph(f"Generated on: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}", normal_style))
        elements.append(Paragraph(f"Production Date: {prod_date.strftime('%d/%m/%Y')}", normal_style))
        elements.append(Spacer(1, 20))
        
        elements.append(Paragraph("Summary Metrics", heading_style))
        
        summary_data = [
            ["Metric", "Value"],
            ["Planned Production Date", prod_date.strftime('%d/%m/%Y')],
            ["Producible FG Types", str(len(ready_fgs))],
            ["Total Production Volume", f"{total_volume:,.1f} Kg"],
            ["Delayed Purchase Orders", str(delayed_pos)]
        ]
        
        summary_table = Table(summary_data, colWidths=[200, 150])
        summary_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 12),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
            ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
        ]))
        elements.append(summary_table)
        elements.append(Spacer(1, 20))
        
        elements.append(Paragraph("Production Capability List", heading_style))
        
        if results:
            table_data = [["FG Code", "Expected", "Max (Kg)", "Actual (Kg)", "Status", "Missing RM", "Batches"]]
            
            for item in results:
                table_data.append([
                    item['FG'],
                    item['Expected'],
                    item['Max'],
                    item['Actual'],
                    item['Status'],
                    item['Missing'],
                    str(item['Batches'])
                ])
            
            prod_table = Table(table_data, colWidths=[70, 60, 60, 60, 60, 70, 50])
            prod_table.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                ('FONTSIZE', (0, 0), (-1, 0), 10),
                ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                ('BACKGROUND', (0, 1), (-1, -1), colors.whitesmoke),
                ('GRID', (0, 0), (-1, -1), 1, colors.black),
                ('FONTSIZE', (0, 1), (-1, -1), 9),
            ]))
            elements.append(prod_table)
            elements.append(Spacer(1, 20))
        
        shortage_exists = False
        for fg in shortage_details:
            if shortage_details[fg]:
                shortage_exists = True
                break
        
        if shortage_exists:
            elements.append(Paragraph("Shortage Details", heading_style))
            
            for fg in shortage_details:
                if shortage_details[fg]:
                    elements.append(Paragraph(f"FG Code: {fg}", styles['Heading3']))
                    for item in shortage_details[fg]:
                        elements.append(Paragraph(f"• {item}", normal_style))
                    elements.append(Spacer(1, 10))
            elements.append(Spacer(1, 20))
        
        if po_status is not None and not po_status.empty:
            elements.append(Paragraph("Purchase Order Delay Status", heading_style))
            
            po_data = [["RM Code", "Quantity", "Arrival Date", "Status"]]
            
            for _, row in po_status.iterrows():
                po_data.append([
                    str(row['RM Code']),
                    f"{row['Quantity']:,.4f} Kg",
                    row['Arrival Date'].strftime('%d/%m/%Y') if hasattr(row['Arrival Date'], 'strftime') else str(row['Arrival Date']),
                    row['Status']
                ])
            
            po_table = Table(po_data, colWidths=[80, 80, 80, 80])
            po_table.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                ('FONTSIZE', (0, 0), (-1, 0), 10),
                ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                ('BACKGROUND', (0, 1), (-1, -1), colors.whitesmoke),
                ('GRID', (0, 0), (-1, -1), 1, colors.black),
                ('FONTSIZE', (0, 1), (-1, -1), 9),
            ]))
            elements.append(po_table)
            elements.append(Spacer(1, 20))
        
        elements.append(Paragraph("System Settings", heading_style))
        settings_text = f"""
        • Decimal Precision: {st.session_state.calculation_margin} places
        • FIFO Order: {', '.join(st.session_state.fg_analysis_order.keys()) if st.session_state.fg_analysis_order else 'Not set'}
        • Report Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}
        """
        elements.append(Paragraph(settings_text, normal_style))
        
        elements.append(Spacer(1, 30))
        elements.append(Paragraph("RGI - Supply Chain Department", ParagraphStyle(
            'CompanyFooter',
            parent=styles['Normal'],
            fontSize=12,
            alignment=1,
            textColor=colors.HexColor('#3498db'),
            spaceBefore=20
        )))
        
        elements.append(Spacer(1, 10))
        elements.append(Paragraph("Report generated by MRP Dashboard System", ParagraphStyle(
            'Footer',
            parent=styles['Normal'],
            fontSize=10,
            alignment=1,
            textColor=colors.grey
        )))
        elements.append(Paragraph("--- End of Report ---", ParagraphStyle(
            'Footer',
            parent=styles['Normal'],
            fontSize=10,
            alignment=1,
            textColor=colors.grey
        )))
        
        doc.build(elements)
        
        pdf_data = buffer.getvalue()
        buffer.close()
        
        return pdf_data, "pdf"
    
    except ImportError:
        html_content = generate_html_report(results, shortage_details, prod_date, total_volume, ready_fgs, delayed_pos, po_status)
        return html_content.encode('utf-8'), "html"

# Function to generate Excel with PDF format
def generate_pdf_format_excel(shortage_details, results, prod_date, calculation_margin):
    """Generate Excel file with shortage details formatted like PDF report"""
    output = io.BytesIO()
    
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        # Sheet 1: Executive Summary
        exec_summary = pd.DataFrame({
            'Report Type': ['MRP Production Planning Summary Report'],
            'Production Date': [prod_date.strftime('%d/%m/%Y')],
            'Report Generated': [datetime.now().strftime('%Y-%m-%d %H:%M:%S')],
            'Decimal Precision': [calculation_margin]
        })
        exec_summary.to_excel(writer, sheet_name='Executive Summary', index=False)
        
        # Sheet 2: Production Capability List (same as PDF)
        if results:
            prod_data = []
            for item in results:
                prod_data.append({
                    'FG Code': item['FG'],
                    'Expected': item['Expected'],
                    'Max (Kg)': item['Max'],
                    'Actual (Kg)': item['Actual'],
                    'Status': item['Status'],
                    'Missing RM': item['Missing'],
                    'Batches': item['Batches']
                })
            
            prod_df = pd.DataFrame(prod_data)
            prod_df.to_excel(writer, sheet_name='Production Capability List', index=False)
        
        # Sheet 3: Shortage Details (formatted exactly like PDF)
        shortage_data = []
        
        # Process shortage details to match PDF format
        for fg_code, items in shortage_details.items():
            if items:  # Only include FGs with shortages
                for item in items:
                    shortage_data.append({
                        'FG Code': fg_code,
                        'Shortage Details': item
                    })
        
        if shortage_data:
            shortage_df = pd.DataFrame(shortage_data)
            shortage_df.to_excel(writer, sheet_name='Shortage Details', index=False)
        
        # Sheet 4: Settings
        settings_data = pd.DataFrame({
            'Setting': ['Decimal Precision', 'FIFO Order', 'Report Generated'],
            'Value': [
                f"{calculation_margin} places",
                ', '.join(st.session_state.fg_analysis_order.keys()) if st.session_state.fg_analysis_order else 'Not set',
                datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            ]
        })
        settings_data.to_excel(writer, sheet_name='System Settings', index=False)
    
    output.seek(0)
    return output

# Function to add footer to all tabs
def add_footer():
    st.markdown("---")
    st.markdown(
        """
        <div style="text-align: center; color: #666; font-size: 12px; padding: 10px;">
            <strong>RGI - Supply Chain Department</strong> | MRP Dashboard System v2.0
        </div>
        """,
        unsafe_allow_html=True
    )

# --- TAB 1: STOCK & PO MANAGEMENT ---
with tab1:
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("📦 Raw Material Stock (RM)")
        
        if st.button("🔄 Clear RM Stock", key="clear_rm", help="Clear all RM stock data"):
            st.session_state.rm_stock = pd.DataFrame(columns=['RM Code', 'Quantity'])
            st.session_state.analysis_completed = False
            st.success("RM stock cleared!")
        
        st.markdown("**Upload RM Stock Excel File**")
        rm_file = st.file_uploader(
            "Choose Excel file", 
            type=['xlsx'], 
            key="rm_up",
            help="Upload Excel file with RM Code and Quantity columns"
        )
        
        if rm_file is not None:
            try:
                df = pd.read_excel(rm_file)
                df.columns = df.columns.str.strip()
                
                column_mapping = {}
                
                # Find RM Code column
                for col in df.columns:
                    col_lower = str(col).lower()
                    if 'rm' in col_lower and ('code' in col_lower or 'id' in col_lower):
                        column_mapping['RM Code'] = col
                        break
                
                # Find Quantity column
                for col in df.columns:
                    col_lower = str(col).lower()
                    if 'quantity' in col_lower or 'qty' in col_lower or 'amount' in col_lower:
                        column_mapping['Quantity'] = col
                        break
                
                if 'RM Code' in column_mapping and 'Quantity' in column_mapping:
                    processed_df = pd.DataFrame()
                    processed_df['RM Code'] = df[column_mapping['RM Code']].apply(preserve_8char_code)
                    processed_df['Quantity'] = pd.to_numeric(df[column_mapping['Quantity']], errors='coerce').fillna(0)
                    
                    # Filter out empty codes
                    processed_df = processed_df[processed_df['RM Code'] != '']
                    processed_df = processed_df[processed_df['RM Code'] != 'nan']
                    processed_df = processed_df.dropna(subset=['RM Code'])
                    
                    if not processed_df.empty:
                        st.session_state.rm_stock = processed_df
                        st.success(f"✅ Successfully loaded {len(processed_df)} RM stock records!")
                    else:
                        st.warning("No valid data found in the uploaded file")
                        
                else:
                    missing_cols = []
                    if 'RM Code' not in column_mapping:
                        missing_cols.append("RM Code")
                    if 'Quantity' not in column_mapping:
                        missing_cols.append("Quantity")
                    st.error(f"Missing columns: {', '.join(missing_cols)}")
                    
            except Exception as e:
                st.error(f"Error processing RM file: {str(e)}")
        
        # Display current stock
        if not st.session_state.rm_stock.empty:
            st.write("### 📊 Current Stock Inventory")
            
            display_df = st.session_state.rm_stock.copy()
            display_df['Quantity'] = display_df['Quantity'].apply(lambda x: f"{x:,.4f} Kg")
            
            st.dataframe(
                display_df,
                use_container_width=True,
                height=min(300, len(display_df) * 35 + 40),
                hide_index=True
            )
            
            # Summary metrics
            total_stock = st.session_state.rm_stock['Quantity'].sum()
            unique_rms = st.session_state.rm_stock['RM Code'].nunique()
            
            col_metric1, col_metric2 = st.columns(2)
            with col_metric1:
                st.metric("Unique RM Codes", unique_rms)
            with col_metric2:
                st.metric("Total Stock Quantity", f"{total_stock:,.2f} Kg")
            
            st.write("### 🔍 View RM Details")
            if not st.session_state.rm_stock.empty:
                sel_rm = st.selectbox(
                    "Select RM Code to view details:",
                    st.session_state.rm_stock['RM Code'].unique(),
                    key="rm_select"
                )
                if sel_rm:
                    qty_row = st.session_state.rm_stock[st.session_state.rm_stock['RM Code'] == sel_rm]
                    if not qty_row.empty:
                        qty = qty_row['Quantity'].values[0]
                        st.metric(label=f"Available Stock for {sel_rm}", value=f"{qty:,.4f} Kg")
        else:
            st.info("📁 No RM stock data loaded yet. Please upload an Excel file.")

    with col2:
        st.subheader("📅 RM in Purchase Orders (PO)")
        
        if st.button("🔄 Clear RM PO", key="clear_po", help="Clear all PO data"):
            st.session_state.rm_po = pd.DataFrame(columns=['RM Code', 'Quantity', 'Arrival Date'])
            st.session_state.analysis_completed = False
            st.success("PO data cleared!")
        
        st.markdown("**Upload RM PO Excel File**")
        po_file = st.file_uploader(
            "Choose Excel file", 
            type=['xlsx'], 
            key="po_up",
            help="Upload Excel file with RM Code, Quantity, and Arrival Date columns"
        )
        
        if po_file is not None:
            try:
                df_po = pd.read_excel(po_file)
                df_po.columns = df_po.columns.str.strip()
                
                column_mapping = {}
                
                # Find RM Code column
                for col in df_po.columns:
                    col_lower = str(col).lower()
                    if 'rm' in col_lower and ('code' in col_lower or 'id' in col_lower):
                        column_mapping['RM Code'] = col
                        break
                
                # Find Quantity column
                for col in df_po.columns:
                    col_lower = str(col).lower()
                    if 'quantity' in col_lower or 'qty' in col_lower or 'amount' in col_lower:
                        column_mapping['Quantity'] = col
                        break
                
                # Find Arrival Date column
                for col in df_po.columns:
                    col_lower = str(col).lower()
                    if 'arrival' in col_lower or 'date' in col_lower or 'delivery' in col_lower:
                        column_mapping['Arrival Date'] = col
                        break
                
                if all(col in column_mapping for col in ['RM Code', 'Quantity', 'Arrival Date']):
                    processed_df = pd.DataFrame()
                    processed_df['RM Code'] = df_po[column_mapping['RM Code']].apply(preserve_8char_code)
                    processed_df['Quantity'] = pd.to_numeric(df_po[column_mapping['Quantity']], errors='coerce').fillna(0)
                    
                    # Parse date column
                    date_col = df_po[column_mapping['Arrival Date']]
                    try:
                        processed_df['Arrival Date'] = pd.to_datetime(date_col, dayfirst=True, errors='coerce')
                    except:
                        try:
                            processed_df['Arrival Date'] = pd.to_datetime(date_col, errors='coerce')
                        except:
                            st.error("Could not parse date column")
                            processed_df['Arrival Date'] = pd.NaT
                    
                    # Filter out empty rows
                    processed_df = processed_df[processed_df['RM Code'] != '']
                    processed_df = processed_df[processed_df['RM Code'] != 'nan']
                    processed_df = processed_df.dropna(subset=['RM Code', 'Arrival Date'])
                    
                    if not processed_df.empty:
                        st.session_state.rm_po = processed_df
                        st.success(f"✅ Successfully loaded {len(processed_df)} PO records!")
                    else:
                        st.warning("No valid data found in the uploaded file")
                        
                else:
                    missing_cols = []
                    for col in ['RM Code', 'Quantity', 'Arrival Date']:
                        if col not in column_mapping:
                            missing_cols.append(col)
                    st.error(f"Missing columns: {', '.join(missing_cols)}")
                    
            except Exception as e:
                st.error(f"Error processing PO file: {str(e)}")
        
        # Display PO data
        if not st.session_state.rm_po.empty:
            st.write("### 📅 PO Schedule")
            
            display_po = st.session_state.rm_po.copy()
            display_po = display_po.sort_values(by='Arrival Date')
            display_po['Quantity'] = display_po['Quantity'].apply(lambda x: f"{x:,.4f} Kg")
            display_po['Arrival Date'] = display_po['Arrival Date'].dt.strftime('%d/%m/%Y')
            
            st.dataframe(
                display_po,
                use_container_width=True,
                height=min(300, len(display_po) * 35 + 40),
                hide_index=True
            )
            
            # PO summary
            total_po_qty = st.session_state.rm_po['Quantity'].sum()
            earliest_date = st.session_state.rm_po['Arrival Date'].min().strftime('%d/%m/%Y')
            latest_date = st.session_state.rm_po['Arrival Date'].max().strftime('%d/%m/%Y')
            
            col_po1, col_po2, col_po3 = st.columns(3)
            with col_po1:
                st.metric("Total PO Quantity", f"{total_po_qty:,.2f} Kg")
            with col_po2:
                st.metric("Earliest Arrival", earliest_date)
            with col_po3:
                st.metric("Latest Arrival", latest_date)
            
            st.write("### 📊 Total RM in PO by Code")
            if not st.session_state.rm_po.empty:
                total_po = st.session_state.rm_po.groupby('RM Code')['Quantity'].sum().reset_index()
                total_po['Quantity'] = total_po['Quantity'].apply(lambda x: f"{x:,.4f} Kg")
                
                st.dataframe(
                    total_po,
                    use_container_width=True,
                    hide_index=True
                )
        else:
            st.info("📁 No PO data loaded yet. Please upload an Excel file.")
    
    add_footer()

# --- TAB 2: FG FORMULAS & SETTINGS ---
with tab2:
    col1, col2 = st.columns([2, 1])
    
    with col1:
        st.subheader("🧪 Finished Goods Formula Management")
        
        st.markdown("**Upload FG Formulas Excel File(s)**")
        fg_files = st.file_uploader(
            "Choose Excel files (multiple allowed)", 
            type=['xlsx'], 
            accept_multiple_files=True,
            key="fg_uploader",
            help="Upload Excel files with FG Code, RM Code, and Quantity columns"
        )
        
        if fg_files:
            total_loaded = 0
            for f in fg_files:
                try:
                    new_fg = pd.read_excel(f)
                    new_fg.columns = new_fg.columns.str.strip()
                    
                    column_mapping = {}
                    
                    # Find FG Code column
                    for col in new_fg.columns:
                        col_lower = str(col).lower()
                        if 'fg' in col_lower and ('code' in col_lower or 'id' in col_lower):
                            column_mapping['FG Code'] = col
                            break
                    
                    # Find RM Code column
                    for col in new_fg.columns:
                        col_lower = str(col).lower()
                        if 'rm' in col_lower and ('code' in col_lower or 'id' in col_lower):
                            column_mapping['RM Code'] = col
                            break
                    
                    # Find Quantity column
                    for col in new_fg.columns:
                        col_lower = str(col).lower()
                        if 'quantity' in col_lower or 'qty' in col_lower:
                            column_mapping['Quantity'] = col
                            break
                    
                    if all(col in column_mapping for col in ['FG Code', 'RM Code', 'Quantity']):
                        processed_fg = pd.DataFrame()
                        processed_fg['FG Code'] = new_fg[column_mapping['FG Code']].apply(preserve_8char_code)
                        processed_fg['RM Code'] = new_fg[column_mapping['RM Code']].apply(preserve_8char_code)
                        processed_fg['Quantity'] = pd.to_numeric(new_fg[column_mapping['Quantity']], errors='coerce').fillna(0)
                        
                        # Filter out empty rows
                        processed_fg = processed_fg[
                            (processed_fg['FG Code'] != '') & 
                            (processed_fg['FG Code'] != 'nan') &
                            (processed_fg['RM Code'] != '') &
                            (processed_fg['RM Code'] != 'nan')
                        ]
                        processed_fg = processed_fg.dropna(subset=['FG Code', 'RM Code'])
                        
                        if not processed_fg.empty:
                            if st.session_state.fg_formulas.empty:
                                st.session_state.fg_formulas = processed_fg
                            else:
                                combined = pd.concat([st.session_state.fg_formulas, processed_fg])
                                st.session_state.fg_formulas = combined.drop_duplicates(
                                    subset=['FG Code', 'RM Code'], 
                                    keep='first'
                                ).reset_index(drop=True)
                            
                            # Assign colors to new FG codes
                            for fg_code in processed_fg['FG Code'].unique():
                                if fg_code not in st.session_state.fg_colors:
                                    get_fg_color(fg_code)
                            
                            total_loaded += len(processed_fg)
                            st.success(f"✅ Loaded {len(processed_fg)} formula entries from {f.name}")
                        else:
                            st.warning(f"No valid data found in {f.name}")
                            
                    else:
                        missing_cols = []
                        for col in ['FG Code', 'RM Code', 'Quantity']:
                            if col not in column_mapping:
                                missing_cols.append(col)
                        st.error(f"{f.name}: Missing columns {', '.join(missing_cols)}")
                        
                except Exception as e:
                    st.error(f"Error reading {f.name}: {str(e)}")
            
            if total_loaded > 0:
                st.success(f"✅ Total: Loaded {total_loaded} formula entries from {len(fg_files)} file(s)")
        
        # Display current formulas
        if not st.session_state.fg_formulas.empty:
            st.divider()
            st.write("### 📋 Current FG Formulas")
            
            display_fg = st.session_state.fg_formulas.copy()
            display_fg['Quantity'] = display_fg['Quantity'].apply(lambda x: f"{x:,.4f} Kg")
            
            # Summary metrics
            total_fgs = display_fg['FG Code'].nunique()
            total_rms = display_fg['RM Code'].nunique()
            total_entries = len(display_fg)
            
            col_fg1, col_fg2, col_fg3 = st.columns(3)
            with col_fg1:
                st.metric("Unique FG Codes", total_fgs)
            with col_fg2:
                st.metric("Unique RM Codes", total_rms)
            with col_fg3:
                st.metric("Total Entries", total_entries)
            
            st.dataframe(
                display_fg,
                use_container_width=True,
                height=min(400, len(display_fg) * 35 + 40),
                hide_index=True
            )
            
            st.divider()
            st.write("### 🔍 Select FG Codes for Production Analysis")
            
            fg_codes = sorted(st.session_state.fg_formulas['FG Code'].unique().tolist())
            
            # Select All functionality
            col_select, col_button = st.columns([4, 1])
            
            with col_select:
                # Get current selection
                if st.session_state.select_all_trigger:
                    current_selection = fg_codes.copy()
                else:
                    current_selection = list(st.session_state.fg_analysis_order.keys())
                
                # Create the multiselect widget with dynamic key
                selected_fgs = st.multiselect(
                    "Select FG Codes (order determines FIFO priority):",
                    options=fg_codes,
                    default=current_selection,
                    key=f"fg_analysis_select_{st.session_state.multiselect_key}",
                    help="Select FG codes to include in production planning. Order matters for FIFO allocation."
                )
            
            with col_button:
                st.write("")  # Spacing
                st.write("")  # Spacing
                if st.button("📋 Select All", key="select_all_fg_button", type="secondary"):
                    # Set the trigger and force widget refresh
                    st.session_state.select_all_trigger = True
                    st.session_state.multiselect_key += 1
                    st.rerun()
            
            # Update the analysis order based on current selection
            if selected_fgs:
                # Sort selected FGs alphabetically/numerically
                sorted_selected_fgs = sorted(selected_fgs)
                
                # Create new order preserving sorted order
                new_order = OrderedDict()
                for i, fg in enumerate(sorted_selected_fgs):
                    new_order[fg] = i
                
                # Update session state
                st.session_state.fg_analysis_order = new_order
                st.session_state.analysis_completed = True
                
                # Reset select_all_trigger if not all are selected
                if set(selected_fgs) != set(fg_codes):
                    st.session_state.select_all_trigger = False
            else:
                # Clear if nothing selected
                st.session_state.fg_analysis_order = OrderedDict()
                st.session_state.analysis_completed = False
                st.session_state.select_all_trigger = False
            
            # Display the current FIFO order
            if st.session_state.fg_analysis_order:
                st.write("**🎯 FIFO Order (First to Last):**")
                fifo_list = ""
                for i, (fg, _) in enumerate(st.session_state.fg_analysis_order.items(), 1):
                    fifo_list += f"{i}. {fg}\n"
                st.text(fifo_list)
                
                st.write("### 📊 View Formula Details")
                sel_fg_view = st.selectbox(
                    "Select FG to view details:",
                    list(st.session_state.fg_analysis_order.keys()),
                    key="fg_view_select"
                )
                
                if sel_fg_view:
                    formula_view = st.session_state.fg_formulas[
                        st.session_state.fg_formulas['FG Code'] == sel_fg_view
                    ].copy()
                    formula_view['Quantity'] = formula_view['Quantity'].apply(lambda x: f"{x:,.4f} Kg")
                    
                    # Calculate total RM required for this FG
                    total_rm_qty = formula_view['Quantity'].str.replace(' Kg', '').str.replace(',', '').astype(float).sum()
                    
                    col_detail1, col_detail2 = st.columns(2)
                    with col_detail1:
                        st.metric(f"RM Components for {sel_fg_view}", len(formula_view))
                    with col_detail2:
                        st.metric("Total RM Quantity", f"{total_rm_qty:,.4f} Kg")
                    
                    st.dataframe(
                        formula_view,
                        use_container_width=True,
                        height=min(300, len(formula_view) * 35 + 40),
                        hide_index=True
                    )
            else:
                if fg_codes:  # Only show if there are FGs available
                    st.info("ℹ️ Select FG codes above to analyze production planning in Tab 5")
        
        else:
            st.info("📁 No FG formulas loaded yet. Please upload Excel files.")
    
    with col2:
        st.subheader("⚙️ Calculation Settings")
        
        st.write("### 🔢 Decimal Precision")
        margin = st.number_input(
            "Decimal places for calculations:",
            min_value=0,
            max_value=6,
            value=st.session_state.calculation_margin,
            help="Controls rounding precision for all calculations (0-6 decimal places)"
        )
        
        if margin != st.session_state.calculation_margin:
            st.session_state.calculation_margin = int(margin)
            st.success(f"Decimal precision set to {margin} places")
        
        st.divider()
        
        st.write("### 🗑️ Data Management")
        
        if st.button("🗑️ Clear All FG Formulas", type="secondary", help="Remove all FG formulas and reset settings"):
            st.session_state.fg_formulas = pd.DataFrame(columns=['FG Code', 'RM Code', 'Quantity'])
            st.session_state.fg_analysis_order = OrderedDict()
            st.session_state.fg_expected_capacity = {}
            st.session_state.fg_colors = {}
            st.session_state.analysis_completed = False
            st.session_state.select_all_trigger = False
            st.session_state.multiselect_key = 0
            # Also clear modified formulas
            st.session_state.modified_fg_formulas = pd.DataFrame(columns=['FG Code', 'RM Code', 'Quantity'])
            st.session_state.formulas_modified = False
            st.success("All FG formulas cleared!")
        
        st.divider()
        
        # Delete specific FG
        if not st.session_state.fg_formulas.empty:
            st.write("### 🎯 Delete Specific FG")
            fg_codes = st.session_state.fg_formulas['FG Code'].unique().tolist()
            to_delete = st.multiselect("Select FG to delete:", fg_codes, key="fg_delete_select")
            
            if st.button("🗑️ Delete Selected FG", type="primary") and to_delete:
                # Store original count
                original_count = len(st.session_state.fg_formulas)
                
                # Delete from formulas
                st.session_state.fg_formulas = st.session_state.fg_formulas[
                    ~st.session_state.fg_formulas['FG Code'].isin(to_delete)
                ]
                
                # Delete from analysis order
                for fg in to_delete:
                    if fg in st.session_state.fg_analysis_order:
                        del st.session_state.fg_analysis_order[fg]
                
                # Delete from expected capacity
                for fg in to_delete:
                    if fg in st.session_state.fg_expected_capacity:
                        del st.session_state.fg_expected_capacity[fg]
                
                # Delete from colors
                for fg in to_delete:
                    if fg in st.session_state.fg_colors:
                        del st.session_state.fg_colors[fg]
                
                # Update analysis flag
                if not st.session_state.fg_analysis_order:
                    st.session_state.analysis_completed = False
                
                # Reset select all trigger
                st.session_state.select_all_trigger = False
                st.session_state.multiselect_key += 1
                
                st.success(f"Deleted {len(to_delete)} FG code(s). Remaining: {len(st.session_state.fg_formulas)} entries")
        
        # Display system info
        st.divider()
        st.write("### 📊 System Information")
        
        info_data = {
            "Metric": ["Loaded FGs", "Selected for Analysis", "Replacement Rules", "Dilution Rules", "Modified Formulas"],
            "Value": [
                len(st.session_state.fg_formulas['FG Code'].unique()) if not st.session_state.fg_formulas.empty else 0,
                len(st.session_state.fg_analysis_order),
                len(st.session_state.rm_replacement_rules),
                len(st.session_state.rm_dilution_rules),
                "Yes" if st.session_state.formulas_modified else "No"
            ]
        }
        
        info_df = pd.DataFrame(info_data)
        st.dataframe(info_df, use_container_width=True, hide_index=True)
    
    add_footer()

# --- TAB 3: RM REPLACEMENT ---
with tab3:
    st.subheader("🔄 RM Code Replacement Rules")
    
    st.markdown("""
    **Instructions:**
    1. Upload an Excel file containing RM replacement rules
    2. Columns should include: **Old RM Code** and **New RM Code**
    3. Click 'Apply Replacement Rules' to generate temporary modified formulas
    4. Modified formulas will be used in Tab 5 (Production Planning)
    """)
    
    col_upload, col_preview = st.columns([2, 1])
    
    with col_upload:
        # File upload for replacement rules
        replacement_file = st.file_uploader(
            "Upload RM Replacement Excel File", 
            type=['xlsx'], 
            key="replacement_upload",
            help="Excel file with Old RM Code and New RM Code columns"
        )
        
        if replacement_file is not None:
            try:
                df_replace = pd.read_excel(replacement_file)
                df_replace.columns = df_replace.columns.str.strip()
                
                column_mapping = {}
                
                # Find Old RM Code column
                for col in df_replace.columns:
                    col_lower = str(col).lower()
                    if ('old' in col_lower or 'from' in col_lower) and 'rm' in col_lower:
                        column_mapping['Old RM Code'] = col
                        break
                
                # Find New RM Code column
                for col in df_replace.columns:
                    col_lower = str(col).lower()
                    if ('new' in col_lower or 'to' in col_lower) and 'rm' in col_lower:
                        column_mapping['New RM Code'] = col
                        break
                
                if all(col in column_mapping for col in ['Old RM Code', 'New RM Code']):
                    processed_replace = pd.DataFrame()
                    processed_replace['Old RM Code'] = df_replace[column_mapping['Old RM Code']].apply(preserve_8char_code)
                    processed_replace['New RM Code'] = df_replace[column_mapping['New RM Code']].apply(preserve_8char_code)
                    
                    # Remove empty rows
                    processed_replace = processed_replace[
                        (processed_replace['Old RM Code'] != '') & 
                        (processed_replace['Old RM Code'] != 'nan') &
                        (processed_replace['New RM Code'] != '') &
                        (processed_replace['New RM Code'] != 'nan')
                    ]
                    
                    if not processed_replace.empty:
                        st.session_state.rm_replacement_rules = processed_replace
                        st.success(f"✅ Successfully loaded {len(processed_replace)} replacement rules!")
                    else:
                        st.warning("No valid replacement rules found in the file")
                        
                else:
                    missing_cols = []
                    for col in ['Old RM Code', 'New RM Code']:
                        if col not in column_mapping:
                            missing_cols.append(col)
                    st.error(f"Missing columns: {', '.join(missing_cols)}")
                    
            except Exception as e:
                st.error(f"Error processing replacement file: {str(e)}")
    
    with col_preview:
        if not st.session_state.rm_replacement_rules.empty:
            st.write("### 📋 Current Replacement Rules")
            
            display_replace = st.session_state.rm_replacement_rules.copy()
            
            st.dataframe(
                display_replace,
                use_container_width=True,
                height=min(300, len(display_replace) * 35 + 40),
                hide_index=True
            )
    
    # Apply Replacement Button
    col_apply, col_clear = st.columns([1, 1])
    
    with col_apply:
        if st.button("🔄 Apply Replacement Rules", type="primary", key="apply_replacement", use_container_width=True):
            if st.session_state.rm_replacement_rules.empty:
                st.warning("⚠️ Please upload replacement rules first")
            elif st.session_state.fg_formulas.empty:
                st.warning("⚠️ Please upload FG formulas in Tab 2 first")
            else:
                # Apply replacement rules
                modified_formulas = apply_rm_replacement()
                st.session_state.modified_fg_formulas = modified_formulas
                st.session_state.formulas_modified = True
                st.session_state.dilution_applied = False  # Reset dilution flag
                
                st.success(f"✅ Replacement rules applied successfully!")
                st.info(f"**Modified {len(modified_formulas)} formula entries**")
                
                # Show comparison
                original_count = len(st.session_state.fg_formulas)
                modified_count = len(modified_formulas)
                
                col_comp1, col_comp2 = st.columns(2)
                with col_comp1:
                    st.metric("Original Formulas", original_count)
                with col_comp2:
                    st.metric("Modified Formulas", modified_count)
    
    with col_clear:
        if st.button("🗑️ Clear Replacement Rules", type="secondary", use_container_width=True):
            st.session_state.rm_replacement_rules = pd.DataFrame(columns=['Old RM Code', 'New RM Code'])
            st.success("Replacement rules cleared!")
    
    # Display status
    st.divider()
    st.write("### 📊 Current Status")
    
    col_status1, col_status2, col_status3 = st.columns(3)
    
    with col_status1:
        st.metric("Original Formulas", 
                 len(st.session_state.fg_formulas) if not st.session_state.fg_formulas.empty else 0)
    
    with col_status2:
        if st.session_state.formulas_modified:
            st.metric("Modified Formulas", 
                     len(st.session_state.modified_fg_formulas), 
                     delta="✓ Applied", delta_color="normal")
        else:
            st.metric("Modified Formulas", "Not applied", delta="Pending", delta_color="off")
    
    with col_status3:
        st.metric("Replacement Rules", len(st.session_state.rm_replacement_rules))
    
    # Show summary of what will be replaced
    if not st.session_state.rm_replacement_rules.empty and not st.session_state.fg_formulas.empty:
        st.divider()
        st.write("### 📋 Replacement Summary")
        
        # Find which RMs in formulas will be replaced
        original_rms = set(st.session_state.fg_formulas['RM Code'].apply(preserve_8char_code))
        replacement_map = {}
        for _, row in st.session_state.rm_replacement_rules.iterrows():
            replacement_map[preserve_8char_code(row['Old RM Code'])] = preserve_8char_code(row['New RM Code'])
        
        rms_to_replace = [rm for rm in original_rms if rm in replacement_map]
        
        if rms_to_replace:
            st.write(f"**RMs to be replaced:** {len(rms_to_replace)}")
            for i, rm in enumerate(rms_to_replace[:10], 1):  # Show first 10
                st.write(f"{i}. {rm} → {replacement_map[rm]}")
            if len(rms_to_replace) > 10:
                st.write(f"... and {len(rms_to_replace) - 10} more")
        else:
            st.info("No matching RMs found in current formulas for replacement")
    
    # Preview modified formulas
    if st.session_state.formulas_modified:
        st.divider()
        st.write("### 🔍 Preview Modified Formulas")
        
        with st.expander("View Modified Formulas", expanded=False):
            display_modified = st.session_state.modified_fg_formulas.copy()
            display_modified['Quantity'] = display_modified['Quantity'].apply(lambda x: f"{x:,.4f} Kg")
            
            st.dataframe(
                display_modified,
                use_container_width=True,
                height=min(300, len(display_modified) * 35 + 40),
                hide_index=True
            )
    
    add_footer()

# --- TAB 4: RM DILUTION ---
with tab4:
    st.subheader("💧 RM Dilution Rules")
    
    st.markdown("""
    **Instructions:**
    1. Upload an Excel file containing RM dilution rules
    2. Columns should include: **RM Code**, **Component RM Code**, and **Percentage**
    3. Percentage should be a number (e.g., 50 for 50%)
    4. Click 'Apply Dilution Rules' to break down diluted RMs into components
    5. **Strict Rule:** If any calculation results in 0.0000, it will be set to 0.0001
    """)
    
    col_dil_upload, col_dil_preview = st.columns([2, 1])
    
    with col_dil_upload:
        # File upload for dilution rules
        dilution_file = st.file_uploader(
            "Upload RM Dilution Excel File", 
            type=['xlsx'], 
            key="dilution_upload",
            help="Excel file with RM Code, Component RM Code, and Percentage columns"
        )
        
        if dilution_file is not None:
            try:
                df_dilution = pd.read_excel(dilution_file)
                df_dilution.columns = df_dilution.columns.str.strip()
                
                column_mapping = {}
                
                # Find RM Code column
                for col in df_dilution.columns:
                    col_lower = str(col).lower()
                    if 'rm' in col_lower and ('code' in col_lower or 'id' in col_lower):
                        if 'component' not in col_lower:
                            column_mapping['RM Code'] = col
                            break
                
                # Find Component RM Code column
                for col in df_dilution.columns:
                    col_lower = str(col).lower()
                    if 'component' in col_lower and 'rm' in col_lower:
                        column_mapping['Component RM Code'] = col
                        break
                
                # Find Percentage column
                for col in df_dilution.columns:
                    col_lower = str(col).lower()
                    if 'percentage' in col_lower or 'percent' in col_lower or '%' in col_lower:
                        column_mapping['Percentage'] = col
                        break
                
                if all(col in column_mapping for col in ['RM Code', 'Component RM Code', 'Percentage']):
                    processed_dilution = pd.DataFrame()
                    processed_dilution['RM Code'] = df_dilution[column_mapping['RM Code']].apply(preserve_8char_code)
                    processed_dilution['Component RM Code'] = df_dilution[column_mapping['Component RM Code']].apply(preserve_8char_code)
                    
                    # Convert percentage to numeric
                    processed_dilution['Percentage'] = pd.to_numeric(
                        df_dilution[column_mapping['Percentage']], 
                        errors='coerce'
                    ).fillna(0)
                    
                    # Remove empty rows
                    processed_dilution = processed_dilution[
                        (processed_dilution['RM Code'] != '') & 
                        (processed_dilution['RM Code'] != 'nan') &
                        (processed_dilution['Component RM Code'] != '') &
                        (processed_dilution['Component RM Code'] != 'nan') &
                        (processed_dilution['Percentage'] > 0)
                    ]
                    
                    if not processed_dilution.empty:
                        st.session_state.rm_dilution_rules = processed_dilution
                        st.success(f"✅ Successfully loaded {len(processed_dilution)} dilution rules!")
                        
                        # Check if percentages sum to 100% for each RM
                        percentage_summary = processed_dilution.groupby('RM Code')['Percentage'].sum().reset_index()
                        invalid_rms = percentage_summary[abs(percentage_summary['Percentage'] - 100) > 0.01]
                        
                        if not invalid_rms.empty:
                            st.warning(f"⚠️ Some RMs don't sum to 100%: {', '.join(invalid_rms['RM Code'].tolist())}")
                    else:
                        st.warning("No valid dilution rules found in the file")
                        
                else:
                    missing_cols = []
                    for col in ['RM Code', 'Component RM Code', 'Percentage']:
                        if col not in column_mapping:
                            missing_cols.append(col)
                    st.error(f"Missing columns: {', '.join(missing_cols)}")
                    
            except Exception as e:
                st.error(f"Error processing dilution file: {str(e)}")
    
    with col_dil_preview:
        if not st.session_state.rm_dilution_rules.empty:
            st.write("### 📋 Current Dilution Rules")
            
            display_dilution = st.session_state.rm_dilution_rules.copy()
            display_dilution['Percentage'] = display_dilution['Percentage'].apply(lambda x: f"{x:.2f}%")
            
            st.dataframe(
                display_dilution,
                use_container_width=True,
                height=min(300, len(display_dilution) * 35 + 40),
                hide_index=True
            )
    
    # Apply Dilution Button
    col_dil_apply, col_dil_clear = st.columns([1, 1])
    
    with col_dil_apply:
        if st.button("💧 Apply Dilution Rules", type="primary", key="apply_dilution", use_container_width=True):
            if st.session_state.rm_dilution_rules.empty:
                st.warning("⚠️ Please upload dilution rules first")
            else:
                # Determine which formulas to use
                if st.session_state.formulas_modified:
                    formulas_to_dilute = st.session_state.modified_fg_formulas
                    source_name = "modified formulas"
                elif not st.session_state.fg_formulas.empty:
                    formulas_to_dilute = st.session_state.fg_formulas
                    source_name = "original formulas"
                else:
                    st.warning("⚠️ Please upload FG formulas in Tab 2 first")
                    formulas_to_dilute = pd.DataFrame()
                    source_name = ""
                
                if not formulas_to_dilute.empty:
                    # Apply dilution rules
                    diluted_formulas = apply_dilution_rules(formulas_to_dilute)
                    
                    # Store the diluted formulas (they will replace the modified formulas)
                    st.session_state.modified_fg_formulas = diluted_formulas
                    st.session_state.dilution_applied = True
                    
                    st.success(f"✅ Dilution rules applied successfully to {source_name}!")
                    
                    # Count how many RMs were diluted
                    original_rms = set(formulas_to_dilute['RM Code'].apply(preserve_8char_code))
                    diluted_rms = [rm for rm in original_rms if rm in set(st.session_state.rm_dilution_rules['RM Code'])]
                    
                    if diluted_rms:
                        st.info(f"**Diluted {len(diluted_rms)} RM(s):** {', '.join(diluted_rms[:5])}{'...' if len(diluted_rms) > 5 else ''}")
                    else:
                        st.info("No RMs were diluted (no matching RM codes found)")
                    
                    # Show example of 0.0000 → 0.0001 conversion if applicable
                    example_conversions = []
                    for _, row in diluted_formulas.iterrows():
                        if 0.0000 <= row['Quantity'] < 0.00005:
                            example_conversions.append(f"{row['FG Code']} - {row['RM Code']}: {row['Quantity']:.6f} → 0.0001")
                    
                    if example_conversions:
                        st.write("**0.0000 → 0.0001 Conversions:**")
                        for conv in example_conversions[:3]:  # Show first 3 examples
                            st.write(f"• {conv}")
                        if len(example_conversions) > 3:
                            st.write(f"... and {len(example_conversions) - 3} more")
    
    with col_dil_clear:
        if st.button("🗑️ Clear Dilution Rules", type="secondary", use_container_width=True):
            st.session_state.rm_dilution_rules = pd.DataFrame(columns=['RM Code', 'Component RM Code', 'Percentage'])
            st.success("Dilution rules cleared!")
    
    # Display status
    st.divider()
    st.write("### 📊 Current Status")
    
    col_dil_status1, col_dil_status2, col_dil_status3 = st.columns(3)
    
    with col_dil_status1:
        if st.session_state.formulas_modified:
            base_count = len(st.session_state.modified_fg_formulas)
        else:
            base_count = len(st.session_state.fg_formulas)
        st.metric("Base Formulas", base_count)
    
    with col_dil_status2:
        if st.session_state.dilution_applied:
            st.metric("Diluted Formulas", 
                     len(st.session_state.modified_fg_formulas), 
                     delta="✓ Applied", delta_color="normal")
        else:
            st.metric("Diluted Formulas", "Not applied", delta="Pending", delta_color="off")
    
    with col_dil_status3:
        st.metric("Dilution Rules", len(st.session_state.rm_dilution_rules))
    
    # Show which RMs can be diluted
    if not st.session_state.rm_dilution_rules.empty:
        st.divider()
        st.write("### 📋 Available Dilutions")
        
        # Group by RM Code
        dilution_summary = st.session_state.rm_dilution_rules.groupby('RM Code').agg({
            'Component RM Code': lambda x: ', '.join(x),
            'Percentage': lambda x: ', '.join([f"{p:.2f}%" for p in x])
        }).reset_index()
        
        dilution_summary.columns = ['RM Code', 'Components', 'Percentages']
        
        st.dataframe(
            dilution_summary,
            use_container_width=True,
            height=min(300, len(dilution_summary) * 35 + 40),
            hide_index=True
        )
    
    # Preview diluted formulas
    if st.session_state.dilution_applied:
        st.divider()
        st.write("### 🔍 Preview Diluted Formulas")
        
        with st.expander("View Diluted Formulas", expanded=False):
            display_diluted = st.session_state.modified_fg_formulas.copy()
            display_diluted['Quantity'] = display_diluted['Quantity'].apply(lambda x: f"{x:,.4f} Kg")
            
            st.dataframe(
                display_diluted,
                use_container_width=True,
                height=min(300, len(display_diluted) * 35 + 40),
                hide_index=True
            )
    
    add_footer()

# --- TAB 5: PRODUCTION PLANNING ---
with tab5:
    col_date, col_export = st.columns([2, 1])
    
    with col_date:
        st.subheader("📊 Production Planning Summary")
        prod_date = st.date_input(
            "Select Planned Production Date", 
            datetime.now(), 
            key="prod_date",
            help="Select the date for which production planning is being done"
        )
    
    with col_export:
        st.write("")
        st.write("")
        if st.button("📈 Generate Production Analysis", type="primary", key="generate_analysis", use_container_width=True):
            if st.session_state.fg_analysis_order:
                st.session_state.analysis_completed = True
                st.success("✅ Production analysis generated! Scroll down to see results.")
            else:
                st.warning("⚠️ Please select FG codes in Tab 2 first.")
    
    # Check if required data is ready
    data_ready = True
    warning_messages = []
    
    if st.session_state.rm_stock.empty:
        warning_messages.append("RM Stock")
        data_ready = False
    
    if st.session_state.fg_formulas.empty:
        warning_messages.append("FG Formulas")
        data_ready = False
    
    if not st.session_state.fg_analysis_order:
        warning_messages.append("FG selection for analysis")
        data_ready = False
    
    if not data_ready:
        st.warning(f"⚠️ Please complete the following in previous tabs: {', '.join(warning_messages)}")
    else:
        # Show analysis if completed
        if st.session_state.analysis_completed:
            decimal_places = st.session_state.calculation_margin
            
            stock_dict = st.session_state.rm_stock.set_index('RM Code')['Quantity'].to_dict()
            stock_dict = {k: round(float(v), decimal_places) for k, v in stock_dict.items()}
            
            # Create copies for calculation
            initial_stock = stock_dict.copy()
            allocated_stock = stock_dict.copy()
            
            # Determine which formulas to use
            if not st.session_state.modified_fg_formulas.empty:
                formulas_to_use = st.session_state.modified_fg_formulas
                formula_source = "Modified Formulas"
            else:
                formulas_to_use = st.session_state.fg_formulas
                formula_source = "Original Formulas"
            
            results = []
            shortage_details = {}
            
            # Process each FG in FIFO order
            for fg in st.session_state.fg_analysis_order.keys():
                formula = formulas_to_use[formulas_to_use['FG Code'] == fg]
                
                if formula.empty:
                    continue
                
                expected_capacity = st.session_state.fg_expected_capacity.get(fg, 0)
                
                # Calculate MAX capacity first (using initial stock, not allocated stock)
                max_possible_batches_list = []
                max_shortage_breakdown = []
                
                for _, row in formula.iterrows():
                    rm = preserve_8char_code(row['RM Code'])
                    req_per_batch = round(float(row['Quantity']), decimal_places)
                    avail = initial_stock.get(rm, 0)  # Use initial stock for max calculation
                    
                    if req_per_batch <= 0 or avail <= 0:
                        max_possible_batches_list.append(0)
                        if avail <= 0:
                            max_shortage_breakdown.append(f"{rm}: Available 0.0000 Kg")
                    else:
                        max_batches_for_rm = int(avail // req_per_batch)
                        max_possible_batches_list.append(max_batches_for_rm)
                
                max_possible_batches = min(max_possible_batches_list) if max_possible_batches_list else 0
                max_capacity = max_possible_batches * 25
                
                # Now calculate ACTUAL capacity based on expected and allocated stock
                possible_batches = []
                missing_rms = []
                shortage_breakdown = []
                
                # Calculate based on expected capacity
                if expected_capacity > 0:
                    # Calculate required batches based on expected capacity
                    expected_batches = max(1, int(expected_capacity // 25))
                    actual_expected_capacity = expected_batches * 25
                    
                    for _, row in formula.iterrows():
                        rm = preserve_8char_code(row['RM Code'])
                        req_per_batch = round(float(row['Quantity']), decimal_places)
                        avail = allocated_stock.get(rm, 0)
                        
                        # Calculate total required for expected batches
                        total_required = req_per_batch * expected_batches
                        
                        if total_required <= 0:
                            shortage_breakdown.append(f"{rm}: Invalid requirement ({req_per_batch:.{decimal_places}f} Kg per batch)")
                            possible_batches.append(0)
                        elif avail <= 0:
                            possible_batches.append(0)
                            missing_rms.append(rm)
                            shortage_breakdown.append(f"{rm}: Required {total_required:.{decimal_places}f} Kg, Available 0.0000 Kg")
                        else:
                            # Check if we have enough for expected batches
                            if avail >= total_required:
                                max_batches_for_rm = expected_batches
                                possible_batches.append(max_batches_for_rm)
                            else:
                                # Calculate how many batches we can make
                                max_batches_for_rm = int(avail // req_per_batch)
                                possible_batches.append(max_batches_for_rm)
                                
                                if max_batches_for_rm < expected_batches:
                                    missing_rms.append(rm)
                                    shortage = total_required - avail
                                    shortage_breakdown.append(f"{rm}: Required {total_required:.{decimal_places}f} Kg for {expected_batches} batches, Available {avail:.{decimal_places}f} Kg, Shortage {shortage:.{decimal_places}f} Kg")
                else:
                    # If no expected capacity, calculate maximum possible
                    for _, row in formula.iterrows():
                        rm = preserve_8char_code(row['RM Code'])
                        req_per_batch = round(float(row['Quantity']), decimal_places)
                        avail = allocated_stock.get(rm, 0)
                        
                        if req_per_batch <= 0:
                            possible_batches.append(0)
                            shortage_breakdown.append(f"{rm}: Invalid requirement ({req_per_batch:.{decimal_places}f} Kg)")
                        elif avail <= 0:
                            possible_batches.append(0)
                            missing_rms.append(rm)
                            shortage_breakdown.append(f"{rm}: Required {req_per_batch:.{decimal_places}f} Kg per batch, Available 0.0000 Kg")
                        else:
                            max_batches_for_rm = int(avail // req_per_batch)
                            possible_batches.append(max_batches_for_rm)
                            
                            if max_batches_for_rm == 0:
                                missing_rms.append(rm)
                                shortage = req_per_batch - avail
                                shortage_breakdown.append(f"{rm}: Required {req_per_batch:.{decimal_places}f} Kg per batch, Available {avail:.{decimal_places}f} Kg, Shortage {shortage:.{decimal_places}f} Kg")
                
                # Determine actual batches and capacity
                if expected_capacity > 0:
                    # For expected capacity mode
                    max_possible_for_actual = min(possible_batches) if possible_batches else 0
                    actual_batches = min(expected_batches, max_possible_for_actual)
                    actual_capacity = actual_batches * 25
                else:
                    # For auto mode
                    max_possible_for_actual = min(possible_batches) if possible_batches else 0
                    actual_batches = max_possible_for_actual
                    actual_capacity = max_possible_for_actual * 25
                
                # Determine status
                if actual_capacity >= 25:
                    status = "✅ Ready"
                else:
                    status = "❌ Shortage"
                
                shortage_details[fg] = shortage_breakdown
                
                if missing_rms:
                    missing_display = f"{len(missing_rms)} RM(s)"
                else:
                    missing_display = "None"
                
                # Allocate stock for production
                if actual_batches > 0 and status == "✅ Ready":
                    for _, row in formula.iterrows():
                        rm = preserve_8char_code(row['RM Code'])
                        req_total = round(row['Quantity'] * actual_batches, decimal_places)
                        if rm in allocated_stock:
                            allocated_stock[rm] = round(allocated_stock[rm] - req_total, decimal_places)
                
                # Add to results
                results.append({
                    "FG": fg,
                    "Expected": f"{expected_capacity:,.1f} Kg" if expected_capacity > 0 else "Auto",
                    "Max": f"{max_capacity:,.1f} Kg",
                    "Actual": f"{actual_capacity:,.1f} Kg",
                    "Status": status,
                    "Missing": missing_display,
                    "Batches": actual_batches
                })
            
            # Prepare PO status for report
            if not st.session_state.rm_po.empty:
                po_status_for_report = st.session_state.rm_po.copy()
                po_status_for_report['Status'] = po_status_for_report['Arrival Date'].apply(
                    lambda x: "Delayed" if x.date() < prod_date else "Incoming"
                )
                delayed_pos = len(po_status_for_report[po_status_for_report['Status'] == "Delayed"])
            else:
                po_status_for_report = None
                delayed_pos = 0
            
            # Calculate totals
            ready_fgs = [r for r in results if "✅" in r['Status']]
            total_volume = sum(float(r['Actual'].replace(' Kg', '').replace(',', '')) 
                              for r in results if r['Actual'] != "0.0 Kg")
            
            # Display formula source info
            st.info(f"**Using:** {formula_source}")
            if st.session_state.formulas_modified:
                st.info("**Note:** Formulas have been modified by replacement/dilution rules")
            
            # Expected capacity settings
            st.write("### 🎯 Set Expected Capacities")
            capacity_col1, capacity_col2 = st.columns(2)
            
            fg_list = list(st.session_state.fg_analysis_order.keys())
            mid_point = (len(fg_list) + 1) // 2
            
            with capacity_col1:
                for fg in fg_list[:mid_point]:
                    current_val = st.session_state.fg_expected_capacity.get(fg, 0)
                    new_capacity = st.number_input(
                        f"{fg} (Kg, min 25):",
                        min_value=0.0,
                        value=float(current_val),
                        step=25.0,
                        format="%.1f",
                        key=f"exp_cap_{fg}"
                    )
                    if new_capacity != current_val:
                        st.session_state.fg_expected_capacity[fg] = new_capacity
            
            with capacity_col2:
                for fg in fg_list[mid_point:]:
                    current_val = st.session_state.fg_expected_capacity.get(fg, 0)
                    new_capacity = st.number_input(
                        f"{fg} (Kg, min 25):",
                        min_value=0.0,
                        value=float(current_val),
                        step=25.0,
                        format="%.1f",
                        key=f"exp_cap_{fg}_2"
                    )
                    if new_capacity != current_val:
                        st.session_state.fg_expected_capacity[fg] = new_capacity
            
            # Production summary
            st.divider()
            st.write("### 📊 Production Summary")
            
            col1, col2, col3, col4 = st.columns(4)
            col1.metric("Producible FG", len(ready_fgs))
            col2.metric("Total Volume", f"{total_volume:,.1f} Kg")
            col3.metric("Total Batches", sum(r['Batches'] for r in results))
            col4.metric("Delayed POs", delayed_pos)
            
            # Production capability list
            st.divider()
            st.write("### 📋 Production Capability List")
            
            res_df = pd.DataFrame(results)
            
            st.dataframe(
                res_df,
                use_container_width=True,
                height=min(400, len(res_df) * 35 + 40),
                hide_index=True,
                column_config={
                    "FG": st.column_config.TextColumn("FG Code", width="small"),
                    "Expected": st.column_config.TextColumn("Expected", width="small"),
                    "Max": st.column_config.TextColumn("Max Cap", width="small"),
                    "Actual": st.column_config.TextColumn("Actual Cap", width="small"),
                    "Status": st.column_config.TextColumn("Status", width="small"),
                    "Missing": st.column_config.TextColumn("Missing RM", width="small"),
                    "Batches": st.column_config.NumberColumn("Batches", width="small")
                }
            )
            
            # Shortage details
            st.divider()
            st.write("### 🔍 Shortage Details")
            
            shortage_exists = False
            for fg in shortage_details:
                if shortage_details[fg]:
                    shortage_exists = True
                    with st.expander(f"❌ {fg} - RM Shortage Breakdown", expanded=False):
                        for item in shortage_details[fg]:
                            st.write(f"• {item}")
            
            if not shortage_exists:
                st.info("✅ No shortages detected for selected FGs")
            
            # PO Delay Tracker
            if not st.session_state.rm_po.empty:
                st.divider()
                st.write("### ⏰ PO Delay Tracker")
                
                po_display = po_status_for_report.copy()
                po_display['Quantity'] = po_display['Quantity'].apply(lambda x: f"{x:,.4f} Kg")
                po_display['Arrival Date'] = po_display['Arrival Date'].dt.strftime('%d/%m/%Y')
                
                st.dataframe(
                    po_display,
                    use_container_width=True,
                    height=min(300, len(po_display) * 35 + 40),
                    hide_index=True,
                    column_config={
                        "RM Code": st.column_config.TextColumn("RM Code", width="small"),
                        "Quantity": st.column_config.TextColumn("Quantity", width="small"),
                        "Arrival Date": st.column_config.TextColumn("Arrival", width="small"),
                        "Status": st.column_config.TextColumn("Status", width="small")
                    }
                )
            
            # Visualizations
            st.divider()
            st.write("### 📈 Production Capacity Visualization")
            
            chart_col1, chart_col2 = st.columns(2)
            
            with chart_col1:
                if len(results) > 0:
                    res_df = pd.DataFrame(results)
                    chart_data = res_df.copy()
                    chart_data['Actual_Num'] = chart_data['Actual'].str.replace(' Kg', '').str.replace(',', '').astype(float)
                    chart_data['Max_Num'] = chart_data['Max'].str.replace(' Kg', '').str.replace(',', '').astype(float)
                    
                    # Create stacked bar chart for Actual vs Max
                    fig1_data = []
                    for _, row in chart_data.iterrows():
                        fig1_data.append({'FG': row['FG'], 'Capacity': row['Actual_Num'], 'Type': 'Actual'})
                        fig1_data.append({'FG': row['FG'], 'Capacity': row['Max_Num'] - row['Actual_Num'], 'Type': 'Available'})
                    
                    fig1_df = pd.DataFrame(fig1_data)
                    
                    fig1 = px.bar(
                        fig1_df,
                        x='FG',
                        y='Capacity',
                        title="Actual vs Maximum Capacity",
                        color='Type',
                        color_discrete_map={'Actual': '#2ca02c', 'Available': '#aec7e8'},
                        text=fig1_df['Capacity'].apply(lambda x: f"{x:,.1f}" if x > 0 else ''),
                        hover_data=['Type']
                    )
                    
                    fig1.update_traces(
                        texttemplate='%{text}',
                        textposition='outside',
                        hovertemplate='<b>%{x}</b><br>' +
                                    'Type: %{customdata[0]}<br>' +
                                    'Capacity: %{y:,.1f} Kg<br>' +
                                    '<extra></extra>'
                    )
                    
                    fig1.update_layout(
                        yaxis_title="Capacity (Kg)",
                        showlegend=True,
                        height=400,
                        xaxis_tickangle=-45,
                        legend_title="Capacity Type",
                        legend=dict(
                            orientation="h",
                            yanchor="bottom",
                            y=1.02,
                            xanchor="right",
                            x=1
                        )
                    )
                    st.plotly_chart(fig1, use_container_width=True)
            
            with chart_col2:
                if len(res_df) > 1:
                    pie_data = res_df.copy()
                    pie_data['Actual_Num'] = pie_data['Actual'].str.replace(' Kg', '').str.replace(',', '').astype(float)
                    pie_data = pie_data[pie_data['Actual_Num'] > 0]
                    
                    if len(pie_data) > 0:
                        pie_colors = [get_fg_color(fg) for fg in pie_data['FG']]
                        
                        fig2 = px.pie(
                            pie_data,
                            values='Actual_Num',
                            names='FG',
                            title="Capacity Distribution",
                            color='FG',
                            color_discrete_sequence=pie_colors,
                            hole=0.3,
                            hover_data=['Status']
                        )
                        
                        fig2.update_traces(
                            hovertemplate='<b>%{label}</b><br>' +
                                        'Capacity: %{value:,.1f} Kg<br>' +
                                        'Percentage: %{percent}<br>' +
                                        '<extra></extra>'
                        )
                        
                        fig2.update_layout(
                            height=400,
                            legend_title="FG Code",
                            showlegend=True
                        )
                        st.plotly_chart(fig2, use_container_width=True)
            
            # Settings info
            st.info(
                f"**Current Settings:**\n"
                f"• Decimal Precision: {st.session_state.calculation_margin} places\n"
                f"• FIFO Order: {', '.join(st.session_state.fg_analysis_order.keys())}\n"
                f"• Formula Source: {formula_source}"
            )
            
            # --- EXPORT REPORTS SECTION ---
            st.divider()
            st.write("### 📤 Export Reports")
            
            # Create 3 columns for export buttons
            export_col1, export_col2, export_col3 = st.columns(3)
            
            with export_col1:
                # PDF/HTML Report Button
                try:
                    report_data, report_type = generate_report(
                        results, 
                        shortage_details, 
                        prod_date, 
                        total_volume, 
                        ready_fgs, 
                        delayed_pos,
                        po_status_for_report
                    )
                    
                    if report_type == "pdf":
                        mime_type = "application/pdf"
                        file_name = f"MRP_Production_Report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf"
                        btn_label = "⬇️ PDF Report"
                    else:
                        mime_type = "text/html"
                        file_name = f"MRP_Production_Report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.html"
                        btn_label = "⬇️ HTML Report"
                    
                    st.download_button(
                        label=btn_label,
                        data=report_data,
                        file_name=file_name,
                        mime=mime_type,
                        key="pdf_html_download",
                        help="Download report in PDF format (or HTML if PDF generation fails)"
                    )
                    
                except Exception as e:
                    st.error(f"Error generating report: {str(e)}")
            
            with export_col2:
                # Excel Export with Detailed RM Analysis
                if not results:
                    st.info("No production data")
                else:
                    try:
                        # Create Excel with detailed RM analysis
                        output = io.BytesIO()
                        with pd.ExcelWriter(output, engine='openpyxl') as writer:
                            # Sheet 1: Production Results
                            results_df = pd.DataFrame(results)
                            results_df.to_excel(writer, sheet_name='Production Results', index=False)
                            
                            # Sheet 2: Detailed RM Analysis (FG Code, RM Code, Required, Available, Shortage)
                            rm_analysis_data = []
                            
                            # Calculate actual batches for each FG
                            fg_batches = {}
                            for item in results:
                                fg_batches[item['FG']] = item['Batches']
                            
                            # Process each FG that has actual production
                            for fg_code, batches in fg_batches.items():
                                if batches > 0:
                                    # Get formula for this FG
                                    formula = formulas_to_use[formulas_to_use['FG Code'] == fg_code]
                                    
                                    for _, row in formula.iterrows():
                                        rm_code = preserve_8char_code(row['RM Code'])
                                        req_per_batch = round(float(row['Quantity']), decimal_places)
                                        total_required = req_per_batch * batches
                                        available = allocated_stock.get(rm_code, 0) + (initial_stock.get(rm_code, 0) - allocated_stock.get(rm_code, 0))
                                        shortage = max(0, total_required - available)
                                        
                                        rm_analysis_data.append({
                                            'FG Code': fg_code,
                                            'RM Code': rm_code,
                                            'Required (Kg)': total_required,
                                            'Available (Kg)': available,
                                            'Shortage (Kg)': shortage,
                                            'Batches': batches,
                                            'Req per Batch (Kg)': req_per_batch
                                        })
                            
                            # Also include FGs with shortages (batches = 0)
                            for fg_code in shortage_details:
                                if shortage_details[fg_code]:  # Has shortages
                                    formula = formulas_to_use[formulas_to_use['FG Code'] == fg_code]
                                    
                                    # Find which RM caused shortage
                                    shortage_rms = set()
                                    for shortage_item in shortage_details[fg_code]:
                                        # Parse RM code from shortage details
                                        match = re.search(r'([A-Z0-9]+):', shortage_item)
                                        if match:
                                            shortage_rms.add(preserve_8char_code(match.group(1)))
                                    
                                    for _, row in formula.iterrows():
                                        rm_code = preserve_8char_code(row['RM Code'])
                                        req_per_batch = round(float(row['Quantity']), decimal_places)
                                        
                                        # For shortage FGs, calculate what would be needed for at least 1 batch
                                        min_batches = 1
                                        total_required = req_per_batch * min_batches
                                        available = initial_stock.get(rm_code, 0)
                                        shortage = max(0, total_required - available)
                                        
                                        # Only add if this RM is causing shortage or if no specific shortage RMs identified
                                        if not shortage_rms or rm_code in shortage_rms:
                                            rm_analysis_data.append({
                                                'FG Code': fg_code,
                                                'RM Code': rm_code,
                                                'Required (Kg)': total_required,
                                                'Available (Kg)': available,
                                                'Shortage (Kg)': shortage,
                                                'Batches': 0,
                                                'Req per Batch (Kg)': req_per_batch,
                                                'Status': 'Shortage'
                                            })
                            
                            if rm_analysis_data:
                                rm_analysis_df = pd.DataFrame(rm_analysis_data)
                                # Reorder columns
                                rm_analysis_df = rm_analysis_df[['FG Code', 'RM Code', 'Req per Batch (Kg)', 'Batches', 
                                                                'Required (Kg)', 'Available (Kg)', 'Shortage (Kg)', 
                                                                'Status']]
                                rm_analysis_df.to_excel(writer, sheet_name='RM Analysis', index=False)
                            
                            # Sheet 3: Shortage Details (if any)
                            if shortage_details:
                                shortage_data = []
                                for fg_code, items in shortage_details.items():
                                    if items:
                                        for item in items:
                                            shortage_data.append({
                                                'FG Code': fg_code,
                                                'Shortage Details': item
                                            })
                                
                                if shortage_data:
                                    shortage_df = pd.DataFrame(shortage_data)
                                    shortage_df.to_excel(writer, sheet_name='Shortage Details', index=False)
                            
                            # Sheet 4: Settings
                            settings_df = pd.DataFrame({
                                'Setting': ['Production Date', 'Decimal Precision', 'FIFO Order', 'Formula Source', 'Report Generated'],
                                'Value': [
                                    prod_date.strftime('%d/%m/%Y'),
                                    f"{st.session_state.calculation_margin} places",
                                    ', '.join(st.session_state.fg_analysis_order.keys()) if st.session_state.fg_analysis_order else 'Not set',
                                    formula_source,
                                    datetime.now().strftime('%Y-%m-d %H:%M:%S')
                                ]
                            })
                            settings_df.to_excel(writer, sheet_name='Settings', index=False)
                            
                            # Sheet 5: Production Summary
                            summary_data = pd.DataFrame({
                                'Metric': ['Producible FG Types', 'Total Production Volume', 'Total Batches', 'Delayed POs'],
                                'Value': [len(ready_fgs), f"{total_volume:,.1f} Kg", sum(r['Batches'] for r in results), delayed_pos]
                            })
                            summary_data.to_excel(writer, sheet_name='Production Summary', index=False)
                        
                        output.seek(0)
                        
                        st.download_button(
                            label="📊 Detailed Excel Report",
                            data=output,
                            file_name=f"MRP_Detailed_Report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            key="detailed_excel_download",
                            help="Includes RM Analysis sheet with FG Code, RM Code, Required, Available, Shortage"
                        )
                        
                    except Exception as e:
                        st.error(f"Error generating Excel file: {str(e)}")
            
            with export_col3:
                # Simple Excel Export (Basic version)
                if not results:
                    st.info("No production data")
                else:
                    try:
                        # Create a simple Excel with basic data
                        output = io.BytesIO()
                        with pd.ExcelWriter(output, engine='openpyxl') as writer:
                            # Production results
                            results_df = pd.DataFrame(results)
                            results_df.to_excel(writer, sheet_name='Production Results', index=False)
                            
                            # Settings
                            settings_df = pd.DataFrame({
                                'Setting': ['Production Date', 'Decimal Precision', 'FIFO Order', 'Report Generated'],
                                'Value': [
                                    prod_date.strftime('%d/%m/%Y'),
                                    f"{st.session_state.calculation_margin} places",
                                    ', '.join(st.session_state.fg_analysis_order.keys()) if st.session_state.fg_analysis_order else 'Not set',
                                    datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                                ]
                            })
                            settings_df.to_excel(writer, sheet_name='Settings', index=False)
                        
                        output.seek(0)
                        
                        st.download_button(
                            label="📈 Basic Excel Report",
                            data=output,
                            file_name=f"MRP_Basic_Report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            key="basic_excel_download",
                            help="Basic report with production results and settings"
                        )
                        
                    except Exception as e:
                        st.error(f"Error generating Excel file: {str(e)}")
            
            # Display RM Analysis data preview
            if results:
                st.divider()
                st.write("### 📋 RM Analysis Data Preview")
                
                # Create preview of RM Analysis data
                rm_analysis_preview = []
                
                # Calculate actual batches for each FG
                fg_batches_preview = {}
                for item in results:
                    fg_batches_preview[item['FG']] = item['Batches']
                
                # Process FGs with production
                for fg_code, batches in fg_batches_preview.items():
                    if batches > 0:
                        formula = formulas_to_use[formulas_to_use['FG Code'] == fg_code]
                        
                        for _, row in formula.iterrows():
                            rm_code = preserve_8char_code(row['RM Code'])
                            req_per_batch = round(float(row['Quantity']), decimal_places)
                            total_required = req_per_batch * batches
                            available = allocated_stock.get(rm_code, 0) + (initial_stock.get(rm_code, 0) - allocated_stock.get(rm_code, 0))
                            shortage = max(0, total_required - available)
                            
                            rm_analysis_preview.append({
                                'FG Code': fg_code,
                                'RM Code': rm_code,
                                'Required': f"{total_required:.4f} Kg",
                                'Available': f"{available:.4f} Kg",
                                'Shortage': f"{shortage:.4f} Kg",
                                'Batches': batches
                            })
                
                if rm_analysis_preview:
                    preview_df = pd.DataFrame(rm_analysis_preview)
                    
                    st.dataframe(
                        preview_df,
                        use_container_width=True,
                        height=min(300, len(preview_df) * 35 + 40),
                        hide_index=True,
                        column_config={
                            "FG Code": st.column_config.TextColumn("FG Code", width="small"),
                            "RM Code": st.column_config.TextColumn("RM Code", width="small"),
                            "Required": st.column_config.TextColumn("Required", width="small"),
                            "Available": st.column_config.TextColumn("Available", width="small"),
                            "Shortage": st.column_config.TextColumn("Shortage", width="small"),
                            "Batches": st.column_config.NumberColumn("Batches", width="small")
                        }
                    )
                    
                    st.caption("This data will be exported in the 'RM Analysis' sheet of the Detailed Excel Report")
                else:
                    st.info("No RM analysis data available for preview")
            
            # Display shortage details for export
            if shortage_details:
                shortage_exists = False
                for fg in shortage_details:
                    if shortage_details[fg]:
                        shortage_exists = True
                        break
                
                if shortage_exists:
                    st.divider()
                    st.write("### 📋 Shortage Details for Export")
                    
                    # Display in expanders
                    for fg in shortage_details:
                        if shortage_details[fg]:
                            with st.expander(f"📋 {fg} - Shortage Details", expanded=False):
                                for item in shortage_details[fg]:
                                    st.write(f"• {item}")
                    
                    st.caption("These shortage details will be exported in the 'Shortage Details' sheet")
        else:
            st.info("👆 Click 'Generate Production Analysis' button above to see production planning results.")
    
    add_footer()

# Sidebar information
with st.sidebar:
    st.image("https://cdn-icons-png.flaticon.com/512/3067/3067256.png", width=100)
    st.title("RGI MRP System")
    st.markdown("---")
    
    st.markdown("### 📋 System Status")
    
    # Status indicators
    status_data = {
        "Component": ["RM Stock", "PO Data", "FG Formulas", "FG Selected", "Replacement Rules", "Dilution Rules", "Analysis Ready"],
        "Status": [
            "✅" if not st.session_state.rm_stock.empty else "❌",
            "✅" if not st.session_state.rm_po.empty else "❌",
            "✅" if not st.session_state.fg_formulas.empty else "❌",
            "✅" if st.session_state.fg_analysis_order else "❌",
            "✅" if not st.session_state.rm_replacement_rules.empty else "❌",
            "✅" if not st.session_state.rm_dilution_rules.empty else "❌",
            "✅" if st.session_state.analysis_completed else "❌"
        ],
        "Count": [
            len(st.session_state.rm_stock),
            len(st.session_state.rm_po),
            len(st.session_state.fg_formulas['FG Code'].unique()) if not st.session_state.fg_formulas.empty else 0,
            len(st.session_state.fg_analysis_order),
            len(st.session_state.rm_replacement_rules),
            len(st.session_state.rm_dilution_rules),
            "Yes" if st.session_state.analysis_completed else "No"
        ]
    }
    
    status_df = pd.DataFrame(status_data)
    st.dataframe(status_df, use_container_width=True, hide_index=True)
    
    st.markdown("---")
    
    st.markdown("### 🚀 Quick Actions")
    
    if st.button("🔄 Reset All Data", type="secondary", use_container_width=True):
        for key in list(st.session_state.keys()):
            del st.session_state[key]
        st.rerun()
    
    if st.button("📊 Go to Production Planning", type="primary", use_container_width=True):
        st.switch_page("?tab=5")
    
    st.markdown("---")
    
    st.markdown("### 📚 How to Use")
    
    with st.expander("Workflow Guide"):
        st.markdown("""
        1. **Tab 1**: Upload RM Stock & PO data
        2. **Tab 2**: Upload FG Formulas & select FGs
        3. **Tab 3**: Upload RM replacement rules → Apply them
        4. **Tab 4**: Upload RM dilution rules → Apply them
        5. **Tab 5**: Generate production planning with reports
        """)
    
    st.markdown("---")
    st.markdown("### 📅 Version Info")
    st.markdown("**MRP Dashboard v2.0**")
    st.markdown(f"Last Updated: {datetime.now().strftime('%Y-%m-%d')}")
    st.markdown("© RGI Supply Chain Department")

# Footer for entire app
st.markdown("---")
st.markdown(
    """
    <div style="text-align: center; color: #666; font-size: 12px; padding: 10px;">
        <strong>RGI MRP System</strong> | Version 2.0 | © 2024 Supply Chain Department
    </div>
    """,
    unsafe_allow_html=True
)