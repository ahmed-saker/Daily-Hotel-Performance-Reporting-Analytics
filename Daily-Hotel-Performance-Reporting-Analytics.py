import pandas as pd
import os
import time
import win32com.client
import codecs
import logging
from datetime import datetime, timedelta
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import matplotlib.pyplot as plt
import base64
import io
import numpy as np
import warnings
import sys
import io as sys_io

# IMPORTANT: Add ADOMD path before importing Pyadomd
from sys import path
path.append(r'C:\Program Files\Microsoft.NET\ADOMD.NET\110')
from pyadomd import Pyadomd

sys.stdout = sys_io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
warnings.filterwarnings("ignore")

# ============================================
# Settings
# ============================================
optimization_report_folder = r"E:\Users\Administrator\Handover V2\Contracting\Optimization Report"
reports_folder = r"E:\Users\Administrator\Handover V2\Contracting\Performance Report"
emails_file = os.path.join(optimization_report_folder, "ContractingEmails.xlsx")

# Password for cube connection
password = 'xxxxxxxxxxxx'

for folder in [reports_folder]:
    if not os.path.exists(folder):
        os.makedirs(folder)

# ============================================
# Cube Connection String
# ============================================
conn_str_cube = f'Provider=MSOLAP.8;Password={password};Persist Security Info=True;User ID=xxxxxx;Initial Catalog=xxxxx - Sales Per Owned Hotels;Data Source=bi.xxxxxxx:1564;Update Isolation Level=2'

# ============================================
# Logging Setup
# ============================================
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.StreamHandler()
    ]
)

# ============================================
# Functions
# ============================================

def get_active_sheet_file():
    """Get YESTERDAY's Active Sheet"""
    yesterday = datetime.now() - timedelta(days=1)
    primary_file = os.path.join(optimization_report_folder, f"Active Sheet {yesterday.strftime('%d')} {yesterday.strftime('%b')} 2026.xlsx")
    fallback_file = os.path.join(optimization_report_folder, "Active Sheet 2026.xlsx")
    
    if os.path.exists(primary_file):
        return primary_file
    return fallback_file

def get_cube_data_for_hotels(hotel_codes, start_date, end_date):
    """
    Fetch Sales, Profit, Bookings, Room Nights from Cube
    """
    if not hotel_codes:
        return pd.DataFrame()
    
    # Clean codes
    hotel_codes_clean = []
    for code in hotel_codes:
        code_str = str(code).strip()
        if '|' in code_str:
            code_str = code_str.split('|')[-1]
        hotel_codes_clean.append(code_str)
    
    # Build MDX set
    hotel_set = '{' + ','.join([f'[Hotels].[Code].&[{code}]' for code in hotel_codes_clean]) + '}'
    
    query = f"""SELECT {{[Measures].[Total Sales], [Measures].[Total Profit], 
                         [Measures].[Bookings], [Measures].[Total Room Nights]}} on columns,
                NON EMPTY {hotel_set} on rows
                from [Sales Per Owned Hotels]
                where ([Booking Date].[Booking Date].&[{start_date}T00:00:00] : [Booking Date].[Booking Date].&[{end_date}T00:00:00],
                       {{[Status].[Status].&[Pag], [Status].[Status].&[Con]}})
                """
    
    try:
        with Pyadomd(conn_str_cube) as conn:
            with conn.cursor().execute(query) as cur:
                result = cur.fetchall()
                if result:
                    df = pd.DataFrame(result, columns=['HotelCode', 'Sales', 'Profit', 'Bookings', 'RoomNights'])
                    df['HotelCode'] = df['HotelCode'].astype(str).str.replace('HBE|', '', regex=False).str.strip()
                    
                    df['Sales'] = pd.to_numeric(df['Sales'], errors='coerce').fillna(0)
                    df['Profit'] = pd.to_numeric(df['Profit'], errors='coerce').fillna(0)
                    df['Bookings'] = pd.to_numeric(df['Bookings'], errors='coerce').fillna(0)
                    df['RoomNights'] = pd.to_numeric(df['RoomNights'], errors='coerce').fillna(0)
                    
                    logging.info(f"[Cube] Data retrieved for {len(df)} hotels on {start_date}")
                    return df
                else:
                    logging.warning(f"[Cube] No data returned for {start_date}")
                    return pd.DataFrame()
    except Exception as e:
        logging.error(f"[Cube] Error: {e}")
        return pd.DataFrame()

def format_excel_performance(file_path, staff_name):
    """Format Excel with performance data for single staff"""
    try:
        workbook = load_workbook(file_path)
        sheet = workbook.active
        
        # Header formatting
        header_fill = PatternFill(start_color="1E4E79", end_color="1E4E79", fill_type="solid")
        for cell in sheet[1]:
            cell.fill = header_fill
            cell.font = Font(color="FFFFFF", bold=True, size=11)
        
        # Column widths
        for column in sheet.columns:
            column_letter = get_column_letter(column[0].column)
            if column[0].value in ['Accommodation Name', 'Product Group', 'Department', 'Country']:
                sheet.column_dimensions[column_letter].width = 25
            else:
                sheet.column_dimensions[column_letter].width = 15
        
        # Borders and alignment
        border = Border(left=Side(style='thin'), right=Side(style='thin'),
                       top=Side(style='thin'), bottom=Side(style='thin'))
        alignment = Alignment(horizontal="center", vertical="center")
        
        for row in sheet.iter_rows(min_row=2):
            for cell in row:
                cell.border = border
                cell.alignment = alignment
        
        # Color coding for performance metrics
        col_indices = {}
        for idx, cell in enumerate(sheet[1], 1):
            if cell.value:
                col_indices[cell.value] = idx
        
        # Color coding for Yesterday metrics
        y_metrics = ['Bookings', 'RoomNights', 'Sales', 'Profit']
        for metric in y_metrics:
            if metric in col_indices:
                col_idx = col_indices[metric]
                max_val = 0
                
                for row in sheet.iter_rows(min_row=2, max_row=sheet.max_row):
                    if row[col_idx-1].value:
                        try:
                            val = float(row[col_idx-1].value)
                            if val > max_val:
                                max_val = val
                        except:
                            pass
                
                for row in sheet.iter_rows(min_row=2, max_row=sheet.max_row):
                    cell = row[col_idx-1]
                    if cell.value and max_val > 0:
                        try:
                            val = float(cell.value)
                            pct = val / max_val
                            if pct >= 0.8:
                                cell.fill = PatternFill(start_color="70AD47", end_color="70AD47", fill_type="solid")
                                cell.font = Font(color="FFFFFF", bold=True)
                            elif pct >= 0.5:
                                cell.fill = PatternFill(start_color="FFC000", end_color="FFC000", fill_type="solid")
                            elif pct >= 0.2:
                                cell.fill = PatternFill(start_color="FF8C00", end_color="FF8C00", fill_type="solid")
                            else:
                                cell.fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
                                cell.font = Font(color="FFFFFF", bold=True)
                        except:
                            pass
        
        # 🆕 Color coding for Profit %
        if 'Profit %' in col_indices:
            col_idx = col_indices['Profit %']
            for row in sheet.iter_rows(min_row=2, max_row=sheet.max_row):
                cell = row[col_idx-1]
                if cell.value:
                    try:
                        val = float(cell.value)
                        if val >= 15:
                            cell.fill = PatternFill(start_color="70AD47", end_color="70AD47", fill_type="solid")
                            cell.font = Font(color="FFFFFF", bold=True)
                        elif val >= 10:
                            cell.fill = PatternFill(start_color="FFC000", end_color="FFC000", fill_type="solid")
                        elif val >= 5:
                            cell.fill = PatternFill(start_color="FF8C00", end_color="FF8C00", fill_type="solid")
                        else:
                            cell.fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
                            cell.font = Font(color="FFFFFF", bold=True)
                    except:
                        pass
        
        # Color coding for Difference columns (WoW)
        diff_cols = ['Bookings Diff (WoW)', 'RoomNights Diff (WoW)', 'Sales Diff (WoW)', 'Profit Diff (WoW)']
        for metric in diff_cols:
            if metric in col_indices:
                col_idx = col_indices[metric]
                for row in sheet.iter_rows(min_row=2, max_row=sheet.max_row):
                    cell = row[col_idx-1]
                    if cell.value:
                        try:
                            val = float(cell.value)
                            if val > 0:
                                cell.fill = PatternFill(start_color="70AD47", end_color="70AD47", fill_type="solid")
                                cell.font = Font(color="FFFFFF", bold=True)
                            elif val < 0:
                                cell.fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
                                cell.font = Font(color="FFFFFF", bold=True)
                            else:
                                cell.fill = PatternFill(start_color="FFC000", end_color="FFC000", fill_type="solid")
                        except:
                            pass
        
        workbook.save(file_path)
        logging.info(f"[Excel] Formatted: {os.path.basename(file_path)}")
        
    except Exception as e:
        logging.error(f"[Excel] Error formatting: {e}")

def format_excel_summary(file_path):
    """Format Excel summary sheet"""
    try:
        workbook = load_workbook(file_path)
        sheet = workbook.active
        
        # Header formatting
        header_fill = PatternFill(start_color="1E4E79", end_color="1E4E79", fill_type="solid")
        for cell in sheet[1]:
            cell.fill = header_fill
            cell.font = Font(color="FFFFFF", bold=True, size=11)
        
        # Column widths
        for column in sheet.columns:
            column_letter = get_column_letter(column[0].column)
            header_value = column[0].value
            if header_value in ['Staff', 'Total Hotels', 'Hotels with Production']:
                sheet.column_dimensions[column_letter].width = 22
            elif 'Producing From' in str(header_value):
                sheet.column_dimensions[column_letter].width = 28
            elif header_value in ['Hotels Produce Both Days', 'Not Produce Both Days']:
                sheet.column_dimensions[column_letter].width = 22
            else:
                sheet.column_dimensions[column_letter].width = 18
        
        # Borders and alignment
        border = Border(left=Side(style='thin'), right=Side(style='thin'),
                       top=Side(style='thin'), bottom=Side(style='thin'))
        alignment = Alignment(horizontal="center", vertical="center")
        
        for row in sheet.iter_rows(min_row=2):
            for cell in row:
                cell.border = border
                cell.alignment = alignment
        
        # Color coding
        col_indices = {}
        for idx, cell in enumerate(sheet[1], 1):
            if cell.value:
                col_indices[cell.value] = idx
        
        # Color coding for Producing From Staff Hotels %
        if 'Producing From Staff Hotels %' in col_indices:
            col_idx = col_indices['Producing From Staff Hotels %']
            for row in sheet.iter_rows(min_row=2, max_row=sheet.max_row):
                cell = row[col_idx-1]
                if cell.value:
                    try:
                        val = float(cell.value)
                        if val >= 80:
                            cell.fill = PatternFill(start_color="70AD47", end_color="70AD47", fill_type="solid")
                            cell.font = Font(color="FFFFFF", bold=True)
                        elif val >= 50:
                            cell.fill = PatternFill(start_color="FFC000", end_color="FFC000", fill_type="solid")
                        elif val >= 20:
                            cell.fill = PatternFill(start_color="FF8C00", end_color="FF8C00", fill_type="solid")
                        else:
                            cell.fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
                            cell.font = Font(color="FFFFFF", bold=True)
                    except:
                        pass
        
        # Color coding for Producing From Total Hotels Producing %
        if 'Producing From Total Hotels Producing %' in col_indices:
            col_idx = col_indices['Producing From Total Hotels Producing %']
            for row in sheet.iter_rows(min_row=2, max_row=sheet.max_row):
                cell = row[col_idx-1]
                if cell.value:
                    try:
                        val = float(cell.value)
                        if val >= 80:
                            cell.fill = PatternFill(start_color="70AD47", end_color="70AD47", fill_type="solid")
                            cell.font = Font(color="FFFFFF", bold=True)
                        elif val >= 50:
                            cell.fill = PatternFill(start_color="FFC000", end_color="FFC000", fill_type="solid")
                        elif val >= 20:
                            cell.fill = PatternFill(start_color="FF8C00", end_color="FF8C00", fill_type="solid")
                        else:
                            cell.fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
                            cell.font = Font(color="FFFFFF", bold=True)
                    except:
                        pass
        
        # Color coding for New Hotels (Green)
        if '🟢 New Hotels' in col_indices:
            col_idx = col_indices['🟢 New Hotels']
            for row in sheet.iter_rows(min_row=2, max_row=sheet.max_row):
                cell = row[col_idx-1]
                if cell.value:
                    try:
                        val = float(cell.value)
                        if val > 0:
                            cell.fill = PatternFill(start_color="70AD47", end_color="70AD47", fill_type="solid")
                            cell.font = Font(color="FFFFFF", bold=True)
                    except:
                        pass
        
        # Color coding for Dropped Hotels (Red)
        if '🔴 Dropped Hotels' in col_indices:
            col_idx = col_indices['🔴 Dropped Hotels']
            for row in sheet.iter_rows(min_row=2, max_row=sheet.max_row):
                cell = row[col_idx-1]
                if cell.value:
                    try:
                        val = float(cell.value)
                        if val > 0:
                            cell.fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
                            cell.font = Font(color="FFFFFF", bold=True)
                    except:
                        pass
        
        # Color coding for Not Produce Both Days (Orange)
        if 'Not Produce Both Days' in col_indices:
            col_idx = col_indices['Not Produce Both Days']
            for row in sheet.iter_rows(min_row=2, max_row=sheet.max_row):
                cell = row[col_idx-1]
                if cell.value:
                    try:
                        val = float(cell.value)
                        if val > 0:
                            cell.fill = PatternFill(start_color="FF8C00", end_color="FF8C00", fill_type="solid")
                            cell.font = Font(color="FFFFFF", bold=True)
                    except:
                        pass
        
        # Highlight high performers (Bookings > 50)
        if 'Bookings' in col_indices:
            col_idx = col_indices['Bookings']
            for row in sheet.iter_rows(min_row=2, max_row=sheet.max_row):
                cell = row[col_idx-1]
                if cell.value:
                    try:
                        val = float(cell.value)
                        if val >= 50:
                            cell.fill = PatternFill(start_color="70AD47", end_color="70AD47", fill_type="solid")
                            cell.font = Font(color="FFFFFF", bold=True)
                        elif val >= 20:
                            cell.fill = PatternFill(start_color="FFC000", end_color="FFC000", fill_type="solid")
                    except:
                        pass
        
        workbook.save(file_path)
        logging.info(f"[Excel] Summary formatted: {os.path.basename(file_path)}")
        
    except Exception as e:
        logging.error(f"[Excel] Error formatting summary: {e}")

def create_charts_for_staff(df, staff_name, max_hotels=10):
    """Create charts for top hotels - Highest at TOP"""
    charts = []
    
    required_cols = ['Accommodation Name', 'Bookings', 'RoomNights', 'Sales', 'Profit']
    missing_cols = [col for col in required_cols if col not in df.columns]
    
    if missing_cols:
        logging.warning(f"[Charts] Missing columns for {staff_name}: {missing_cols}")
        return charts
    
    metrics = [
        ('Bookings', f'Top {max_hotels} Hotels by Bookings'),
        ('RoomNights', f'Top {max_hotels} Hotels by Room Nights'),
        ('Sales', f'Top {max_hotels} Hotels by Sales'),
        ('Profit', f'Top {max_hotels} Hotels by Profit')
    ]
    
    for metric, title in metrics:
        try:
            df[metric] = pd.to_numeric(df[metric], errors='coerce').fillna(0)
            top_n = df.nlargest(max_hotels, metric)[['Accommodation Name', metric]]
            
            if top_n.empty or top_n[metric].sum() == 0:
                continue
            
            top_n = top_n.sort_values(metric, ascending=True)
            
            fig, ax = plt.subplots(figsize=(12, 6))
            colors = plt.cm.Greens(np.linspace(0.9, 0.3, len(top_n)))[::-1]
            
            bars = ax.barh(top_n['Accommodation Name'], top_n[metric], color=colors)
            
            for bar, value in zip(bars, top_n[metric]):
                if metric in ['Sales', 'Profit']:
                    value_text = f'${value:,.0f}'
                else:
                    value_text = f'{value:,.0f}'
                ax.text(value, bar.get_y() + bar.get_height()/2, 
                       value_text, va='center', ha='left', fontsize=9, fontweight='bold')
            
            ax.set_xlabel(metric, fontsize=12, fontweight='bold')
            ax.set_title(f'{title}\n{staff_name}', fontsize=14, fontweight='bold', pad=20)
            ax.set_ylabel('Hotel Name', fontsize=12, fontweight='bold')
            ax.grid(axis='x', alpha=0.3)
            
            plt.tight_layout()
            
            buf = io.BytesIO()
            plt.savefig(buf, format='png', dpi=100, bbox_inches='tight')
            buf.seek(0)
            img_base64 = base64.b64encode(buf.read()).decode('utf-8')
            plt.close()
            
            charts.append({
                'title': title,
                'metric': metric,
                'image': img_base64
            })
            
        except Exception as e:
            logging.error(f"[Charts] Error creating chart for {metric}: {e}")
            continue
    
    return charts

def send_performance_email(staff_name, staff_email, file_path, today, charts_data=None, summary=None, comparison_summary=None):
    """Send email with performance report and charts"""
    try:
        signature_code = ""
        try:
            signature_path = os.path.join(os.environ['USERPROFILE'], r'AppData\Roaming\Microsoft\Signatures\Amr_files\\')
            html_doc = os.path.join(os.environ['USERPROFILE'], r'AppData\Roaming\Microsoft\Signatures\Amr.htm')
            
            if os.path.exists(html_doc):
                html_file = codecs.open(html_doc, 'r', 'utf-8', errors='ignore')
                signature_code = html_file.read()
                signature_code = signature_code.replace('Amr_files/', signature_path)
                html_file.close()
        except:
            pass
        
        outlook = win32com.client.Dispatch("Outlook.Application")
        send_acc = ''
        accounts = outlook.Session.Accounts
        
        for acc in range(len(accounts)):
            if accounts[acc].DisplayName.lower() == 'reports@eetglobal.net'.lower():
                send_acc = accounts[acc]
                break
        
        newMail = outlook.CreateItem(0)
        if send_acc:
            newMail._oleobj_.Invoke(*(64209, 0, 8, 0, send_acc))
        
        newMail.Attachments.Add(file_path)
        newMail.To = staff_email
        newMail.CC = 'xxxxxxxx'
        yesterday_str = (datetime.now() - timedelta(days=1)).strftime('%Y%m%d')
        newMail.Subject = f"{staff_name} - Hotels Performance Report For Yesterday - {yesterday_str}"
        newMail.BodyFormat = 2
        
        yesterday_display = (datetime.now() - timedelta(days=1)).strftime('%d %b %Y')
        last_week_display = (datetime.now() - timedelta(days=8)).strftime('%d %b %Y')
        
        # Build HTML Body
        html_body = f"""
        <html>
        <head>
            <style>
                body {{ font-family: Arial, sans-serif; }}
                .header {{ background-color: #1E4E79; color: white; padding: 10px; border-radius: 5px; }}
                .summary {{ background-color: #f5f5f5; padding: 15px; margin: 10px 0; border-radius: 5px; }}
                .comparison {{ background-color: #e8f0fe; padding: 15px; margin: 10px 0; border-radius: 5px; }}
                .chart-container {{ margin: 25px 0; }}
                .chart-title {{ font-size: 16px; font-weight: bold; color: #333; margin-bottom: 10px; }}
                .chart-img {{ width: 100%; max-width: 800px; border: 1px solid #ddd; border-radius: 5px; }}
                .footer {{ margin-top: 30px; }}
                .stat-box {{ display: inline-block; padding: 10px; margin: 5px; background: #e8f0fe; border-radius: 5px; }}
                .up {{ color: green; font-weight: bold; }}
                .down {{ color: red; font-weight: bold; }}
                .neutral {{ color: orange; font-weight: bold; }}
                .hotel-list {{ background-color: #f9f9f9; padding: 10px; margin: 5px 0; border-radius: 5px; }}
                .new-hotel {{ background-color: #d4edda; padding: 5px; margin: 3px 0; border-radius: 3px; }}
                .dropped-hotel {{ background-color: #f8d7da; padding: 5px; margin: 3px 0; border-radius: 3px; }}
            </style>
        </head>
        <body>
            <div class="header">
                <h2>📊 Hotels Performance Report - {staff_name}</h2>
            </div>
        """
        
        # Summary statistics
        if summary:
            html_body += f"""
            <div class="summary">
                <h3>📋 Summary - Yesterday</h3>
                <div class="stat-box"><strong>🏨 Total Hotels:</strong> {summary.get('total_hotels', 0)}</div>
                <div class="stat-box"><strong>✅ Hotels with Production:</strong> {summary.get('hotels_with_production', 0)}</div>
                <div class="stat-box"><strong>📅 Bookings:</strong> {summary.get('total_bookings', 0):,.0f}</div>
                <div class="stat-box"><strong>🛏️ Room Nights:</strong> {summary.get('total_room_nights', 0):,.0f}</div>
                <div class="stat-box"><strong>💰 Sales:</strong> ${summary.get('total_sales', 0):,.0f}</div>
                <div class="stat-box"><strong>📈 Profit:</strong> ${summary.get('total_profit', 0):,.0f}</div>
                <div class="stat-box"><strong>📊 Profit %:</strong> {summary.get('profit_pct', 0):,.2f}%</div>
            </div>
            """
        
        # Week-over-Week Comparison Section
        if comparison_summary:
            html_body += f"""
            <div class="comparison">
                <h3>📊 Week-over-Week Comparison ({last_week_display} vs {yesterday_display})</h3>
                <div class="stat-box"><strong>📅 Hotels with Bookings:</strong> {comparison_summary.get('hotels_with_bookings_yesterday', 0)} (Yesterday) | {comparison_summary.get('hotels_with_bookings_last_week', 0)} (Last Week)</div>
                <div class="stat-box"><strong>🟢 New Hotels (Started Selling):</strong> <span class="up">{comparison_summary.get('new_hotels_count', 0)}</span></div>
                <div class="stat-box"><strong>🔴 Dropped Hotels (Stopped Selling):</strong> <span class="down">{comparison_summary.get('dropped_hotels_count', 0)}</span></div>
                <div class="stat-box"><strong>Hotels Produce Both Days:</strong> {comparison_summary.get('active_hotels_count', 0)}</div>
                <div class="stat-box"><strong>Not Produce Both Days:</strong> {comparison_summary.get('inactive_hotels_count', 0)}</div>
            """
            
            # New Hotels List
            if comparison_summary.get('new_hotels'):
                html_body += """
                <div class="hotel-list">
                    <h4>🟢 New Hotels (Didn't sell last week, Sold Yesterday):</h4>
                """
                for hotel in comparison_summary['new_hotels'][:20]:
                    html_body += f'<div class="new-hotel">🏨 {hotel}</div>'
                if len(comparison_summary['new_hotels']) > 20:
                    html_body += f"<div><em>... and {len(comparison_summary['new_hotels']) - 20} more</em></div>"
                html_body += "</div>"
            
            # Dropped Hotels List
            if comparison_summary.get('dropped_hotels'):
                html_body += """
                <div class="hotel-list">
                    <h4>🔴 Dropped Hotels (Sold last week, Didn't sell Yesterday):</h4>
                """
                for hotel in comparison_summary['dropped_hotels'][:20]:
                    html_body += f'<div class="dropped-hotel">🏨 {hotel}</div>'
                if len(comparison_summary['dropped_hotels']) > 20:
                    html_body += f"<div><em>... and {len(comparison_summary['dropped_hotels']) - 20} more</em></div>"
                html_body += "</div>"
            
            html_body += "</div>"
        
        # Charts
        if charts_data:
            html_body += """
            <hr>
            <div style="margin: 20px 0;">
                <h3>📈 Top 10 Hotels Performance (Yesterday)</h3>
            </div>
            """
            for chart in charts_data:
                html_body += f"""
                <div class="chart-container">
                    <div class="chart-title">{chart['title']}</div>
                    <img src="data:image/png;base64,{chart['image']}" class="chart-img">
                </div>
                """
        
        html_body += f"""
            <div class="footer">
                <hr>
                <p>Thanks & Best Regards,</p>
                <p><strong>Saker</strong></p>
            </div>
            {signature_code}
        </body>
        </html>
        """
        
        newMail.HTMLBody = html_body
        time.sleep(2)
        newMail.Send()
        logging.info(f"[Email] Sent to {staff_name} ({staff_email})")
        return True
        
    except Exception as e:
        logging.error(f"[Email] Error sending to {staff_name}: {e}")
        return False

def send_summary_email(file_path, today, summary_df):
    """Send summary email to Farouk and Heba with summary table in body"""
    try:
        signature_code = ""
        try:
            signature_path = os.path.join(os.environ['USERPROFILE'], r'AppData\Roaming\Microsoft\Signatures\Amr_files\\')
            html_doc = os.path.join(os.environ['USERPROFILE'], r'AppData\Roaming\Microsoft\Signatures\Amr.htm')
            
            if os.path.exists(html_doc):
                html_file = codecs.open(html_doc, 'r', 'utf-8', errors='ignore')
                signature_code = html_file.read()
                signature_code = signature_code.replace('Amr_files/', signature_path)
                html_file.close()
        except:
            pass
        
        outlook = win32com.client.Dispatch("Outlook.Application")
        send_acc = ''
        accounts = outlook.Session.Accounts
        
        for acc in range(len(accounts)):
            if accounts[acc].DisplayName.lower() == 'reports@eetglobal.net'.lower():
                send_acc = accounts[acc]
                break
        
        newMail = outlook.CreateItem(0)
        if send_acc:
            newMail._oleobj_.Invoke(*(64209, 0, 8, 0, send_acc))
        
        newMail.Attachments.Add(file_path)
        newMail.To = 'xxxxxxx'
        newMail.CC = 'xxxxxxxxx'
        yesterday_str = (datetime.now() - timedelta(days=1)).strftime('%Y%m%d')
        newMail.Subject = f"Hotels Performance Summary Report - {yesterday_str}"
        newMail.BodyFormat = 2
        
        # Build HTML table for summary
        html_table = """
        <table style="border-collapse: collapse; width: 100%; font-family: Arial; font-size: 11px;">
            <tr style="background-color: #1E4E79; color: white;">
                <th style="padding: 8px; border: 1px solid #ddd;">Staff</th>
                <th style="padding: 8px; border: 1px solid #ddd;">Total Hotels</th>
                <th style="padding: 8px; border: 1px solid #ddd;">Hotels with Production</th>
                <th style="padding: 8px; border: 1px solid #ddd;">Producing From Staff Hotels %</th>
                <th style="padding: 8px; border: 1px solid #ddd;">Producing From Total Hotels Producing %</th>
                <th style="padding: 8px; border: 1px solid #ddd; background-color: #28a745;">🟢 New Hotels</th>
                <th style="padding: 8px; border: 1px solid #ddd; background-color: #dc3545;">🔴 Dropped Hotels</th>
                <th style="padding: 8px; border: 1px solid #ddd;">Hotels Produce Both Days</th>
                <th style="padding: 8px; border: 1px solid #ddd;">Not Produce Both Days</th>
                <th style="padding: 8px; border: 1px solid #ddd;">Bookings</th>
                <th style="padding: 8px; border: 1px solid #ddd;">Room Nights</th>
                <th style="padding: 8px; border: 1px solid #ddd;">Sales</th>
                <th style="padding: 8px; border: 1px solid #ddd;">Profit</th>
            </tr>
        """
        
        for _, row in summary_df.iterrows():
            html_table += f"""
            <tr>
                <td style="padding: 6px; border: 1px solid #ddd;"><strong>{row['Staff']}</strong></td>
                <td style="padding: 6px; border: 1px solid #ddd; text-align: center;">{row['Total Hotels']}</td>
                <td style="padding: 6px; border: 1px solid #ddd; text-align: center;">{row['Hotels with Production']}</td>
                <td style="padding: 6px; border: 1px solid #ddd; text-align: center;">{row['Producing From Staff Hotels %']}%</td>
                <td style="padding: 6px; border: 1px solid #ddd; text-align: center;">{row['Producing From Total Hotels Producing %']}%</td>
                <td style="padding: 6px; border: 1px solid #ddd; text-align: center; background-color: #d4edda;"><strong>{row['🟢 New Hotels']}</strong></td>
                <td style="padding: 6px; border: 1px solid #ddd; text-align: center; background-color: #f8d7da;"><strong>{row['🔴 Dropped Hotels']}</strong></td>
                <td style="padding: 6px; border: 1px solid #ddd; text-align: center;">{row['Hotels Produce Both Days']}</td>
                <td style="padding: 6px; border: 1px solid #ddd; text-align: center;">{row['Not Produce Both Days']}</td>
                <td style="padding: 6px; border: 1px solid #ddd; text-align: center;">{row['Bookings']:,.0f}</td>
                <td style="padding: 6px; border: 1px solid #ddd; text-align: center;">{row['Room Nights']:,.0f}</td>
                <td style="padding: 6px; border: 1px solid #ddd; text-align: center;">${row['Sales']:,.0f}</td>
                <td style="padding: 6px; border: 1px solid #ddd; text-align: center;">${row['Profit']:,.0f}</td>
            </tr>
            """
        
        html_table += "</table>"
        
        html_body = f"""
        <html>
        <head>
            <style>
                body {{ font-family: Arial, sans-serif; }}
                .header {{ background-color: #1E4E79; color: white; padding: 10px; border-radius: 5px; }}
                .footer {{ margin-top: 30px; }}
            </style>
        </head>
        <body>
            <div class="header">
                <h2>📊 Hotels Performance Summary Report</h2>
            </div>
            
            <div style="margin: 20px 0;">
                <p>Dear Mr. Farouk & Ms. Heba,</p>
                <p>Please find attached the summary report for all staff members.</p>
                <p><strong>Report Date:</strong> {yesterday_str}</p>
                <p><strong>Summary:</strong></p>
            </div>
            
            {html_table}
            
            <div style="margin: 20px 0;">
                <p><strong>Legend:</strong></p>
                <ul>
                    <li>🟢 <strong>New Hotels:</strong> Started selling yesterday (didn't sell last week same day)</li>
                    <li>🔴 <strong>Dropped Hotels:</strong> Stopped selling yesterday (sold last week same day)</li>
                    <li><strong>Hotels Produce Both Days:</strong> Selling both days</li>
                    <li><strong>Not Produce Both Days:</strong> No sales both days</li>
                </ul>
            </div>
            
            <div class="footer">
                <hr>
                <p>Thanks & Best Regards,</p>
                <p><strong>Saker</strong></p>
            </div>
            {signature_code}
        </body>
        </html>
        """
        
        newMail.HTMLBody = html_body
        time.sleep(2)
        newMail.Send()
        logging.info(f"[Email] Summary sent to Farouk & Heba")
        return True
        
    except Exception as e:
        logging.error(f"[Email] Error sending summary: {e}")
        return False

# ============================================
# Main Run Function
# ============================================

def run_performance_report():
    """Generate performance report for all staff"""
    
    start_time = datetime.now()
    
    print(f"\n{'='*60}")
    print(f"🚀 Starting Hotels Performance Report")
    print(f"📅 Date & Time: {start_time.strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"{'='*60}")
    
    # Step 1: Load YESTERDAY's Active Sheet
    print("\n[Step 1] Loading YESTERDAY's Active Sheet...")
    file_name = get_active_sheet_file()
    if not os.path.exists(file_name):
        print(f"❌ Active sheet not found: {file_name}")
        return
    
    active_sheet = pd.read_excel(file_name)
    yesterday = datetime.now() - timedelta(days=1)
    print(f"✅ Loaded {len(active_sheet)} hotels from Active Sheet {yesterday.strftime('%d')} {yesterday.strftime('%b')} 2026.xlsx")
    
    # Step 2: Load Emails
    print("\n[Step 2] Loading Staff Emails...")
    if not os.path.exists(emails_file):
        print(f"❌ Emails file not found: {emails_file}")
        return
    
    emails_df = pd.read_excel(emails_file)
    staff_list = emails_df['Staff'].unique()
    print(f"✅ Loaded {len(staff_list)} staff members")
    
    # Step 3: Get Cube Data (Yesterday + Last Week Same Day)
    print("\n[Step 3] Fetching Cube Data...")
    today_date = datetime.now()
    yesterday_date = (today_date - timedelta(days=1)).strftime('%Y-%m-%d')
    last_week_date = (today_date - timedelta(days=8)).strftime('%Y-%m-%d')
    
    print(f"   📅 Yesterday: {yesterday_date}")
    print(f"   📅 Last Week Same Day: {last_week_date}")
    
    all_hotel_codes = active_sheet['Code'].astype(str).str.strip().tolist()
    
    # Get data for both days
    cube_yesterday = get_cube_data_for_hotels(all_hotel_codes, yesterday_date, yesterday_date)
    cube_last_week = get_cube_data_for_hotels(all_hotel_codes, last_week_date, last_week_date)
    
    if not cube_yesterday.empty:
        print(f"   📊 Yesterday Total Bookings: {cube_yesterday['Bookings'].sum()}")
    if not cube_last_week.empty:
        print(f"   📊 Last Week Total Bookings: {cube_last_week['Bookings'].sum()}")
    
    # Clean active sheet
    active_sheet['CleanCode'] = active_sheet['Code'].astype(str).str.strip()
    active_sheet['CleanCode'] = active_sheet['CleanCode'].str.replace('HBE|', '', regex=False)
    active_sheet = active_sheet.drop_duplicates(subset=['CleanCode'], keep='first')
    print(f"✅ Active Sheet cleaned: {len(active_sheet)} unique hotels")
    
    # Merge Yesterday Data
    if not cube_yesterday.empty:
        cube_yesterday['HotelCode'] = cube_yesterday['HotelCode'].astype(str).str.strip()
        cube_yesterday = cube_yesterday.drop_duplicates(subset=['HotelCode'], keep='first')
        
        active_sheet = active_sheet.merge(
            cube_yesterday[['HotelCode', 'Sales', 'Profit', 'Bookings', 'RoomNights']],
            left_on='CleanCode',
            right_on='HotelCode',
            how='left'
        )
        active_sheet['Sales'] = active_sheet['Sales'].fillna(0)
        active_sheet['Profit'] = active_sheet['Profit'].fillna(0)
        active_sheet['Bookings'] = active_sheet['Bookings'].fillna(0)
        active_sheet['RoomNights'] = active_sheet['RoomNights'].fillna(0)
        active_sheet = active_sheet.drop('HotelCode', axis=1)
        print(f"✅ Yesterday data merged")
    else:
        active_sheet['Sales'] = 0
        active_sheet['Profit'] = 0
        active_sheet['Bookings'] = 0
        active_sheet['RoomNights'] = 0
        print(f"⚠️ No yesterday data")
    
    # Merge Last Week Data
    if not cube_last_week.empty:
        cube_last_week['HotelCode'] = cube_last_week['HotelCode'].astype(str).str.strip()
        cube_last_week = cube_last_week.drop_duplicates(subset=['HotelCode'], keep='first')
        
        active_sheet = active_sheet.merge(
            cube_last_week[['HotelCode', 'Sales', 'Profit', 'Bookings', 'RoomNights']],
            left_on='CleanCode',
            right_on='HotelCode',
            how='left',
            suffixes=('', '_LW')
        )
        active_sheet['Sales_LW'] = active_sheet['Sales_LW'].fillna(0)
        active_sheet['Profit_LW'] = active_sheet['Profit_LW'].fillna(0)
        active_sheet['Bookings_LW'] = active_sheet['Bookings_LW'].fillna(0)
        active_sheet['RoomNights_LW'] = active_sheet['RoomNights_LW'].fillna(0)
        active_sheet = active_sheet.drop('HotelCode', axis=1)
        print(f"✅ Last Week data merged")
    else:
        active_sheet['Sales_LW'] = 0
        active_sheet['Profit_LW'] = 0
        active_sheet['Bookings_LW'] = 0
        active_sheet['RoomNights_LW'] = 0
        print(f"⚠️ No last week data")
    
    # 🆕 Calculate Profit % for Yesterday
    active_sheet['Profit %'] = active_sheet.apply(
        lambda x: round((x['Profit'] / x['Sales']) * 100, 2) if x['Sales'] > 0 else 0, 
        axis=1
    )
    
    # 🆕 Calculate Profit % for Last Week
    active_sheet['Profit % (Last Week)'] = active_sheet.apply(
        lambda x: round((x['Profit_LW'] / x['Sales_LW']) * 100, 2) if x['Sales_LW'] > 0 else 0, 
        axis=1
    )
    
    # Calculate Week-over-Week Differences
    active_sheet['Bookings_Diff_WoW'] = active_sheet['Bookings'] - active_sheet['Bookings_LW']
    active_sheet['RoomNights_Diff_WoW'] = active_sheet['RoomNights'] - active_sheet['RoomNights_LW']
    active_sheet['Sales_Diff_WoW'] = active_sheet['Sales'] - active_sheet['Sales_LW']
    active_sheet['Profit_Diff_WoW'] = active_sheet['Profit'] - active_sheet['Profit_LW']
    
    # Rename columns
    active_sheet = active_sheet.rename(columns={
        'Bookings_LW': 'Bookings (Last Week)',
        'RoomNights_LW': 'RoomNights (Last Week)',
        'Sales_LW': 'Sales (Last Week)',
        'Profit_LW': 'Profit (Last Week)',
        'Bookings_Diff_WoW': 'Bookings Diff (WoW)',
        'RoomNights_Diff_WoW': 'RoomNights Diff (WoW)',
        'Sales_Diff_WoW': 'Sales Diff (WoW)',
        'Profit_Diff_WoW': 'Profit Diff (WoW)'
    })
    
    if 'CleanCode' in active_sheet.columns:
        active_sheet = active_sheet.drop('CleanCode', axis=1)
    
    print(f"✅ Data merged successfully")
    
    # Step 4: Generate Reports for each Staff
    print(f"\n[Step 4] Generating Reports for {len(staff_list)} Staff Members...")
    today_str = datetime.now().strftime('%Y%m%d')
    
    # Create summary dataframe
    summary_rows = []
    
    for idx, staff_name in enumerate(staff_list, 1):
        print(f"\n   [{idx}/{len(staff_list)}] Processing: {staff_name}")
        
        df = active_sheet[active_sheet['Staff'] == staff_name].copy()
        if df.empty:
            print(f"      ⚠️ No hotels found - skipping")
            continue
        
        staff_row = emails_df[emails_df['Staff'] == staff_name]
        if staff_row.empty:
            print(f"      ⚠️ No email found - skipping")
            continue
        staff_email = staff_row.iloc[0]['Email']
        print(f"      📧 Email: {staff_email}")
        print(f"      🏨 Hotels: {len(df)}")
        
        # Prepare output data
        columns = ['Code', 'Accommodation Name', 'Product Group', 'Department', 'Country',
                  'Bookings', 'RoomNights', 'Sales', 'Profit', 'Profit %',
                  'Bookings (Last Week)', 'RoomNights (Last Week)', 'Sales (Last Week)', 'Profit (Last Week)', 'Profit % (Last Week)',
                  'Bookings Diff (WoW)', 'RoomNights Diff (WoW)', 'Sales Diff (WoW)', 'Profit Diff (WoW)']
        available_cols = [col for col in columns if col in df.columns]
        df_output = df[available_cols].copy()
        
        if 'Bookings' in df_output.columns:
            df_output = df_output.sort_values('Bookings', ascending=False)
        
        output_file = os.path.join(reports_folder, f"{staff_name} - Hotels Performance Report - {today_str}.xlsx")
        df_output.to_excel(output_file, index=False, sheet_name='Hotels Performance')
        print(f"      📁 Saved: {os.path.basename(output_file)}")
        
        format_excel_performance(output_file, staff_name)
        
        print(f"      📊 Creating charts...")
        charts_data = create_charts_for_staff(df, staff_name, max_hotels=10)
        print(f"      ✅ {len(charts_data)} charts created")
        
        # Summary
        hotels_with_production = len(df[df['Bookings'] > 0])
        total_sales = df['Sales'].sum()
        total_profit = df['Profit'].sum()
        profit_pct = round((total_profit / total_sales) * 100, 2) if total_sales > 0 else 0
        
        summary = {
            'total_hotels': len(df),
            'hotels_with_production': hotels_with_production,
            'total_bookings': df['Bookings'].sum(),
            'total_room_nights': df['RoomNights'].sum(),
            'total_sales': total_sales,
            'total_profit': total_profit,
            'profit_pct': profit_pct
        }
        
        # Comparison Summary (WoW)
        hotels_yesterday = df[df['Bookings'] > 0]['Accommodation Name'].tolist()
        hotels_last_week = df[df['Bookings (Last Week)'] > 0]['Accommodation Name'].tolist()
        
        new_hotels = [h for h in hotels_yesterday if h not in hotels_last_week]
        dropped_hotels = [h for h in hotels_last_week if h not in hotels_yesterday]
        active_hotels = [h for h in hotels_yesterday if h in hotels_last_week]
        inactive_hotels = [h for h in df['Accommodation Name'].tolist() if h not in hotels_yesterday and h not in hotels_last_week]
        
        comparison_summary = {
            'hotels_with_bookings_yesterday': len(hotels_yesterday),
            'hotels_with_bookings_last_week': len(hotels_last_week),
            'new_hotels_count': len(new_hotels),
            'dropped_hotels_count': len(dropped_hotels),
            'active_hotels_count': len(active_hotels),
            'inactive_hotels_count': len(inactive_hotels),
            'new_hotels': new_hotels,
            'dropped_hotels': dropped_hotels
        }
        
        print(f"      📊 WoW: New={len(new_hotels)}, Dropped={len(dropped_hotels)}, Active={len(active_hotels)}, Inactive={len(inactive_hotels)}")
        print(f"      📊 Profit %: {profit_pct}%")
        
        # Add to summary rows
        production_from_staff_pct = round((summary['hotels_with_production'] / summary['total_hotels']) * 100, 2) if summary['total_hotels'] > 0 else 0
        
        summary_rows.append({
            'Staff': staff_name,
            'Total Hotels': summary['total_hotels'],
            'Hotels with Production': summary['hotels_with_production'],
            'Producing From Staff Hotels %': production_from_staff_pct,
            'Producing From Total Hotels Producing %': 0,
            '🟢 New Hotels': comparison_summary['new_hotels_count'],
            '🔴 Dropped Hotels': comparison_summary['dropped_hotels_count'],
            'Hotels Produce Both Days': comparison_summary['active_hotels_count'],
            'Not Produce Both Days': comparison_summary['inactive_hotels_count'],
            'Bookings': summary['total_bookings'],
            'Room Nights': summary['total_room_nights'],
            'Sales': summary['total_sales'],
            'Profit': summary['total_profit']
        })
        
        print(f"      📧 Sending email...")
        send_performance_email(staff_name, staff_email, output_file, today_str, charts_data, summary, comparison_summary)
        print(f"      ✅ Done")
    
    # Step 5: Create Summary Sheet
    print(f"\n[Step 5] Creating Summary Report...")
    
    if summary_rows:
        summary_df = pd.DataFrame(summary_rows)
        
        # Calculate Producing From Total Hotels Producing %
        total_producing = summary_df['Hotels with Production'].sum()
        
        if total_producing > 0:
            summary_df['Producing From Total Hotels Producing %'] = summary_df['Hotels with Production'].apply(
                lambda x: round((x / total_producing) * 100, 2)
            )
        else:
            summary_df['Producing From Total Hotels Producing %'] = 0
        
        summary_file = os.path.join(reports_folder, f"Hotels Performance Summary - {today_str}.xlsx")
        summary_df.to_excel(summary_file, index=False, sheet_name='Summary')
        
        format_excel_summary(summary_file)
        print(f"✅ Summary saved: {os.path.basename(summary_file)}")
        
        print(f"📧 Sending summary email to Farouk & Heba...")
        send_summary_email(summary_file, today_str, summary_df)
        print(f"✅ Summary email sent")
    
    # Final Summary
    end_time = datetime.now()
    duration = end_time - start_time
    
    print(f"\n{'='*60}")
    print(f"✅ Report Completed Successfully!")
    print(f"📅 Start Time: {start_time.strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"📅 End Time: {end_time.strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"⏱️ Duration: {duration.total_seconds():.2f} seconds")
    print(f"📁 Reports saved to: {reports_folder}")
    print(f"{'='*60}")

def run_scheduled():
    """Run with schedule check"""
    print(f"\n{'='*60}")
    print(f"🕐 Scheduled Mode - Hotels Performance Report")
    print(f"⏰ Will run daily at 9:00 AM")
    print(f"📁 Reports saved to: {reports_folder}")
    print(f"{'='*60}")
    print("Waiting for scheduled time... Press Ctrl+C to stop")
    print("="*60)
    
    last_run_date = None
    SCHEDULE_HOUR = 9
    
    while True:
        now = datetime.now()
        current_hour = now.hour
        today_date = now.date()
        
        if current_hour == SCHEDULE_HOUR:
            if last_run_date != today_date:
                print(f"\n{'='*60}")
                print(f"🕐 {now.strftime('%H:%M:%S')} - Starting scheduled execution!")
                print(f"{'='*60}")
                
                try:
                    run_performance_report()
                    last_run_date = today_date
                    print(f"\n✅ Completed for {today_date}")
                except Exception as e:
                    logging.error(f"Error in scheduled run: {e}")
                    print(f"❌ Error: {e}")
                
                time.sleep(3600)
        
        time.sleep(30)

# ============================================
# Main
# ============================================

if __name__ == "__main__":
    # Run once (for testing)
    #run_performance_report()
    
    # Or run scheduled (daily at 9 AM)
    run_scheduled()



