import streamlit as st
import pandas as pd
import re
import io
import sqlite3
from datetime import datetime
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl import Workbook

# ==========================================
# --- DATABASE MANAGEMENT ---
# ==========================================
DB_NAME = "/app/doms_history.db"

def init_db():
    conn = sqlite3.connect(DB_NAME)
    c = conn.cursor()
    c.execute('''
        CREATE TABLE IF NOT EXISTS run_history (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            run_date TEXT,
            filename TEXT,
            record_count INTEGER,
            file_data BLOB
        )
    ''')
    conn.commit()
    conn.close()

def save_run_to_db(filename, record_count, file_data):
    conn = sqlite3.connect(DB_NAME)
    c = conn.cursor()
    run_date = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    c.execute('''
        INSERT INTO run_history (run_date, filename, record_count, file_data)
        VALUES (?, ?, ?, ?)
    ''', (run_date, filename, record_count, file_data))
    conn.commit()
    conn.close()

def get_history_metadata():
    conn = sqlite3.connect(DB_NAME)
    df = pd.read_sql_query("SELECT id, run_date, filename, record_count FROM run_history ORDER BY run_date DESC", conn)
    conn.close()
    return df

def get_file_from_db(record_id):
    conn = sqlite3.connect(DB_NAME)
    c = conn.cursor()
    c.execute("SELECT file_data, filename FROM run_history WHERE id = ?", (record_id,))
    result = c.fetchone()
    conn.close()
    return result

# Initialize the DB when the app starts
init_db()

# ==========================================
# --- CORE PROCESSING LOGIC ---
# ==========================================
def extract_first_number(area_str):
    if not isinstance(area_str, str):
        return None
    match = re.search(r'\d+', area_str)
    return int(match.group()) if match else None

def process_files(file1, file2):
    if file1.name.endswith('.csv'):
        df1 = pd.read_csv(file1)
    else:
        df1 = pd.read_excel(file1)
        
    df1.rename(columns={'Resolution Notes': 'Resolution Note'}, inplace=True)
    
    for col in ['Description', 'Cause Category', 'Cause Code']:
        if col not in df1.columns:
            df1[col] = ''
    
    xls = pd.ExcelFile(file2)
    df2 = pd.read_excel(file2, sheet_name=xls.sheet_names[0])
    df2.columns = df2.columns.str.strip()
    df2['Station No.'] = pd.to_numeric(df2['Station No.'], errors='coerce').astype('Int64')
    
    df1['Extracted_Station_No'] = df1['Area'].apply(extract_first_number).astype('Int64')
    
    merged_df = pd.merge(df1, df2, left_on='Extracted_Station_No', right_on='Station No.', how='inner')
    
    merged_df['Created_dt'] = pd.to_datetime(merged_df['Created'], errors='coerce')
    merged_df['Date_dt'] = pd.to_datetime(merged_df['Date'], errors='coerce')
    filtered_df = merged_df[merged_df['Created_dt'] >= merged_df['Date_dt']]
    
    res_note = filtered_df['Resolution Note'].fillna('').str.lower()
    summary = filtered_df['Summary'].fillna('').str.lower()
    desc = filtered_df['Description'].fillna('').str.lower()
    
    has_rfid = res_note.str.contains('rfid') | summary.str.contains('rfid') | desc.str.contains('rfid')
    has_doms_pump = res_note.str.contains('doms|pump') | summary.str.contains('doms|pump') | desc.str.contains('doms|pump')
    
    target_mask = has_doms_pump & ~has_rfid
    filtered_df = filtered_df[target_mask]
    
    final_cols = [
        'Station No.', 'Number', 'Priority', 'Created', 'Summary', 
        'Resolution Note', 'Status', 'Cause Category', 'Cause Code', 
        'DOMS Model', 'Date'
    ]
    
    final_cols = [col for col in final_cols if col in filtered_df.columns]
    final_df = filtered_df[final_cols]
    
    output = io.BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.title = "Comparison Results"
    
    for r in dataframe_to_rows(final_df, index=False, header=True):
        ws.append(r)
        
    header_fill = PatternFill(start_color="2F5597", end_color="2F5597", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True)
    thin_border = Border(left=Side(style='thin', color='D9D9D9'),
                         right=Side(style='thin', color='D9D9D9'),
                         top=Side(style='thin', color='D9D9D9'),
                         bottom=Side(style='thin', color='D9D9D9'))
                         
    for col_idx, cell in enumerate(ws[1], 1):
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = thin_border
        
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
        for cell in row:
            cell.border = thin_border
            cell.alignment = Alignment(vertical="top", wrap_text=True)
            
    for col in ws.columns:
        max_length = max((len(str(cell.value).split('\n')[0]) for cell in col if cell.value), default=0)
        ws.column_dimensions[col[0].column_letter].width = max(min(max_length + 2, 65), 12)
        
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    
    wb.save(output)
    output.seek(0)
    
    return output, final_df 

# ==========================================
# --- ENTERPRISE WEB INTERFACE & UI/UX ---
# ==========================================
st.set_page_config(page_title="DOMS Recon", page_icon="⚙️", layout="wide")

st.markdown("""
    <style>
    .stDownloadButton button {
        width: 100%;
        background-color: #2F5597;
        color: white;
        border-radius: 5px;
        font-weight: bold;
    }
    .stDownloadButton button:hover {
        background-color: #1e3a68;
        color: white;
    }
    .history-row {
        padding: 10px 0px;
        border-bottom: 1px solid #333;
    }
    </style>
""", unsafe_allow_html=True)

st.title("⚙️ DOMS Data Merge & Filter Tool")
st.markdown("Automated incident reconciliation and filtering for DOMS and Pump hardware tickets.")

with st.expander("ℹ️ Operating Instructions & Filtering Rules"):
    st.markdown("""
    * **Input:** Requires the raw Incidents Report and the DOMS Rollout schedule.
    * **Strict Filtering:** Automatically isolates tickets explicitly mentioning `DOMS` or `Pumps` while systematically rejecting any ticket referencing `RFID`.
    """)

with st.sidebar:
    st.header("📂 Data Input")
    file1 = st.file_uploader("1. Incidents Report (CSV/XLSX)", type=["csv", "xlsx"])
    file2 = st.file_uploader("2. DOMS Rollout Schedule (XLSX)", type=["xlsx"])
    st.markdown("---")
    st.caption("Environment: Node 2 (App Server)")

if file1 and file2:
    with st.spinner('Reconciling datasets and applying exclusion rules...'):
        try:
            excel_data, final_df = process_files(file1, file2)
            row_count = len(final_df)
            
            if row_count == 0:
                st.warning("⚠️ No records found matching the DOMS/PUMPS criteria, or all records were excluded by the RFID filter.")
            else:
                st.success(f"✅ Processing Complete: Successfully isolated {row_count} validated records.")
                
                # Generate dynamic filename
                timestamp_str = datetime.now().strftime("%Y%m%d_%H%M%S")
                export_filename = f"DOMS_Recon_{timestamp_str}.xlsx"
                
                # Save to DB instantly
                save_run_to_db(export_filename, row_count, excel_data.getvalue())
                
                col1, col2 = st.columns([4, 1])
                with col1:
                    st.subheader("📊 Live Data Preview")
                    st.dataframe(final_df.head(15), use_container_width=True)
                    st.caption(f"Showing top 15 of {row_count} records. Download the Excel file to view the complete dataset.")
                    
                with col2:
                    st.subheader("Export")
                    st.download_button(
                        label="📥 Download Report",
                        data=excel_data,
                        file_name=export_filename,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
        except Exception as e:
            st.error(f"❌ A structural error occurred. Error: {e}")
else:
    st.info("👈 Please upload both datasets in the sidebar menu to initiate the reconciliation process.")
    
    # --- DB HISTORY SECTION (Renders when waiting for input) ---
    st.markdown("<br><br>", unsafe_allow_html=True)
    st.header("🗄️ Extraction History")
    st.caption("Previous reconciliation runs are securely stored in the local SQLite database. You can download historical files below.")
    
    history_df = get_history_metadata()
    
    if history_df.empty:
        st.write("No historical runs found in the database.")
    else:
        # Create a clean header row
        h_col1, h_col2, h_col3, h_col4 = st.columns([2, 3, 2, 2])
        h_col1.markdown("**Execution Date**")
        h_col2.markdown("**Generated Filename**")
        h_col3.markdown("**Records Found**")
        h_col4.markdown("**Action**")
        st.markdown("---")
        
        # Display the top 10 most recent runs
        for index, row in history_df.head(10).iterrows():
            r_col1, r_col2, r_col3, r_col4 = st.columns([2, 3, 2, 2])
            
            with r_col1:
                st.write(row['run_date'])
            with r_col2:
                st.write(row['filename'])
            with r_col3:
                st.write(f"{row['record_count']} Validated Tickets")
            with r_col4:
                # Fetch the BLOB file data from the DB for the download button
                db_file_data, db_filename = get_file_from_db(row['id'])
                if db_file_data:
                    st.download_button(
                        label="📥 Download",
                        data=db_file_data,
                        file_name=db_filename,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key=f"dl_btn_{row['id']}" # Unique key required by Streamlit
                    )