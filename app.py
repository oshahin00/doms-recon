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
# --- DATABASE MANAGEMENT (DOMS) ---
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

init_db()

# ==========================================
# --- CORE PROCESSING LOGIC (DOMS) ---
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
    has_offline = res_note.str.contains(r'pss?\s*5000\s*offline\s*\(\s*1100\s*\)') | \
                  summary.str.contains(r'pss?\s*5000\s*offline\s*\(\s*1100\s*\)') | \
                  desc.str.contains(r'pss?\s*5000\s*offline\s*\(\s*1100\s*\)')
    
    target_mask = has_doms_pump & ~has_rfid & ~has_offline
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
st.set_page_config(page_title="DOMS & MYF Tools", page_icon="⚙️", layout="wide")

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

st.title("⚙️ Enterprise Data Hub")

tab1, tab2 = st.tabs(["DOMS Data Merge Tool", "MYF Data & Incident Formatting"])

# -------------------------------------------------------------------------
# TAB 1: DOMS RECONCILIATION
# -------------------------------------------------------------------------
with tab1:
    st.markdown("Automated incident reconciliation and filtering for DOMS and Pump hardware tickets.")
    
    with st.expander("ℹ️ Operating Instructions & Filtering Rules"):
        st.markdown("""
        * **Input:** Requires the raw Incidents Report and the DOMS Rollout schedule.
        * **Strict Filtering:** Automatically isolates tickets explicitly mentioning `DOMS` or `Pumps` while systematically rejecting any ticket referencing `RFID`.
        """)

    col_input1, col_input2 = st.columns(2)
    with col_input1:
        file1 = st.file_uploader("1. Incidents Report (CSV/XLSX)", type=["csv", "xlsx"], key="doms_f1")
    with col_input2:
        file2 = st.file_uploader("2. DOMS Rollout Schedule (XLSX)", type=["xlsx"], key="doms_f2")

    if file1 and file2:
        with st.spinner('Reconciling datasets and applying exclusion rules...'):
            try:
                excel_data, final_df = process_files(file1, file2)
                row_count = len(final_df)
                
                if row_count == 0:
                    st.warning("⚠️ No records found matching the DOMS/PUMPS criteria, or all records were excluded by the filters.")
                else:
                    st.success(f"✅ Processing Complete: Successfully isolated {row_count} validated records.")
                    
                    timestamp_str = datetime.now().strftime("%Y%m%d_%H%M%S")
                    export_filename = f"DOMS_Recon_{timestamp_str}.xlsx"
                    
                    save_run_to_db(export_filename, row_count, excel_data.getvalue())
                    
                    col1, col2 = st.columns([4, 1])
                    with col1:
                        st.subheader("📊 Live Data Preview")
                        st.dataframe(final_df.head(15), width='stretch')
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
        st.info("👈 Please upload both datasets to initiate the reconciliation process.")
        
        st.markdown("<br><br>", unsafe_allow_html=True)
        st.header("🗄️ Extraction History")
        st.caption("Previous reconciliation runs are securely stored in the local SQLite database.")
        
        history_df = get_history_metadata()
        
        if history_df.empty:
            st.write("No historical runs found in the database.")
        else:
            h_col1, h_col2, h_col3, h_col4 = st.columns([2, 3, 2, 2])
            h_col1.markdown("**Execution Date**")
            h_col2.markdown("**Generated Filename**")
            h_col3.markdown("**Records Found**")
            h_col4.markdown("**Action**")
            st.markdown("---")
            
            for index, row in history_df.head(10).iterrows():
                r_col1, r_col2, r_col3, r_col4 = st.columns([2, 3, 2, 2])
                with r_col1: st.write(row['run_date'])
                with r_col2: st.write(row['filename'])
                with r_col3: st.write(f"{row['record_count']} Validated Tickets")
                with r_col4:
                    db_file_data, db_filename = get_file_from_db(row['id'])
                    if db_file_data:
                        st.download_button(
                            label="📥 Download",
                            data=db_file_data,
                            file_name=db_filename,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            key=f"dl_btn_{row['id']}" 
                        )

# -------------------------------------------------------------------------
# TAB 2: MYF DATA & INCIDENT FORMATTING
# -------------------------------------------------------------------------
with tab2:
    st.header("Option 1: MYF Extraction Result Editor")
    st.caption("Upload the Excel file generated by your local Outlook `.bat` script to review and edit the data.")
    
    myf_file = st.file_uploader("Upload MYF Results (Excel)", type=["xlsx"], key="myf_upload")
    
    # Store parsed MYF dataframe in session state so Option 2 can access it
    if 'myf_reference_df' not in st.session_state:
        st.session_state.myf_reference_df = None

    if myf_file:
        try:
            myf_bytes = io.BytesIO(myf_file.read())
            xls_myf = pd.ExcelFile(myf_bytes)
            sheet_names = xls_myf.sheet_names
            
            best_idx = 0
            max_cols = 0
            for idx, s in enumerate(sheet_names):
                temp_df = pd.read_excel(myf_bytes, sheet_name=s)
                if len(temp_df.columns) > max_cols:
                    max_cols = len(temp_df.columns)
                    best_idx = idx
            
            selected_sheet = st.selectbox("Select Excel Sheet", sheet_names, index=best_idx, key="myf_sheet_select")
            
            df_myf = pd.read_excel(myf_bytes, sheet_name=selected_sheet)
            st.session_state.myf_reference_df = df_myf
            
            date_cols = ['Received Time', 'Action Start Time', 'Action End Time']
            for col in date_cols:
                if col in df_myf.columns:
                    df_myf[col] = pd.to_datetime(df_myf[col], format='mixed', errors='coerce')
            
            unique_count = 0
            for col in df_myf.columns:
                if 'incident' in col.lower() and 'number' in col.lower():
                    unique_count = df_myf[col].dropna().nunique()
                    break
            
            st.metric(label="📊 Total Unique Tickets", value=unique_count)
            st.info(f"💡 Loaded sheet: **{selected_sheet}**. Click directly into Date/Time cells to use the interactive date/time picker widget.")
            
            column_configurations = {}
            for col in date_cols:
                if col in df_myf.columns:
                    column_configurations[col] = st.column_config.DatetimeColumn(
                        col,
                        format="DD-MM-YYYY HH:mm:ss",
                        step=1,
                    )
            
            edited_myf_df = st.data_editor(
                df_myf, 
                num_rows="dynamic", 
                width='stretch',
                column_config=column_configurations,
                key="myf_editor"
            )
            
            myf_output = io.BytesIO()
            with pd.ExcelWriter(myf_output, engine='openpyxl') as writer:
                edited_myf_df.to_excel(writer, index=False, sheet_name='Engineers Updated')
                    
            st.download_button(
                label="📥 Export Updated MYF Data",
                data=myf_output.getvalue(),
                file_name=f"MYF_Updated_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="myf_download_btn"
            )
        except Exception as e:
            st.error(f"❌ Failed to load MYF Excel file: {e}")

    st.markdown("---")
    
    st.header("Option 2: Incident File Formatter & Cross-Matcher")
    st.caption("Upload a raw incident report. It will automatically match tickets against Option 1, remove unmatched records, and enforce strict column ordering.")
    
    incident_file = st.file_uploader("Upload Raw Incident Report (CSV/XLSX)", type=["csv", "xlsx"], key="inc_upload")
    
    if incident_file:
        if st.button("🚀 Match & Format Incident Report", type="primary", key="match_inc_btn"):
            try:
                if incident_file.name.endswith('.csv'):
                    df_raw_inc = pd.read_csv(incident_file)
                else:
                    df_raw_inc = pd.read_excel(incident_file)
                    
                df_raw_inc.rename(columns={'Resolution Note': 'Resolution Notes'}, inplace=True)
                
                # Check if Option 1 reference data is available
                if st.session_state.myf_reference_df is None:
                    st.warning("⚠️ Please upload and load your MYF Results file in **Option 1** first so the engine knows which tickets to match against!")
                else:
                    myf_df = st.session_state.myf_reference_df
                    
                    # Extract unique incident numbers from MYF reference data
                    myf_incident_col = None
                    for col in myf_df.columns:
                        if 'incident' in col.lower() and 'number' in col.lower():
                            myf_incident_col = col
                            break
                    
                    # Find incident number column in raw incident report ('Number' or 'Incident Number')
                    inc_number_col = None
                    for candidate in ['Number', 'Incident Number', 'Incident']:
                        if candidate in df_raw_inc.columns:
                            inc_number_col = candidate
                            break
                    
                    if not myf_incident_col or not inc_number_col:
                        st.error(f"❌ Could not locate Incident Number columns. MYF col found: {myf_incident_col}, Incident Report col found: {inc_number_col}")
                    else:
                        # Normalize strings for clean matching
                        valid_tickets = set(myf_df[myf_incident_col].dropna().astype(str).str.strip().str.upper())
                        
                        # Filter raw incidents to KEEP ONLY those present in Option 1 (remove unmatched)
                        mask_match = df_raw_inc[inc_number_col].dropna().astype(str).str.strip().str.upper().isin(valid_tickets)
                        df_matched = df_raw_inc[mask_match]
                        
                        mandated_columns = [
                            "Number", "Priority", "Created", "Area", "Assigned To", "Summary", 
                            "Description", "Resolution Notes", "State", "On Hold Reason", "Work Notes", 
                            "Reported by", "Status", "Assignment Group", "Resolved", 
                            "Additional Comments", "Cause Category", "Cause Code", "First Assignment Group"
                        ]
                        
                        final_ordered_cols = [col for col in mandated_columns if col in df_matched.columns]
                        df_formatted_inc = df_matched[final_ordered_cols]
                        
                        matched_count = len(df_formatted_inc)
                        st.success(f"✅ Matching Complete: Found and retained {matched_count} matching incident(s)!")
                        st.dataframe(df_formatted_inc.head(15), width='stretch')
                        
                        inc_output = io.BytesIO()
                        with pd.ExcelWriter(inc_output, engine='openpyxl') as writer:
                            df_formatted_inc.to_excel(writer, index=False, sheet_name='Formatted_Incidents')
                        
                        st.download_button(
                            label="📥 Download Matched & Formatted Incidents",
                            data=inc_output.getvalue(),
                            file_name=f"Matched_Incidents_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            key="inc_download_btn"
                        )
            except Exception as e:
                st.error(f"❌ Failed to process and match incident file: {e}")