import streamlit as st
import pandas as pd
import re
import io
import os
import yaml
import time
from datetime import datetime
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl import Workbook
from sqlalchemy import create_engine, text
from sqlalchemy.engine import URL

# ==========================================
# --- PAGE CONFIGURATION & DESIGN SYSTEM ---
# ==========================================
st.set_page_config(page_title="OpsFlow Studio", page_icon="⚡", layout="wide")

# Inject Custom SaaS CSS (Typography, Cards, Buttons, Micro-interactions)
st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600&display=swap');
    
    html, body, [class*="css"] {
        font-family: 'Inter', sans-serif;
    }
    
    /* Button Styling */
    .stDownloadButton button, .stButton button {
        width: 100%;
        border-radius: 6px !important;
        font-weight: 500 !important;
        transition: all 0.15s ease-in-out !important;
    }
    
    .stButton button:hover {
        transform: translateY(-1px);
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1), 0 2px 4px -1px rgba(0, 0, 0, 0.06) !important;
    }
    
    /* Expander & Container Borders */
    [data-testid="stExpander"] {
        border-radius: 8px !important;
        border: 1px solid #e5e7eb !important;
        box-shadow: 0 1px 2px 0 rgba(0, 0, 0, 0.05) !important;
    }
    
    /* Hide default Streamlit top padding and footer for a cleaner app feel */
    .block-container {
        padding-top: 2rem !important;
    }
    header {visibility: hidden;}
    footer {visibility: hidden;}
    
    /* Typography adjustments */
    h1 { font-weight: 600 !important; letter-spacing: -0.025em !important; }
    h2, h3 { font-weight: 500 !important; letter-spacing: -0.025em !important; }
    </style>
""", unsafe_allow_html=True)

# ==========================================
# --- CONFIGURATION & DATABASE ---
# ==========================================
try:
    with open("config.yaml", "r") as file:
        config = yaml.safe_load(file)
    db_cfg = config['database']
except FileNotFoundError:
    st.error("Configuration file 'config.yaml' not found. Please create it.")
    st.stop()

url_object = URL.create(
    drivername="postgresql+psycopg2",
    username=db_cfg['username'],
    password=db_cfg['password'],
    host=db_cfg['host'],
    port=db_cfg['port'],
    database=db_cfg['dbname']
)

engine = create_engine(url_object)

def init_db():
    try:
        with engine.begin() as conn:
            conn.execute(text('''
                CREATE TABLE IF NOT EXISTS run_history (
                    id SERIAL PRIMARY KEY,
                    report_type VARCHAR(50),
                    run_date VARCHAR(50),
                    filename VARCHAR(255),
                    record_count INT,
                    file_data BYTEA
                )
            '''))
    except Exception as e:
        print(f"Database Initialization Error: {e}")

def save_run_to_db(report_type, filename, record_count, file_data):
    run_date = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with engine.begin() as conn:
        conn.execute(
            text("""
                INSERT INTO run_history (report_type, run_date, filename, record_count, file_data) 
                VALUES (:report_type, :run_date, :filename, :record_count, :file_data)
            """),
            {
                "report_type": report_type, 
                "run_date": run_date, 
                "filename": filename, 
                "record_count": record_count, 
                "file_data": file_data
            }
        )

def get_history_metadata(report_type):
    try:
        query = text("SELECT id, run_date, filename, record_count FROM run_history WHERE report_type = :rtype ORDER BY run_date DESC")
        df = pd.read_sql(query, engine, params={"rtype": report_type})
        return df
    except Exception as e:
        return pd.DataFrame(columns=['id', 'run_date', 'filename', 'record_count'])

def get_file_from_db(record_id):
    try:
        query = text("SELECT file_data, filename FROM run_history WHERE id = :rid")
        with engine.connect() as conn:
            result = conn.execute(query, {"rid": record_id}).fetchone()
            if result:
                return bytes(result[0]), result[1]
    except Exception as e:
        pass
    return None, None

def delete_record_from_db(record_id):
    try:
        with engine.begin() as conn:
            conn.execute(text("DELETE FROM run_history WHERE id = :rid"), {"rid": record_id})
        return True
    except Exception as e:
        return False

def export_to_local_output(data, filename):
    os.makedirs("output", exist_ok=True)
    filepath = os.path.join("output", filename)
    with open(filepath, "wb") as f:
        if isinstance(data, io.BytesIO):
            f.write(data.getvalue())
        else:
            f.write(data)
    return filepath

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
    
    return final_df

def generate_excel_bytes(df):
    output = io.BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.title = "Comparison Results"
    
    for r in dataframe_to_rows(df, index=False, header=True):
        ws.append(r)
        
    header_fill = PatternFill(start_color="111827", end_color="111827", fill_type="solid") # Modern Dark SaaS Header
    header_font = Font(color="FFFFFF", bold=True)
    thin_border = Border(left=Side(style='thin', color='E5E7EB'),
                         right=Side(style='thin', color='E5E7EB'),
                         top=Side(style='thin', color='E5E7EB'),
                         bottom=Side(style='thin', color='E5E7EB'))
                         
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
    return output


# ==========================================
# --- UI: SIDEBAR & NAVIGATION ---
# ==========================================
st.title("⚡ OpsFlow Studio")
st.markdown("Automated ticket reconciliation and operational data modeling.")

with st.sidebar:
    st.header("Workspace Guide")
    st.markdown("---")
    st.markdown("""
    **DOMS Reconciliation**
    *   **Strict Filtering:** Automatically isolates tickets explicitly mentioning `DOMS` or `Pumps`.
    *   **Exclusions:** Systematically rejects any ticket referencing `RFID`.
    
    **MYF & Incidents**
    *   Ensure exact ticket ID matches when cross-referencing raw incident reports against audited MYF datasets.
    """)
    st.markdown("---")
    st.caption("Engine: PostgreSQL Desktop | Deployed via Docker")

# Dynamic Reset Counters
if "doms_counter" not in st.session_state: st.session_state.doms_counter = 0
if "myf_counter" not in st.session_state: st.session_state.myf_counter = 0
if "inc_counter" not in st.session_state: st.session_state.inc_counter = 0

tab1, tab2 = st.tabs(["DOMS Data Merge", "MYF & Incident Management"])

# ==========================================
# --- TAB 1: DOMS RECONCILIATION ---
# ==========================================
with tab1:
    st.markdown("<br>", unsafe_allow_html=True)
    
    # CARD: Upload Zone
    with st.container(border=True):
        st.subheader("1. Source Data")
        col_input1, col_input2 = st.columns(2)
        with col_input1:
            file1 = st.file_uploader("Raw Incidents Report (CSV/XLSX)", type=["csv", "xlsx"], key=f"doms_f1_{st.session_state.doms_counter}")
        with col_input2:
            file2 = st.file_uploader("DOMS Rollout Schedule (XLSX)", type=["xlsx"], key=f"doms_f2_{st.session_state.doms_counter}")

    if file1 and file2:
        with st.spinner('Reconciling datasets...'):
            try:
                raw_final_df = process_files(file1, file2)
                row_count_initial = len(raw_final_df)
                
                if row_count_initial == 0:
                    st.warning("⚠️ No matching DOMS/PUMPS records found based on exclusion rules.")
                else:
                    st.toast(f"✅ Reconciled {row_count_initial} valid records.", icon="✅")
                    
                    # CARD: Live Editor
                    with st.container(border=True):
                        st.subheader("2. Audit & Verify")
                        st.caption("Review the filtered data. Select '🗑️ Remove' to discard any anomalies before finalizing.")
                        
                        raw_final_df.insert(0, '🗑️ Remove', False)
                        edited_preview_df = st.data_editor(raw_final_df, num_rows="fixed", width='stretch', hide_index=True, key="doms_editor")
                        
                        clean_preview_df = edited_preview_df[~edited_preview_df['🗑️ Remove']].drop(columns=['🗑️ Remove'])
                        current_row_count = len(clean_preview_df)
                        excel_data = generate_excel_bytes(clean_preview_df)
                        timestamp_str = datetime.now().strftime("%Y%m%d_%H%M%S")
                        export_filename = f"DOMS_Recon_{timestamp_str}.xlsx"
                        
                        st.markdown("<br>", unsafe_allow_html=True)
                        col_act1, col_act2 = st.columns([3, 1])
                        with col_act1:
                            if st.button("💾 Commit & Save to Database", key="save_doms_db", type="primary"):
                                save_run_to_db("DOMS", export_filename, current_row_count, excel_data.getvalue())
                                st.toast("Commit successful. Pipeline updated.", icon="✅")
                                time.sleep(1)
                                st.session_state.doms_counter += 1
                                st.rerun()
                        with col_act2:
                            if st.button("📥 Direct Export", key="doms_dl_new"):
                                filepath = export_to_local_output(excel_data, export_filename)
                                st.toast(f"Exported to /output directory.", icon="🚀")
                                time.sleep(1)
                                st.session_state.doms_counter += 1
                                st.rerun()
            except Exception as e:
                st.error(f"Execution Error: {e}")
                
    st.markdown("<br>", unsafe_allow_html=True)
    
    # CARD: History
    with st.container(border=True):
        st.subheader("🗄️ Extraction History")
        history_df = get_history_metadata("DOMS")
        
        if history_df.empty:
            st.caption("No historical deployments found.")
        else:
            h_col1, h_col2, h_col3, h_col4, h_col5, h_col6 = st.columns([2, 3, 1, 1, 1, 1])
            h_col1.markdown("**Date (AST)**")
            h_col2.markdown("**Filename**")
            h_col3.markdown("**Rows**")
            st.markdown("---")
            
            for index, row in history_df.iterrows():
                r_id = row['id']
                r_col1, r_col2, r_col3, r_col4, r_col5, r_col6 = st.columns([2, 3, 1, 1, 1, 1], vertical_alignment="center")
                with r_col1: st.caption(str(row['run_date']))
                with r_col2: st.caption(str(row['filename']))
                with r_col3: st.caption(f"{row['record_count']}")
                
                db_file_data, db_filename = get_file_from_db(r_id)
                
                with r_col4:
                    if db_file_data:
                        if st.button("📥 Export", key=f"dl_doms_{r_id}", use_container_width=True):
                            filepath = export_to_local_output(db_file_data, db_filename)
                            st.toast(f"Exported to local volume.", icon="🚀")
                            time.sleep(1)
                            st.rerun()
                with r_col5:
                    if st.button("👁️ View", key=f"rev_doms_{r_id}", use_container_width=True):
                        st.session_state[f"show_review_{r_id}"] = not st.session_state.get(f"show_review_{r_id}", False)
                with r_col6:
                    if st.button("🗑️ Drop", key=f"del_doms_{r_id}", use_container_width=True):
                        delete_record_from_db(r_id)
                        st.rerun()
                
                if st.session_state.get(f"show_review_{r_id}", False) and db_file_data:
                    st.info(f"Viewing Snapshot: {db_filename}")
                    rev_df = pd.read_excel(io.BytesIO(db_file_data))
                    st.dataframe(rev_df, width='stretch')
                
                st.divider()

# ==========================================
# --- TAB 2: MYF & INCIDENT MANAGEMENT ---
# ==========================================
with tab2:
    st.markdown("<br>", unsafe_allow_html=True)
    subtab1, subtab2 = st.tabs(["Step 1: MYF Context Editor", "Step 2: Incident Cross-Match"])
    
    # -------------------------------------------------------------------------
    # SUB-TAB 1: MYF AUDIT
    # -------------------------------------------------------------------------
    with subtab1:
        with st.container(border=True):
            st.subheader("MYF Initialization")
            myf_file = st.file_uploader("Upload Generated MYF Extraction (Excel)", type=["xlsx"], key=f"myf_upload_{st.session_state.myf_counter}")
            
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
                    
                    selected_sheet = st.selectbox("Target Worksheet", sheet_names, index=best_idx, key="myf_sheet_select")
                    
                    df_myf = pd.read_excel(myf_bytes, sheet_name=selected_sheet)
                    st.session_state.myf_reference_df = df_myf.copy()
                    
                    df_myf.insert(0, '🗑️ Remove', False)
                    
                    date_cols = ['Received Time', 'Action Start Time', 'Action End Time']
                    for col in date_cols:
                        if col in df_myf.columns:
                            df_myf[col] = pd.to_datetime(df_myf[col], format='mixed', errors='coerce')
                    
                    unique_count = 0
                    for col in df_myf.columns:
                        if 'incident' in col.lower() and 'number' in col.lower():
                            unique_count = df_myf[col].dropna().nunique()
                            break
                    
                    st.metric(label="Total Unique Tickets Loaded", value=unique_count)
                    
                    column_configurations = {
                        '🗑️ Remove': st.column_config.CheckboxColumn("🗑️ Remove", default=False)
                    }
                    for col in date_cols:
                        if col in df_myf.columns:
                            column_configurations[col] = st.column_config.DatetimeColumn(
                                col,
                                format="DD-MM-YYYY HH:mm:ss",
                                step=1,
                            )
                    
                    edited_myf_df = st.data_editor(
                        df_myf, 
                        num_rows="fixed", 
                        width='stretch',
                        hide_index=True,
                        column_config=column_configurations,
                        key="myf_editor"
                    )
                    
                    clean_myf_df = edited_myf_df[~edited_myf_df['🗑️ Remove']].drop(columns=['🗑️ Remove'])
                    
                    myf_output = io.BytesIO()
                    with pd.ExcelWriter(myf_output, engine='openpyxl') as writer:
                        clean_myf_df.to_excel(writer, index=False, sheet_name='Engineers Updated')
                    
                    myf_filename = f"MYF_Updated_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
                    
                    st.markdown("<br>", unsafe_allow_html=True)
                    col_m1, col_m2 = st.columns([3, 1])
                    with col_m1:
                        if st.button("💾 Commit MYF to Database", key="save_myf_db", type="primary"):
                            save_run_to_db("MYF_Update", myf_filename, len(clean_myf_df), myf_output.getvalue())
                            st.toast("MYF context saved to history.", icon="✅")
                            time.sleep(1)
                            st.session_state.myf_counter += 1
                            st.rerun()
                    with col_m2:
                        if st.button("📥 Direct Export", key="myf_download_btn"):
                            filepath = export_to_local_output(myf_output, myf_filename)
                            st.toast(f"Exported to /output directory.", icon="🚀")
                            time.sleep(1)
                            st.session_state.myf_counter += 1
                            st.rerun()
                except Exception as e:
                    st.error(f"Data Binding Error: {e}")

    # -------------------------------------------------------------------------
    # SUB-TAB 2: INCIDENT CROSS-MATCHER
    # -------------------------------------------------------------------------
    with subtab2:
        with st.container(border=True):
            st.subheader("Incident Cross-Reference Engine")
            
            if st.session_state.myf_reference_df is None:
                st.info("🔒 Prerequisite: Upload and configure the MYF Results in Step 1 to unlock the Cross-Matcher.")
            else:
                incident_file = st.file_uploader("Raw Incident Report (CSV/XLSX)", type=["csv", "xlsx"], key=f"inc_upload_{st.session_state.inc_counter}")
                
                if incident_file:
                    try:
                        if incident_file.name.endswith('.csv'):
                            df_raw_inc = pd.read_csv(incident_file)
                        else:
                            df_raw_inc = pd.read_excel(incident_file)
                            
                        df_raw_inc.rename(columns={'Resolution Note': 'Resolution Notes'}, inplace=True)
                        
                        myf_df = st.session_state.myf_reference_df
                        myf_incident_col = None
                        
                        for col in myf_df.columns:
                            if 'incident' in col.lower() and 'number' in col.lower():
                                myf_incident_col = col
                                break
                        
                        inc_number_col = None
                        for candidate in ['Number', 'Incident Number', 'Incident']:
                            if candidate in df_raw_inc.columns:
                                inc_number_col = candidate
                                break
                        
                        if not myf_incident_col or not inc_number_col:
                            st.error("Schema Mismatch: Could not locate standard Incident Number keys.")
                        else:
                            valid_tickets = set(myf_df[myf_incident_col].dropna().astype(str).str.strip().str.upper())
                            match_mask = df_raw_inc[inc_number_col].astype(str).str.strip().str.upper().isin(valid_tickets).to_numpy()
                            df_matched = df_raw_inc[match_mask]
                            
                            mandated_columns = [
                                "Number", "Priority", "Created", "Area", "Assigned To", "Summary", 
                                "Description", "Resolution Notes", "State", "On Hold Reason", "Work Notes", 
                                "Reported by", "Status", "Assignment Group", "Resolved", 
                                "Additional Comments", "Cause Category", "Cause Code", "First Assignment Group"
                            ]
                            
                            final_ordered_cols = [col for col in mandated_columns if col in df_raw_inc.columns]
                            df_formatted_inc = df_matched[final_ordered_cols]
                            
                            st.toast(f"✅ Match Complete: {len(df_formatted_inc)} incidents retained.", icon="✅")
                            
                            df_formatted_inc.insert(0, '🗑️ Remove', False)
                            edited_matched_inc = st.data_editor(df_formatted_inc, num_rows="fixed", width='stretch', hide_index=True, key="matched_inc_editor")
                            
                            clean_matched_inc = edited_matched_inc[~edited_matched_inc['🗑️ Remove']].drop(columns=['🗑️ Remove'])
                            
                            inc_output = io.BytesIO()
                            with pd.ExcelWriter(inc_output, engine='openpyxl') as writer:
                                clean_matched_inc.to_excel(writer, index=False, sheet_name='Formatted_Incidents')
                            
                            inc_filename = f"Matched_Incidents_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
                            
                            st.markdown("<br>", unsafe_allow_html=True)
                            col_i1, col_i2 = st.columns([3, 1])
                            with col_i1:
                                if st.button("💾 Commit Matched Report to Database", key="save_matched_db", type="primary"):
                                    save_run_to_db("Matched_Incident", inc_filename, len(clean_matched_inc), inc_output.getvalue())
                                    st.toast("Matched report saved to history.", icon="✅")
                                    time.sleep(1)
                                    st.session_state.inc_counter += 1
                                    st.rerun()
                            with col_i2:
                                if st.button("📥 Direct Export", key="inc_download_btn"):
                                    filepath = export_to_local_output(inc_output, inc_filename)
                                    st.toast(f"Exported to /output directory.", icon="🚀")
                                    time.sleep(1)
                                    st.session_state.inc_counter += 1
                                    st.rerun()
                    except Exception as e:
                        st.error(f"Processing Error: {e}")

    # --- TAB 2 HISTORY SECTION ---
    st.markdown("<br>", unsafe_allow_html=True)
    with st.container(border=True):
        st.subheader("🗄️ MYF & Match History")
        
        history_myf = get_history_metadata("MYF_Update")
        history_match = get_history_metadata("Matched_Incident")
        history_combined = pd.concat([history_myf, history_match])
        
        if not history_combined.empty:
            history_combined = history_combined.sort_values(by="run_date", ascending=False)

        if history_combined.empty:
            st.caption("No historical configurations found.")
        else:
            h_col1, h_col2, h_col3, h_col4, h_col5, h_col6 = st.columns([2, 3, 1, 1, 1, 1])
            h_col1.markdown("**Date (AST)**")
            h_col2.markdown("**Filename**")
            h_col3.markdown("**Rows**")
            st.markdown("---")
            
            for index, row in history_combined.iterrows():
                r_id = row['id']
                r_col1, r_col2, r_col3, r_col4, r_col5, r_col6 = st.columns([2, 3, 1, 1, 1, 1], vertical_alignment="center")
                with r_col1: st.caption(str(row['run_date']))
                with r_col2: st.caption(str(row['filename']))
                with r_col3: st.caption(f"{row['record_count']}")
                
                db_file_data, db_filename = get_file_from_db(r_id)
                
                with r_col4:
                    if db_file_data:
                        if st.button("📥 Export", key=f"dl_tab2_{r_id}", use_container_width=True):
                            filepath = export_to_local_output(db_file_data, db_filename)
                            st.toast(f"Exported to local volume.", icon="🚀")
                            time.sleep(1)
                            st.rerun()
                with r_col5:
                    if st.button("👁️ View", key=f"rev_tab2_{r_id}", use_container_width=True):
                        st.session_state[f"show_review_{r_id}"] = not st.session_state.get(f"show_review_{r_id}", False)
                with r_col6:
                    if st.button("🗑️ Drop", key=f"del_tab2_{r_id}", use_container_width=True):
                        delete_record_from_db(r_id)
                        st.rerun()
                
                if st.session_state.get(f"show_review_{r_id}", False) and db_file_data:
                    st.info(f"Viewing Snapshot: {db_filename}")
                    rev_df = pd.read_excel(io.BytesIO(db_file_data))
                    st.dataframe(rev_df, width='stretch')
                    
                st.divider()