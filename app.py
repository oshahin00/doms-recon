import streamlit as st
import pandas as pd
import re
import io
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl import Workbook

def extract_first_number(area_str):
    if not isinstance(area_str, str):
        return None
    match = re.search(r'\d+', area_str)
    return int(match.group()) if match else None

def process_files(file1, file2):
    # Read the first file (handling both CSV and Excel just in case)
    if file1.name.endswith('.csv'):
        df1 = pd.read_csv(file1)
    else:
        df1 = pd.read_excel(file1)
        
    df1.rename(columns={'Resolution Notes': 'Resolution Note'}, inplace=True)
    
    # Safely initialize columns if they don't exist to prevent KeyErrors
    for col in ['Description', 'Cause Category', 'Cause Code']:
        if col not in df1.columns:
            df1[col] = ''
    
    # Read the second file
    xls = pd.ExcelFile(file2)
    df2 = pd.read_excel(file2, sheet_name=xls.sheet_names[0])
    df2.columns = df2.columns.str.strip()
    df2['Station No.'] = pd.to_numeric(df2['Station No.'], errors='coerce').astype('Int64')
    
    # Extract station number from the Area column
    df1['Extracted_Station_No'] = df1['Area'].apply(extract_first_number).astype('Int64')
    
    # Merge
    merged_df = pd.merge(df1, df2, left_on='Extracted_Station_No', right_on='Station No.', how='inner')
    
    # Date Filtering (Created >= Date)
    merged_df['Created_dt'] = pd.to_datetime(merged_df['Created'], errors='coerce')
    merged_df['Date_dt'] = pd.to_datetime(merged_df['Date'], errors='coerce')
    filtered_df = merged_df[merged_df['Created_dt'] >= merged_df['Date_dt']]
    
    # --- STRICT ADVANCED FILTERING ---
    res_note = filtered_df['Resolution Note'].fillna('').str.lower()
    summary = filtered_df['Summary'].fillna('').str.lower()
    desc = filtered_df['Description'].fillna('').str.lower()
    
    # 1. Identify ANY record that mentions 'rfid'
    has_rfid = res_note.str.contains('rfid') | summary.str.contains('rfid') | desc.str.contains('rfid')
    
    # 2. Identify ANY record that mentions 'doms' or 'pump'
    has_doms_pump = res_note.str.contains('doms|pump') | summary.str.contains('doms|pump') | desc.str.contains('doms|pump')
    
    # 3. Apply the strict rule
    target_mask = has_doms_pump & ~has_rfid
    filtered_df = filtered_df[target_mask]
    
    # --- COLUMN SELECTION & ORDERING ---
    final_cols = [
        'Station No.', 'Number', 'Priority', 'Created', 'Summary', 
        'Resolution Note', 'Status', 'Cause Category', 'Cause Code', 
        'DOMS Model', 'Date'
    ]
    
    final_cols = [col for col in final_cols if col in filtered_df.columns]
    final_df = filtered_df[final_cols]
    
    # Create Excel file in memory
    output = io.BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.title = "Comparison Results"
    
    for r in dataframe_to_rows(final_df, index=False, header=True):
        ws.append(r)
        
    # Styling
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
    
    # Notice we now return the dataframe as well to build the live preview
    return output, final_df 

# ==========================================
# --- ENTERPRISE WEB INTERFACE & UI/UX ---
# ==========================================

# 1. Page Configuration (Must be the first Streamlit command)
st.set_page_config(page_title="DOMS Recon", page_icon="⚙️", layout="wide")

# 2. Custom CSS for button styling
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
    </style>
""", unsafe_allow_html=True)

# 3. Main Header and Instructions
st.title("⚙️ DOMS Data Merge & Filter Tool")
st.markdown("Automated incident reconciliation and filtering for DOMS and Pump hardware tickets.")

with st.expander("ℹ️ Operating Instructions & Filtering Rules"):
    st.markdown("""
    * **Input:** Requires the raw Incidents Report and the DOMS Rollout schedule.
    * **Logic:** Matches Station Numbers and filters out records where the Incident Creation Date precedes the Rollout Date.
    * **Strict Filtering:** Automatically isolates tickets explicitly mentioning `DOMS` or `Pumps` while systematically rejecting any ticket referencing `RFID`.
    """)

# 4. Sidebar for Inputs
with st.sidebar:
    st.header("📂 Data Input")
    file1 = st.file_uploader("1. Incidents Report (CSV/XLSX)", type=["csv", "xlsx"])
    file2 = st.file_uploader("2. DOMS Rollout Schedule (XLSX)", type=["xlsx"])
    
    st.markdown("---")
    st.caption("Environment: Node 2 (App Server)")

# 5. Execution Logic
if file1 and file2:
    with st.spinner('Reconciling datasets and applying exclusion rules...'):
        try:
            excel_data, final_df = process_files(file1, file2)
            row_count = len(final_df)
            
            if row_count == 0:
                st.warning("⚠️ No records found matching the DOMS/PUMPS criteria, or all records were excluded by the RFID filter.")
            else:
                st.success(f"✅ Processing Complete: Successfully isolated {row_count} validated records.")
                
                # Split the UI into two columns for the preview and the download button
                col1, col2 = st.columns([4, 1])
                
                with col1:
                    st.subheader("📊 Live Data Preview")
                    # Display the first 15 rows interactively
                    st.dataframe(final_df.head(15), use_container_width=True)
                    st.caption(f"Showing top 15 of {row_count} records. Download the Excel file to view the complete dataset.")
                    
                with col2:
                    st.subheader("Export")
                    st.download_button(
                        label="📥 Download Report",
                        data=excel_data,
                        file_name="DOMS_Reconciliation_Result.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                    
        except Exception as e:
            st.error(f"❌ A structural error occurred during processing. Please verify column names match the system requirements. Error: {e}")
else:
    # Default landing screen
    st.info("👈 Please upload both datasets in the sidebar menu to initiate the reconciliation process.")