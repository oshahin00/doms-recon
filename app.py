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
    
    # --- ADVANCED FILTERING: DOMS & PUMPS vs RFID ---
    res_note = filtered_df['Resolution Note'].fillna('').str.lower()
    summary = filtered_df['Summary'].fillna('').str.lower()
    desc = filtered_df['Description'].fillna('').str.lower()
    
    # 1. Check if the Resolution Note explicitly fixes a DOMS or Pump issue
    is_doms_pump_res = res_note.str.contains('doms|pump', regex=True)
    
    # 2. Check if the Resolution Note explicitly blames an RFID issue (False Positive)
    is_rfid_res = res_note.str.contains('rfid', regex=True)
    
    # 3. Check if the Summary or Description mentioned DOMS or Pump initially
    is_doms_pump_initial = summary.str.contains('doms|pump', regex=True) | desc.str.contains('doms|pump', regex=True)
    
    # Keep the row IF: 
    # The resolution explicitly says doms/pump, OR 
    # (It was reported as doms/pump initially AND the resolution didn't explicitly blame RFID)
    target_mask = is_doms_pump_res | (is_doms_pump_initial & ~is_rfid_res)
    
    filtered_df = filtered_df[target_mask]
    
    # --- COLUMN SELECTION & ORDERING ---
    final_cols = [
        'Station No.', 'Number', 'Priority', 'Created', 'Summary', 
        'Resolution Note', 'Status', 'Cause Category', 'Cause Code', 
        'DOMS Model', 'Date'
    ]
    
    # Ensure only existing columns are selected just in case to prevent crashes
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
    
    return output, len(final_df)

# --- Web Interface ---
st.set_page_config(page_title="Data Merger", layout="centered")

st.title("DOMS Data Merge & Filter Tool")
st.write("Upload your incident report and rollout schedule to generate a formatted comparison sheet.")

file1 = st.file_uploader("Upload First Sheet (Incidents - CSV or XLSX)", type=["csv", "xlsx"])
file2 = st.file_uploader("Upload Second Sheet (DOMS Rollout - XLSX)", type=["xlsx"])

if file1 and file2:
    st.success("Files uploaded successfully. Processing...")
    
    try:
        excel_data, row_count = process_files(file1, file2)
        
        if row_count == 0:
            st.warning("No records found matching the DOMS/PUMPS criteria or date filters.")
        else:
            st.write(f"**Processing Complete:** Found {row_count} matching records related to DOMS/Pumps.")
            
        # Provide the download button
        st.download_button(
            label="Download Formatted Excel File",
            data=excel_data,
            file_name="Comparison_Result.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        st.error(f"An error occurred during processing. Please verify column names match the requirements. Error: {e}")