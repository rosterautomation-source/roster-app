import streamlit as st
import pandas as pd
import io
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload

# ==========================================
# 1. CONFIGURATION
# ==========================================
st.set_page_config(page_title="Roster Automation", layout="wide")

DRIVE_FOLDER_ID = "1pcZWYGXCC1axVDXWtXp1YyQJ79WVeivr" 
TEMPLATE_FILE = "Template.xlsx"

SEQ = ['C', 'C', 'B', 'B', 'A', 'A', 'W/O']
MONTH_NAMES = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"]

# ==========================================
# 2. GOOGLE DRIVE CONNECTION (WITH WORKSPACE BYPASS)
# ==========================================
def get_drive_service():
    creds = service_account.Credentials.from_service_account_info(
        st.secrets["gcp_service_account"],
        scopes=["https://www.googleapis.com/auth/drive"]
    )
    return build('drive', 'v3', credentials=creds)

def get_latest_roster(service):
    query = f"'{DRIVE_FOLDER_ID}' in parents and mimeType='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'"
    # Added includeItemsFromAllDrives and supportsAllDrives to bypass Workspace restrictions
    results = service.files().list(
        q=query, 
        orderBy="createdTime desc", 
        pageSize=1,
        includeItemsFromAllDrives=True,
        supportsAllDrives=True
    ).execute()
    
    items = results.get('files', [])
    if not items: return None
    request = service.files().get_media(fileId=items[0]['id'])
    return io.BytesIO(request.execute()), items[0]['name']

def upload_to_drive(service, file_stream, filename):
    file_metadata = {'name': filename, 'parents': [DRIVE_FOLDER_ID]}
    media = MediaIoBaseUpload(file_stream, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    # Added supportsAllDrives=True to bypass Workspace upload restrictions
    service.files().create(
        body=file_metadata, 
        media_body=media,
        supportsAllDrives=True
    ).execute()

# ==========================================
# 3. ROSTER ENGINE
# ==========================================
def get_state(row):
    last_val = None
    prev_val = None
    for d in range(31, 20, -1):
        col = str(d)
        if col in row and pd.notna(row[col]):
            if last_val is None: last_val = row[col]
            elif prev_val is None: 
                prev_val = row[col]
                break
    
    if last_val == 'C': return 2 if prev_val == 'C' else 1
    if last_val == 'B': return 4 if prev_val == 'B' else 3
    if last_val == 'A': return 6 if prev_val == 'A' else 5
    return 0

# ==========================================
# 4. WEB INTERFACE
# ==========================================
st.title("📅 Monthly Roster Generator")

service = get_drive_service()
latest_file = get_latest_roster(service)

if not latest_file:
    st.error("Missing baseline! Please upload an initial Excel file to your Google Drive folder first.")
    st.stop()

st.info(f"Connected to Google Drive. Last file found: **{latest_file[1]}**")

# --- SIDEBAR CONTROLS ---
st.sidebar.header("1. Roster Settings")
target_month_name = st.sidebar.selectbox("Target Month", MONTH_NAMES, index=3) # Default April
target_year = st.sidebar.number_input("Target Year", min_value=2024, max_value=2050, value=2026)

target_month_num = MONTH_NAMES.index(target_month_name) + 1
days_in_month = pd.Period(f'{target_year}-{target_month_num:02d}-01').days_in_month

st.sidebar.markdown("---")
st.sidebar.header("2. Add Planned Leaves")
if 'leaves' not in st.session_state: st.session_state.leaves = {}

df_prev = pd.read_excel(latest_file[0], skiprows=2)
df_prev.columns = ['S No', 'NAME'] + [str(i) for i in range(1, len(df_prev.columns)-1)]
names = df_prev['NAME'].dropna().unique().tolist()

sel_name = st.sidebar.selectbox("Employee", names)
sel_days = st.sidebar.text_input("Days (e.g. 5, 6, 7)")

if st.sidebar.button("Add Leave"):
    day_list = [int(d.strip()) for d in sel_days.split(',') if d.strip().isdigit()]
    st.session_state.leaves[sel_name] = day_list

if st.session_state.leaves:
    st.sidebar.write("Currently Registered Leaves:", st.session_state.leaves)

# --- GENERATE BUTTON ---
if st.button(f"Generate Roster for {target_month_name} {target_year} ({days_in_month} Days)", type="primary"):
    with st.spinner("Calculating shifts and generating dynamic Excel..."):
        
        df_clean = df_prev.dropna(subset=['NAME']).reset_index(drop=True)
        employees = df_clean['NAME'].tolist()
        num_employees = len(employees)
        emp_state = {row['NAME']: get_state(row) for _, row in df_clean.iterrows()}
        
        roster = {emp: {d: None for d in range(1, days_in_month + 1)} for emp in employees}

        # Apply Leaves
        for emp, days in st.session_state.leaves.items():
            days.sort()
            if len(days) <= 2:
                if days[0] > 1: roster[emp][days[0]-1] = 'X'
            else:
                if days[0] > 2: roster[emp][days[0]-2] = 'W/O'
                if days[0] > 1: roster[emp][days[0]-1] = 'X'
            for d in days: 
                if d <= days_in_month: roster[emp][d] = 'L'
            if days[-1] < days_in_month: 
                roster[emp][days[-1]+1] = 'C'
                emp_state[emp] = 1 

        # Fill Shifts
        for d in range(1, days_in_month + 1):
            assigned = [e for e in employees if roster[e][d] is not None]
            c_rem = 8 - sum(1 for e in assigned if roster[e][d] == 'C')
            targ = {'C': c_rem, 'B': 8, 'A': 8}
            avail = [e for e in employees if roster[e][d] is None]
            
            for s in ['C', 'B', 'A']:
                for e in avail[:]:
                    if targ[s] > 0 and SEQ[emp_state[e]] == s:
                        roster[e][d], targ[s] = s, targ[s]-1
                        emp_state[e] = (emp_state[e]+1)%7
                        avail.remove(e)
            
            for s in ['C', 'B', 'A']:
                while targ[s] > 0 and avail:
                    e = avail.pop(0)
                    roster[e][d], targ[s] = s, targ[s]-1
                    emp_state[e] = (SEQ.index(s)+1)%7

            for e in avail:
                if SEQ[emp_state[e]] == 'W/O':
                    roster[e][d], emp_state[e] = 'W/O', 0
                else: roster[e][d] = 'X'

        # ==========================================
        # 5. DYNAMIC EXCEL WRITING
        # ==========================================
        wb = load_workbook(TEMPLATE_FILE)
        ws = wb.active
        
        month_abbr = target_month_name[:3].upper()
        ws['A1'] = f"DUTY ROSTER FOR THE MONTH OF {month_abbr} {target_year}"
        
        for d in range(1, days_in_month + 1):
            ws.cell(row=3, column=2+d, value=d)
            
        calc_col_start = 3 + days_in_month
        calc_headers = ['TOTAL', 'A', 'B', 'C', 'W/O', 'X', 'L', 'G']
        for idx, header in enumerate(calc_headers):
            ws.cell(row=3, column=calc_col_start + idx, value=header)

        last_day_col = get_column_letter(2 + days_in_month)
        tot_col = get_column_letter(calc_col_start)
        a_col = get_column_letter(calc_col_start + 1)
        b_col = get_column_letter(calc_col_start + 2)
        c_col = get_column_letter(calc_col_start + 3)
        wo_col = get_column_letter(calc_col_start + 4)
        x_col = get_column_letter(calc_col_start + 5)
        l_col = get_column_letter(calc_col_start + 6)
        g_col = get_column_letter(calc_col_start + 7)

        last_employee_row = 3 + num_employees

        for i, emp in enumerate(employees):
            r = 4 + i
            ws.cell(row=r, column=1, value=i+1)
            ws.cell(row=r, column=2, value=emp)
            
            for d in range(1, days_in_month + 1): 
                ws.cell(row=r, column=2+d, value=roster[emp][d])
                
            ws[f'{tot_col}{r}'] = f'=SUM({a_col}{r}:{c_col}{r})'
            ws[f'{a_col}{r}'] = f'=COUNTIF(C{r}:{last_day_col}{r},"A*")'
            ws[f'{b_col}{r}'] = f'=COUNTIF(C{r}:{last_day_col}{r},"B*")'
            ws[f'{c_col}{r}'] = f'=COUNTIF(C{r}:{last_day_col}{r},"C*")'
            ws[f'{wo_col}{r}'] = f'=COUNTIF(C{r}:{last_day_col}{r},"W/O*")'
            ws[f'{x_col}{r}'] = f'=COUNTIF(C{r}:{last_day_col}{r},"X*")'
            ws[f'{l_col}{r}'] = f'=COUNTIF(C{r}:{last_day_col}{r},"L*")'
            ws[f'{g_col}{r}'] = f'=COUNTIF(C{r}:{last_day_col}{r},"G*")'

        bottom_start = last_employee_row + 2
        ws[f'B{bottom_start}'] = "A"
        ws[f'B{bottom_start+1}'] = "B"
        ws[f'B{bottom_start+2}'] = "C"
        
        for d in range(1, days_in_month + 1):
            col_letter = get_column_letter(2 + d)
            ws[f'{col_letter}{bottom_start}'] = f'=COUNTIF({col_letter}4:{col_letter}{last_employee_row},"A*")'
            ws[f'{col_letter}{bottom_start+1}'] = f'=COUNTIF({col_letter}4:{col_letter}{last_employee_row},"B*")'
            ws[f'{col_letter}{bottom_start+2}'] = f'=COUNTIF({col_letter}4:{col_letter}{last_employee_row},"C*")'

        out_stream = io.BytesIO()
        wb.save(out_stream)
        out_stream.seek(0)
        
        new_name = f"ROSTER OF {month_abbr} {target_year} - Calculation.xlsx"
        upload_to_drive(service, out_stream, new_name)
        
        st.balloons()
        st.success(f"🔥 Successfully saved to Drive as: **{new_name}**")
        st.download_button("Download Now", data=out_stream, file_name=new_name)
