import streamlit as st
import pandas as pd
import io
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import PatternFill, Alignment

# ==========================================
# 1. CONFIGURATION
# ==========================================
st.set_page_config(page_title="Roster Automation", layout="wide")

DRIVE_FOLDER_ID = "1pcZWYGXCC1axVDXWtXp1YyQJ79WVeivr" 
TEMPLATE_FILE = "Template.xlsx"

SEQ = ['C', 'C', 'B', 'B', 'A', 'A', 'W/O']
MONTH_NAMES = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"]

# ==========================================
# 2. GOOGLE DRIVE CONNECTION
# ==========================================
def get_drive_service():
    from google.oauth2 import service_account
    from googleapiclient.discovery import build
    creds = service_account.Credentials.from_service_account_info(
        st.secrets["gcp_service_account"],
        scopes=["https://www.googleapis.com/auth/drive"]
    )
    return build('drive', 'v3', credentials=creds)

def get_latest_roster(service):
    query = f"'{DRIVE_FOLDER_ID}' in parents and mimeType='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'"
    results = service.files().list(q=query, orderBy="createdTime desc", pageSize=1, includeItemsFromAllDrives=True, supportsAllDrives=True).execute()
    items = results.get('files', [])
    if not items: return None
    request = service.files().get_media(fileId=items[0]['id'])
    return io.BytesIO(request.execute()), items[0]['name']

# ==========================================
# 3. ENGINE
# ==========================================
def get_state(row):
    last_val, prev_val = None, None
    for d in range(31, 1, -1):
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
# 4. WEB UI
# ==========================================
st.title("📅 Monthly Roster Generator")

service = get_drive_service()
latest_file = get_latest_roster(service)

if not latest_file:
    st.error("Upload a file to Drive folder first.")
    st.stop()

st.sidebar.header("Settings")
target_month_name = st.sidebar.selectbox("Month", MONTH_NAMES, index=3)
target_year = st.sidebar.number_input("Year", min_value=2024, max_value=2050, value=2026)
target_month_num = MONTH_NAMES.index(target_month_name) + 1
days_in_month = pd.Period(f'{target_year}-{target_month_num:02d}-01').days_in_month

if 'leaves' not in st.session_state: st.session_state.leaves = {}
df_prev = pd.read_excel(latest_file[0], skiprows=2)
df_prev.columns = ['S No', 'NAME'] + [str(i) for i in range(1, len(df_prev.columns)-1)]
names = df_prev['NAME'].dropna().unique().tolist()

if st.button(f"Generate Roster ({days_in_month} Days)", type="primary"):
    with st.spinner("Writing Excel..."):
        employees = df_prev.dropna(subset=['NAME'])['NAME'].tolist()
        emp_state = {row['NAME']: get_state(row) for _, row in df_prev.dropna(subset=['NAME']).iterrows()}
        roster = {emp: {d: None for d in range(1, days_in_month + 1)} for emp in employees}

        for d in range(1, days_in_month + 1):
            for emp in employees:
                shift = SEQ[emp_state[emp]]
                roster[emp][d] = shift
                emp_state[emp] = (emp_state[emp] + 1) % 7

        wb = load_workbook(TEMPLATE_FILE)
        ws = wb.active
        
        # --- FIX: HARD ERASE GHOST DATA ---
        # Instead of looping, we clear a specific range to avoid AttributeError
        for row in ws.iter_rows(min_row=2, max_row=50, min_col=3, max_col=50):
            for cell in row:
                cell.value = None
                cell.fill = PatternFill(fill_type=None)

        # Title
        ws['B1'] = f"DUTY ROSTER FOR THE MONTH OF {target_month_name[:3].upper()} {target_year}"
        
        # Colors
        yellow = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
        peach = PatternFill(start_color="FFCC99", end_color="FFCC99", fill_type="solid")

        # Set Day Headers
        for d in range(1, days_in_month + 1):
            col = d + 2
            ws.cell(row=3, column=col, value=d)
            ws.column_dimensions[get_column_letter(col)].width = 5

        # --- FIX: EXACT TOTAL SHIFT POSITIONING ---
        start_totals = days_in_month + 3
        headers = ['TOTAL', 'A', 'B', 'C', 'W/O', 'X', 'L', 'G']
        
        ws.merge_cells(start_row=2, start_column=start_totals, end_row=2, end_column=start_totals+7)
        ws.cell(row=2, column=start_totals, value="TOTAL SHIFTS").alignment = Alignment(horizontal='center')

        for i, h in enumerate(headers):
            col = start_totals + i
            ws.cell(row=3, column=col, value=h)
            ws.column_dimensions[get_column_letter(col)].width = 10 if h == 'TOTAL' else 5

        # Write Rows
        for idx, emp in enumerate(employees):
            r = idx + 4
            ws.cell(row=r, column=1, value=idx+1)
            ws.cell(row=r, column=2, value=emp)
            
            for d in range(1, days_in_month + 1):
                ws.cell(row=r, column=d+2, value=roster[emp][d])
            
            # Column Letters for Formulas
            last_d_ltr = get_column_letter(days_in_month + 2)
            tot_ltr = get_column_letter(start_totals)
            a_ltr = get_column_letter(start_totals + 1)
            b_ltr = get_column_letter(start_totals + 2)
            c_ltr = get_column_letter(start_totals + 3)
            wo_ltr = get_column_letter(start_totals + 4)
            x_ltr = get_column_letter(start_totals + 5)
            l_ltr = get_column_letter(start_totals + 6)
            g_ltr = get_column_letter(start_totals + 7)

            ws[f'{tot_ltr}{r}'] = f'=SUM({a_ltr}{r}:{c_ltr}{r})'
            ws[f'{a_ltr}{r}'] = f'=COUNTIF(C{r}:{last_d_ltr}{r},"A*")'
            ws[f'{b_ltr}{r}'] = f'=COUNTIF(C{r}:{last_d_ltr}{r},"B*")'
            ws[f'{c_ltr}{r}'] = f'=COUNTIF(C{r}:{last_d_ltr}{r},"C*")'
            ws[f'{wo_ltr}{r}'] = f'=COUNTIF(C{r}:{last_d_ltr}{r},"W/O*")'
            ws[f'{x_ltr}{r}'] = f'=COUNTIF(C{r}:{last_d_ltr}{r},"X*")'
            ws[f'{l_ltr}{r}'] = f'=COUNTIF(C{r}:{last_d_ltr}{r},"L*")'
            ws[f'{g_ltr}{r}'] = f'=COUNTIF(C{r}:{last_d_ltr}{r},"G*")'

            # Final Paint
            ws[f'{tot_ltr}{r}'].fill = yellow
            for col_ltr in [a_ltr, b_ltr, c_ltr, wo_ltr, x_ltr, l_ltr, g_ltr]:
                ws[f'{col_ltr}{r}'].fill = peach

        # Save
        out = io.BytesIO()
        wb.save(out)
        st.balloons()
        st.download_button("Download Final Roster", out.getvalue(), f"ROSTER_{target_month_name}.xlsx")
