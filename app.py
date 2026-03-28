import streamlit as st
import pandas as pd
import io
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import PatternFill, Alignment, Border, Side

# ==========================================
# 1. CONFIGURATION
# ==========================================
st.set_page_config(page_title="Roster Automation", layout="wide")

DRIVE_FOLDER_ID = "1pcZWYGXCC1axVDXWtXp1YyQJ79WVeivr" 
TEMPLATE_FILE = "Template.xlsx"

SEQ = ['C', 'C', 'B', 'B', 'A', 'A', 'W/O']
MONTH_NAMES = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"]

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
# 2. WEB UI & SIDEBAR
# ==========================================
st.title("📅 Monthly Roster Generator")

service = get_drive_service()
latest_file = get_latest_roster(service)

if not latest_file:
    st.error("Upload a file to Drive folder first.")
    st.stop()

st.sidebar.header("1. Settings")
target_month_name = st.sidebar.selectbox("Month", MONTH_NAMES, index=3)
target_year = st.sidebar.number_input("Year", min_value=2024, max_value=2050, value=2026)
target_month_num = MONTH_NAMES.index(target_month_name) + 1
days_in_month = pd.Period(f'{target_year}-{target_month_num:02d}-01').days_in_month

st.sidebar.markdown("---")
st.sidebar.header("2. Leaves")
if 'leaves' not in st.session_state: st.session_state.leaves = {}

df_prev = pd.read_excel(latest_file[0], skiprows=2)
df_prev.columns = ['S No', 'NAME'] + [str(i) for i in range(1, len(df_prev.columns)-1)]
names_list = df_prev['NAME'].dropna().unique().tolist()

sel_name = st.sidebar.selectbox("Select Employee", names_list)
sel_days = st.sidebar.text_input("Enter Days (e.g., 5, 12)")

if st.sidebar.button("Register Leave"):
    st.session_state.leaves[sel_name] = [int(d.strip()) for d in sel_days.split(',') if d.strip().isdigit()]
    st.sidebar.success("Added!")

# ==========================================
# 3. GENERATION
# ==========================================
if st.button(f"Generate Roster ({days_in_month} Days)", type="primary"):
    with st.spinner("Fixing Borders and Merges..."):
        employees = df_prev.dropna(subset=['NAME'])['NAME'].tolist()
        emp_state = {name: get_state(row) for name, row in zip(employees, df_prev.dropna(subset=['NAME']).to_dict('records'))}
        
        roster = {emp: {d: None for d in range(1, days_in_month + 1)} for emp in employees}
        for d in range(1, days_in_month + 1):
            for emp in employees:
                if emp in st.session_state.leaves and d in st.session_state.leaves[emp]:
                    roster[emp][d] = 'L'
                else:
                    shift = SEQ[emp_state[emp]]
                    roster[emp][d] = shift
                    emp_state[emp] = (emp_state[emp] + 1) % 7

        wb = load_workbook(TEMPLATE_FILE)
        ws = wb.active
        
        # 1. CLEAN SLATE: Remove all old merges and borders to avoid AO Column Ghost
        for m in list(ws.merged_cells.ranges): ws.unmerge_cells(str(m))
        for r in ws.iter_rows(min_row=1, max_row=100, min_col=1, max_col=60):
            for cell in r:
                if cell.column > 2: # Keep S No and Name headers
                    cell.value = None
                    cell.fill = PatternFill(fill_type=None)
                cell.border = Border()

        # 2. STYLES
        thin = Side(border_style="thin", color="000000")
        thick_blue = Side(border_style="thick", color="0000FF")
        yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
        peach_fill = PatternFill(start_color="FFCC99", end_color="FFCC99", fill_type="solid")
        center = Alignment(horizontal='center', vertical='center')

        # 3. DYNAMIC MERGES
        start_totals = days_in_month + 3
        end_totals = start_totals + 7
        
        # Title
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=end_totals)
        ws.cell(row=1, column=1, value=f"DUTY ROSTER FOR THE MONTH OF {target_month_name[:3].upper()} {target_year}").alignment = center

        # Attendance Header
        ws.merge_cells(start_row=2, start_column=3, end_row=2, end_column=days_in_month+2)
        ws.cell(row=2, column=3, value="ATTENDANCE").alignment = center

        # Total Shifts Header
        ws.merge_cells(start_row=2, start_column=start_totals, end_row=2, end_column=end_totals)
        ws.cell(row=2, column=start_totals, value="TOTAL SHIFTS").alignment = center

        # 4. WRITE HEADERS
        for d in range(1, days_in_month + 1):
            ws.cell(row=3, column=d+2, value=d).alignment = center
            ws.column_dimensions[get_column_letter(d+2)].width = 5
            
        headers = ['TOTAL', 'A', 'B', 'C', 'W/O', 'X', 'L', 'G']
        for i, h in enumerate(headers):
            col = start_totals + i
            ws.cell(row=3, column=col, value=h).alignment = center
            ws.column_dimensions[get_column_letter(col)].width = 10 if h == 'TOTAL' else 5

        # 5. WRITE DATA & BORDERS
        for idx, emp in enumerate(employees):
            r = idx + 4
            ws.cell(row=r, column=1, value=idx+1)
            ws.cell(row=r, column=2, value=emp)
            
            for d in range(1, days_in_month + 1):
                ws.cell(row=r, column=d+2, value=roster[emp][d]).alignment = center
            
            # Column Mapping
            tot_ltr = get_column_letter(start_totals)
            last_d_ltr = get_column_letter(days_in_month + 2)
            shift_letters = [get_column_letter(start_totals + i) for i in range(1, 8)]

            # Formulas
            ws[f'{tot_ltr}{r}'] = f'=SUM({get_column_letter(start_totals+1)}{r}:{get_column_letter(start_totals+3)}{r})'
            ws[f'{get_column_letter(start_totals+1)}{r}'] = f'=COUNTIF(C{r}:{last_d_ltr}{r},"A*")'
            ws[f'{get_column_letter(start_totals+2)}{r}'] = f'=COUNTIF(C{r}:{last_d_ltr}{r},"B*")'
            ws[f'{get_column_letter(start_totals+3)}{r}'] = f'=COUNTIF(C{r}:{last_d_ltr}{r},"C*")'
            ws[f'{get_column_letter(start_totals+4)}{r}'] = f'=COUNTIF(C{r}:{last_d_ltr}{r},"W/O*")'
            ws[f'{get_column_letter(start_totals+5)}{r}'] = f'=COUNTIF(C{r}:{last_d_ltr}{r},"X*")'
            ws[f'{get_column_letter(start_totals+6)}{r}'] = f'=COUNTIF(C{r}:{last_d_ltr(r)},"L*")' # Fixed typo here
            ws[f'{get_column_letter(start_totals+6)}{r}'] = f'=COUNTIF(C{r}:{last_d_ltr}{r},"L*")'
            ws[f'{get_column_letter(start_totals+7)}{r}'] = f'=COUNTIF(C{r}:{last_d_ltr}{r},"G*")'

            # Paint
            ws[f'{tot_ltr}{r}'].fill = yellow_fill
            for ltr in shift_letters: ws[f'{ltr}{r}'].fill = peach_fill

            # Apply Borders to the entire row (1 to end_totals)
            for c in range(1, end_totals + 1):
                cell = ws.cell(row=r, column=c)
                right_side = thick_blue if c == end_totals else thin
                left_side = thick_blue if c == 1 else thin
                cell.border = Border(left=left_side, right=right_side, top=thin, bottom=thin)

        # Apply borders to headers (Row 1-3)
        for r_h in range(1, 4):
            for c_h in range(1, end_totals + 1):
                cell = ws.cell(row=r_h, column=c_h)
                right_side = thick_blue if c_h == end_totals else thin
                left_side = thick_blue if c_h == 1 else thin
                cell.border = Border(left=left_side, right=right_side, top=thin, bottom=thin)

        out = io.BytesIO()
        wb.save(out)
        st.balloons()
        st.download_button("Download Final Roster", out.getvalue(), f"ROSTER_{target_month_name}.xlsx")
