import streamlit as st
import pandas as pd
import io
import math
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import PatternFill, Alignment, Border, Side, Font

# ==========================================
# 1. CONFIGURATION
# ==========================================
st.set_page_config(page_title="Roster Automation", layout="wide")

DRIVE_FOLDER_ID = "1pcZWYGXCC1axVDXWtXp1YyQJ79WVeivr" 
TEMPLATE_FILE = "Template.xlsx"

# The Master Sequence
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
    for d in range(31, 0, -1):
        col = str(d)
        if col in row and pd.notna(row[col]):
            val = str(row[col]).strip().upper()
            if val in ['A', 'B', 'C', 'W/O', 'X', 'L']:
                if last_val is None: last_val = val
                elif prev_val is None: 
                    prev_val = val
                    break
    if last_val == 'C': return 2 if prev_val == 'C' else 1
    if last_val == 'B': return 4 if prev_val == 'B' else 3
    if last_val == 'A': return 6 if prev_val == 'A' else 5
    return 0

# ==========================================
# 2. WEB UI & DATA LOADING
# ==========================================
st.title("📅 Monthly Roster Generator")

service = get_drive_service()
latest_file = get_latest_roster(service)

if not latest_file:
    st.error("Baseline file not found.")
    st.stop()

st.sidebar.header("1. Settings")
target_month_name = st.sidebar.selectbox("Month", MONTH_NAMES, index=3)
target_year = st.sidebar.number_input("Year", min_value=2024, max_value=2050, value=2026)
target_month_num = MONTH_NAMES.index(target_month_name) + 1
days_in_month = pd.Period(f'{target_year}-{target_month_num:02d}-01').days_in_month

df_raw = pd.read_excel(latest_file[0], skiprows=2)
total_col_idx = next((i for i, c in enumerate(df_raw.columns) if "TOTAL" in str(c).upper()), None)

employees = []
emp_data_map = {}
prev_totals = {}

for _, row in df_raw.iterrows():
    raw_name = str(row.iloc[1]).strip()
    if raw_name and raw_name.lower() not in ["nan", "a", "b", "c", "total"]:
        employees.append(raw_name)
        emp_data_map[raw_name] = row
        val = row.iloc[total_col_idx] if total_col_idx is not None else 24
        prev_totals[raw_name] = val if pd.notna(val) else 24

st.sidebar.markdown("---")
st.sidebar.header("2. Leaves")
if 'leaves' not in st.session_state: st.session_state.leaves = {}
sel_name = st.sidebar.selectbox("Select Employee", employees)
sel_days = st.sidebar.text_input("Enter Days (e.g., 5, 12)")
if st.sidebar.button("Register Leave"):
    st.session_state.leaves[sel_name] = [int(d.strip()) for d in sel_days.split(',') if d.strip().isdigit()]
    st.sidebar.success("Added!")

# ==========================================
# 3. THE REWRITTEN ROTATION ENGINE
# ==========================================
if st.button(f"Generate Fair Roster ({days_in_month} Days)", type="primary"):
    with st.spinner("Balancing Rotation & Fairness..."):
        emp_state = {name: get_state(emp_data_map[name]) for name in employees}
        duty_history = {name: prev_totals[name] for name in employees}
        this_month_duties = {name: 0 for name in employees}
        c_counts = {name: 0 for name in employees}
        roster = {name: {d: None for d in range(1, days_in_month + 1)} for name in employees}
        
        c_limits = {name: min(8, math.ceil((days_in_month - len(st.session_state.leaves.get(name, []))) / 3)) for name in employees}

        for d in range(1, days_in_month + 1):
            counts = {'A': 0, 'B': 0, 'C': 0}
            available = []

            # 1. Identify Leaves
            for emp in employees:
                if emp in st.session_state.leaves and d in st.session_state.leaves[emp]:
                    roster[emp][d] = 'L'
                else:
                    available.append(emp)

            # 2. Determine "Preferred" shift based on SEQ
            # We sort 'available' by duty_history so those in "debt" get picked first if multiple people want a shift
            available.sort(key=lambda x: duty_history[x])

            unassigned = []
            
            # 3. Main Pass: Give people what they are supposed to have in the rotation
            for emp in available:
                pref = SEQ[emp_state[emp]]
                if pref in counts and counts[pref] < 8:
                    if pref == 'C' and c_counts[emp] >= c_limits[emp] and this_month_duties[emp] < 24:
                        # Skip C if they hit their monthly limit (unless they are low on duties)
                        unassigned.append(emp)
                    else:
                        roster[emp][d] = pref
                        counts[pref] += 1
                        duty_history[emp] += 1
                        this_month_duties[emp] += 1
                        if pref == 'C': c_counts[emp] += 1
                        emp_state[emp] = (emp_state[emp] + 1) % 7
                elif pref == 'W/O':
                    roster[emp][d] = 'W/O'
                    emp_state[emp] = 0 # Restart cycle after W/O
                else:
                    # Shift full or not preferred
                    unassigned.append(emp)

            # 4. Fill Gaps: If a shift is less than 8, fill it with unassigned people
            for s in ['C', 'B', 'A']:
                while counts[s] < 8 and unassigned:
                    # Pick person from unassigned with lowest history
                    unassigned.sort(key=lambda x: duty_history[x])
                    emp = unassigned.pop(0)
                    roster[emp][d] = s
                    counts[s] += 1
                    duty_history[emp] += 1
                    this_month_duties[emp] += 1
                    if s == 'C': c_counts[emp] += 1
                    # Update state to the shift they were forced into
                    emp_state[emp] = (SEQ.index(s) + 1) % 7

            # 5. Final W/O: Everyone else left gets a forced day off (state doesn't advance)
            for emp in unassigned:
                roster[emp][d] = 'W/O'

        # ==========================================
        # 4. EXCEL WRITING
        # ==========================================
        wb = load_workbook(TEMPLATE_FILE)
        ws = wb.active
        for m in list(ws.merged_cells.ranges): ws.unmerge_cells(str(m))
        for r in ws.iter_rows(min_row=1, max_row=120, min_col=1, max_col=65):
            for cell in r:
                if cell.column > 2: cell.value = None
                cell.fill = PatternFill(fill_type=None); cell.border = Border()

        thin = Side(border_style="thin"); thick_blue = Side(border_style="thick", color="0000FF")
        all_border = Border(left=thin, right=thin, top=thin, bottom=thin)
        yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
        peach_fill = PatternFill(start_color="FFCC99", end_color="FFCC99", fill_type="solid")
        center = Alignment(horizontal='center', vertical='center')
        
        start_totals = days_in_month + 3
        end_totals = start_totals + 7
        ws.column_dimensions['A'].width = 6.43
        ws.column_dimensions['B'].width = 25

        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=end_totals)
        t_cell = ws.cell(row=1, column=1, value=f"DUTY ROSTER FOR THE MONTH OF {target_month_name[:3].upper()} {target_year}")
        t_cell.font = Font(bold=True, size=20); t_cell.alignment = center
        
        ws.merge_cells('A2:A3'); ws['A2'] = "S No"; ws['A2'].font = Font(bold=True, size=16); ws['A2'].alignment = center
        ws.merge_cells('B2:B3'); ws['B2'] = "NAME"; ws['B2'].font = Font(bold=True, size=16); ws['B2'].alignment = center
        ws.merge_cells(start_row=2, start_column=3, end_row=2, end_column=days_in_month+2)
        ws.cell(row=2, column=3, value="ATTENDANCE").font = Font(bold=True, size=16); ws.cell(row=2, column=3).alignment = center
        ws.merge_cells(start_row=2, start_column=start_totals, end_row=2, end_column=end_totals)
        ws.cell(row=2, column=start_totals, value="TOTAL SHIFTS").font = Font(bold=True, size=16); ws.cell(row=2, column=start_totals).alignment = center

        for d in range(1, days_in_month + 1):
            ws.cell(row=3, column=d+2, value=d).alignment = center
            ws.column_dimensions[get_column_letter(d+2)].width = 5
        for i, h in enumerate(['TOTAL', 'A', 'B', 'C', 'W/O', 'X', 'L', 'G']):
            col = start_totals + i
            ws.cell(row=3, column=col, value=h).alignment = center
            ws.column_dimensions[get_column_letter(col)].width = 10 if h == 'TOTAL' else 5

        for idx, emp in enumerate(employees):
            r = idx + 4
            ws.cell(row=r, column=1, value=idx+1)
            ws.cell(row=r, column=2, value=emp)
            for d in range(1, days_in_month + 1):
                ws.cell(row=r, column=d+2, value=roster[emp][d]).alignment = center
            
            t_ltr = get_column_letter(start_totals)
            last_d_ltr = get_column_letter(days_in_month + 2)
            ws[f'{t_ltr}{r}'] = f'=SUM({get_column_letter(start_totals+1)}{r}:{get_column_letter(start_totals+3)}{r})'
            for i_h, h_code in enumerate(['A*', 'B*', 'C*', 'W/O*', 'X*', 'L*', 'G*']):
                ws.cell(row=r, column=start_totals+1+i_h, value=f'=COUNTIF(C{r}:{last_d_ltr}{r},"{h_code}")')

            ws[f'{t_ltr}{r}'].fill = yellow_fill; ws[f'{t_ltr}{r}'].font = Font(bold=True)
            for cp in range(start_totals+1, end_totals+1): ws.cell(row=r, column=cp).fill = peach_fill
            for c in range(1, end_totals + 1):
                ws.cell(row=r, column=c).border = Border(left=thick_blue if c==1 else thin, right=thick_blue if c==end_totals else thin, top=thin, bottom=thin if idx<len(employees)-1 else thick_blue)

        # Summary Table
        s_row = len(employees) + 6
        for i, stype in enumerate(["A", "B", "C"]):
            curr_r = s_row + i
            ws.cell(row=curr_r, column=2, value=stype).font = Font(bold=True)
            ws.cell(row=curr_r, column=2).fill = yellow_fill; ws.cell(row=curr_r, column=2).alignment = center; ws.cell(row=curr_r, column=2).border = all_border
            for d in range(1, days_in_month + 1):
                col_i = d + 2
                cltr = get_column_letter(col_i)
                ws.cell(row=curr_r, column=col_i, value=f'=COUNTIF({cltr}4:{cltr}{len(employees)+3},"{stype}*")').fill = yellow_fill
                ws.cell(row=curr_r, column=col_i).alignment = center; ws.cell(row=curr_r, column=col_i).border = all_border

        out = io.BytesIO()
        wb.save(out)
        st.balloons()
        st.download_button("Download Fair Balanced Roster", out.getvalue(), f"ROSTER_{target_month_name}.xlsx")
