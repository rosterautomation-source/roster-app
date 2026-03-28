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
            if last_val is None: last_val = str(row[col]).strip()
            elif prev_val is None: 
                prev_val = str(row[col]).strip()
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
    st.error("Upload a baseline file to Drive first.")
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
names_list = [n for n in df_prev['NAME'].dropna().unique().tolist() if str(n).strip() not in ["A", "B", "C"]]

sel_name = st.sidebar.selectbox("Select Employee", names_list)
sel_days = st.sidebar.text_input("Enter Days (e.g., 5, 12)")

if st.sidebar.button("Register Leave"):
    st.session_state.leaves[sel_name] = [int(d.strip()) for d in sel_days.split(',') if d.strip().isdigit()]
    st.sidebar.success("Added!")

# ==========================================
# 3. ADVANCED BALANCING ENGINE
# ==========================================
if st.button(f"Generate Roster ({days_in_month} Days)", type="primary"):
    with st.spinner("Calculating fair C-shift distribution..."):
        employees = names_list
        emp_state = {name: get_state(row) for name, row in zip(employees, df_prev[df_prev['NAME'].isin(employees)].to_dict('records'))}
        
        # Track counts
        duty_total = {emp: 0 for emp in employees}
        c_counts = {emp: 0 for emp in employees}
        roster = {emp: {d: None for d in range(1, days_in_month + 1)} for emp in employees}

        # Calculate C-Limits per person (Pro-rated for leaves)
        c_limits = {}
        for emp in employees:
            present_days = days_in_month - len(st.session_state.leaves.get(emp, []))
            # Standard: 8 C's for 24 duties. So Max C = (Present Days / 3)
            c_limits[emp] = math.ceil(present_days / 3)

        for d in range(1, days_in_month + 1):
            daily_targets = {'C': 8, 'B': 8, 'A': 8}
            available = []

            for emp in employees:
                if emp in st.session_state.leaves and d in st.session_state.leaves[emp]:
                    roster[emp][d] = 'L'
                else:
                    available.append(emp)

            # Sort: Prioritize those with lowest total duties
            available.sort(key=lambda x: duty_total[x])

            # --- STEP 1: FILL C SHIFTS (The constrained one) ---
            c_assigned = 0
            # Priority 1: Needs C + Prefers C + Under Limit
            for emp in available[:]:
                if c_assigned < 8 and SEQ[emp_state[emp]] == 'C' and (c_counts[emp] < c_limits[emp] or duty_total[emp] >= 24):
                    roster[emp][d] = 'C'
                    c_counts[emp] += 1
                    duty_total[emp] += 1
                    c_assigned += 1
                    emp_state[emp] = (emp_state[emp] + 1) % 7
                    available.remove(emp)

            # Priority 2: Anyone under Limit
            for emp in available[:]:
                if c_assigned < 8 and (c_counts[emp] < c_limits[emp] or duty_total[emp] >= 24):
                    roster[emp][d] = 'C'
                    c_counts[emp] += 1
                    duty_total[emp] += 1
                    c_assigned += 1
                    emp_state[emp] = (SEQ.index('C') + 1) % 7
                    available.remove(emp)
            
            # Priority 3: Forced C (If still not 8)
            for emp in available[:]:
                if c_assigned < 8:
                    roster[emp][d] = 'C'
                    c_counts[emp] += 1
                    duty_total[emp] += 1
                    c_assigned += 1
                    available.remove(emp)

            # --- STEP 2: FILL B AND A (Fair distribution) ---
            for s in ['B', 'A']:
                s_assigned = 0
                # Try preferred first
                for emp in available[:]:
                    if s_assigned < 8 and SEQ[emp_state[emp]] == s:
                        roster[emp][d] = s
                        duty_total[emp] += 1
                        s_assigned += 1
                        emp_state[emp] = (emp_state[emp] + 1) % 7
                        available.remove(emp)
                # Fill remaining
                while s_assigned < 8 and available:
                    emp = available.pop(0)
                    roster[emp][d] = s
                    duty_total[emp] += 1
                    s_assigned += 1
                    emp_state[emp] = (SEQ.index(s) + 1) % 7

            # --- STEP 3: W/O ---
            for emp in available:
                roster[emp][d] = 'W/O'
                emp_state[emp] = 0

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

        # Title/Headers
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=end_totals)
        t_cell = ws.cell(row=1, column=1, value=f"DUTY ROSTER FOR THE MONTH OF {target_month_name[:3].upper()} {target_year}")
        t_cell.font = Font(bold=True, size=20); t_cell.alignment = center
        
        ws.merge_cells('A2:A3'); ws['A2'] = "S No"; ws['A2'].font = Font(bold=True, size=16); ws['A2'].alignment = center
        ws.merge_cells('B2:B3'); ws['B2'] = "NAME"; ws['B2'].font = Font(bold=True, size=16); ws['B2'].alignment = center
        ws.merge_cells(start_row=2, start_column=3, end_row=2, end_column=days_in_month+2)
        ws.cell(row=2, column=3, value="ATTENDANCE").font = Font(bold=True, size=16); ws.cell(row=2, column=3).alignment = center
        ws.merge_cells(start_row=2, start_column=start_totals, end_row=2, end_column=end_totals)
        ws.cell(row=2, column=start_totals, value="TOTAL SHIFTS").font = Font(bold=True, size=16); ws.cell(row=2, column=start_totals).alignment = center

        # Header Details
        for d in range(1, days_in_month + 1):
            ws.cell(row=3, column=d+2, value=d).alignment = center
            ws.column_dimensions[get_column_letter(d+2)].width = 5
        for i, h in enumerate(['TOTAL', 'A', 'B', 'C', 'W/O', 'X', 'L', 'G']):
            col = start_totals + i
            ws.cell(row=3, column=col, value=h).alignment = center
            ws.column_dimensions[get_column_letter(col)].width = 10 if h == 'TOTAL' else 5

        # Data Rows
        num_emp = len(employees)
        for idx, emp in enumerate(employees):
            r = idx + 4
            ws.cell(row=r, column=1, value=idx+1)
            ws.cell(row=r, column=2, value=emp)
            for d in range(1, days_in_month + 1):
                ws.cell(row=r, column=d+2, value=roster[emp][d]).alignment = center
            
            # Formulas
            t_ltr = get_column_letter(start_totals)
            last_d_ltr = get_column_letter(days_in_month + 2)
            ws[f'{t_ltr}{r}'] = f'=SUM({get_column_letter(start_totals+1)}{r}:{get_column_letter(start_totals+3)}{r})'
            for i_h, h_code in enumerate(['A*', 'B*', 'C*', 'W/O*', 'X*', 'L*', 'G*']):
                ws.cell(row=r, column=start_totals+1+i_h, value=f'=COUNTIF(C{r}:{last_d_ltr}{r},"{h_code}")')

            # Formatting
            ws[f'{t_ltr}{r}'].fill = yellow_fill; ws[f'{t_ltr}{r}'].font = Font(bold=True)
            for cp in range(start_totals+1, end_totals+1): ws.cell(row=r, column=cp).fill = peach_fill
            for c in range(1, end_totals + 1):
                ws.cell(row=r, column=c).border = Border(left=thick_blue if c==1 else thin, right=thick_blue if c==end_totals else thin, top=thin, bottom=thin if idx<num_emp-1 else thick_blue)

        # Bottom Summary Box
        s_row = num_emp + 6
        for i, stype in enumerate(["A", "B", "C"]):
            curr_r = s_row + i
            ws.cell(row=curr_r, column=2, value=stype).font = Font(bold=True)
            ws.cell(row=curr_r, column=2).fill = yellow_fill; ws.cell(row=curr_r, column=2).alignment = center; ws.cell(row=curr_r, column=2).border = all_border
            for d in range(1, days_in_month + 1):
                cltr = get_column_letter(d+2)
                c_cell = ws.cell(row=curr_r, column=cltr, value=f'=COUNTIF({cltr}4:{cltr}{num_emp+3},"{stype}*")')
                c_cell.fill = yellow_fill; c_cell.alignment = center; c_cell.border = all_border

        out = io.BytesIO()
        wb.save(out)
        st.balloons()
        st.download_button("Download Intelligent Roster", out.getvalue(), f"ROSTER_{target_month_name}.xlsx")
