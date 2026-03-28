import streamlit as st
import pandas as pd
import io
import random
import math
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import PatternFill, Alignment, Border, Side, Font

# ==========================================
# 1. CONFIGURATION & CONSTANTS
# ==========================================
st.set_page_config(page_title="Pro Roster Automation", layout="wide")

DRIVE_FOLDER_ID = "1pcZWYGXCC1axVDXWtXp1YyQJ79WVeivr"
TEMPLATE_FILE = "Template.xlsx"

SEQ = ['C', 'C', 'B', 'B', 'A', 'A', 'W/O']
MONTH_NAMES = ["January", "February", "March", "April", "May", "June",
               "July", "August", "September", "October", "November", "December"]

# ==========================================
# 2. DRIVE & DATA HELPERS
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
    if not items:
        return None
    request = service.files().get_media(fileId=items[0]['id'])
    return io.BytesIO(request.execute()), items[0]['name']

def get_state(row):
    last_val, prev_val = None, None
    for d in range(31, 0, -1):
        col = str(d)
        if col in row and pd.notna(row[col]):
            val = str(row[col]).strip().upper()
            if val in ['A', 'B', 'C', 'W/O', 'X', 'L']:
                if last_val is None:
                    last_val = val
                elif prev_val is None:
                    prev_val = val
                    break
    if last_val == 'C': return 2 if prev_val == 'C' else 1
    if last_val == 'B': return 4 if prev_val == 'B' else 3
    if last_val == 'A': return 6 if prev_val == 'A' else 5
    return 0

# ==========================================
# 3. WEB UI
# ==========================================
st.title("📅 Professional Roster Generator")

service = get_drive_service()
latest_file = get_latest_roster(service)

if not latest_file:
    st.error("Baseline file not found in Drive.")
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
    name = str(row.iloc[1]).strip()
    if name and name.lower() not in ["nan", "a", "b", "c", "total", "none"]:
        employees.append(name)
        emp_data_map[name] = row
        val = row.iloc[total_col_idx] if total_col_idx is not None else 24
        prev_totals[name] = float(val) if pd.notna(val) else 24.0

st.sidebar.header("2. Leaves")
if 'leaves' not in st.session_state:
    st.session_state.leaves = {}

sel_name = st.sidebar.selectbox("Select Employee", employees)
sel_days = st.sidebar.text_input("Enter Days (e.g., 5, 12)")

if st.sidebar.button("Register Leave"):
    st.session_state.leaves[sel_name] = [int(d.strip()) for d in sel_days.split(',') if d.strip().isdigit()]
    st.sidebar.success(f"Added for {sel_name}")

# ==========================================
# 4. ENGINE (Your Refined Logic)
# ==========================================
if st.button("Generate Fair & Balanced Roster", type="primary"):
    with st.spinner("Executing Production Scheduling Engine & Formatting..."):

        TARGET_TOTAL = (days_in_month * 24) // len(employees)
        MIN_WO = 4
        MAX_CONSECUTIVE = 5

        emp_state = {n: get_state(emp_data_map[n]) for n in employees}
        this_month_duties = {n: 0 for n in employees}
        c_counts = {n: 0 for n in employees}

        this_month_A = {n: 0 for n in employees}
        this_month_B = {n: 0 for n in employees}
        target_C = {n: 0 for n in employees}

        last_wo_day = {n: -10 for n in employees}
        wo_counts = {n: 0 for n in employees}
        x_counts = {n: 0 for n in employees}
        consecutive_days = {n: 0 for n in employees}

        roster = {n: {d: "" for d in range(1, days_in_month + 1)} for n in employees}

        for d in range(1, days_in_month + 1):

            available = []

            for emp in employees:
                if emp in st.session_state.leaves and d in st.session_state.leaves[emp]:
                    roster[emp][d] = 'L'
                    consecutive_days[emp] = 0
                else:
                    available.append(emp)

            # Prioritize workers with the lowest current duties
            available.sort(key=lambda x: this_month_duties[x])

            workers_today = available[:24]
            off_today = available[24:]

            for emp in off_today:
                if wo_counts[emp] < MIN_WO and (d - last_wo_day[emp]) >= 5:
                    roster[emp][d] = 'W/O'
                    wo_counts[emp] += 1
                    last_wo_day[emp] = d
                else:
                    roster[emp][d] = 'X'
                    x_counts[emp] += 1
                consecutive_days[emp] = 0

            shift_slots = {'C': 8, 'B': 8, 'A': 8}
            assigned_today = set()

            def shift_score(emp, shift):
                score = (TARGET_TOTAL - this_month_duties[emp]) * 5

                if shift == 'C':
                    score += (target_C[emp] - c_counts[emp]) * 10
                    if c_counts[emp] >= target_C[emp]:
                        score -= 100
                elif shift == 'A':
                    score += (this_month_duties[emp] - this_month_A[emp])
                elif shift == 'B':
                    score += (this_month_duties[emp] - this_month_B[emp])

                score += random.uniform(0, 1)
                return score

            for shift in ['C', 'B', 'A']:
                for _ in range(8):

                    candidates = [e for e in workers_today if e not in assigned_today]

                    valid = []
                    for emp in candidates:
                        if shift == 'C' and c_counts[emp] >= target_C[emp]:
                            continue
                        valid.append(emp)

                    if not valid:
                        valid = candidates

                    valid.sort(key=lambda x: shift_score(x, shift), reverse=True)
                    chosen = valid[0]

                    roster[chosen][d] = shift
                    assigned_today.add(chosen)

                    this_month_duties[chosen] += 1
                    target_C[chosen] = round(this_month_duties[chosen] / 3)
                    consecutive_days[chosen] += 1

                    if shift == 'C':
                        c_counts[chosen] += 1
                    elif shift == 'A':
                        this_month_A[chosen] += 1
                    elif shift == 'B':
                        this_month_B[chosen] += 1

                    emp_state[chosen] = (emp_state[chosen] + 1) % 7

        # ==========================================
        # 5. EXCEL FORMATTING (PRO BLOCKS & BORDERS)
        # ==========================================
        wb = load_workbook(TEMPLATE_FILE)
        ws = wb.active

        # Clean slate: unmerge and clear existing data
        for m in list(ws.merged_cells.ranges):
            ws.unmerge_cells(str(m))

        for r in ws.iter_rows(min_row=1, max_row=120, min_col=1, max_col=65):
            for cell in r:
                if cell.column > 2:
                    cell.value = None
                cell.fill = PatternFill(fill_type=None)
                cell.border = Border()

        # Define Styles
        thin = Side(border_style="thin")
        thick_blue = Side(border_style="thick", color="0000FF")
        all_border = Border(left=thin, right=thin, top=thin, bottom=thin)
        yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
        peach_fill = PatternFill(start_color="FFCC99", end_color="FFCC99", fill_type="solid")
        center = Alignment(horizontal='center', vertical='center')

        # Dimensions & Bounds
        ws.column_dimensions['A'].width = 6.43
        ws.column_dimensions['B'].width = 25
        start_totals = days_in_month + 3
        end_totals = start_totals + 7

        # Main Title Merge & Format
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=end_totals)
        t_cell = ws.cell(row=1, column=1, value=f"DUTY ROSTER FOR THE MONTH OF {target_month_name[:3].upper()} {target_year}")
        t_cell.font = Font(bold=True, size=20)
        t_cell.alignment = center

        # Header Merges
        ws.merge_cells('A2:A3')
        ws['A2'] = "S No"
        ws['A2'].font = Font(bold=True, size=16)
        ws['A2'].alignment = center

        ws.merge_cells('B2:B3')
        ws['B2'] = "NAME"
        ws['B2'].font = Font(bold=True, size=16)
        ws['B2'].alignment = center

        ws.merge_cells(start_row=2, start_column=3, end_row=2, end_column=days_in_month+2)
        ws.cell(row=2, column=3, value="ATTENDANCE").font = Font(bold=True, size=16)
        ws.cell(row=2, column=3).alignment = center

        ws.merge_cells(start_row=2, start_column=start_totals, end_row=2, end_column=end_totals)
        ws.cell(row=2, column=start_totals, value="TOTAL SHIFTS").font = Font(bold=True, size=16)
        ws.cell(row=2, column=start_totals).alignment = center

        # Day Columns and Totals Columns Setup
        for d in range(1, days_in_month + 1):
            ws.cell(row=3, column=d+2, value=d).alignment = center
            ws.column_dimensions[get_column_letter(d+2)].width = 5

        h_labels = ['TOTAL', 'A', 'B', 'C', 'W/O', 'X', 'L', 'G']
        for i, h in enumerate(h_labels):
            col = start_totals + i
            ws.cell(row=3, column=col, value=h).alignment = center
            ws.column_dimensions[get_column_letter(col)].width = 10 if h == 'TOTAL' else 5

        # Write Roster Data & Calculated Fields
        num_emp = len(employees)
        for idx, emp in enumerate(employees):
            r = idx + 4
            ws.cell(row=r, column=1, value=idx+1).alignment = center
            ws.cell(row=r, column=2, value=emp)

            for d in range(1, days_in_month + 1):
                ws.cell(row=r, column=d+2, value=roster[emp][d]).alignment = center

            t_ltr = get_column_letter(start_totals)
            last_d_ltr = get_column_letter(days_in_month + 2)
            
            # Formulas
            ws[f'{t_ltr}{r}'] = f'=SUM({get_column_letter(start_totals+1)}{r}:{get_column_letter(start_totals+3)}{r})'
            for i_h, h_code in enumerate(['A*', 'B*', 'C*', 'W/O*', 'X*', 'L*', 'G*']):
                ws.cell(row=r, column=start_totals+1+i_h, value=f'=COUNTIF(C{r}:{last_d_ltr}{r},"{h_code}")')

            # Row Paints
            ws[f'{t_ltr}{r}'].fill = yellow_fill
            ws[f'{t_ltr}{r}'].font = Font(bold=True)
            for cp in range(start_totals+1, end_totals+1):
                ws.cell(row=r, column=cp).fill = peach_fill

            # Apply Borders to each cell in the row
            for c in range(1, end_totals + 1):
                ws.cell(row=r, column=c).border = Border(
                    left=thick_blue if c==1 else thin,
                    right=thick_blue if c==end_totals else thin,
                    top=thin,
                    bottom=thin if idx < num_emp - 1 else thick_blue
                )

        # Apply Header Borders (Rows 1-3)
        for rh in range(1, 4):
            for ch in range(1, end_totals + 1):
                ws.cell(row=rh, column=ch).border = Border(
                    left=thick_blue if ch==1 else thin,
                    right=thick_blue if ch==end_totals else thin,
                    top=thick_blue if rh==1 else thin,
                    bottom=thin
                )

        # Bottom Summary Box (8-8-8 Output)
        s_row = num_emp + 6
        for i, stype in enumerate(["A", "B", "C"]):
            curr_r = s_row + i
            ws.cell(row=curr_r, column=2, value=stype).font = Font(bold=True)
            ws.cell(row=curr_r, column=2).fill = yellow_fill
            ws.cell(row=curr_r, column=2).alignment = center
            ws.cell(row=curr_r, column=2).border = all_border
            
            for d in range(1, days_in_month + 1):
                col_i = d + 2
                cltr = get_column_letter(col_i)
                ws.cell(row=curr_r, column=col_i, value=f'=COUNTIF({cltr}4:{cltr}{num_emp+3},"{stype}*")').fill = yellow_fill
                ws.cell(row=curr_r, column=col_i).alignment = center
                ws.cell(row=curr_r, column=col_i).border = all_border

        # Save and Download
        out = io.BytesIO()
        wb.save(out)

        st.balloons()
        st.success("Production Roster Generated Successfully with Full Formatting!")
        st.download_button("Download Final Roster", out.getvalue(), f"ROSTER_{target_month_name}.xlsx")
