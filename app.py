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
# 4. ENGINE
# ==========================================
if st.button("Generate Fair & Balanced Roster", type="primary"):
    with st.spinner("Executing Production Scheduling Engine..."):

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

        wb = load_workbook(TEMPLATE_FILE)
        ws = wb.active

        for m in list(ws.merged_cells.ranges):
            ws.unmerge_cells(str(m))

        for r in ws.iter_rows(min_row=1, max_row=120, min_col=1, max_col=65):
            for cell in r:
                if cell.column > 2:
                    cell.value = None

        for idx, emp in enumerate(employees):
            r = idx + 4
            ws.cell(row=r, column=1, value=idx+1)
            ws.cell(row=r, column=2, value=emp)

            for d in range(1, days_in_month + 1):
                ws.cell(row=r, column=d+2, value=roster[emp][d])

        out = io.BytesIO()
        wb.save(out)

        st.success("Roster Generated Successfully")
        st.download_button("Download Roster", out.getvalue(), "roster.xlsx")
