import streamlit as st
import pandas as pd
import io
import random
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import PatternFill, Alignment, Border, Side, Font

# ==========================================
# CONFIGURATION
# ==========================================
st.set_page_config(page_title="Roster Automation", layout="wide")

DRIVE_FOLDER_ID = "1pcZWYGXCC1axVDXWtXp1YyQJ79WVeivr"
TEMPLATE_FILE = "Template.xlsx"

SEQ = ['C', 'C', 'B', 'B', 'A', 'A', 'W/O']
MONTH_NAMES = ["January", "February", "March", "April", "May", "June",
               "July", "August", "September", "October", "November", "December"]

MAX_CONSECUTIVE = 5
TARGET_TOTAL = 24
MAX_C_SHIFTS = 9

# ==========================================
# GOOGLE DRIVE
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
    results = service.files().list(q=query, orderBy="createdTime desc",
                                  pageSize=1, includeItemsFromAllDrives=True,
                                  supportsAllDrives=True).execute()
    items = results.get('files', [])
    if not items:
        return None
    request = service.files().get_media(fileId=items[0]['id'])
    return io.BytesIO(request.execute()), items[0]['name']

# ==========================================
# STATE DETECTION
# ==========================================
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
# UI
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

# ==========================================
# LOAD DATA
# ==========================================
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

# ==========================================
# LEAVES
# ==========================================
st.sidebar.header("2. Leaves")
if 'leaves' not in st.session_state:
    st.session_state.leaves = {}

sel_name = st.sidebar.selectbox("Select Employee", employees)
sel_days = st.sidebar.text_input("Enter Days (e.g., 5, 12)")

if st.sidebar.button("Register Leave"):
    st.session_state.leaves[sel_name] = [
        int(d.strip()) for d in sel_days.split(',') if d.strip().isdigit()
    ]
    st.sidebar.success(f"Leave added for {sel_name}")

# ==========================================
# MAIN ENGINE
# ==========================================
if st.button("Generate Fair Roster", type="primary"):

    emp_state = {n: get_state(emp_data_map[n]) for n in employees}
    duty_history = {n: prev_totals[n] for n in employees}
    this_month_duties = {n: 0 for n in employees}
    c_counts = {n: 0 for n in employees}
    consecutive_days = {n: 0 for n in employees}
    wo_counts = {n: 0 for n in employees}

    roster = {n: {d: "" for d in range(1, days_in_month + 1)} for n in employees}

    for d in range(1, days_in_month + 1):

        available = []

        for emp in employees:
            if emp in st.session_state.leaves and d in st.session_state.leaves[emp]:
                roster[emp][d] = 'L'
                consecutive_days[emp] = 0
            else:
                available.append(emp)

        # SCORING
        def score(emp):
            total = duty_history[emp] + this_month_duties[emp]

            s = 0
            s += (TARGET_TOTAL - total) * 5
            s -= c_counts[emp] * 2

            if consecutive_days[emp] >= MAX_CONSECUTIVE:
                s -= 50

            expected = SEQ[emp_state[emp]]
            if expected in ['A', 'B', 'C']:
                s += 3

            s += random.uniform(0, 2)
            return s

        available.sort(key=lambda x: score(x), reverse=True)

        if len(available) < 24:
            st.error(f"Not enough employees on day {d}")
            break

        workers_today = available[:24]
        off_today = available[24:]

        # OFF assignment
        for emp in off_today:
            if wo_counts[emp] < 4:
                roster[emp][d] = 'W/O'
                wo_counts[emp] += 1
            else:
                roster[emp][d] = 'X'
            consecutive_days[emp] = 0

        # SHIFT assignment
        shift_slots = {'C': 8, 'B': 8, 'A': 8}
        unassigned = workers_today[:]

        # PASS 1 (rotation friendly)
        for sft in ['C', 'B', 'A']:
            for emp in unassigned[:]:
                if shift_slots[sft] == 0:
                    break

                expected = SEQ[emp_state[emp]]

                if sft == 'C' and c_counts[emp] >= MAX_C_SHIFTS:
                    continue

                if expected == sft:
                    roster[emp][d] = sft
                    shift_slots[sft] -= 1
                    this_month_duties[emp] += 1
                    consecutive_days[emp] += 1

                    if sft == 'C':
                        c_counts[emp] += 1

                    emp_state[emp] = (emp_state[emp] + 1) % 7
                    unassigned.remove(emp)

        # PASS 2 (fair fill)
        for sft in ['C', 'B', 'A']:
            while shift_slots[sft] > 0 and unassigned:
                unassigned.sort(key=lambda x: (
                    c_counts[x] if sft == 'C' else 0,
                    this_month_duties[x]
                ))
                emp = unassigned.pop(0)

                roster[emp][d] = sft
                shift_slots[sft] -= 1
                this_month_duties[emp] += 1
                consecutive_days[emp] += 1

                if sft == 'C':
                    c_counts[emp] += 1

                emp_state[emp] = (emp_state[emp] + 1) % 7

    # ==========================================
    # EXCEL GENERATION
    # ==========================================
    wb = load_workbook(TEMPLATE_FILE)
    ws = wb.active

    for m in list(ws.merged_cells.ranges):
        ws.unmerge_cells(str(m))

    for r in ws.iter_rows(min_row=1, max_row=120, min_col=1, max_col=65):
        for cell in r:
            if cell.column > 2:
                cell.value = None

    center = Alignment(horizontal='center', vertical='center')
    yellow = PatternFill(start_color="FFFF00", fill_type="solid")

    start_totals = days_in_month + 3

    for idx, emp in enumerate(employees):
        r = idx + 4
        ws.cell(row=r, column=1, value=idx+1)
        ws.cell(row=r, column=2, value=emp)

        for d in range(1, days_in_month + 1):
            ws.cell(row=r, column=d+2, value=roster[emp][d]).alignment = center

    out = io.BytesIO()
    wb.save(out)

    st.success("Roster Generated Successfully")
    st.download_button("Download Roster", out.getvalue(), f"ROSTER_{target_month_name}.xlsx")
