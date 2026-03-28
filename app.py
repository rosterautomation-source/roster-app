import streamlit as st
import pandas as pd
import io
import random
import os
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import PatternFill, Alignment, Border, Side, Font

# ==========================================

# CONFIG

# ==========================================

st.set_page_config(page_title="Pro Roster Automation", layout="wide")

TEMPLATE_FILE = "Template.xlsx"
SAVE_PATH = "latest_roster.xlsx"

SEQ = ['C', 'C', 'B', 'B', 'A', 'A', 'W/O']
MONTH_NAMES = ["January", "February", "March", "April", "May", "June",
"July", "August", "September", "October", "November", "December"]

# ==========================================

# UI

# ==========================================

st.title("📅 Professional Roster Generator")

# Upload only first time

uploaded_file = st.file_uploader("Upload Previous Month Roster (First Time Only)", type=["xlsx"])

if uploaded_file is not None:
with open(SAVE_PATH, "wb") as f:
f.write(uploaded_file.getbuffer())
st.success("File saved successfully ✅")

# Load latest roster

if os.path.exists(SAVE_PATH):
df_raw = pd.read_excel(SAVE_PATH, skiprows=2)
st.success("Using latest saved roster ✅")
else:
st.warning("Upload previous month roster first time")
st.stop()

# ==========================================

# SETTINGS

# ==========================================

st.sidebar.header("1. Settings")
target_month_name = st.sidebar.selectbox("Month", MONTH_NAMES, index=3)
target_year = st.sidebar.number_input("Year", min_value=2024, max_value=2050, value=2026)

target_month_num = MONTH_NAMES.index(target_month_name) + 1
days_in_month = pd.Period(f'{target_year}-{target_month_num:02d}-01').days_in_month

# ==========================================

# DATA PREP

# ==========================================

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

# STATE LOGIC

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

# LEAVES

# ==========================================

st.sidebar.header("2. Leaves")
if 'leaves' not in st.session_state:
st.session_state.leaves = {}

sel_name = st.sidebar.selectbox("Select Employee", employees)
sel_days = st.sidebar.text_input("Enter Days (e.g., 5, 12)")

if st.sidebar.button("Register Leave"):
st.session_state.leaves[sel_name] = [int(d.strip()) for d in sel_days.split(',') if d.strip().isdigit()]
st.sidebar.success(f"Added for {sel_name}")

# ==========================================

# GENERATION ENGINE

# ==========================================

if st.button("Generate Roster", type="primary"):

```
TARGET_TOTAL = 24

emp_state = {n: get_state(emp_data_map[n]) for n in employees}
this_month_duties = {n: 0 for n in employees}
c_counts = {n: 0 for n in employees}
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

    available = sorted(available, key=lambda x: this_month_duties[x])

    workers = available[:24]
    off = available[24:]

    for emp in off:
        roster[emp][d] = 'W/O' if SEQ[emp_state[emp]] == 'W/O' else 'X'
        consecutive_days[emp] = 0

    assigned = set()

    for shift in ['C', 'B', 'A']:
        for _ in range(8):

            candidates = [e for e in workers if e not in assigned]

            valid = []
            for emp in candidates:
                if shift == 'A' and d > 1 and roster[emp][d-1] == 'C':
                    continue
                valid.append(emp)

            if not valid:
                valid = candidates

            valid.sort(key=lambda x: (this_month_duties[x], random.random()))
            chosen = valid[0]

            roster[chosen][d] = shift
            assigned.add(chosen)

            this_month_duties[chosen] += 1
            consecutive_days[chosen] += 1

            if SEQ[emp_state[chosen]] == shift:
                emp_state[chosen] = (emp_state[chosen] + 1) % 7
            else:
                emp_state[chosen] = (SEQ.index(shift) + 1) % 7

# ==========================================
# EXCEL OUTPUT
# ==========================================
wb = load_workbook(TEMPLATE_FILE)
ws = wb.active

for i, emp in enumerate(employees):
    row = i + 4
    ws.cell(row=row, column=2, value=emp)
    for d in range(1, days_in_month + 1):
        ws.cell(row=row, column=d+2, value=roster[emp][d])

output = io.BytesIO()
wb.save(output)
output.seek(0)

# 🔥 SAVE FOR NEXT MONTH
with open(SAVE_PATH, "wb") as f:
    f.write(output.getbuffer())

st.success("Roster Generated & Saved for next month ✅")

st.download_button(
    "Download Roster",
    output,
    f"ROSTER_{target_month_name}.xlsx"
)
```
