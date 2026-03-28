import streamlit as st
import pandas as pd
import io
import random
from openpyxl import load_workbook

# ==========================================

# CONFIG

# ==========================================

st.set_page_config(page_title="Roster Automation", layout="wide")

TEMPLATE_FILE = "Template.xlsx"
SAVE_PATH = "latest_roster.xlsx"

SEQ = ['C', 'C', 'B', 'B', 'A', 'A', 'W/O']
MONTH_NAMES = ["January", "February", "March", "April", "May", "June",
"July", "August", "September", "October", "November", "December"]

# ==========================================

# LOAD FILE (AUTO)

# ==========================================

try:
df_raw = pd.read_excel(SAVE_PATH, skiprows=2)
st.success("Loaded latest roster automatically ✅")
except:
st.error("latest_roster.xlsx not found in repo")
st.stop()

# ==========================================

# SETTINGS

# ==========================================

st.title("📅 Roster Generator")

st.sidebar.header("Settings")
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

for _, row in df_raw.iterrows():
name = str(row.iloc[1]).strip()
if name and name.lower() not in ["nan", "a", "b", "c", "total", "none"]:
employees.append(name)
emp_data_map[name] = row

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

# GENERATE BUTTON

# ==========================================

if st.button("Generate Roster"):

```
emp_state = {n: get_state(emp_data_map[n]) for n in employees}
duties = {n: 0 for n in employees}

roster = {n: {d: "" for d in range(1, days_in_month + 1)} for n in employees}

for d in range(1, days_in_month + 1):

    available = sorted(employees, key=lambda x: duties[x])
    workers = available[:24]
    off = available[24:]

    for emp in off:
        roster[emp][d] = 'W/O'

    assigned = set()

    for shift in ['C', 'B', 'A']:
        for _ in range(8):

            candidates = [e for e in workers if e not in assigned]

            if shift == 'A':
                candidates = [e for e in candidates if not (d > 1 and roster[e][d-1] == 'C')]

            if not candidates:
                candidates = workers

            chosen = sorted(candidates, key=lambda x: (duties[x], random.random()))[0]

            roster[chosen][d] = shift
            duties[chosen] += 1
            assigned.add(chosen)

            if SEQ[emp_state[chosen]] == shift:
                emp_state[chosen] = (emp_state[chosen] + 1) % 7
            else:
                emp_state[chosen] = (SEQ.index(shift) + 1) % 7

# ==========================================
# WRITE TO EXCEL
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

st.success("Roster Generated ✅")

st.download_button(
    "Download Roster",
    output,
    f"ROSTER_{target_month_name}.xlsx"
)
```
