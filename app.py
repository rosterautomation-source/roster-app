import streamlit as st
import pandas as pd

st.set_page_config(page_title="Roster App", layout="wide")
st.title("Roster Generator")
st.write("Loading previous month roster...")
df = pd.read_excel("latest_roster.xlsx", skiprows=2)
st.success("File loaded successfully")

# FIND TOTAL COLUMN
total_col_index = None
for i in range(len(df.columns)):
    if "TOTAL" in str(df.columns[i]).upper():
        total_col_index = i
        break

employees = []
prev_duties = {}
emp_rows = {}

# EXTRACT EMPLOYEES + DUTIES
for i in range(len(df)):
    name = str(df.iloc[i, 1]).strip()
    if name != "" and name.lower() not in ["nan", "a", "b", "c", "total", "none"]:
        employees.append(name)
        emp_rows[name] = df.iloc[i]

        if total_col_index is not None:
            val = df.iloc[i, total_col_index]
            if pd.notna(val):
                prev_duties[name] = float(val)
            else:
                prev_duties[name] = 0
        else:
            prev_duties[name] = 0

st.write("Total Employees:", len(employees))

# =========================
# DETECT LAST SHIFT STATE
# =========================
SEQ = ['C', 'C', 'B', 'B', 'A', 'A', 'W/O']

def get_state(row):
    last_val = None
    prev_val = None

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

    if last_val == 'C':
        if prev_val == 'C':
            return 2
        else:
            return 1

    if last_val == 'B':
        if prev_val == 'B':
            return 4
        else:
            return 3

    if last_val == 'A':
        if prev_val == 'A':
            return 6
        else:
            return 5

    return 0

emp_state = {}
for emp in employees:
    emp_state[emp] = get_state(emp_rows[emp])

st.write("Sample Rotation States:")
sample_state = dict(list(emp_state.items())[:5])
st.write(sample_state)
