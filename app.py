import streamlit as st
import pandas as pd

st.set_page_config(page_title="Roster App", layout="wide")
st.title("Roster Generator")

with st.spinner("Loading previous month roster..."):
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

for i in range(len(df)):
    name = str(df.iloc[i, 1]).strip()
    if name != "" and name.lower() not in ["nan", "a", "b", "c", "total", "none"]:
        employees.append(name)
        emp_rows[name] = df.iloc[i]

        if total_col_index is not None:
            val = df.iloc[i, total_col_index]
            prev_duties[name] = float(val) if pd.notna(val) else 0
        else:
            prev_duties[name] = 0

st.write("Total Employees:", len(employees))

# =========================
# STATE DETECTION
# =========================
def get_last_shift(row):
    for col in reversed(df.columns):
        val = row[col]
        if pd.notna(val):
            v = str(val).strip().upper()
            if v in ['A', 'B', 'C']:
                return v
    return None

last_shift = {}
for emp in employees:
    last_shift[emp] = get_last_shift(emp_rows[emp])

# =========================
# DAY 1 LOGIC WITH RULES
# =========================
sorted_employees = sorted(employees, key=lambda x: prev_duties[x])
workers = sorted_employees[:24]
off_people = sorted_employees[24:]
day1_roster = {}

# FIRST PASS: ASSIGN C (8)
count = 0
for emp in workers:
    if count < 8:
        day1_roster[emp] = "C"
        count += 1

# SECOND PASS: ASSIGN B (8)
count = 0
for emp in workers:
    if emp not in day1_roster:
        if count < 8:
            day1_roster[emp] = "B"
            count += 1

# THIRD PASS: ASSIGN A (8) WITH RULE
count = 0
for emp in workers:
    if emp not in day1_roster:
        if count < 8:
            # 🚨 RULE: C → A NOT ALLOWED
            if last_shift[emp] == "C":
                day1_roster[emp] = "B"
            else:
                day1_roster[emp] = "A"
            count += 1

# OFF PEOPLE
for emp in off_people:
    day1_roster[emp] = "W/O"

st.write("Day 1 Roster Sample:")
sample = dict(list(day1_roster.items())[:10])
st.write(sample)

st.write("Last Shift Sample:")
sample_last = dict(list(last_shift.items())[:5])
st.write(sample_last)
