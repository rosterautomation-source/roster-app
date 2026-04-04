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

# SORT BASED ON FAIRNESS (LOW DUTIES FIRST)
sorted_employees = sorted(employees, key=lambda x: prev_duties[x])

# PICK WORKERS
workers = sorted_employees[:24]
off_people = sorted_employees[24:]

# ASSIGN SHIFTS
day1_roster = {}
for i in range(len(workers)):
    if i < 8:
        day1_roster[workers[i]] = "C"
    elif i < 16:
        day1_roster[workers[i]] = "B"
    else:
        day1_roster[workers[i]] = "A"

# OFF PEOPLE
for emp in off_people:
    day1_roster[emp] = "W/O"

st.write("Day 1 Roster Sample:")
sample = dict(list(day1_roster.items())[:10])
st.write(sample)
