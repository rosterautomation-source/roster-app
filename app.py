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

```
if name != "" and name.lower() not in ["nan", "a", "b", "c", "total", "none"]:

    employees.append(name)
    emp_rows[name] = df.iloc[i]

    if total_col_index is not None:
        val = df.iloc[i, total_col_index]
        prev_duties[name] = float(val) if pd.notna(val) else 0
    else:
        prev_duties[name] = 0
```

st.write("Total Employees:", len(employees))

# =========================

# LAST SHIFT

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

# CURRENT MONTH DUTIES

# =========================

current_duties = {}
for emp in employees:
current_duties[emp] = 0

# =========================

# GENERATE MONTH

# =========================

days = 30
roster = {}

for emp in employees:
roster[emp] = {}

for d in range(1, days + 1):

```
# SORT BY FAIRNESS
sorted_employees = sorted(
    employees,
    key=lambda x: (current_duties[x] + prev_duties[x])
)

workers = sorted_employees[:24]
off_people = sorted_employees[24:]

day_roster = {}

# ASSIGN SHIFTS DYNAMICALLY
c_count = 0
b_count = 0
a_count = 0

for emp in workers:

    if c_count < 8:
        day_roster[emp] = "C"
        c_count += 1

    elif b_count < 8:
        day_roster[emp] = "B"
        b_count += 1

    else:
        # RULE: C → A NOT ALLOWED
        if last_shift[emp] == "C":
            day_roster[emp] = "B"
        else:
            day_roster[emp] = "A"

        a_count += 1

for emp in off_people:
    day_roster[emp] = "W/O"

# UPDATE
for emp in employees:
    roster[emp][d] = day_roster[emp]

    if day_roster[emp] in ["A", "B", "C"]:
        current_duties[emp] += 1
        last_shift[emp] = day_roster[emp]
```

# SHOW RESULT

st.write("Final Roster Sample (First 5 Employees):")

for emp in employees[:5]:
st.write(emp, roster[emp])
