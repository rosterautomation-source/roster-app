import streamlit as st
import pandas as pd
import io
from openpyxl import load_workbook

# BASIC CONFIG

st.set_page_config(page_title="Roster App", layout="wide")

TEMPLATE_FILE = "Template.xlsx"
DATA_FILE = "latest_roster.xlsx"

st.title("Roster Generator")

# LOAD BASE FILE

df = pd.read_excel(DATA_FILE, skiprows=2)

# GET EMPLOYEES

employees = []
for _, row in df.iterrows():
name = str(row.iloc[1]).strip()
if name and name.lower() not in ["nan", "total", "a", "b", "c"]:
employees.append(name)

# SETTINGS

month = st.sidebar.selectbox("Month", ["April", "May", "June", "July"])
year = st.sidebar.number_input("Year", 2024, 2050, 2026)

days = 30

# GENERATE BUTTON

if st.button("Generate Roster"):

```
roster = {}

for emp in employees:
    roster[emp] = {}
    for d in range(1, days+1):
        roster[emp][d] = ""

for d in range(1, days+1):

    workers = employees[:24]
    off = employees[24:]

    for emp in off:
        roster[emp][d] = "W/O"

    i = 0

    for emp in workers:
        if i < 8:
            roster[emp][d] = "C"
        elif i < 16:
            roster[emp][d] = "B"
        else:
            roster[emp][d] = "A"
        i += 1

# WRITE TO TEMPLATE
wb = load_workbook(TEMPLATE_FILE)
ws = wb.active

for i, emp in enumerate(employees):
    row = i + 4
    ws.cell(row=row, column=2, value=emp)

    for d in range(1, days+1):
        ws.cell(row=row, column=d+2, value=roster[emp][d])

output = io.BytesIO()
wb.save(output)
output.seek(0)

st.success("Roster Generated Successfully")

st.download_button(
    "Download Roster",
    output,
    file_name="ROSTER.xlsx"
)
```
