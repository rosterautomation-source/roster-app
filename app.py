import streamlit as st
import pandas as pd
import io
from openpyxl import load_workbook

st.set_page_config(page_title="Roster App", layout="wide")

st.title("Roster Generator")

load file

df = pd.read_excel("latest_roster.xlsx", skiprows=2)

st.write("File Loaded Successfully")

extract employees

employees = []

for i in range(len(df)):
name = str(df.iloc[i, 1]).strip()
if name != "" and name.lower() not in ["nan", "total", "a", "b", "c"]:
employees.append(name)

st.write("Total Employees:", len(employees))

settings

month = st.sidebar.selectbox("Month", ["April", "May", "June", "July"])
year = st.sidebar.number_input("Year", 2024, 2050, 2026)

days = 30

generate

if st.button("Generate Roster"):

roster = {}

for emp in employees:
    roster[emp] = {}
    for d in range(1, days + 1):
        roster[emp][d] = ""

for d in range(1, days + 1):

    workers = employees[:24]
    off = employees[24:]

    for emp in off:
        roster[emp][d] = "W/O"

    index = 0

    for emp in workers:
        if index < 8:
            roster[emp][d] = "C"
        elif index < 16:
            roster[emp][d] = "B"
        else:
            roster[emp][d] = "A"

        index = index + 1

wb = load_workbook("Template.xlsx")
ws = wb.active

for i in range(len(employees)):
    row = i + 4
    ws.cell(row=row, column=2, value=employees[i])

    for d in range(1, days + 1):
        ws.cell(row=row, column=d + 2, value=roster[employees[i]][d])

output = io.BytesIO()
wb.save(output)
output.seek(0)

st.success("Roster Generated Successfully")

st.download_button(
    label="Download Roster",
    data=output,
    file_name="ROSTER.xlsx"
)
