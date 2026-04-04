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

# EXTRACT EMPLOYEES + DUTIES

for i in range(len(df)):
name = str(df.iloc[i, 1]).strip()

```
if name != "" and name.lower() not in ["nan", "a", "b", "c", "total", "none"]:

    employees.append(name)

    if total_col_index is not None:
        val = df.iloc[i, total_col_index]
        if pd.notna(val):
            prev_duties[name] = float(val)
        else:
            prev_duties[name] = 0
    else:
        prev_duties[name] = 0
```

st.write("Total Employees:", len(employees))

st.write("Sample Employees:")
st.write(employees[:5])

st.write("Previous Month Duties (Sample):")
sample_dict = dict(list(prev_duties.items())[:5])
st.write(sample_dict)
