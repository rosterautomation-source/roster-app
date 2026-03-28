import streamlit as st
import pandas as pd

st.title("Roster Debug App")

try:
df = pd.read_excel("latest_roster.xlsx")
st.write("File Loaded Successfully ✅")
st.dataframe(df.head())
except Exception as e:
st.error("Error loading file")
st.write(e)
