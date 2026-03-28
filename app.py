import streamlit as st
import pandas as pd

st.title("Roster Test")

df = pd.read_excel("latest_roster.xlsx")

st.write("File loaded successfully")
st.dataframe(df.head())

if st.button("Test Button"):
st.write("App working correctly")
