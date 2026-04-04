import streamlit as st
import pandas as pd

st.set_page_config(page_title="Roster App", layout="wide")

st.title("Roster Generator")

st.write("Loading previous month roster...")

df = pd.read_excel("latest_roster.xlsx", skiprows=2)

st.success("File loaded successfully")

st.write("Preview of data:")
st.dataframe(df.head())
