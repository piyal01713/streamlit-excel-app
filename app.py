import streamlit as st
import pandas as pd

st.title("📊 Excel Resource Explorer")

uploaded_files = st.file_uploader(
    "Upload Excel files",
    type=["xlsx"],
    accept_multiple_files=True
)

dfs = []

if uploaded_files:
    for f in uploaded_files:
        df = pd.read_excel(f)
        df["source"] = f.name
        dfs.append(df)

    combined = pd.concat(dfs, ignore_index=True)

    st.success("Loaded successfully!")
    st.dataframe(combined)

    st.write("Rows:", len(combined))
