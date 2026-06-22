import os
import pandas as pd
import streamlit as st

# Set page configuration with branding
st.set_page_config(page_title="Excel Insight", layout="wide")

# Main Title
st.title("💡 Excel Insight")

# Initialize session states
if "dataframes" not in st.session_state:
    st.session_state.dataframes = {}

if "excel_url" not in st.session_state:
    st.session_state.excel_url = ""

# Tabs for organization
tab1, tab2 = st.tabs(["📥 Import Data", "📊 Data View"])

with tab1:
    st.subheader("Load Excel Data")

    # URL Input
    excel_url = st.text_input(
        "Excel URL",
        value=st.session_state.excel_url,
        placeholder="https://example.com",
    )

    col1, col2 = st.columns(2)

    with col1:
        if st.button("Load URL"):
            try:
                df = pd.read_excel(excel_url)
                st.session_state.dataframes["URL File"] = df
                st.session_state.excel_url = excel_url
                st.success("Excel loaded successfully from URL.")
            except Exception as e:
                st.error(f"Could not load Excel: {e}")

    with col2:
        if st.button("🔄 Resync URL"):
            if st.session_state.excel_url:
                try:
                    df = pd.read_excel(st.session_state.excel_url)
                    st.session_state.dataframes["URL File"] = df
                    st.success("URL data refreshed.")
                except Exception as e:
                    st.error(f"Refresh failed: {e}")

    st.divider()

    # Multi-file Uploader
    uploaded_files = st.file_uploader(
        "Upload Excel files", type=["xlsx", "xls"], accept_multiple_files=True
    )

    if uploaded_files:
        for file in uploaded_files:
            try:
                df = pd.read_excel(file)
                df["source_file"] = file.name
                st.session_state.dataframes[file.name] = df
                st.success(f"Successfully loaded: {file.name}")
            except Exception as e:
                st.error(f"Error loading {file.name}: {e}")

with tab2:
    st.subheader("🧐 Inspect Loaded Datasets")

    # Guard rail: check if data exists
    if not st.session_state.dataframes:
        st.warning("Please upload a file or load a URL in the 'Import Data' tab first.")
    else:
        # Loop through loaded dataframes to offer full interactive views
        for name, df in st.session_state.dataframes.items():
            with st.expander(f"📁 {name} — ({df.shape[0]} rows, {df.shape[1]} columns)", expanded=True):
                # Interactive streamable data table
                st.dataframe(df, use_container_width=True)
