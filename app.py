import pandas as pd
import streamlit as st
from anthropic import Anthropic
import json
import re
import base64
import io
import requests

# ------------------------------------------------
# CONFIG & INITIALIZATION
# ------------------------------------------------

st.set_page_config(page_title="Excel Insights", layout="wide")

# Claude 3.5 Haiku: High speed, low cost, great at code logic
MODEL = "claude-3-5-haiku-20241022" 

api_key = st.secrets.get("ANTHROPIC_API_KEY")
if not api_key:
    st.error("Missing ANTHROPIC_API_KEY in Streamlit secrets!")
    st.stop()

client = Anthropic(api_key=api_key)

# ------------------------------------------------
# HELPERS: URL CONVERSION
# ------------------------------------------------

def get_direct_download_link(url):
    try:
        if "docs.google.com/spreadsheets" in url:
            file_id = re.search(r'/d/([^/]+)', url).group(1)
            return f'https://docs.google.com/spreadsheets/d/{file_id}/export?format=xlsx'
        elif "drive.google.com" in url:
            file_id = re.search(r'd/([^/]+)', url).group(1)
            return f'https://drive.google.com/uc?export=download&id={file_id}'
        elif "1drv.ms" in url or "onedrive.live.com" in url:
            encoded_url = base64.b64encode(url.encode()).decode().replace('+', '-').replace('/', '_').rstrip('=')
            return f"https://api.onedrive.com/v1.0/shares/u!{encoded_url}/root/content"
        elif "sharepoint.com" in url:
            return url.split("?")[0] + "?download=1" if "?" in url else url + "?download=1"
        return url
    except Exception as e:
        return None

# ------------------------------------------------
# CONTEXT BUILDER (SMART TOKEN USAGE)
# ------------------------------------------------

def build_context(dfs_dict, scope):
    """
    If scope is 'All Sheets', sends summary of everything.
    If scope is a specific name, sends detailed summary of just that sheet.
    """
    if scope == "Analyze All Sheets (Join/Compare)":
        context = "ANALYSIS SCOPE: ALL SHEETS. Use the dictionary `dfs` where keys are sheet names.\n\n"
        for name, df in dfs_dict.items():
            context += f"Sheet: '{name}' | Columns: {list(df.columns)} | Rows: {len(df)}\n"
            context += f"Samples:\n{df.head(2).to_string(index=False)}\n---\n"
    else:
        df = dfs_dict[scope]
        context = f"ANALYSIS SCOPE: SINGLE SHEET ('{scope}'). Use `dfs['{scope}']`.\n\n"
        context += f"Columns: {list(df.columns)}\n"
        context += f"Detailed Samples:\n{df.head(5).to_string(index=False)}\n"
        context += f"Total Rows: {len(df)}"
    
    return context

# ------------------------------------------------
# AI LOGIC
# ------------------------------------------------

def get_analysis_code(user_query, context):
    system_prompt = """You are a Data Science Assistant. You have access to a dictionary of DataFrames named `dfs`.
    Write a Python snippet that produces a single variable named `result`. 
    RESPOND ONLY WITH THE CODE. No conversation."""

    try:
        response = client.messages.create(
            model=MODEL,
            max_tokens=1000,
            system=system_prompt,
            messages=[{"role": "user", "content": f"Data Context:\n{context}\n\nQuestion: {user_query}"}]
        )
        code = re.sub(r"```python\n|```", "", response.content[0].text.strip())
        return code
    except Exception as e:
        st.error(f"AI Planner Error: {e}")
        return None

def generate_natural_answer(user_query, execution_result):
    try:
        response = client.messages.create(
            model=MODEL,
            max_tokens=500,
            system="Explain the data result to the user clearly and professionally.",
            messages=[{"role": "user", "content": f"User asked: {user_query}\nCode Result: {execution_result}"}]
        )
        return response.content[0].text
    except Exception as e:
        return f"Result: {execution_result}"

# ------------------------------------------------
# STREAMLIT UI
# ------------------------------------------------

st.title("Excel Insights")

with st.sidebar:
    st.header("1. Load Data")
    input_method = st.radio("Source:", ["Upload File", "Cloud Link"])
    
    raw_data = None
    all_dfs = {}

    if input_method == "Upload File":
        uploaded_file = st.file_uploader("Upload XLSX", type=["xlsx"])
        if uploaded_file: raw_data = uploaded_file.read()
    else:
        url_input = st.text_input("Paste Link:")
        if url_input:
            with st.spinner("Fetching..."):
                dl_link = get_direct_download_link(url_input)
                try:
                    res = requests.get(dl_link)
                    raw_data = res.content
                except: st.error("Check link permissions.")

    if raw_data:
        try:
            excel_file = pd.ExcelFile(io.BytesIO(raw_data))
            for sheet in excel_file.sheet_names:
                all_dfs[sheet] = excel_file.parse(sheet)
            
            st.success(f"Workbook Loaded ({len(all_dfs)} sheets)")
            
            st.header("2. Set Analysis Scope")
            options = ["Analyze All Sheets (Join/Compare)"] + list(all_dfs.keys())
            scope = st.selectbox("What should the AI look at?", options)
            
            if scope == "Analyze All Sheets (Join/Compare)":
                st.warning("Power Mode: Higher token usage. Best for comparisons.")
            else:
                st.info("Token Saver: Analyzing specific sheet only.")
                
        except Exception as e:
            st.error(f"Error: {e}")

# MAIN AREA
if all_dfs:
    # Preview
    with st.expander("👀 Preview Selected Scope"):
        if scope == "Analyze All Sheets (Join/Compare)":
            st.write("Summary of all sheets available for analysis.")
            st.json({name: list(df.columns) for name, df in all_dfs.items()})
        else:
            st.dataframe(all_dfs[scope].head(10))

    # Chat
    st.subheader("💬 Insights Chat")
    user_query = st.text_input("Ask a question:", placeholder="e.g. 'What is the average price?'")
    
    if user_query:
        context = build_context
