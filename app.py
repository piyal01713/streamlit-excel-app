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

MODEL = "claude-3-5-haiku-20241022" 

api_key = st.secrets.get("ANTHROPIC_API_KEY")
if not api_key:
    st.error("Missing ANTHROPIC_API_KEY in Streamlit secrets!")
    st.stop()

client = Anthropic(api_key=api_key)

# ------------------------------------------------
# HELPERS: URL CONVERSION & SECURITY
# ------------------------------------------------

def get_direct_download_link(url):
    """Specific logic for SharePoint Business/Organization links."""
    try:
        if "docs.google.com/spreadsheets" in url:
            file_id = re.search(r'/d/([^/]+)', url).group(1)
            return f'https://docs.google.com/spreadsheets/d/{file_id}/export?format=xlsx'
        
        elif "drive.google.com" in url:
            file_id = re.search(r'd/([^/]+)', url).group(1)
            return f'https://drive.google.com/uc?export=download&id={file_id}'
        
        elif "sharepoint.com" in url:
            # FIX FOR YOUR LINK: Remove 'viewer' segments and force download
            # Replaces /:x:/r/ or /:w:/r/ with simple /
            clean_url = re.sub(r'/:[a-z]:/r/', '/', url)
            # Remove all query parameters (?d=... &csf=...)
            clean_url = clean_url.split("?")[0]
            # Add the download trigger
            return f"{clean_url}?download=1"

        elif "1drv.ms" in url or "onedrive.live.com" in url:
            encoded_url = base64.b64encode(url.encode()).decode().replace('+', '-').replace('/', '_').rstrip('=')
            return f"https://api.onedrive.com/v1.0/shares/u!{encoded_url}/root/content"

        return url
    except Exception:
        return url

# ------------------------------------------------
# AI & DATA LOGIC
# ------------------------------------------------

def build_context(dfs_dict, scope):
    if scope == "Analyze All Sheets (Join/Compare)":
        context = "SCOPE: ALL SHEETS. Dictionary: `dfs`.\n"
        for name, df in dfs_dict.items():
            context += f"Sheet: '{name}' | Columns: {list(df.columns)}\n"
            context += f"Sample:\n{df.head(2).to_string(index=False)}\n---\n"
    else:
        df = dfs_dict[scope]
        context = f"SCOPE: SHEET '{scope}'. Use `dfs['{scope}']`.\n"
        context += f"Columns: {list(df.columns)}\nSamples:\n{df.head(5).to_string(index=False)}"
    return context

def get_analysis_code(user_query, context):
    system_prompt = "Write Python for a dictionary `dfs`. Output MUST be a variable `result`. RESPOND ONLY WITH CODE."
    try:
        response = client.messages.create(
            model=MODEL,
            max_tokens=1000,
            system=system_prompt,
            messages=[{"role": "user", "content": f"Context: {context}\n\nQuery: {user_query}"}]
        )
        return re.sub(r"```python\n|```", "", response.content[0].text.strip())
    except Exception as e:
        return None

def generate_natural_answer(user_query, execution_result):
    try:
        response = client.messages.create(
            model=MODEL,
            max_tokens=500,
            system="Explain the data result naturally.",
            messages=[{"role": "user", "content": f"Query: {user_query}\nResult: {execution_result}"}]
        )
        return response.content[0].text
    except:
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
        url_input = st.text_input("Paste Shareable Link:")
        if url_input:
            with st.spinner("Attempting to connect..."):
                dl_link = get_direct_download_link(url_input)
                try:
                    res = requests.get(dl_link, timeout=10)
                    raw_data = res.content
                    # CHECK FOR LOGIN PAGE (Security Wall)
                    if b"<!DOCTYPE" in raw_data[:100] or b"<html" in raw_data[:100].lower():
                        st.error("🔒 Link is Restricted")
                        st.warning("""
                        **Why this failed:**
                        Your organization requires a login to see this file. 
                        
                        **How to fix:**
                        1. Change SharePoint settings to **'Anyone with the link'**.
                        2. Or, use the **'Upload File'** option instead.
                        """)
                        raw_data = None
                except Exception as e:
                    st.error(f"Error: {e}")

    if raw_data:
        try:
            excel_file = pd.ExcelFile(io.BytesIO(raw_data))
            for sheet in excel_file.sheet_names:
                all_dfs[sheet] = excel_file.parse(sheet)
            st.success(f"Workbook Loaded!")
            scope = st.selectbox("Scope:", ["Analyze All Sheets (Join/Compare)"] + list(all_dfs.keys()))
        except Exception as e:
            st.error("Could not read file. It might be password protected or encrypted.")

# MAIN INTERFACE
if all_dfs:
    with st.expander("👀 Data Preview"):
        if scope == "Analyze All Sheets (Join/Compare)":
            st.write("Workbook Summary:")
            st.json({k: list(v.columns) for k, v in all_dfs.items()})
        else:
            st.dataframe(all_dfs[scope].head(10))

    user_query = st.text_input("Ask a question:")
    if user_query:
        context = build_context(all_dfs, scope)
        with st.spinner("AI analyzing..."):
            code = get_analysis_code(user_query, context)
            if code:
                try:
                    exec_globals = {'pd': pd, 'dfs': all_dfs}; exec_locals = {}
                    exec(code, exec_globals, exec_locals)
                    calc_res = exec_locals.get('result', "No result")
                    answer = generate_natural_answer(user_query, calc_res)
                    st.markdown(f"### Answer\n{answer}")
                    with st.expander("Logic"): st.code(code)
                except Exception as e:
                    st.error(f"Execution Error: {e}")
else:
    st.info("Load an Excel file to begin.")
