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

# Updated model as per your request
MODEL = "claude-haiku-4-5-20251001" 

# Initialize Session State to keep data across reruns
if 'all_dfs' not in st.session_state:
    st.session_state.all_dfs = None
if 'scope' not in st.session_state:
    st.session_state.scope = "Analyze All Sheets (Join/Compare)"

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
        elif "sharepoint.com" in url:
            clean_url = re.sub(r'/:[a-z]:/r/', '/', url).split("?")[0]
            return f"{clean_url}?download=1"
        elif "1drv.ms" in url or "onedrive.live.com" in url:
            encoded_url = base64.b64encode(url.encode()).decode().replace('+', '-').replace('/', '_').rstrip('=')
            return f"https://api.onedrive.com/v1.0/shares/u!{encoded_url}/root/content"
        return url
    except:
        return url

# ------------------------------------------------
# AI LOGIC
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
    system_prompt = "Write Python code for a dictionary of DataFrames named `dfs`. Result MUST be in a variable `result`. RESPOND ONLY WITH CODE."
    try:
        response = client.messages.create(
            model=MODEL,
            max_tokens=1000,
            system=system_prompt,
            messages=[{"role": "user", "content": f"Context: {context}\n\nQuery: {user_query}"}]
        )
        return re.sub(r"```python\n|```", "", response.content[0].text.strip())
    except Exception as e:
        st.error(f"AI Error: {e}")
        return None

def generate_natural_answer(user_query, execution_result):
    try:
        response = client.messages.create(
            model=MODEL,
            max_tokens=500,
            system="Explain the data result naturally to the user.",
            messages=[{"role": "user", "content": f"Query: {user_query}\nResult: {execution_result}"}]
        )
        return response.content[0].text
    except:
        return f"Result: {execution_result}"

# ------------------------------------------------
# STREAMLIT UI
# ------------------------------------------------

st.title("📊 Excel Insights")

with st.sidebar:
    st.header("1. Load Data")
    input_method = st.radio("Source:", ["Upload File", "Cloud Link"])
    
    raw_data = None
    if input_method == "Upload File":
        uploaded_file = st.file_uploader("Upload XLSX", type=["xlsx"])
        if uploaded_file:
            raw_data = uploaded_file.read()
    else:
        url_input = st.text_input("Paste Shareable Link:")
        if url_input:
            with st.spinner("Connecting..."):
                dl_link = get_direct_download_link(url_input)
                try:
                    res = requests.get(dl_link, timeout=15)
                    raw_data = res.content
                except Exception as e:
                    st.error(f"Error: {e}")

    if raw_data:
        try:
            excel_file = pd.ExcelFile(io.BytesIO(raw_data))
            # Store in session state
            st.session_state.all_dfs = {sheet: excel_file.parse(sheet) for sheet in excel_file.sheet_names}
            st.success("Workbook Loaded!")
        except:
            st.error("Could not read file. Please check permissions.")

    if st.session_state.all_dfs:
        st.session_state.scope = st.selectbox("Scope:", ["Analyze All Sheets (Join/Compare)"] + list(st.session_state.all_dfs.keys()))

# MAIN INTERFACE
if st.session_state.all_dfs:
    with st.expander("👀 Data Preview"):
        if st.session_state.scope == "Analyze All Sheets (Join/Compare)":
            st.write("Workbook Summary:")
            st.json({k: list(v.columns) for k, v in st.session_state.all_dfs.items()})
        else:
            st.dataframe(st.session_state.all_dfs[st.session_state.scope].head(10))

    # Using a form ensures the "Enter" key triggers the submit button
    with st.form("query_form"):
        user_query = st.text_input("Ask a question:")
        submitted = st.form_submit_button("Analyze")

    if submitted and user_query:
        context = build_context(st.session_state.all_dfs, st.session_state.scope)
        with st.spinner("AI analyzing..."):
            code = get_analysis_code(user_query, context)
            if code:
                try:
                    exec_globals = {'pd': pd, 'dfs': st.session_state.all_dfs}
                    exec_locals = {}
                    exec(code, exec_globals, exec_locals)
                    calc_res = exec_locals.get('result', "No result")
                    
                    answer = generate_natural_answer(user_query, calc_res)
                    
                    st.markdown(f"### Answer\n{answer}")
                    
                    if isinstance(calc_res, (pd.DataFrame, pd.Series)):
                        st.dataframe(calc_res)
                        
                    with st.expander("Logic"):
                        st.code(code)
                except Exception as e:
                    st.error(f"Execution Error: {e}")
                    st.code(code)
else:
    st.info("Load an Excel file in the sidebar to begin.")
