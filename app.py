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

# Using the model ID you requested
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
# AI & DATA LOGIC
# ------------------------------------------------

def build_context(dfs_dict, scope):
    if scope == "Analyze All Sheets (Join/Compare)":
        context = "SCOPE: ALL SHEETS. Dictionary: `dfs`.\n"
        for name, df in dfs_dict.items():
            context += f"Sheet: '{name}' | Columns: {list(df.columns)}\n"
            # Show more rows in sample to help AI understand messy headers
            context += f"Sample Data (First 10 rows):\n{df.head(10).to_string(index=False)}\n---\n"
    else:
        df = dfs_dict[scope]
        context = f"SCOPE: SHEET '{scope}'. Use `dfs['{scope}']`.\n"
        context += f"Columns: {list(df.columns)}\nSample Data (First 10 rows):\n{df.head(10).to_string(index=False)}"
    return context

def get_analysis_code(user_query, context):
    # IMPROVED SYSTEM PROMPT: Explicitly tells AI to look inside the data, not just headers.
    system_prompt = (
        "You are a Python data analyst. Write code for a dictionary of DataFrames named `dfs`. "
        "The final result MUST be stored in a variable named `result`. "
        "IMPORTANT: When searching for specific words or values (like 'Ipex'), search across ALL columns and ALL rows. "
        "Use case-insensitive matching and strip whitespace. "
        "If the query asks if something exists, return a clear confirmation or the relevant rows. "
        "RESPOND ONLY WITH CLEAN PYTHON CODE. NO EXPLANATION."
    )
    try:
        response = client.messages.create(
            model=MODEL,
            max_tokens=1000,
            system=system_prompt,
            messages=[{"role": "user", "content": f"Context:\n{context}\n\nUser Query: {user_query}"}]
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
            system="Explain the technical data result in a friendly, natural way. If the result is a dataframe, summarize the key finding.",
            messages=[{"role": "user", "content": f"User Query: {user_query}\nRaw Data Result: {execution_result}"}]
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
            with st.spinner("Connecting to cloud..."):
                dl_link = get_direct_download_link(url_input)
                try:
                    res = requests.get(dl_link, timeout=15)
                    raw_data = res.content
                except Exception as e:
                    st.error(f"Failed to fetch: {e}")

    if raw_data:
        try:
            excel_file = pd.ExcelFile(io.BytesIO(raw_data))
            # Save to session state so it doesn't vanish on refresh
            st.session_state.all_dfs = {sheet: excel_file.parse(sheet) for sheet in excel_file.sheet_names}
            st.success("File Loaded Successfully!")
        except:
            st.error("Error: Could not read Excel file. Check if it is password protected.")

    if st.session_state.all_dfs:
        st.session_state.scope = st.selectbox(
            "Select Analysis Scope:", 
            ["Analyze All Sheets (Join/Compare)"] + list(st.session_state.all_dfs.keys())
        )

# MAIN INTERFACE
if st.session_state.all_dfs:
    # 1. Data Preview
    with st.expander("👀 View Current Data"):
        if st.session_state.scope == "Analyze All Sheets (Join/Compare)":
            st.write("Current Sheets & Columns:")
            st.json({k: list(v.columns) for k, v in st.session_state.all_dfs.items()})
        else:
            st.dataframe(st.session_state.all_dfs[st.session_state.scope])

    # 2. Query Input (Inside form for Enter-key support)
    with st.form("chat_interface"):
        user_query = st.text_input("Ask a question about your data (e.g., 'Is Ipex mentioned in the project row?'):")
        submitted = st.form_submit_button("Analyze")

    # 3. Execution
    if submitted and user_query:
        context = build_context(st.session_state.all_dfs, st.session_state.scope)
        
        with st.spinner("Analyzing spreadsheet..."):
            code = get_analysis_code(user_query, context)
            
            if code:
                try:
                    # Environment setup
                    exec_globals = {'pd': pd, 'dfs': st.session_state.all_dfs}
                    exec_locals = {}
                    
                    # Run the AI-generated code
                    exec(code, exec_globals, exec_locals)
                    calc_res = exec_locals.get('result', "No result found.")
                    
                    # Explain results
                    answer = generate_natural_answer(user_query, calc_res)
                    
                    st.markdown("---")
                    st.markdown(f"### Answer\n{answer}")
                    
                    # If the result is data, show the table
                    if isinstance(calc_res, (pd.DataFrame, pd.Series)) and not (isinstance(calc_res, pd.DataFrame) and calc_res.empty):
                        st.dataframe(calc_res)
                    
                    # Show logic for debugging
                    with st.expander("View AI Logic (Python)"):
                        st.code(code)
                        
                except Exception as e:
                    st.error(f"Execution Error: {e}")
                    st.info("The AI generated code that didn't quite work. Try re-phrasing your question.")
                    with st.expander("See Failed Code"):
                        st.code(code)
else:
    st.info("Waiting for data. Please upload an Excel file or provide a link in the sidebar.")
