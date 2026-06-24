import pandas as pd
import streamlit as st
from anthropic import Anthropic
import json
import re
import base64
import io
import requests
import openpyxl  # Added for Excel formatting detection

# ------------------------------------------------
# CONFIG & INITIALIZATION
# ------------------------------------------------

st.set_page_config(page_title="Excel Insights", layout="wide")

MODEL = "claude-haiku-4-5-20251001" 

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
# HELPERS
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

def build_context(dfs_dict, scope):
    if scope == "Analyze All Sheets (Join/Compare)":
        context = "SCOPE: ALL SHEETS. Dictionary: `dfs`.\n"
        for name, df in dfs_dict.items():
            context += f"Sheet: '{name}' | Columns: {list(df.columns)}\n"
            context += f"Sample Data:\n{df.head(5).to_string(index=False)}\n---\n"
    else:
        df = dfs_dict[scope]
        context = f"SCOPE: SHEET '{scope}'. Use `dfs['{scope}']`.\n"
        context += f"Columns: {list(df.columns)}\nSample Data:\n{df.head(10).to_string(index=False)}"
    return context

# ------------------------------------------------
# AI LOGIC
# ------------------------------------------------

def get_analysis_code(user_query, context):
    system_prompt = (
        "You are a Python data expert. Write code for a dictionary of DataFrames named `dfs`. "
        "The final result MUST be in a variable named `result`. "
        "Search all rows and columns for value matches. "
        "If column names contain 'Unnamed', the code should attempt to find the correct data dynamically. "
        "RESPOND ONLY WITH CODE."
    )
    try:
        response = client.messages.create(
            model=MODEL,
            max_tokens=1000,
            system=system_prompt,
            messages=[{"role": "user", "content": f"Context:\n{context}\n\nQuery: {user_query}"}]
        )
        return re.sub(r"```python\n|```", "", response.content[0].text.strip())
    except Exception as e:
        st.error(f"AI Error: {e}")
        return None

def generate_natural_answer(user_query, execution_result):
    clean_system_prompt = (
        "You are a concise data assistant. Summarize the data result clearly.\n"
        "1. No preamble (don't say 'Here is the answer').\n"
        "2. Use **bolding** for key identifiers and values.\n"
        "3. Use bullet points for lists.\n"
        "4. Be direct and professional."
    )
    try:
        response = client.messages.create(
            model=MODEL,
            max_tokens=500,
            system=clean_system_prompt,
            messages=[{"role": "user", "content": f"User asked: {user_query}\nRaw Result: {execution_result}"}]
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
        if uploaded_file: raw_data = uploaded_file.read()
    else:
        url_input = st.text_input("Paste Shareable Link:")
        if url_input:
            with st.spinner("Connecting..."):
                dl_link = get_direct_download_link(url_input)
                try:
                    res = requests.get(dl_link, timeout=15)
                    raw_data = res.content
                except: st.error("Link Error")

    if raw_data:
        try:
            # --- UPDATED LOGIC TO DETECT FROZEN HEADERS ---
            # 1. Load the workbook structure with openpyxl
            wb = openpyxl.load_workbook(io.BytesIO(raw_data), read_only=True)
            temp_dfs = {}

            for sheet_name in wb.sheetnames:
                ws = wb[sheet_name]
                
                # Default header is the first row (index 0)
                header_idx = 0 
                
                # Check for frozen panes
                # If Excel freezes Row 1, freeze_panes is 'A2'. 
                # If Excel freezes Row 1 & 2 (headers), freeze_panes is 'A3'.
                if ws.freeze_panes and ws.freeze_panes != 'A1':
                    row_match = re.search(r'\d+', ws.freeze_panes)
                    if row_match:
                        # We assume the row immediately ABOVE the freeze line is the header
                        header_idx = int(row_match.group()) - 2
                        if header_idx < 0: header_idx = 0
                
                # 2. Parse with pandas using the detected header row
                # We use io.BytesIO(raw_data) again because openpyxl was in read_only
                df = pd.read_excel(io.BytesIO(raw_data), sheet_name=sheet_name, header=header_idx)
                temp_dfs[sheet_name] = df

            st.session_state.all_dfs = temp_dfs
            st.success("File Loaded (Headers detected!)")
        except Exception as e: 
            st.error(f"Read Error: {e}")

    if st.session_state.all_dfs:
        st.session_state.scope = st.selectbox("Scope:", ["Analyze All Sheets (Join/Compare)"] + list(st.session_state.all_dfs.keys()))

# MAIN INTERFACE
if st.session_state.all_dfs:
    
    with st.expander("👀 View Data Preview"):
        if st.session_state.scope == "Analyze All Sheets (Join/Compare)":
            st.write("Summary of all sheets:")
            for sheet_name, df in st.session_state.all_dfs.items():
                st.markdown(f"**Sheet:** `{sheet_name}`")
                st.dataframe(df.head(3))
        else:
            st.dataframe(st.session_state.all_dfs[st.session_state.scope])

    with st.form("chat_form"):
        user_query = st.text_input("Ask a question about your data:")
        submitted = st.form_submit_button("Analyze")

    if submitted and user_query:
        context = build_context(st.session_state.all_dfs, st.session_state.scope)
        with st.spinner("Processing..."):
            code = get_analysis_code(user_query, context)
            if code:
                try:
                    exec_globals = {'pd': pd, 'dfs': st.session_state.all_dfs}
                    exec_locals = {}
                    exec(code, exec_globals, exec_locals)
                    calc_res = exec_locals.get('result', "No result variable was created.")
                    
                    answer = generate_natural_answer(user_query, calc_res)
                    
                    st.markdown("### 💡 Result")
                    st.markdown(answer)
                    
                    if isinstance(calc_res, (pd.DataFrame, pd.Series)):
                        if not calc_res.empty:
                            st.dataframe(calc_res)
                        
                    with st.expander("Technical Logic"):
                        st.code(code)
                except Exception as e:
                    st.error(f"Error executing code: {e}")
                    with st.expander("View Code"):
                        st.code(code)
else:
    st.info("Please upload an Excel file in the sidebar to begin.")
