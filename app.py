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
            context += f"Sample Data:\n{df.head(10).to_string(index=False)}\n---\n"
    else:
        df = dfs_dict[scope]
        context = f"SCOPE: SHEET '{scope}'. Use `dfs['{scope}']`.\n"
        context += f"Columns: {list(df.columns)}\nSample Data:\n{df.head(10).to_string(index=False)}"
    return context

# ------------------------------------------------
# AI LOGIC (FOR GENERATING CODE)
# ------------------------------------------------

def get_analysis_code(user_query, context):
    system_prompt = (
        "You are a Python data expert. Write code for a dictionary of DataFrames named `dfs`. "
        "The final result MUST be in a variable named `result`. "
        "Search all rows and columns for value matches. "
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

# ------------------------------------------------
# AI LOGIC (FOR CLEAN ANSWERS)
# ------------------------------------------------

def generate_natural_answer(user_query, execution_result):
    """
    This function was updated to provide the 'Cleaner' answer format.
    """
    clean_system_prompt = (
        "You are a concise data assistant. Your job is to summarize the data result for the user. "
        "RULES FOR YOUR ANSWER:\n"
        "1. Be direct. No preamble like 'Based on the data...' or 'I found that...'.\n"
        "2. Use **bolding** for names, totals, and key identifiers.\n"
        "3. Use bullet points if listing multiple items.\n"
        "4. If something is not found, state it simply.\n"
        "5. Use professional, clean Markdown formatting."
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
            excel_file = pd.ExcelFile(io.BytesIO(raw_data))
            st.session_state.all_dfs = {sheet: excel_file.parse(sheet) for sheet in excel_file.sheet_names}
            st.success("File Loaded!")
        except: st.error("Read Error")

    if st.session_state.all_dfs:
        st.session_state.scope = st.selectbox("Scope:", ["Analyze All Sheets (Join/Compare)"] + list(st.session_state.all_dfs.keys()))

# MAIN INTERFACE
if st.session_state.all_dfs:
    with st.expander("👀 View Data"):
        st.dataframe(st.session_state.all_dfs[st.session_state.scope] if st.session_state.scope != "Analyze All Sheets (Join/Compare)" else st.session_state.all_dfs)

    with st.form("chat"):
        user_query = st.text_input("Ask a question:")
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
                    calc_res = exec_locals.get('result', "No result")
                    
                    # GET CLEAN ANSWER
                    answer = generate_natural_answer(user_query, calc_res)
                    
                    st.markdown("### 💡 Result")
                    st.markdown(answer) # This renders the clean Markdown
                    
                    if isinstance(calc_res, (pd.DataFrame, pd.Series)):
                        st.dataframe(calc_res)
                        
                    with st.expander("Debug Logic"):
                        st.code(code)
                except Exception as e:
                    st.error(f"Error: {e}")
else:
    st.info("Upload a file to begin.")
