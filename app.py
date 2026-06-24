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

# Initialize Session State
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
            context += f"Sample Data:\n{df.head(2).to_string(index=False)}\n---\n"
    else:
        df = dfs_dict[scope]
        context = f"SCOPE: SHEET '{scope}'. Use `dfs['{scope}']`.\n"
        context += f"Columns: {list(df.columns)}\nSample Data:\n{df.head(5).to_string(index=False)}"
    return context

def get_analysis_code(user_query, context):
    system_prompt = "Write Python code that uses a dictionary of DataFrames named `dfs`. The final answer must be assigned to a variable named `result`. Output ONLY the code, no explanation."
    try:
        response = client.messages.create(
            model=MODEL,
            max_tokens=1000,
            system=system_prompt,
            messages=[{"role": "user", "content": f"Context: {context}\n\nUser Question: {user_query}"}]
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
            system="The user asked a question and got a raw data result. Explain the result clearly in 1-2 sentences.",
            messages=[{"role": "user", "content": f"Question: {user_query}\nRaw Result: {execution_result}"}]
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
                    st.error(f"Connection failed: {e}")

    # Process data into session state
    if raw_data:
        try:
            excel_file = pd.ExcelFile(io.BytesIO(raw_data))
            dfs = {sheet: excel_file.parse(sheet) for sheet in excel_file.sheet_names}
            st.session_state.all_dfs = dfs
            st.success("File Loaded!")
        except:
            st.error("Error reading file. Check link permissions.")

    if st.session_state.all_dfs:
        st.session_state.scope = st.selectbox(
            "Select Analysis Scope:", 
            ["Analyze All Sheets (Join/Compare)"] + list(st.session_state.all_dfs.keys())
        )

# ------------------------------------------------
# MAIN INTERFACE
# ------------------------------------------------

if st.session_state.all_dfs:
    # 1. Preview
    with st.expander("👀 Data Preview"):
        if st.session_state.scope == "Analyze All Sheets (Join/Compare)":
            st.json({k: list(v.columns) for k, v in st.session_state.all_dfs.items()})
        else:
            st.dataframe(st.session_state.all_dfs[st.session_state.scope].head(10))

    # 2. Input and Submit
    with st.form("chat_form"):
        user_query = st.text_input("Ask a question about your data:")
        submitted = st.form_submit_button("Run Analysis")

    # 3. Execution Logic
    if submitted and user_query:
        context = build_context(st.session_state.all_dfs, st.session_state.scope)
        
        with st.spinner("AI is thinking..."):
            code = get_analysis_code(user_query, context)
            
            if code:
                try:
                    # Setup environment for code execution
                    exec_globals = {'pd': pd, 'dfs': st.session_state.all_dfs}
                    exec_locals = {}
                    
                    exec(code, exec_globals, exec_locals)
                    
                    # Get the result variable defined in the generated code
                    calc_res = exec_locals.get('result', "No 'result' variable produced.")
                    
                    # Generate natural language summary
                    answer = generate_natural_answer(user_query, calc_res)
                    
                    st.markdown("---")
                    st.markdown(f"### 💡 Answer\n{answer}")
                    
                    # Show result visually if it's a dataframe or chart-like object
                    if isinstance(calc_res, (pd.DataFrame, pd.Series)):
                        st.dataframe(calc_res)
                    
                    with st.expander("Show Technical Logic"):
                        st.code(code, language="python")
                        
                except Exception as e:
                    st.error(f"Code Execution Error: {e}")
                    st.code(code)
else:
    st.info("Please upload a file or paste a link in the sidebar to begin.")
