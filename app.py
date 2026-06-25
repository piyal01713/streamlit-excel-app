import pandas as pd
import streamlit as st
from anthropic import Anthropic
import re
import base64
import io
import requests

# ------------------------------------------------
# 1. CONFIG & INITIALIZATION
# ------------------------------------------------

st.set_page_config(page_title="Excel Insights Pro", layout="wide")

# Using the specific model IDs from your previous successful code versions
MODELS = {
    "Claude 4.5 Haiku": "claude-haiku-4-5-20251001",
    "Claude 4.5 Sonnet": "claude-sonnet-4-5-20250929",
}

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
# 2. HELPERS (URL & DATA CONTEXT)
# ------------------------------------------------

def get_direct_download_link(url):
    try:
        # Google Sheets
        if "docs.google.com/spreadsheets" in url:
            file_id = re.search(r'/d/([^/]+)', url).group(1)
            return f'https://docs.google.com/spreadsheets/d/{file_id}/export?format=xlsx'
        # Google Drive File
        elif "drive.google.com" in url:
            file_id = re.search(r'd/([^/]+)', url).group(1)
            return f'https://drive.google.com/uc?export=download&id={file_id}'
        # OneDrive / Sharepoint
        elif "1drv.ms" in url or "onedrive.live.com" in url:
            encoded_url = base64.b64encode(url.encode()).decode().replace('+', '-').replace('/', '_').rstrip('=')
            return f"https://api.onedrive.com/v1.0/shares/u!{encoded_url}/root/content"
        return url
    except Exception:
        return url

def build_context(dfs_dict, scope, depth):
    context = (
        "CRITICAL: DataFrames have NO HEADERS (columns are currently 0, 1, 2...). "
        "The real column names are inside one of the rows. You MUST find and set them before analysis.\n"
        f"ANALYSIS DEPTH: {depth}.\n\n"
    )
    
    selected_sheets = dfs_dict.keys() if scope == "Analyze All Sheets (Join/Compare)" else [scope]
    for name in selected_sheets:
        df = dfs_dict[name]
        context += f"### SHEET: '{name}'\n"
        # Convert to CSV for token efficiency and structure preservation
        if depth == "Quick (Top 100 Rows)":
            context += df.head(100).to_csv(index=False, header=False)
        else:
            context += df.to_csv(index=False, header=False)
        context += "\n---\n"
    return context

# ------------------------------------------------
# 3. AI LOGIC
# ------------------------------------------------

def get_analysis_code(user_query, context, model_id):
    system_prompt = (
        "You are a Senior Python Data Expert. You are working with a dictionary of DataFrames called `dfs`.\n"
        "DataFrames have integer columns (0, 1, 2) initially.\n\n"
        "MANDATORY CLEANING STEPS IN YOUR CODE:\n"
        "1. Identify the header row index by searching for keywords (e.g. 'Project', 'Name', etc).\n"
        "2. Clean the DF: `df.columns = df.iloc[idx].astype(str).str.strip()` then `df = df.iloc[idx+1:]`.\n"
        "3. SAFE INDEXING: Always use `row.iloc[0]` instead of `row[0]` to avoid KeyErrors after column renaming.\n"
        "4. SEARCHING: Use `.str.contains('keyword', na=False, case=False)` to find specific data like 'Ipex'.\n"
        "5. RESULT: Store your final answer (value, string, or DataFrame) in a variable named `result`.\n"
        "RESPOND ONLY WITH CODE."
    )
    try:
        response = client.messages.create(
            model=model_id,
            max_tokens=2500,
            system=system_prompt,
            messages=[{"role": "user", "content": f"Context:\n{context}\n\nQuery: {user_query}"}]
        )
        return re.sub(r"```python\n|```", "", response.content[0].text.strip())
    except Exception as e:
        st.error(f"AI Error: {e}")
        return None

def generate_natural_answer(user_query, execution_result, model_id):
    try:
        response = client.messages.create(
            model=model_id,
            max_tokens=800,
            system="Summarize data results clearly. No preamble. Use bolding.",
            messages=[{"role": "user", "content": f"User asked: {user_query}\nRaw Result: {execution_result}"}]
        )
        return response.content[0].text
    except:
        return f"Result: {execution_result}"

# ------------------------------------------------
# 4. STREAMLIT UI
# ------------------------------------------------

st.title("📊 Excel Insights Pro")

with st.sidebar:
    st.header("1. Intelligence Settings")
    selected_model_name = st.selectbox("Model:", list(MODELS.keys()), index=1)
    target_model_id = MODELS[selected_model_name]
    
    st.divider()
    st.header("2. Analysis Depth")
    analysis_depth = st.radio("Vision Range:", ["Quick (Top 100 Rows)", "Full Dataset (All Rows)"], 
                             help="Use 'Full Dataset' if headers/data are past row 20.")

    st.divider()
    st.header("3. Load Data")
    input_method = st.radio("Source:", ["Upload File", "Cloud Link"])
    
    raw_data = None
    if input_method == "Upload File":
        # Supports both xlsx and xlsm
        uploaded_file = st.file_uploader("Upload Excel", type=["xlsx", "xlsm"])
        if uploaded_file: raw_data = uploaded_file.read()
    else:
        url_input = st.text_input("Paste Google/OneDrive Link:")
        if url_input:
            with st.spinner("Fetching cloud file..."):
                dl_link = get_direct_download_link(url_input)
                try:
                    res = requests.get(dl_link, timeout=25)
                    raw_data = res.content
                except Exception as e:
                    st.error(f"Download failed: {e}")

    if raw_data:
        try:
            # Using openpyxl engine for .xlsm support and header=None for buried headers
            excel_file = pd.ExcelFile(io.BytesIO(raw_data), engine='openpyxl')
            st.session_state.all_dfs = {s: excel_file.parse(s, header=None) for s in excel_file.sheet_names}
            st.success("Workbook Loaded!")
        except Exception as e: 
            st.error(f"Read Error: {e}")

    if st.session_state.all_dfs:
        st.session_state.scope = st.selectbox("Scope:", ["Analyze All Sheets (Join/Compare)"] + list(st.session_state.all_dfs.keys()))

# MAIN INTERFACE
if st.session_state.all_dfs:
    with st.expander("👀 View Raw Structure (First 50 Rows)"):
        # Show first sheet by default or selected sheet
        preview_key = list(st.session_state.all_dfs.keys())[0] if st.session_state.scope == "Analyze All Sheets (Join/Compare)" else st.session_state.scope
        st.dataframe(st.session_state.all_dfs[preview_key].head(50))

    with st.form("chat_form"):
        user_query = st.text_input("Ask a question about the data:")
        submitted = st.form_submit_button("Run Analysis")

    if submitted and user_query:
        context = build_context(st.session_state.all_dfs, st.session_state.scope, analysis_depth)
        with st.spinner("AI Cleaning & Analyzing..."):
            code = get_analysis_code(user_query, context, target_model_id)
            if code:
                try:
                    exec_globals = {'pd': pd, 'dfs': st.session_state.all_dfs}
                    exec_locals = {}
                    exec(code, exec_globals, exec_locals)
                    calc_res = exec_locals.get('result', "No result variable created.")
                    
                    answer = generate_natural_answer(user_query, calc_res, target_model_id)
                    st.markdown("### 💡 AI Findings")
                    st.markdown(answer)
                    
                    if isinstance(calc_res, (pd.DataFrame, pd.Series)):
                        st.dataframe(calc_res)
                    with st.expander("View Logic"):
                        st.code(code)
                except Exception as e:
                    st.error(f"Execution Error: {e}")
                    st.info("Tip: Try switching to 'Full Dataset' if headers are buried deep.")
                    st.code(code)
else:
    st.info("Upload an Excel file or paste a Google Sheet link in the sidebar to begin.")
