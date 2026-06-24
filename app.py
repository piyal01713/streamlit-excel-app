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

# Claude 3.5 Haiku: High speed, low cost, excellent at Python logic
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
    """Converts sharing links to direct download streams."""
    try:
        # 1. Google Sheets (Native) -> Export as XLSX to get all sheets
        if "docs.google.com/spreadsheets" in url:
            file_id = re.search(r'/d/([^/]+)', url).group(1)
            return f'https://docs.google.com/spreadsheets/d/{file_id}/export?format=xlsx'

        # 2. Google Drive (Uploaded File)
        elif "drive.google.com" in url:
            file_id = re.search(r'd/([^/]+)', url).group(1)
            return f'https://drive.google.com/uc?export=download&id={file_id}'

        # 3. OneDrive Personal
        elif "1drv.ms" in url or "onedrive.live.com" in url:
            encoded_url = base64.b64encode(url.encode()).decode().replace('+', '-').replace('/', '_').rstrip('=')
            return f"https://api.onedrive.com/v1.0/shares/u!{encoded_url}/root/content"

        # 4. OneDrive Business / SharePoint
        elif "sharepoint.com" in url:
            clean_url = url.split("?")[0]
            return f"{clean_url}?download=1"

        return url
    except Exception:
        return url

# ------------------------------------------------
# CONTEXT BUILDER (TOKEN OPTIMIZATION)
# ------------------------------------------------

def build_context(dfs_dict, scope):
    """Builds a metadata summary based on the chosen scope."""
    if scope == "Analyze All Sheets (Join/Compare)":
        context = "ANALYSIS SCOPE: ALL SHEETS. Use the dictionary `dfs` where keys are sheet names.\n\n"
        for name, df in dfs_dict.items():
            context += f"Sheet: '{name}' | Columns: {list(df.columns)} | Rows: {len(df)}\n"
            # Just 2 rows of sample to save tokens in multi-sheet mode
            context += f"Samples:\n{df.head(2).to_string(index=False)}\n---\n"
    else:
        df = dfs_dict[scope]
        context = f"ANALYSIS SCOPE: SINGLE SHEET ('{scope}'). Use `dfs['{scope}']`.\n\n"
        context += f"Columns: {list(df.columns)}\n"
        context += f"Detailed Samples:\n{df.head(5).to_string(index=False)}\n"
        context += f"Total Rows: {len(df)}"
    return context

# ------------------------------------------------
# AI AGENT LOGIC
# ------------------------------------------------

def get_analysis_code(user_query, context):
    system_prompt = """You are a Data Science Assistant. You have access to a dictionary of DataFrames named `dfs`.
    Write a Python snippet that produces a single variable named `result`. 
    The result can be a number, string, or DataFrame.
    RESPOND ONLY WITH THE CODE. No conversation, no code blocks."""

    try:
        response = client.messages.create(
            model=MODEL,
            max_tokens=1000,
            system=system_prompt,
            messages=[{"role": "user", "content": f"Data Context:\n{context}\n\nQuestion: {user_query}"}]
        )
        code = response.content[0].text.strip()
        # Clean potential markdown wrappers
        code = re.sub(r"```python\n|```", "", code)
        return code
    except Exception as e:
        st.error(f"AI Planner Error: {e}")
        return None

def generate_natural_answer(user_query, execution_result):
    try:
        response = client.messages.create(
            model=MODEL,
            max_tokens=500,
            system="Explain the data result clearly and concisely. If the result is a table, summarize the key findings.",
            messages=[{"role": "user", "content": f"User asked: {user_query}\nCalculation Result: {execution_result}"}]
        )
        return response.content[0].text
    except Exception as e:
        return f"Calculated Result: {execution_result}"

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
        if uploaded_file:
            raw_data = uploaded_file.read()
    else:
        url_input = st.text_input("Paste Shareable Link:", placeholder="OneDrive, GDrive, or GSheets")
        if url_input:
            with st.spinner("Fetching cloud file..."):
                dl_link = get_direct_download_link(url_input)
                try:
                    res = requests.get(dl_link, timeout=10)
                    raw_data = res.content
                    
                    # SECURITY CHECK: Detect HTML Login Pages
                    if raw_data.strip().startswith(b"<!DOCTYPE") or b"<html" in raw_data[:100].lower():
                        st.error("❌ Access Denied: This link leads to a login page.")
                        st.info("💡 **Fix:** Set the link permissions to **'Anyone with the link can view'**. Your organization is currently requiring a login, which prevents the AI from reaching the data.")
                        raw_data = None
                except Exception as e:
                    st.error(f"Connection Error: {e}")

    # Process Bytes into DataFrames
    if raw_data:
        try:
            excel_file = pd.ExcelFile(io.BytesIO(raw_data))
            for sheet in excel_file.sheet_names:
                all_dfs[sheet] = excel_file.parse(sheet)
            
            st.success(f"Workbook Loaded: {len(all_dfs)} Sheet(s)")
            
            st.header("2. Analysis Scope")
            options = ["Analyze All Sheets (Join/Compare)"] + list(all_dfs.keys())
            scope = st.selectbox("Scope:", options, help="Select 'All Sheets' to ask questions comparing data across different tabs.")
            
        except Exception as e:
            st.error(f"Error reading file: {e}")

# MAIN INTERFACE
if all_dfs:
    # 1. Preview
    with st.expander("👀 Data Preview"):
        if scope == "Analyze All Sheets (Join/Compare)":
            st.write("Current Workbook Structure:")
            summary = [{"Sheet": k, "Columns": ", ".join(list(v.columns)), "Rows": len(v)} for k, v in all_dfs.items()]
            st.table(summary)
        else:
            st.write(f"Previewing: **{scope}**")
            st.dataframe(all_dfs[scope].head(10), use_container_width=True)

    # 2. Chat
    st.subheader("💬 Ask Your Data")
    user_query = st.text_input("Ask a question about the data:", placeholder="e.g., 'Compare total sales between Sheet1 and Sheet2'")
    
    if user_query:
        context = build_context(all_dfs, scope)
        
        with st.spinner("AI is analyzing workbook..."):
            python_code = get_analysis_code(user_query, context)
            
        if python_code:
            try:
                # Prepare secure environment
                exec_globals = {'pd': pd, 'dfs': all_dfs}
                exec_locals = {}
                
                # Run the AI-generated logic
                exec(python_code, exec_globals, exec_locals)
                calc_result = exec_locals.get('result', "No result variable defined by AI.")
                
                # Generate natural language explanation
                with st.spinner("Finalizing answer..."):
                    answer = generate_natural_answer(user_query, calc_result)
                    st.markdown("---")
                    st.markdown(f"### Assistant Answer\n{answer}")
                
                # Debugging/Technical Expanders
                with st.expander("View AI Logic"):
                    st.code(python_code, language="python")
                    st.write("Raw Output:")
                    st.write(calc_result)
                    
            except Exception as e:
                st.error(f"Calculation Error: {e}")
                with st.expander("Show Failed Code"):
                    st.code(python_code)
else:
    # Landing State
    st.info("Waiting for data... Please upload an Excel file or paste a public link in the sidebar.")
    st.image("https://www.gstatic.com/images/branding/product/2x/sheets_2020q4_48dp.png", width=50)
