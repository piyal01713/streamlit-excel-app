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
        # 1. Google Sheets (Native) -> Export as XLSX to get all sheets
        if "docs.google.com/spreadsheets" in url:
            file_id = re.search(r'/d/([^/]+)', url).group(1)
            return f'https://docs.google.com/spreadsheets/d/{file_id}/export?format=xlsx'

        # 2. Google Drive (Uploaded .xlsx)
        elif "drive.google.com" in url:
            file_id = re.search(r'd/([^/]+)', url).group(1)
            return f'https://drive.google.com/uc?export=download&id={file_id}'

        # 3. OneDrive Personal
        elif "1drv.ms" in url or "onedrive.live.com" in url:
            encoded_url = base64.b64encode(url.encode()).decode().replace('+', '-').replace('/', '_').rstrip('=')
            return f"https://api.onedrive.com/v1.0/shares/u!{encoded_url}/root/content"

        # 4. OneDrive Business / SharePoint
        elif "sharepoint.com" in url:
            return url.split("?")[0] + "?download=1" if "?" in url else url + "?download=1"

        return url
    except Exception as e:
        st.error(f"Error parsing link format: {e}")
        return None

# ------------------------------------------------
# GROUNDED DATA OPERATIONS
# ------------------------------------------------

def build_data_context(df):
    return f"""
DATASET STRUCTURE:
Columns: {list(df.columns)}
Column Types: {df.dtypes.to_string()}
Sample Rows:
{df.head(10).to_string(index=False)}
Row Count: {len(df)}
"""

def run_pandas_operation(df, op):
    try:
        operation = op.get("operation")
        if operation == "groupby_sum":
            return df.groupby(op["group"])[op["column"]].sum().to_string()
        if operation == "groupby_mean":
            return df.groupby(op["group"])[op["column"]].mean().to_string()
        if operation == "filter_equals":
            return df[df[op["column"]].astype(str) == str(op["value"])].head(20).to_string(index=False)
        if operation == "filter_contains":
            return df[df[op["column"]].astype(str).str.contains(str(op["value"]), case=False, na=False)].head(20).to_string(index=False)
        if operation == "top_n":
            return df.nlargest(op["n"], op["column"]).to_string(index=False)
        if operation == "describe":
            return df.describe(include="all").to_string()
        return "ERROR: Unsupported operation"
    except Exception as e:
        return f"Execution error: {e}"

# ------------------------------------------------
# LLM LOGIC
# ------------------------------------------------

def get_llm_response(user_query, data_context):
    system_prompt = """You are an Excel AI Assistant. Analyze queries and output ONLY JSON. No conversation."""
    try:
        response = client.messages.create(
            model=MODEL,
            max_tokens=1000,
            system=system_prompt,
            messages=[{"role": "user", "content": f"Context: {data_context}\n\nQuestion: {user_query}"}]
        )
        raw_text = response.content[0].text.strip()
        if "```" in raw_text:
            raw_text = re.sub(r"```[^\n]*\n?", "", raw_text).replace("```", "").strip()
        return json.loads(raw_text)
    except Exception as e:
        return None

def generate_natural_answer(user_query, execution_result):
    try:
        response = client.messages.create(
            model=MODEL,
            max_tokens=500,
            system="Directly and naturally answer based on the data result.",
            messages=[{"role": "user", "content": f"Query: {user_query}\nResult: {execution_result}"}]
        )
        return response.content[0].text
    except Exception as e:
        return f"Error: {e}"

# ------------------------------------------------
# STREAMLIT UI
# ------------------------------------------------

st.title("Excel Insights")

with st.sidebar:
    st.header("1. Data Source")
    input_method = st.radio("Select Method:", ["Upload File", "Cloud Link"])
    
    raw_data = None
    file_ext = ""

    # Step 1: Get the raw bytes
    if input_method == "Upload File":
        uploaded_file = st.file_uploader("Upload CSV or Excel", type=["csv", "xlsx"])
        if uploaded_file:
            file_ext = "csv" if uploaded_file.name.endswith(".csv") else "xlsx"
            raw_data = uploaded_file.read()
    else:
        url_input = st.text_input("Paste Shareable Link:")
        if url_input:
            with st.spinner("Fetching cloud file..."):
                dl_link = get_direct_download_link(url_input)
                try:
                    response = requests.get(dl_link)
                    raw_data = response.content
                    # Determine extension from URL or content-type
                    file_ext = "xlsx" if "xlsx" in dl_link or "spreadsheet" in url_input else "csv"
                except:
                    st.error("Failed to fetch. Check link permissions.")

    # Step 2: Handle Sheet Selection
    df = None
    if raw_data:
        if file_ext == "xlsx":
            try:
                excel_file = pd.ExcelFile(io.BytesIO(raw_data))
                sheet_names = excel_file.sheet_names
                
                selected_sheet = sheet_names[0]
                if len(sheet_names) > 1:
                    st.success(f"Found {len(sheet_names)} sheets!")
                    selected_sheet = st.selectbox("Select Sheet to Analyze:", sheet_names)
                
                df = excel_file.parse(selected_sheet)
            except Exception as e:
                st.error(f"Error parsing Excel: {e}")
        else:
            try:
                df = pd.read_csv(io.BytesIO(raw_data))
            except Exception as e:
                st.error(f"Error parsing CSV: {e}")

# MAIN INTERFACE
if df is not None:
    st.subheader(f"📊 Preview: {selected_sheet if 'selected_sheet' in locals() else 'Data'}")
    st.dataframe(df.head(5), use_container_width=True)

    user_query = st.text_input("Ask a question about this sheet:")
    
    if user_query:
        context = build_data_context(df)
        with st.spinner("Analyzing..."):
            op = get_llm_response(user_query, context)
            if op:
                raw_result = run_pandas_operation(df, op)
                answer = generate_natural_answer(user_query, raw_result)
                st.markdown(f"### Assistant Answer\n{answer}")
                with st.expander("Technical Trace"):
                    st.json(op)
                    st.code(raw_result)
            else:
                st.warning("I couldn't map that to a data operation. Try rephrasing.")
else:
    st.info("Upload a file or provide a link in the sidebar to begin.")
