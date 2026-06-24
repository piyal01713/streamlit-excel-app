import pandas as pd
import streamlit as st
from anthropic import Anthropic
import json
import re
import base64

# ------------------------------------------------
# CONFIG & INITIALIZATION
# ------------------------------------------------

# CHANGED: Updated page title
st.set_page_config(page_title="Excel Insights", layout="wide")

# Using the requested model
MODEL = "claude-haiku-4-5-20251001"

api_key = st.secrets.get("ANTHROPIC_API_KEY")
if not api_key:
    st.error("Missing ANTHROPIC_API_KEY in Streamlit secrets!")
    st.stop()

client = Anthropic(api_key=api_key)

if "dataframes" not in st.session_state:
    st.session_state.dataframes = {}

# ------------------------------------------------
# HELPERS: URL CONVERSION
# ------------------------------------------------

def get_direct_download_link(url):
    """
    Transforms browser/share links into direct data streams.
    Works for Google Sheets, GDrive Files, OneDrive Personal, and OneDrive Business.
    """
    try:
        # 1. Google Sheets (Native format)
        if "docs.google.com/spreadsheets" in url:
            file_id = re.search(r'/d/([^/]+)', url).group(1)
            return f'https://docs.google.com/spreadsheets/d/{file_id}/export?format=csv'

        # 2. Google Drive (Uploaded .xlsx or .csv)
        elif "drive.google.com" in url:
            file_id = re.search(r'd/([^/]+)', url).group(1)
            return f'https://drive.google.com/uc?export=download&id={file_id}'

        # 3. OneDrive Personal
        elif "1drv.ms" in url or "onedrive.live.com" in url:
            encoded_url = base64.b64encode(url.encode()).decode().replace('+', '-').replace('/', '_').rstrip('=')
            return f"https://api.onedrive.com/v1.0/shares/u!{encoded_url}/root/content"

        # 4. OneDrive Business / SharePoint
        elif "sharepoint.com" in url:
            if "?" in url:
                return url.split("?")[0] + "?download=1"
            else:
                return url + "?download=1"

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
Column Types:
{df.dtypes.to_string()}

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
    system_prompt = """You are an Excel AI Assistant. Analyze queries and output ONLY JSON:
    - {"operation": "groupby_sum", "group": "col", "column": "col"}
    - {"operation": "groupby_mean", "group": "col", "column": "col"}
    - {"operation": "filter_equals", "column": "col", "value": "val"}
    - {"operation": "filter_contains", "column": "col", "value": "val"}
    - {"operation": "top_n", "column": "col", "n": 5}
    - {"operation": "describe"}"""

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
        st.error(f"Planner Error: {e}")
        return None

def generate_natural_answer(user_query, execution_result):
    try:
        response = client.messages.create(
            model=MODEL,
            max_tokens=500,
            system="Answer the user question based on the provided data results. Be concise and professional.",
            messages=[{"role": "user", "content": f"Query: {user_query}\nResult: {execution_result}"}]
        )
        return response.content[0].text
    except Exception as e:
        return f"Error: {e}"

# ------------------------------------------------
# STREAMLIT UI
# ------------------------------------------------

# CHANGED: Updated Title
st.title("Excel Insights")

with st.sidebar:
    st.header("1. Data Source")
    input_method = st.radio("Select Method:", ["Upload File", "Cloud Link (GDrive/OneDrive)"])
    
    df = None

    if input_method == "Upload File":
        uploaded_file = st.file_uploader("Upload CSV or Excel", type=["csv", "xlsx"])
        if uploaded_file:
            if uploaded_file.name.endswith(".csv"):
                df = pd.read_csv(uploaded_file)
            else:
                df = pd.read_excel(uploaded_file)
    else:
        url_input = st.text_input("Paste Shareable Link:", placeholder="https://docs.google.com/...")
        st.caption("⚠️ Ensure link is set to 'Anyone with the link can view'")
        if url_input:
            with st.spinner("Connecting to cloud..."):
                dl_link = get_direct_download_link(url_input)
                try:
                    df = pd.read_csv(dl_link)
                except:
                    try:
                        df = pd.read_excel(dl_link)
                    except Exception as e:
                        st.error("Failed to load. Check link permissions.")

# MAIN INTERFACE
if df is not None:
    st.session_state.dataframes["active"] = df
    
    st.subheader("📊 Data Preview")
    st.dataframe(df.head(5), use_container_width=True)

    st.subheader("💬 Ask a Question")
    user_query = st.text_input("Example: 'What are the top 5 sales?'")
    
    if user_query:
        context = build_data_context(df)
        
        with st.spinner("Thinking..."):
            op = get_llm_response(user_query, context)
            
        if op:
            with st.spinner("Calculating..."):
                raw_result = run_pandas_operation(df, op)
                final_answer = generate_natural_answer(user_query, raw_result)
                
            st.markdown("---")
            st.markdown(f"### Assistant Answer\n{final_answer}")
            
            with st.expander("Show Technical Details"):
                st.write("**Planned Operation:**")
                st.json(op)
                st.write("**Raw Data Result:**")
                st.code(raw_result)
else:
    st.info("Please provide a data source in the sidebar to begin.")
