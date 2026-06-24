import pandas as pd
import streamlit as st
from anthropic import Anthropic
import json
import re
import base64  # Needed for OneDrive URL conversion

# ------------------------------------------------
# CONFIG
# ------------------------------------------------

st.set_page_config(page_title="Excel AI Agent (Grounded)", layout="wide")

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
    """Detects provider and converts sharing link to direct download link."""
    try:
        # Google Drive
        if "drive.google.com" in url:
            # Extract file ID
            file_id_match = re.search(r'd/([^/]+)', url)
            if file_id_match:
                return f'https://drive.google.com/uc?export=download&id={file_id_match.group(1)}'
        
        # OneDrive (Personal)
        elif "1drv.ms" in url or "onedrive.live.com" in url:
            # OneDrive requires base64 encoding the share URL
            base64_url = base64.b64encode(url.encode()).decode().replace('+', '-').replace('/', '_').rstrip('=')
            return f"https://api.onedrive.com/v1.0/shares/u!{base64_url}/root/content"
            
        return url # Return as is if it doesn't match known cloud patterns
    except Exception as e:
        st.error(f"Error parsing URL: {e}")
        return None

# ------------------------------------------------
# SESSION STATE
# ------------------------------------------------

if "dataframes" not in st.session_state:
    st.session_state.dataframes = {}

# ------------------------------------------------
# BUILD SAFE DATA CONTEXT
# ------------------------------------------------

def build_data_context(df):
    return f"""
DATASET STRUCTURE:

Columns:
{list(df.columns)}

Column Types:
{df.dtypes.to_string()}

Sample Rows:
{df.head(10).to_string(index=False)}

Row Count: {len(df)}
"""

# ------------------------------------------------
# SAFE EXECUTOR
# ------------------------------------------------

def run_pandas_operation(df, op):
    try:
        operation = op.get("operation")

        if operation == "groupby_sum":
            res = df.groupby(op["group"])[op["column"]].sum()
            return res.to_string()

        if operation == "groupby_mean":
            res = df.groupby(op["group"])[op["column"]].mean()
            return res.to_string()

        if operation == "filter_equals":
            res = df[df[op["column"]].astype(str) == str(op["value"])]
            return res.head(20).to_string(index=False)

        if operation == "filter_contains":
            res = df[df[op["column"]].astype(str).str.contains(str(op["value"]), case=False, na=False)]
            return res.head(20).to_string(index=False)

        if operation == "top_n":
            res = df.nlargest(op["n"], op["column"])
            return res.to_string(index=False)

        if operation == "describe":
            return df.describe(include="all").to_string()

        return "ERROR: Unsupported operation"

    except Exception as e:
        return f"Execution error: {e}"

# ------------------------------------------------
# CLAUDE PLANNER & JSON PARSER
# ------------------------------------------------

def get_llm_response(user_query, data_context):
    system_prompt = """You are an Excel AI Assistant. You must analyze data queries and output a specific JSON instruction matching one of these supported operations:
    - {"operation": "groupby_sum", "group": "column_name", "column": "column_name"}
    - {"operation": "groupby_mean", "group": "column_name", "column": "column_name"}
    - {"operation": "filter_equals", "column": "column_name", "value": "target_value"}
    - {"operation": "filter_contains", "column": "column_name", "value": "search_term"}
    - {"operation": "top_n", "column": "column_name", "n": 5}
    - {"operation": "describe"}

    Respond ONLY with the JSON block. Do not include conversational text or explanations."""

    user_content = f"Data Context:\n{data_context}\n\nUser Question: {user_query}"

    try:
        response = client.messages.create(
            model=MODEL,
            max_tokens=1000,
            system=system_prompt,
            messages=[{"role": "user", "content": user_content}]
        )
        
        raw_text = response.content[0].text.strip()
        
        if raw_text.startswith("```"):
            raw_text = re.sub(r"^```(?:json)?\n", "", raw_text)
            raw_text = re.sub(r"\n```$", "", raw_text).strip()
            
        op_json = json.loads(raw_text)
        return op_json
        
    except Exception as e:
        st.error(f"LLM Error in Planner: {e}")
        return None

# ------------------------------------------------
# CONVERSATIONAL RESPONSE GENERATOR
# ------------------------------------------------

def generate_natural_answer(user_query, execution_result):
    try:
        response = client.messages.create(
            model=MODEL,
            max_tokens=500,
            system="You are a helpful data assistant. Use the provided raw dataset results to directly, cleanly, and naturally answer the user's question. Formulate a polite response. Do not output python code.",
            messages=[{"role": "user", "content": f"User asked: {user_query}\n\nData engine returned this result:\n{execution_result}"}]
        )
        return response.content[0].text
    except Exception as e:
        return f"Could not generate conversational answer due to an error: {e}"

# ------------------------------------------------
# STREAMLIT UI
# ------------------------------------------------

st.title("Excel AI Agent (Grounded)")

# Sidebar for Input Methods
with st.sidebar:
    st.header("Data Source")
    input_method = st.radio("Choose input method:", ["Upload File", "Link (GDrive/OneDrive)"])
    
    df = None

    if input_method == "Upload File":
        uploaded_file = st.file_uploader("Upload an Excel or CSV file", type=["csv", "xlsx"])
        if uploaded_file:
            try:
                if uploaded_file.name.endswith(".csv"):
                    df = pd.read_csv(uploaded_file)
                else:
                    df = pd.read_excel(uploaded_file)
            except Exception as e:
                st.error(f"Error loading file: {e}")

    else:
        url_input = st.text_input("Paste Shareable Link:", placeholder="https://drive.google.com/...")
        if url_input:
            with st.spinner("Fetching cloud file..."):
                direct_link = get_direct_download_link(url_input)
                try:
                    # Note: We try reading as Excel first for cloud links
                    df = pd.read_excel(direct_link)
                except:
                    try:
                        df = pd.read_csv(direct_link)
                    except Exception as e:
                        st.error("Could not read file from link. Ensure the link is public ('Anyone with the link').")

# MAIN AREA
if df is not None:
    st.session_state.dataframes["active"] = df
    
    st.subheader("Data Preview (First 5 Rows)")
    st.dataframe(df.head(5))

    user_query = st.text_input("Ask a question about your data:")
    if user_query:
        context = build_data_context(df)
        
        with st.spinner("Analyzing query..."):
            operation = get_llm_response(user_query, context)
            
        if operation:
            with st.spinner("Executing data operations..."):
                result = run_pandas_operation(df, operation)
                
            with st.spinner("Generating answer..."):
                conversational_answer = generate_natural_answer(user_query, result)
                
            st.subheader("Assistant Response")
            st.info(conversational_answer)
            with st.expander("View Raw Data Operation"):
                st.json(operation)
                st.text(result)
else:
    st.info("Please upload a file or provide a link in the sidebar to begin.")
