import pandas as pd
import streamlit as st
from anthropic import Anthropic
import json

# ------------------------------------------------
# CONFIG
# ------------------------------------------------

st.set_page_config(page_title="Excel AI Agent (Grounded)", layout="wide")

MODEL = "claude-haiku-4-5-20251001"

# CRITICAL FIX: Ensure the key exists before running to prevent empty string API crashes
api_key = st.secrets.get("ANTHROPIC_API_KEY")
if not api_key:
    st.error("Missing ANTHROPIC_API_KEY in Streamlit secrets!")
    st.stop()

client = Anthropic(api_key=api_key)

# ------------------------------------------------
# SESSION STATE
# ------------------------------------------------

if "dataframes" not in st.session_state:
    st.session_state.dataframes = {}

if "messages" not in st.session_state:
    st.session_state.messages = []

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
# SAFE EXECUTOR (ONLY ALLOWED OPS)
# ------------------------------------------------

def run_pandas_operation(df, op):
    try:
        operation = op.get("operation")

        if operation == "groupby_sum":
            res = df.groupby(op["group"])[op["column"]].sum()
            return res.to_string() # Convert to string for safe printing/LLM context

        if operation == "groupby_mean":
            res = df.groupby(op["group"])[op["column"]].mean()
            return res.to_string()

        if operation == "filter_equals":
            res = df[df[op["column"]] == op["value"]]
            return res.head(20).to_string(index=False) # Limit rows returned to LLM

        if operation == "top_n":
            res = df.nlargest(op["n"], op["column"])
            return res.to_string(index=False)

        if operation == "describe":
            return df.describe(include="all").to_string()

        return "ERROR: Unsupported operation"

    except Exception as e:
        return f"Execution error: {e}"

# ------------------------------------------------
# CLAUDE PLANNER & EXECUTION (COMPLETED)
# ------------------------------------------------

def get_llm_response(user_query, data_context):
    system_prompt = """You are an Excel AI Assistant. You must analyze data queries and output a specific JSON instruction matching one of these supported operations:
    - {"operation": "groupby_sum", "group": "column_name", "column": "column_name"}
    - {"operation": "groupby_mean", "group": "column_name", "column": "column_name"}
    - {"operation": "filter_equals", "column": "column_name", "value": "target_value"}
    - {"operation": "top_n", "column": "column_name", "n": 5}
    - {"operation": "describe"}

    Respond ONLY with the JSON block. Do not include conversational text."""

    user_content = f"Data Context:\n{data_context}\n\nUser Question: {user_query}"

    try:
        response = client.messages.create(
            model=MODEL,
            max_tokens=1000,
            system=system_prompt,
            messages=[{"role": "user", "content": user_content}]
        )
        # Parse the JSON block from response text
        op_json = json.loads(response.content[0].text.strip())
        return op_json
    except Exception as e:
        st.error(f"LLM Error: {e}")
        return None

# ------------------------------------------------
# STREAMLIT UI LAUNCH
# ------------------------------------------------

st.title("Excel AI Agent (Grounded)")

uploaded_file = st.file_uploader("Upload an Excel or CSV file", type=["csv", "xlsx"])

if uploaded_file:
    if uploaded_file.name.endswith(".csv"):
        df = pd.read_csv(uploaded_file)
    else:
        df = pd.read_excel(uploaded_file)
        
    st.session_state.dataframes["active"] = df
    st.dataframe(df.head(5))

    user_query = st.text_input("Ask a question about your data:")
    if user_query:
        context = build_data_context(df)
        
        with st.spinner("Analyzing data structure..."):
            operation = get_llm_response(user_query, context)
            
        if operation:
            st.subheader("Executed Operation Parameters")
            st.json(operation)
            
            with st.spinner("Running execution engine..."):
                result = run_pandas_operation(df, operation)
                
            st.subheader("Final Result")
            st.text(result)
