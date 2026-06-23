import pandas as pd
import streamlit as st
from anthropic import Anthropic
import json
import re

# ------------------------------------------------
# CONFIG
# ------------------------------------------------

st.set_page_config(page_title="Excel AI Agent (Grounded)", layout="wide")

MODEL = "claude-haiku-4-5-20251001"

# Check for API key in secrets before proceeding
api_key = st.secrets.get("ANTHROPIC_API_KEY")
if not api_key:
    st.error("Missing ANTHROPIC_API_KEY in Streamlit secrets! Please add it to your .streamlit/secrets.toml file.")
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
            return res.to_string()

        if operation == "groupby_mean":
            res = df.groupby(op["group"])[op["column"]].mean()
            return res.to_string()

        if operation == "filter_equals":
            res = df[df[op["column"]] == op["value"]]
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
        
        # Strip out markdown formatting if Claude returned it
        if raw_text.startswith("```"):
            raw_text = re.sub(r"^```(?:json)?\n", "", raw_text)
            raw_text = re.sub(r"\n```$", "", raw_text).strip()
            
        op_json = json.loads(raw_text)
        return op_json
        
    except json.JSONDecodeError:
        st.error(f"Failed to parse JSON. Claude's raw response was:\n\n{response.content[0].text}")
        return None
    except Exception as e:
        st.error(f"LLM Error: {e}")
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

uploaded_file = st.file_uploader("Upload an Excel or CSV file", type=["csv", "xlsx"])

if uploaded_file:
    if uploaded_file.name.endswith(".csv"):
        df = pd.read_csv(uploaded_file)
    else:
        df = pd.read_excel(uploaded_file)
        
    st.session_state.dataframes["active"] = df
    
    st.subheader("Data Preview (First 5 Rows)")
    st.dataframe(df.head(5))

    user_query = st.text_input("Ask a question about your data:")
    if user_query:
        context = build_data_context(df)
        
        # Step 1: Intercept query and extract JSON operation
        with st.spinner("Analyzing query and parsing operations..."):
            operation = get_llm_response(user_query, context)
            
        if operation:
            # Step 2: Execute operation securely inside Pandas
            with st.spinner("Running secure data lookup engine..."):
                result = run_pandas_operation(df, operation)
                
            # Step 3: Rewrite table back into regular human text
            with st.spinner("Synthesizing final natural language answer..."):
                conversational_answer = generate_natural_answer(user_query, result)
                
            # Display the final conversational response to the user
            st.subheader("Assistant Response")
            st.write(conversational_answer)
