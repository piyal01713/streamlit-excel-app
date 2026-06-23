import pandas as pd
import streamlit as st
from anthropic import Anthropic
import json

# ------------------------------------------------
# CONFIG
# ------------------------------------------------

st.set_page_config(page_title="Excel AI Agent (Grounded)", layout="wide")

MODEL = "claude-haiku-4-5-20251001"

client = Anthropic(api_key=st.secrets.get("ANTHROPIC_API_KEY", ""))

# ------------------------------------------------
# SESSION STATE
# ------------------------------------------------

if "dataframes" not in st.session_state:
    st.session_state.dataframes = {}

if "messages" not in st.session_state:
    st.session_state.messages = []

# ------------------------------------------------
# BUILD SAFE DATA CONTEXT (IMPORTANT FIX)
# ------------------------------------------------

def build_data_context(df):
    return f"""
DATASET STRUCTURE:

Columns:
{list(df.columns)}

Column Types:
{df.dtypes.to_string()}

Sample Rows (VERY IMPORTANT):
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
            return df.groupby(op["group"])[op["column"]].sum()

        if operation == "groupby_mean":
            return df.groupby(op["group"])[op["column"]].mean()

        if operation == "filter_equals":
            return df[df[op["column"]] == op["value"]]

        if operation == "top_n":
            return df.nlargest(op["n"], op["column"])

        if operation == "describe":
            return df.describe(include="all")

        return "ERROR: Unsupported operation"

    except Exception as e:
        return f"Execution error: {e}"


# ------------------------------------------------
# CLAUDE PLANNER (STRICT + GROUNDED)
# ------------------------------------------------

def get_plan(question, df):

    system_prompt = f"""
You are a STRICT data analysis planner.

You MUST follow these rules:
- You can ONLY use the dataset columns provided
- NEVER assume extra columns exist
- NEVER mention SQL
- NEVER guess missing data
- Output ONLY valid JSON

DATASET INFO:
Columns:
{list(df.columns)}

Sample Data:
{df.head(10).to_string(index=False)}

SUPPORTED OPERATIONS:

1. groupby_sum
2. groupby_mean
3. filter_equals
4. top_n
5. describe

JSON FORMAT ONLY:

Examples:

User: total sales by month
{{
  "operation": "groupby_sum",
  "group": "month",
  "column": "sales"
}}

User: top 5 customers by revenue
{{
  "operation": "top_n",
  "column": "revenue",
  "n": 5
}}
"""

    response = client.messages.create(
        model=MODEL,
        max_tokens=400,
        temperature=0,
        system=system_prompt,
        messages=[{"role": "user", "content": question}]
    )

    try:
        text = response.content[0].text.strip()
        text = text.replace("```json", "").replace("```", "").strip()
        return json.loads(text)

    except:
        return None


# ------------------------------------------------
# CLAUDE EXPLAINER (NO HALLUCINATION RULES)
# ------------------------------------------------

def explain_result(question, result):

    response = client.messages.create(
        model=MODEL,
        max_tokens=800,
        temperature=0.2,
        system="""
You are a strict data analyst.

RULES:
- ONLY explain the computed result given
- NEVER assume missing columns
- NEVER mention SQL
- If result is empty or invalid, say "No valid data found in dataset"
- Do NOT hallucinate or guess

Be concise and factual.
""",
        messages=[{
            "role": "user",
            "content": f"""
Question: {question}

Computed Result:
{result}
"""
        }]
    )

    return response.content[0].text


# ------------------------------------------------
# UI
# ------------------------------------------------

st.title("📊 Excel AI Agent (Fully Grounded)")

tab1, tab2 = st.tabs(["📥 Import", "🤖 Chat"])

# ------------------------------------------------
# IMPORT
# ------------------------------------------------

with tab1:

    uploaded_files = st.file_uploader(
        "Upload Excel files",
        type=["xlsx", "xls"],
        accept_multiple_files=True
    )

    if uploaded_files:

        for file in uploaded_files:
            df = pd.read_excel(file)
            st.session_state.dataframes[file.name] = df
            st.success(f"✅ {file.name} uploaded successfully")

    st.info(f"Total files: {len(st.session_state.dataframes)}")

# ------------------------------------------------
# CHAT
# ------------------------------------------------

with tab2:

    if not st.session_state.dataframes:
        st.warning("Please upload Excel file first")
        st.stop()

    df = list(st.session_state.dataframes.values())[0]

    # chat history
    for msg in st.session_state.messages:
        with st.chat_message(msg["role"]):
            st.markdown(msg["content"])

    prompt = st.chat_input("Ask your Excel question...")

    if prompt:

        st.session_state.messages.append({
            "role": "user",
            "content": prompt
        })

        with st.chat_message("user"):
            st.markdown(prompt)

        # ------------------------------------------------
        # STEP 1: PLAN (GROUNDED)
        # ------------------------------------------------

        plan = get_plan(prompt, df)

        if not plan:
            with st.chat_message("assistant"):
                st.error("Could not understand question based on dataset.")
            st.stop()

        # ------------------------------------------------
        # STEP 2: EXECUTE
        # ------------------------------------------------

        result = run_pandas_operation(df, plan)

        # safety check
        if result is None or str(result).strip() == "":
            result = "No valid data found in dataset"

        # ------------------------------------------------
        # STEP 3: EXPLAIN
        # ------------------------------------------------

        explanation = explain_result(prompt, result)

        with st.chat_message("assistant"):
            st.markdown(explanation)

        st.session_state.messages.append({
            "role": "assistant",
            "content": explanation
        })
