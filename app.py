import pandas as pd
import streamlit as st
from anthropic import Anthropic
import json

# ------------------------------------------------
# CONFIG
# ------------------------------------------------

st.set_page_config(page_title="Excel AI Agent", layout="wide")

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
# SAFE TOOL LAYER (PANDAS EXECUTOR)
# ------------------------------------------------

def run_pandas_operation(df, op):
    """
    Safe execution layer (NO raw Python execution)
    """

    try:
        operation = op.get("operation")

        # -------------------------
        # GROUPBY
        # -------------------------
        if operation == "groupby_sum":
            return df.groupby(op["group"])[op["column"]].sum()

        if operation == "groupby_mean":
            return df.groupby(op["group"])[op["column"]].mean()

        # -------------------------
        # FILTER
        # -------------------------
        if operation == "filter_equals":
            return df[df[op["column"]] == op["value"]]

        # -------------------------
        # TOP N
        # -------------------------
        if operation == "top_n":
            return df.nlargest(op["n"], op["column"])

        # -------------------------
        # DESCRIBE
        # -------------------------
        if operation == "describe":
            return df.describe(include="all")

        return f"Unknown operation: {operation}"

    except Exception as e:
        return f"Execution error: {e}"


# ------------------------------------------------
# CLAUDE PLANNER (STEP 1)
# ------------------------------------------------

def get_plan(question, df_columns):
    """
    Claude converts natural language → structured JSON
    """

    system_prompt = f"""
You are a data analyst planner.

Convert user questions into structured JSON operations.

Available columns:
{list(df_columns)}

Rules:
- ONLY output valid JSON
- No explanations
- No markdown

Supported operations:

1. groupby_sum
2. groupby_mean
3. filter_equals
4. top_n
5. describe

Format examples:

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
        max_tokens=300,
        temperature=0,
        system=system_prompt,
        messages=[{"role": "user", "content": question}]
    )

    try:
        return json.loads(response.content[0].text)
    except:
        return None


# ------------------------------------------------
# CLAUDE EXPLAINER (STEP 3)
# ------------------------------------------------

def explain_result(question, result):
    """
    Claude turns computed output into insights
    """

    response = client.messages.create(
        model=MODEL,
        max_tokens=800,
        temperature=0.2,
        system="You are an expert data analyst. Explain results clearly and provide business insights.",
        messages=[
            {
                "role": "user",
                "content": f"""
Question: {question}

Result:
{result}
"""
            }
        ]
    )

    return response.content[0].text


# ------------------------------------------------
# UI
# ------------------------------------------------

st.title("📊 Excel AI Agent (v2)")

tab1, tab2 = st.tabs(["📥 Import", "🤖 Chat"])

# ------------------------------------------------
# IMPORT
# ------------------------------------------------

with tab1:

    files = st.file_uploader(
        "Upload Excel",
        type=["xlsx", "xls"],
        accept_multiple_files=True
    )

    for file in files:
        df = pd.read_excel(file)
        st.session_state.dataframes[file.name] = df
        st.success(f"Loaded {file.name}")

    for name, df in st.session_state.dataframes.items():
        with st.expander(name):
            st.dataframe(df.head(10))

# ------------------------------------------------
# CHAT
# ------------------------------------------------

with tab2:

    if not st.session_state.dataframes:
        st.warning("Upload Excel first")
        st.stop()

    df = list(st.session_state.dataframes.values())[0]

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

        # -------------------------------
        # STEP 1: PLAN (Claude)
        # -------------------------------
        plan = get_plan(prompt, df.columns)

        if not plan:
            st.error("Could not generate plan")
            st.stop()

        # -------------------------------
        # STEP 2: EXECUTE (Python safe layer)
        # -------------------------------
        result = run_pandas_operation(df, plan)

        # -------------------------------
        # STEP 3: EXPLAIN (Claude)
        # -------------------------------
        explanation = explain_result(prompt, result)

        with st.chat_message("assistant"):
            st.markdown(explanation)

        st.session_state.messages.append({
            "role": "assistant",
            "content": explanation
        })
