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
# SAFE PANDAS EXECUTOR
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

        return f"Unknown operation: {operation}"

    except Exception as e:
        return f"Execution error: {e}"

# ------------------------------------------------
# CLAUDE PLANNER (STRICT JSON FIXED)
# ------------------------------------------------

def get_plan(question, df_columns):

    system_prompt = f"""
You are a STRICT JSON generator.

Return ONLY valid JSON.

NO markdown.
NO explanation.
NO extra text.

Supported operations:
- groupby_sum
- groupby_mean
- filter_equals
- top_n
- describe

Available columns:
{list(df_columns)}

Rules:
- Only use valid columns
- Output JSON only

Example:
{{
  "operation": "groupby_sum",
  "group": "month",
  "column": "sales"
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
        text = response.content[0].text.strip()
        text = text.replace("```json", "").replace("```", "").strip()
        return json.loads(text)

    except Exception as e:
        st.error("❌ Could not generate plan from Claude")
        st.write("RAW OUTPUT:", response.content[0].text)
        return None

# ------------------------------------------------
# CLAUDE EXPLAINER
# ------------------------------------------------

def explain_result(question, result):

    response = client.messages.create(
        model=MODEL,
        max_tokens=800,
        temperature=0.2,
        system="You are an expert data analyst. Explain results clearly and give business insights.",
        messages=[{
            "role": "user",
            "content": f"""
Question: {question}

Result:
{result}
"""
        }]
    )

    return response.content[0].text

# ------------------------------------------------
# UI
# ------------------------------------------------

st.title("📊 Excel AI Agent (Stable Version)")

tab1, tab2 = st.tabs(["📥 Import", "🤖 Chat"])

# ------------------------------------------------
# IMPORT TAB
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

    st.info(f"Total files uploaded: {len(st.session_state.dataframes)}")

# ------------------------------------------------
# CHAT TAB
# ------------------------------------------------

with tab2:

    if not st.session_state.dataframes:
        st.warning("Please upload Excel file first")
        st.stop()

    df = list(st.session_state.dataframes.values())[0]

    # show chat history
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

        # -------------------------
        # STEP 1: PLAN
        # -------------------------
        plan = get_plan(prompt, df.columns)

        if not plan:
            st.stop()

        # -------------------------
        # STEP 2: EXECUTE
        # -------------------------
        result = run_pandas_operation(df, plan)

        # -------------------------
        # STEP 3: EXPLAIN
        # -------------------------
        explanation = explain_result(prompt, result)

        with st.chat_message("assistant"):
            st.markdown(explanation)

        st.session_state.messages.append({
            "role": "assistant",
            "content": explanation
        })
