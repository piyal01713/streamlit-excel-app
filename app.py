import pandas as pd
import streamlit as st
from anthropic import Anthropic
import re
import io

# ------------------------------------------------
# CONFIG & INITIALIZATION
# ------------------------------------------------

st.set_page_config(page_title="Excel Insights Pro", layout="wide")

MODELS = {
    "Claude 4.5 Haiku (Fast/Cheap)": "claude-haiku-4-5-20251001",
    "Claude 3.5 Sonnet (Strong Reasoning)": "claude-sonnet-20241022",
}

if 'all_dfs' not in st.session_state:
    st.session_state.all_dfs = None

api_key = st.secrets.get("ANTHROPIC_API_KEY")
if not api_key:
    st.error("Missing ANTHROPIC_API_KEY in Streamlit secrets!")
    st.stop()

client = Anthropic(api_key=api_key)

# ------------------------------------------------
# HELPERS
# ------------------------------------------------

def build_context(dfs_dict, scope, depth):
    """
    Constructs the data map based on user-selected depth.
    """
    context = (
        f"MODE: {depth} analysis.\n"
        "The DataFrames were loaded with NO HEADERS. Your first job is to locate the real data.\n"
        "Excel sheets often have frozen rows, titles, or empty spaces at the top.\n\n"
    )
    
    selected_sheets = dfs_dict.keys() if scope == "Analyze All Sheets (Join/Compare)" else [scope]
    
    for name in selected_sheets:
        df = dfs_dict[name]
        context += f"### SHEET: '{name}'\n"
        
        if depth == "Quick (Top 100 Rows)":
            # Send first 100 rows as CSV string
            context += df.head(100).to_csv(index=False, header=False)
        else:
            # FULL DATASET: Send everything
            # Using to_csv is more token-efficient than the standard DataFrame string
            context += df.to_csv(index=False, header=False)
            
        context += "\n---\n"
    
    return context

# ------------------------------------------------
# AI LOGIC
# ------------------------------------------------

def get_analysis_code(user_query, context, model_id):
    system_prompt = (
        "You are a Senior Data Engineer. You will receive raw Excel data in CSV format.\n"
        "1. ANALYZE the structure: Identify where the actual table(s) start.\n"
        "2. CLEAN: Write code to drop junk rows, promote the correct row to header, and handle merged cells.\n"
        "3. SOLVE: Perform the user's request using the cleaned data.\n"
        "4. OUTPUT: Save the final result in a variable named `result`.\n"
        "Input variable: `dfs` (dictionary of DataFrames).\n"
        "RESPOND ONLY WITH PYTHON CODE."
    )
    try:
        response = client.messages.create(
            model=model_id,
            max_tokens=2000,
            system=system_prompt,
            messages=[{"role": "user", "content": f"Context:\n{context}\n\nQuery: {user_query}"}]
        )
        return re.sub(r"```python\n|```", "", response.content[0].text.strip())
    except Exception as e:
        st.error(f"AI Error: {e}")
        return None

def generate_natural_answer(user_query, execution_result, model_id):
    clean_system_prompt = "You are a direct data assistant. Summarize the findings. Use **bolding** for values."
    try:
        response = client.messages.create(
            model=model_id,
            max_tokens=800,
            system=clean_system_prompt,
            messages=[{"role": "user", "content": f"User asked: {user_query}\nRaw Result: {execution_result}"}]
        )
        return response.content[0].text
    except:
        return f"Result: {execution_result}"

# ------------------------------------------------
# STREAMLIT UI
# ------------------------------------------------

st.title("📊 Excel Insights Pro (Deep Search)")

with st.sidebar:
    st.header("1. Intelligence Settings")
    selected_model_name = st.selectbox("Model:", list(MODELS.keys()), index=1)
    target_model_id = MODELS[selected_model_name]
    
    # NEW: Complexity / Depth Option
    st.divider()
    st.header("2. Analysis Depth")
    analysis_depth = st.radio(
        "How much data should Claude see?",
        ["Quick (Top 100 Rows)", "Full Dataset (All Rows)"],
        help="Use 'Full Dataset' if your headers/data are buried deep in the sheet (e.g. past row 20) or if you have multiple tables."
    )
    if analysis_depth == "Full Dataset (All Rows)":
        st.warning("⚠️ Full Dataset mode uses significantly more tokens and may increase API costs.")

    st.divider()
    st.header("3. Load Data")
    uploaded_file = st.file_uploader("Upload Excel (.xlsx, .xlsm)", type=["xlsx", "xlsm"])
    
    if uploaded_file:
        try:
            # We load with header=None so we don't accidentally lose a header row
            excel_file = pd.ExcelFile(io.BytesIO(uploaded_file.read()), engine='openpyxl')
            st.session_state.all_dfs = {
                sheet: excel_file.parse(sheet, header=None) 
                for sheet in excel_file.sheet_names
            }
            st.success("Workbook Ready!")
        except Exception as e: 
            st.error(f"Read Error: {e}")

    if st.session_state.all_dfs:
        st.session_state.scope = st.selectbox("Scope:", ["Analyze All Sheets (Join/Compare)"] + list(st.session_state.all_dfs.keys()))

# MAIN INTERFACE
if st.session_state.all_dfs:
    
    with st.expander("👀 Inspect Raw Structure"):
        st.info("Showing first 50 rows. Use the 'Analysis Depth' in sidebar to control what the AI sees.")
        if st.session_state.scope == "Analyze All Sheets (Join/Compare)":
            for sheet_name, df in st.session_state.all_dfs.items():
                st.markdown(f"**Sheet:** `{sheet_name}`")
                st.dataframe(df.head(50))
        else:
            st.dataframe(st.session_state.all_dfs[st.session_state.scope].head(50))

    with st.form("chat_form"):
        user_query = st.text_input("Ask a question about this data:", placeholder="e.g. 'Calculate the total margin for the table starting at row 25'")
        submitted = st.form_submit_button("Run Deep Analysis")

    if submitted and user_query:
        context = build_context(st.session_state.all_dfs, st.session_state.scope, analysis_depth)
        
        with st.spinner(f"Analyzing {analysis_depth}..."):
            code = get_analysis_code(user_query, context, target_model_id)
            if code:
                try:
                    # Execute generated logic
                    exec_globals = {'pd': pd, 'dfs': st.session_state.all_dfs}
                    exec_locals = {}
                    exec(code, exec_globals, exec_locals)
                    calc_res = exec_locals.get('result', "No result variable created.")
                    
                    answer = generate_natural_answer(user_query, calc_res, target_model_id)
                    
                    st.markdown("### 💡 AI Findings")
                    st.markdown(answer)
                    
                    if isinstance(calc_res, (pd.DataFrame, pd.Series)):
                        st.dataframe(calc_res)
                        
                    with st.expander("View AI Cleaning Logic"):
                        st.code(code)
                except Exception as e:
                    st.error(f"Execution Error: {e}")
                    with st.expander("View Generated Code"):
                        st.code(code)
else:
    st.info("Please upload your .xlsm or .xlsx file to start.")
