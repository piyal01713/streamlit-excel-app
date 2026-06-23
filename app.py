import pandas as pd
import streamlit as st
from anthropic import Anthropic

# ------------------------------------------------
# PAGE CONFIG
# ------------------------------------------------

st.set_page_config(
    page_title="Excel Insight",
    layout="wide"
)

# ------------------------------------------------
# CUSTOM CSS
# ------------------------------------------------

st.markdown("""
<style>

.block-container {
    padding-top: 1rem !important;
}

[data-testid="stHeader"] {
    height: 0px !important;
    background: transparent !important;
}

div[data-testid="stVBox"] > div:has(div[data-testid="stChatMessage"]) {
    max-height: 72vh !important;
    overflow-y: auto !important;
}

button[data-testid="stMarkdownContainer"] p {
    font-size: 16px !important;
    font-weight: 600 !important;
}

div[data-testid="stChatMessage"]:has(div[aria-label="chat-message-assistant"])
div[data-testid="stMarkdownContainer"] {

    font-size:16px !important;
    line-height:1.6 !important;
}

</style>
""", unsafe_allow_html=True)

# ------------------------------------------------
# TITLE
# ------------------------------------------------

st.title("💡 Excel Insight")

# ------------------------------------------------
# SESSION STATE
# ------------------------------------------------

if "dataframes" not in st.session_state:
    st.session_state.dataframes = {}

if "messages" not in st.session_state:
    st.session_state.messages = []

if "excel_url" not in st.session_state:
    st.session_state.excel_url = ""

# ------------------------------------------------
# SIDEBAR
# ------------------------------------------------

st.sidebar.header("🔑 Claude Configuration")

anthropic_api_key = st.secrets.get(
    "ANTHROPIC_API_KEY",
    ""
)

if anthropic_api_key:
    st.sidebar.success("Claude API Key Loaded")
else:
    st.sidebar.warning("ANTHROPIC_API_KEY missing")

st.sidebar.divider()

if st.sidebar.button("🗑️ Clear Chat History"):
    st.session_state.messages = []
    st.rerun()

# ------------------------------------------------
# TABS
# ------------------------------------------------

tab1, tab2 = st.tabs(
    [
        "📥 Import Data",
        "🤖 AI Insight Chat"
    ]
)

# ------------------------------------------------
# IMPORT DATA TAB
# ------------------------------------------------

with tab1:

    st.subheader("Load Excel Data")

    excel_url = st.text_input(
        "Excel URL",
        value=st.session_state.excel_url,
        placeholder="https://example.com/file.xlsx"
    )

    col1, col2 = st.columns(2)

    with col1:

        if st.button("Load URL"):

            try:

                df = pd.read_excel(excel_url)

                st.session_state.dataframes["URL File"] = df

                st.session_state.excel_url = excel_url

                st.success("Excel loaded successfully.")

            except Exception as e:

                st.error(e)

    with col2:

        if st.button("🔄 Resync URL"):

            try:

                if st.session_state.excel_url:

                    df = pd.read_excel(
                        st.session_state.excel_url
                    )

                    st.session_state.dataframes["URL File"] = df

                    st.success("URL refreshed.")

            except Exception as e:

                st.error(e)

    st.divider()

    uploaded_files = st.file_uploader(
        "Upload Excel Files",
        type=["xlsx", "xls"],
        accept_multiple_files=True
    )

    if uploaded_files:

        for file in uploaded_files:

            try:

                df = pd.read_excel(file)

                st.session_state.dataframes[file.name] = df

                st.success(f"Loaded {file.name}")

            except Exception as e:

                st.error(e)

    if st.session_state.dataframes:

        st.subheader("📋 Data Preview")

        for name, df in st.session_state.dataframes.items():

            with st.expander(
                f"{name} | Rows:{df.shape[0]} Columns:{df.shape[1]}"
            ):

                st.dataframe(df.head(20))

# ------------------------------------------------
# CHAT TAB
# ------------------------------------------------

with tab2:

    st.subheader("💬 Chat with your Excel Data")

    if not st.session_state.dataframes:

        st.warning(
            "Please upload an Excel file first."
        )

    elif not anthropic_api_key:

        st.error(
            "Missing ANTHROPIC_API_KEY."
        )

    else:

        client = Anthropic(
            api_key=anthropic_api_key
        )

        # Build dataset context

        data_context = ""

        for name, df in st.session_state.dataframes.items():

            preview_rows = min(50, len(df))

            data_context += f"""

FILE: {name}

ROWS: {df.shape[0]}

COLUMNS:

{', '.join(df.columns.astype(str).tolist())}

FIRST {preview_rows} ROWS:

{df.head(preview_rows).to_string()}

SUMMARY:

{df.describe(include='all').to_string()}

------------------------------------------------

"""

        # Chat history

        for message in st.session_state.messages:

            with st.chat_message(message["role"]):

                st.markdown(message["content"])

        # User input

        prompt = st.chat_input(
            "Ask something about your Excel..."
        )

        if prompt:

            st.session_state.messages.append(
                {
                    "role":"user",
                    "content":prompt
                }
            )

            with st.chat_message("user"):

                st.markdown(prompt)

            with st.chat_message("assistant"):

                try:

                    system_prompt = f"""

You are an expert Excel analyst.

You help users:

- Analyze spreadsheets
- Find trends
- Calculate statistics
- Compare values
- Detect anomalies
- Explain business insights

Answer ONLY from the uploaded Excel data.

Dataset:

{data_context}

"""

                    response = client.messages.create(

                        model="claude-3-5-haiku-latest",

                        max_tokens=4000,

                        temperature=0.1,

                        system=system_prompt,

                        messages=[

                            {
                                "role":"user",
                                "content":prompt
                            }

                        ]

                    )

                    answer = response.content[0].text

                    st.markdown(answer)

                    st.session_state.messages.append(
                        {
                            "role":"assistant",
                            "content":answer
                        }
                    )

                except Exception as e:

                    st.error(f"Error: {e}")
