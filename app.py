import os
import pandas as pd
import streamlit as st
from openai import OpenAI  

# 1. Page Configuration
st.set_page_config(page_title="Excel Insight", layout="wide")

# 2. Complete CSS Custom Styling Injection
st.markdown(
    """
    <style>
    /* Maximize vertical window space by removing top margins */
    .block-container {
        padding-top: 1rem !important;
        padding-bottom: 0rem !important;
        margin-top: 0rem !important;
    }
    /* Hide top header background bar space */
    [data-testid="stHeader"] {
        height: 0px !important;
        background: transparent !important;
    }
    /* Tighter padding layouts for tabs */
    [data-testid="stTab"] {
        padding-top: 2px !important;
        padding-bottom: 2px !important;
    }
    /* VIEWPORT MARGIN: Enlarges the chat history scrollable area */
    div[data-testid="stVBox"] > div:has(div[data-testid="stChatMessage"]) {
        max-height: 72vh !important;
        overflow-y: auto !important;
    }
    
    /* Boost font size and weight of tab text (Import Data & AI Insight Chat) */
    button[data-testid="stMarkdownContainer"] p {
        font-size: 16px !important;
        font-weight: 600 !important;
    }

    /* Enlarge the AI assistant's chat responses, bullet lists, and tables */
    div[data-testid="stChatMessage"]:has(div[aria-label="chat-message-assistant"]) div[data-testid="stMarkdownContainer"] {
        font-size: 16px !important;
        line-height: 1.6 !important;
    }
    </style>
    """,
    unsafe_allow_html=True
)

# 3. Main Title
st.title("💡 Excel Insight")

# 4. Initialize Session States
if "dataframes" not in st.session_state:
    st.session_state.dataframes = {}

if "messages" not in st.session_state:
    st.session_state.messages = []

if "excel_url" not in st.session_state:
    st.session_state.excel_url = ""

# 5. Sidebar Configuration (Automated OpenAI / Proxy Secrets Extraction)
st.sidebar.header("🔑 AI Configuration")

# Extract key and custom endpoint URL from your Streamlit dashboard secrets
openai_api_key = st.secrets.get("OPENAI_API_KEY", "")
openai_base_url = st.secrets.get("OPENAI_BASE_URL", None)

if openai_api_key:
    if openai_base_url:
        st.sidebar.success("🔒 API Key & Custom URL loaded securely!")
    else:
        st.sidebar.success("🔒 API Key loaded securely from Secrets!")
else:
    st.sidebar.warning("⚠️ OPENAI_API_KEY not detected in Streamlit Secrets.")

model_choice = st.sidebar.selectbox(
    "Choose Model", ["gpt-4o-mini", "gpt-4o"]
)

st.sidebar.divider()
if st.sidebar.button("🗑️ Clear Chat History"):
    st.session_state.messages = []
    st.rerun()

# 6. Layout Tabs
tab1, tab2 = st.tabs(["📥 Import Data", "🤖 AI Insight Chat"])

with tab1:
    st.subheader("Load Excel Data")

    excel_url = st.text_input(
        "Excel URL",
        value=st.session_state.excel_url,
        placeholder="https://example.com",
    )

    col1, col2 = st.columns(2)

    with col1:
        if st.button("Load URL"):
            try:
                df = pd.read_excel(excel_url)
                st.session_state.dataframes["URL File"] = df
                st.session_state.excel_url = excel_url
                st.success("Excel loaded successfully from URL.")
            except Exception as e:
                st.error(f"Could not load Excel: {e}")

    with col2:
        if st.button("🔄 Resync URL"):
            if st.session_state.excel_url:
                try:
                    df = pd.read_excel(st.session_state.excel_url)
                    st.session_state.dataframes["URL File"] = df
                    st.success("URL data refreshed.")
                except Exception as e:
                    st.error(f"Refresh failed: {e}")

    st.divider()

    uploaded_files = st.file_uploader(
        "Upload Excel files", type=["xlsx", "xls"], accept_multiple_files=True
    )

    if uploaded_files:
        for file in uploaded_files:
            try:
                df = pd.read_excel(file)
                df["source_file"] = file.name
                st.session_state.dataframes[file.name] = df
                st.success(f"Successfully loaded: {file.name}")
            except Exception as e:
                st.error(f"Error loading {file.name}: {e}")

    if st.session_state.dataframes:
        st.subheader("📋 Active Data Previews")
        for name, df in st.session_state.dataframes.items():
            with st.expander(f"View {name} ({df.shape} rows)"):
                st.dataframe(df.head(10))

with tab2:
    st.subheader("💬 Chat with your Excel Data")

    if not st.session_state.dataframes:
        st.warning("Please upload a file or load a URL in the 'Import Data' tab first.")
    elif not openai_api_key:
        st.error("Missing API Key! Please verify OPENAI_API_KEY setup in your Streamlit application dashboard secrets.")
    else:
        # Construct dataset snapshot payload context
        data_context = "You are an analytical assistant exploring these complete spreadsheet datasets:\n\n"
        for name, df in st.session_state.dataframes.items():
            data_context += f"--- File Name: {name} ---\n"
            data_context += f"Columns: {', '.join(df.columns.astype(str).tolist())}\n"
            data_context += f"Full Data Rows:\n{df.to_string()}\n\n"

        # Chat history scroll window container
        chat_history_space = st.container(height=620)

        with chat_history_space:
            for message in st.session_state.messages:
                with st.chat_message(message["role"]):
                    st.markdown(message["content"])

        # Pinned bottom chat input interaction field
        if prompt := st.chat_input("Ask a question about your data..."):
            
            st.session_state.messages.append({"role": "user", "content": prompt})
            
            with chat_history_space.chat_message("user"):
                st.markdown(prompt)

            with chat_history_space.chat_message("assistant"):
                try:
                    # Instantiates OpenAI endpoint engine client, passing custom proxy base url if provided
                    client = OpenAI(
                        api_key=openai_api_key,
                        base_url=openai_base_url if openai_base_url else None
                    )
                    
                    system_instruction = (
                        f"{data_context}\nAnswer user queries based comprehensively on the full data rows provided. "
                        "Do not limit answers to samples. If operations require structured analysis, walk the user through step-by-step calculations."
                    )
                    
                    # Convert history state data payload sequence matching OpenAI's list system structures
                    openai_messages = [
                        {"role": "system", "content": system_instruction}
                    ] + [
                        {"role": m["role"], "content": m["content"]}
                        for m in st.session_state.messages
                    ]

                    # Trigger API stream parameter execution call
                    response_stream = client.chat.completions.create(
                        model=model_choice,
                        messages=openai_messages,
                        temperature=0.1,
                        stream=True  
                    )

                    # Dynamic text token chunk data parser
                    def response_generator():
                        for chunk in response_stream:
                            if chunk.choices.delta.content is not None:
                                yield chunk.choices.delta.content

                    # Write chunks to application tab with smooth automatic scrolling adjustments
                    full_response = st.write_stream(response_generator())
                    
                    st.session_state.messages.append({"role": "assistant", "content": full_response})
                    st.rerun()
                    
                except Exception as e:
                    st.error(f"🛑 AI Engine Connection Failure: {e}")
