import os
import pandas as pd
import streamlit as st
from openai import AzureOpenAI  # Handled the explicit Azure OpenAI library import

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

# 5. Sidebar Configuration (Azure OpenAI Secrets Extraction)
st.sidebar.header("🔑 AI Configuration")

# Extract variables safely from your Streamlit dashboard secrets
azure_api_key = st.secrets.get("AZURE_OPENAI_API_KEY", "")
azure_endpoint = st.secrets.get("AZURE_OPENAI_ENDPOINT", "")
azure_deployment = st.secrets.get("AZURE_OPENAI_DEPLOYMENT", "")
azure_version = st.secrets.get("AZURE_OPENAI_VERSION", "")

# Verify that all required variables are present
if azure_api_key and azure_endpoint and azure_deployment:
    st.sidebar.success("🔒 Azure OpenAI credentials loaded securely!")
else:
    st.sidebar.warning("⚠️ Azure OpenAI secrets missing in Streamlit Configuration.")

# Display configuration specs for verification
st.sidebar.text(f"Deployment: {azure_deployment}")
st.sidebar.text(f"API Version: {azure_version}")

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
    elif not azure_api_key or not azure_endpoint:
        st.error("Missing Azure Credentials! Please verify your .streamlit/secrets.toml config.")
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
            
            # Clean container target blocks (fixed typo bug)
            with chat_history_space:
                with st.chat_message("user"):
                    st.markdown(prompt)

            with chat_history_space:
                with st.chat_message("assistant"):
                    try:
                        # Instantiates the explicit Azure OpenAI Engine Client
                        client = AzureOpenAI(
                            azure_endpoint=azure_endpoint,
                            azure_deployment=azure_deployment,
                            api_version=azure_version,
                            api_key=azure_api_key
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

                        # Trigger API stream parameter execution call using Azure deployment name
                        response_stream = client.chat.completions.create(
                            model=azure_deployment, 
                            messages=openai_messages,
                            temperature=0.1,
                            stream=True  
                        )

                        # Write streaming outputs dynamically to the app UI
                        st.write_stream(response_stream)

                    except Exception as e:
                        st.error(f"API Error encountered: {e}")
