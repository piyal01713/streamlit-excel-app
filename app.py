import os
import pandas as pd
import streamlit as st
from openai import OpenAI

# Updated app name in page configuration
st.set_page_config(page_title="Excel Insight", layout="wide")

# Updated main title
st.title("💡 Excel Insight")

# Initialize session states
if "dataframes" not in st.session_state:
    st.session_state.dataframes = {}

if "messages" not in st.session_state:
    st.session_state.messages = []

if "excel_url" not in st.session_state:
    st.session_state.excel_url = ""

# Sidebar for API Configuration
st.sidebar.header("🔑 LLM Configuration")
openai_api_key = st.sidebar.text_input("OpenAI API Key", type="password")
model_choice = st.sidebar.selectbox(
    "Choose Model", ["gpt-4o-mini", "gpt-4o"]
)

# New optimized visual anchor icons for the tabs
tab1, tab2 = st.tabs(["📥 Import Data", "🤖 AI Insight Chat"])

with tab1:
    st.subheader("Load Excel Data")

    # URL Input
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
        st.info("Please add your OpenAI API key in the sidebar to start chatting.")
    else:
        # Construct the context from loaded dataframes
        data_context = "You are analyzing the following spreadsheet data:\n\n"
        for name, df in st.session_state.dataframes.items():
            data_context += f"--- Dataset Name: {name} ---\n"
            data_context += f"Columns: {', '.join(df.columns.tolist())}\n"
            data_context += f"Data Sample (First 5 rows):\n{df.head(5).to_string()}\n\n"

        for message in st.session_state.messages:
            with st.chat_message(message["role"]):
                st.markdown(message["content"])

        if prompt := st.chat_input("Ask a question about your data..."):
            st.session_state.messages.append({"role": "user", "content": prompt})
            with st.chat_message("user"):
                st.markdown(prompt)

            with st.chat_message("assistant"):
                message_placeholder = st.empty()
                
                try:
                    client = OpenAI(api_key=openai_api_key)
                    
                    system_prompt = (
                        f"{data_context}\nAnswer user queries based strictly on this data. "
                        "If calculations are complex or require heavy math, explain your reasoning steps clearly."
                    )
                    
                    messages_input = [
                        {"role": "system", "content": system_prompt}
                    ] + [
                        {"role": m["role"], "content": m["content"]}
                        for m in st.session_state.messages
                    ]

                    response = client.chat.completions.create(
                        model=model_choice,
                        messages=messages_input,
                    )
                    
                    full_response = response.choices.message.content
                    message_placeholder.markdown(full_response)
                    
                    st.session_state.messages.append({"role": "assistant", "content": full_response})
                    
                except Exception as e:
                    st.error(f"API Error: {e}")
