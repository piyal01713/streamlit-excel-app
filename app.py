import os
import pandas as pd
import streamlit as st
from google import genai
from google.genai import types

# Set page configuration with branding
st.set_page_config(page_title="Excel Insight", layout="wide")

# Main Title
st.title("💡 Excel Insight")

# Initialize session states
if "dataframes" not in st.session_state:
    st.session_state.dataframes = {}

if "messages" not in st.session_state:
    st.session_state.messages = []

if "excel_url" not in st.session_state:
    st.session_state.excel_url = ""

# --- Sidebar Configuration (Automated Secrets Extraction) ---
st.sidebar.header("🔑 Gemini Configuration")

# Securely extract key from Streamlit's secrets engine
gemini_api_key = st.secrets.get("GEMINI_API_KEY", "")

if gemini_api_key:
    st.sidebar.success("🔒 API Key loaded securely from Secrets!")
else:
    st.sidebar.warning("⚠️ GEMINI_API_KEY not detected in Streamlit Secrets.")

model_choice = st.sidebar.selectbox(
    "Choose Model", ["gemini-2.5-flash", "gemini-2.5-pro"]
)

# Option to wipe chat history
st.sidebar.divider()
if st.sidebar.button("🗑️ Clear Chat History"):
    st.session_state.messages = []
    st.rerun()

# Tabs for organization
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

    # Multi-file Uploader
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

    # Guard rails
    if not st.session_state.dataframes:
        st.warning("Please upload a file or load a URL in the 'Import Data' tab first.")
    elif not gemini_api_key:
        st.error("Missing API Key! Please verify GEMINI_API_KEY setup in your Streamlit application dashboard.")
    else:
        # Construct raw dataset snapshots for context
        data_context = "You are an analytical assistant exploring these spreadsheet datasets:\n\n"
        for name, df in st.session_state.dataframes.items():
            data_context += f"--- File Name: {name} ---\n"
            data_context += f"Columns: {', '.join(df.columns.astype(str).tolist())}\n"
            data_context += f"Sample Data Rows:\n{df.head(5).to_string()}\n\n"

        # Display history
        for message in st.session_state.messages:
            with st.chat_message(message["role"]):
                st.markdown(message["content"])

        # Handle user prompt inputs
        if prompt := st.chat_input("Ask a question about your data..."):
            st.session_state.messages.append({"role": "user", "content": prompt})
            with st.chat_message("user"):
                st.markdown(prompt)

            with st.chat_message("assistant"):
                try:
                    # Construct client via Google GenAI Python SDK using the safe secret variable
                    client = genai.Client(api_key=gemini_api_key)
                    
                    # Package chat context and systemic guide bounds
                    system_instruction = (
                        f"{data_context}Answer user queries based strictly on the data columns and sample rows provided. "
                        "If operations require structured analysis, walk the user through step-by-step calculations."
                    )
                    
                    # Convert internal message lists into standard formats
                    contents_input = []
                    for m in st.session_state.messages:
                        role_name = "user" if m["role"] == "user" else "model"
                        contents_input.append(
                            types.Content(
                                role=role_name,
                                parts=[types.Part.from_text(text=m["content"])]
                            )
                        )

                    # Inner generator function to stream response chunks to UI
                    def response_generator():
                        response_stream = client.models.generate_content_stream(
                            model=model_choice,
                            contents=contents_input,
                            config=types.GenerateContentConfig(
                                system_instruction=system_instruction,
                                temperature=0.1
                            )
                        )
                        for chunk in response_stream:
                            if chunk.text:
                                yield chunk.text

                    # Write the generator content live to the app page with streaming animation
                    full_response = st.write_stream(response_generator())
                    
                    # Save assistant answers into states
                    st.session_state.messages.append({"role": "assistant", "content": full_response})
                    
                except Exception as e:
                    st.error(f"Gemini API Engine Error: {e}")
