import streamlit as st
import pandas as pd

st.set_page_config(
    page_title="Excel Resource Explorer",
    layout="wide"
)

st.title("📊 Excel Resource Explorer")

# Session state initialization
if "dataframes" not in st.session_state:
    st.session_state.dataframes = {}

if "messages" not in st.session_state:
    st.session_state.messages = []

if "excel_url" not in st.session_state:
    st.session_state.excel_url = ""

# Tabs
tab1, tab2 = st.tabs(["📁 Data", "💬 Chat"])

with tab1:

    st.subheader("Load Excel Data")

    # URL Input
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
        "Upload Excel files",
        type=["xlsx", "xls"],
        accept_multiple_files=True
    )

    if uploaded_files:

        for file in uploaded_files:
            try:
                df = pd.read_excel(file)

                df["source_file"] = file.name

                st.session_state.dataframes[file.name] = df

            except Exception as e:
                st.error(f"Error loading {file.name}: {e}")

    # Display loaded files
    if st.session_state.dataframes:

        st.subheader("Loaded Files")

        for filename in st.session_state.dataframes.keys():
            st.write(f"✅ {filename}")

        combined = pd.concat(
            st.session_state.dataframes.values(),
            ignore_index=True
        )

        st.divider()

        st.subheader("Combined Dataset")

        st.metric("Total Rows", len(combined))
        st.metric("Total Columns", len(combined.columns))

        st.dataframe(
            combined,
            use_container_width=True
        )

    else:
        st.info("Upload an Excel file or provide a URL.")

with tab2:

    st.subheader("Chat With Your Data")

    if not st.session_state.dataframes:
        st.info("Load Excel data first.")
    else:

        combined = pd.concat(
            st.session_state.dataframes.values(),
            ignore_index=True
        )

        for message in st.session_state.messages:

            with st.chat_message(message["role"]):
                st.write(message["content"])

        question = st.chat_input(
            "Ask about your Excel data..."
        )

        if question:

            st.session_state.messages.append(
                {
                    "role": "user",
                    "content": question
                }
            )

            question_lower = question.lower()

            # Simple placeholder responses
            if "rows" in question_lower:
                answer = f"There are {len(combined)} rows."

            elif "columns" in question_lower:
                answer = "Columns:\n\n" + ", ".join(combined.columns)

            elif "files" in question_lower:
                answer = "Loaded files:\n\n" + "\n".join(
                    st.session_state.dataframes.keys()
                )

            else:
                answer = """
AI integration not added yet.

For now you can ask:
- How many rows?
- What columns exist?
- Which files are loaded?

Next step is connecting OpenAI.
"""

            st.session_state.messages.append(
                {
                    "role": "assistant",
                    "content": answer
                }
            )

            st.rerun()
