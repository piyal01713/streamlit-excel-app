import pandas as pd
import streamlit as st
from anthropic import Anthropic
import json
import re
import base64
import io
import requests
import openpyxl

# ------------------------------------------------
# CONFIG & INITIALIZATION
# ------------------------------------------------

st.set_page_config(page_title="Excel Insights", layout="wide")

# Using your specific model ID
MODEL = "claude-haiku-4-5-20251001" 

if 'all_dfs' not in st.session_state:
    st.session_state.all_dfs = None
if 'scope' not in st.session_state:
    st.session_state.scope = "Analyze All Sheets (Join/Compare)"

api_key = st.secrets.get("ANTHROPIC_API_KEY")
if not api_key:
    st.error("Missing ANTHROPIC_API_KEY in Streamlit secrets!")
    st.stop()

client = Anthropic(api_key=api_key)

# ------------------------------------------------
# HELPERS
# ------------------------------------------------

def get_direct_download_link(url):
    try:
        if "docs.google.com/spreadsheets" in url:
            file_id = re.search(r'/d/([^/]+)', url).group(1)
            return f'https://docs.google.com/spreadsheets/d/{file_id}/export?format=xlsx'
        elif "drive.google.com" in url:
            file_id = re.search(r'd/([^/]+)', url).group(1)
            return f'https://drive.google.com/uc?export=download&id={file_id}'
        elif "sharepoint.com" in url:
            clean_url = re.sub(r'/:[a-z]:/r/', '/', url).split("?")[0]
            return f"{clean_url}?download=1"
        elif "1drv.ms" in url or "onedrive.live.com" in url:
            encoded_url = base64.b64encode(url.encode()).decode().replace('+', '-').replace('/', '_').rstrip('=')
            return f"https://api.onedrive.com/v1.0/shares/u!{encoded_url}/root/content"
        return url
    except:
        return url

def build_context(dfs_dict, scope):
    if scope == "Analyze All Sheets (Join/Compare)":
        context = "SCOPE: ALL SHEETS. Dictionary: `dfs`.\n"
        for name, df in dfs_dict.items():
            context += f"Sheet: '{name}' | Columns: {list(df.columns)}\n"
            context += f"Sample Data:\n{df.head(5).to_string(index=False)}\n---\n"
    else:
        df = dfs_dict[scope]
        context = f"SCOPE: SHEET '{scope}'. Use `dfs['{scope}']`.\n"
        context += f"Columns: {list(df.columns)}\nSample Data:\n{df.head(10).to_string(index=False)}"
    return context

# ------------------------------------------------
# AI LOGIC
# ------------------------------------------------

def get_analysis_code(user_query, context):
    system_prompt = (
        "You are a Python data expert. Write code for a dictionary of DataFrames named `dfs`. "
        "The final result MUST be in a variable named `result`. "
        "Search all rows and columns for value matches. "
        "If column names contain 'Unnamed', the code should attempt to find the correct data dynamically. "
        "RESPOND ONLY WITH CODE."
    )
    try:
        response = client.messages.create(
            model=MODEL,
            max_tokens=1000,
            system=system_prompt,
            messages=[{"role": "user", "content": f"Context:\n{context}\n\nQuery: {user_query}"}]
        )
        return re.sub(r"```python\n|```", "", response.content[0].text.strip())
    except Exception as e:
        st.error(f"AI Error: {e}")
        return None

def generate_natural_answer(user_query, execution_result):
    clean_system_prompt = (
        "You are a concise data assistant. Summarize the data result clearly.\n"
        "1. No preamble (don't say 'Here is the answer').\n"
        "2. Use **bolding** for key identifiers and values.\n"
        "3. Use bullet points for lists.\n"
        "4. Be direct and professional."
    )
    try:
        response = client.messages.create(
            model=MODEL,
            max_tokens=500,
            system=clean_system_prompt,
            messages=[{"role": "user", "content": f"User asked: {user_query}\nRaw Result: {execution_result}"}]
        )
        return response.content[0].text
    except:
        return f"Result: {execution_result}"

# ------------------------------------------------
# STREAMLIT UI
# ------------------------------------------------

st.title("📊 Excel Insights")

with st.sidebar:
    st.header("1. Load Data")
    input_method = st.radio("Source:", ["Upload File", "Cloud Link"])
    
    raw_data = None
    if
