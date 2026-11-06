
import streamlit as st
import pandas as pd
import json
import httpx
import os
import re

st.set_page_config(page_title="Ai xlcex", page_icon="📊", layout="wide")

st.title("📊 Ai xlcex - Smart Spreadsheet & Analytics")
st.caption("AI-assisted smart spreadsheet system with Auto-Charts")

# Initialize grid with clean strings
if "df" not in st.session_state:
    st.session_state.df = pd.DataFrame(
        [["" for _ in range(26)] for _ in range(15)],
        columns=[chr(65 + i) for i in range(26)]
    )

# Sidebar: Settings
st.sidebar.header("⚙️ Settings & Tools")
api_key = st.sidebar.text_input("Gemini API Key", type="password", value=os.getenv("GEMINI_API_KEY", "")).strip()

uploaded_file = st.sidebar.file_uploader("📂 Upload CSV or Excel file", type=["csv", "xlsx"])
if uploaded_file is not None:
    try:
        if uploaded_file.name.endswith(".csv"):
            st.session_state.df = pd.read_csv(uploaded_file).fillna("")
        else:
            st.session_state.df = pd.read_excel(uploaded_file).fillna("")
        st.sidebar.success("File uploaded successfully!")
    except Exception as e:
        st.sidebar.error(f"Error loading file: {e}")

# Prompt Area
prompt = st.text_area("🤖 Enter AI instruction (Paste unorganized notes, lists, or formulas):", height=100)

if st.button("Run AI Prompt"):
    if not prompt.strip():
        st.warning("Please enter a prompt first.")
    elif not api_key:
        st.info("Please enter your Gemini API Key in the left sidebar.")
    else:
        with st.spinner("AI is converting data into spreadsheet..."):
            try:
                url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-3.8-flash:generateContent?key={api_key}"
                
                # Pass clean string representation of current table
                current_table_str = st.session_state.df.to_csv(index=False)
                
                system_prompt = (
                    "You are a spreadsheet parser and calculation engine. "
                    "Analyze the user request and extract/generate a structured table. "
                    "You must output ONLY a raw JSON object with the following schema: \n"
                    "{\n"
                    '  "columns": ["Col1", "Col2", "Col3"],\n'
                    '  "rows": [["val1", "val2", "val3"], ["val4", "val5", "val6"]]\n'
                    "}\n"
                    "Rules:\n"
                    "1. Calculate any requested totals, sums, or formulas directly into the values.\n"
                    "2. Return ONLY JSON. Do not include markdown codeblocks, notes, or explanation."
                )

                payload = {
                    "contents": [{
                        "parts": [{
                            "text": f"{system_prompt}\n\nCurrent Table:\n{current_table_str}\n\nUser Instruction:\n{prompt}"
                        }]
                    }]
                }
                
                headers = {"Content-Type": "application/json"}
                resp = httpx.post(url, headers=headers, json=payload, timeout=90.0)
                
                if resp.status_code != 200:
                    st.error(f"API Error {resp.status_code}: {resp.text}")
                else:
                    res_json = resp.json()
                    raw = res_json['candidates'][0]['content']['parts'][0]['text'].strip()
                    
                    # Extract JSON block cleanly
                    json_match = re.search(r'\{[\s\S]*\}', raw)
                    if json_match:
                        raw = json_match.group(0)
                        
                    parsed = json.loads(raw)
                    cols = parsed.get("columns", [chr(65 + i) for i in range(len(parsed.get("rows", [[]])[0]))])
                    rows = parsed.get("rows", [])
                    
                    st.session_state.df = pd.DataFrame(rows, columns=cols)
                    st.success("Spreadsheet generated successfully!")
                    st.rerun()
            except Exception as e:
                st.error(f"Execution Error: {e}")

# Sheet Editor
st.subheader("📋 Sheet Editor")
edited_df = st.data_editor(st.session_state.df, num_rows="dynamic", use_container_width=True)
st.session_state.df = edited_df

# Actions
col1, col2 = st.columns(2)
with col1:
    csv = st.session_state.df.to_csv(index=False).encode('utf-8')
    st.download_button(
        label="📥 Download Data as CSV",
        data=csv,
        file_name="ai_xlcex_export.csv",
        mime="text/csv",
    )

# Visualizations
st.divider()
st.subheader("📈 Quick Data Visualization")
numeric_df = st.session_state.df.apply(pd.to_numeric, errors='coerce')
numeric_cols = [c for c in numeric_df.columns if numeric_df[c].notnull().sum() > 0]

if len(numeric_cols) > 0:
    selected_col = st.selectbox("Select Column to visualize:", numeric_cols)
    chart_type = st.radio("Chart Type:", ["Bar Chart", "Line Chart"], horizontal=True)
    if chart_type == "Bar Chart":
        st.bar_chart(numeric_df[selected_col].dropna())
    else:
        st.line_chart(numeric_df[selected_col].dropna())
else:
    st.caption("Tip: Once the AI outputs numbers, you will automatically get interactive charts here.")
