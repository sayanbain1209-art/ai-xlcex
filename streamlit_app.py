import streamlit as st
import pandas as pd
import json
import httpx
import os

st.set_page_config(page_title="Ai xlcex", page_icon="📊", layout="wide")

st.title("📊 Ai xlcex - Smart Spreadsheet & Analytics")
st.caption("AI-assisted smart spreadsheet system with Auto-Charts")

# Initialize grid in session state
if "df" not in st.session_state:
    st.session_state.df = pd.DataFrame(
        [["" for _ in range(6)] for _ in range(10)],
        columns=[chr(65 + i) for i in range(6)]
    )

# Sidebar: API Key & File Upload
st.sidebar.header("⚙️ Settings & Tools")
api_key = st.sidebar.text_input("Gemini API Key", type="password", value=os.getenv("GEMINI_API_KEY", ""))

uploaded_file = st.sidebar.file_uploader("📂 Upload CSV or Excel file", type=["csv", "xlsx"])
if uploaded_file is not None:
    try:
        if uploaded_file.name.endswith(".csv"):
            st.session_state.df = pd.read_csv(uploaded_file)
        else:
            st.session_state.df = pd.read_excel(uploaded_file)
        st.sidebar.success("File uploaded successfully!")
    except Exception as e:
        st.sidebar.error(f"Error loading file: {e}")

# AI Prompt Bar
prompt = st.text_input("🤖 Enter AI instruction (e.g., 'Fill Column A with 1 to 5', 'Calculate column sums'):")

if st.button("Run AI Prompt"):
    if not prompt:
        st.warning("Please enter a prompt first.")
    else:
        with st.spinner("AI is processing..."):
            if api_key:
                try:
                    url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key={api_key}"
                    system_prompt = (
                        "You are an AI spreadsheet engine. Given this 2D table as JSON, apply the user request. "
                        "Return ONLY valid JSON with structure: {\"updated_data\": [[...]]}."
                    )
                    payload = {
                        "contents": [{
                            "parts": [{
                                "text": f"{system_prompt}\nUser request: {prompt}\nData:\n{st.session_state.df.to_json(orient='values')}"
                            }]
                        }]
                    }
                    resp = httpx.post(url, json=payload, timeout=30.0)
                    resp.raise_for_status()
                    res_json = resp.json()
                    raw = res_json['candidates'][0]['content']['parts'][0]['text'].strip()
                    if raw.startswith("```"):
                        raw = raw.split("\n", 1)[-1].rsplit("\n", 1)[0]
                    parsed = json.loads(raw)
                    st.session_state.df = pd.DataFrame(parsed["updated_data"], columns=st.session_state.df.columns)
                    st.success("Updated successfully!")
                except Exception as e:
                    st.error(f"Error: {e}")
            else:
                st.info("Demo mode: Add Gemini API Key in the left sidebar to execute real AI changes.")

# Main Editable Spreadsheet Table
st.subheader("📋 Sheet Editor")
edited_df = st.data_editor(st.session_state.df, num_rows="dynamic", use_container_width=True)
st.session_state.df = edited_df

# Action Buttons
col1, col2 = st.columns(2)
with col1:
    csv = st.session_state.df.to_csv(index=False).encode('utf-8')
    st.download_button(
        label="📥 Download Data as CSV",
        data=csv,
        file_name="ai_xlcex_export.csv",
        mime="text/csv",
    )

# Visual Chart Generator
st.divider()
st.subheader("📈 Quick Data Visualization")
numeric_columns = st.session_state.df.select_dtypes(include=['float64', 'int64']).columns.tolist()

if len(numeric_columns) > 0:
    selected_col = st.selectbox("Select Column to visualize:", numeric_columns)
    chart_type = st.radio("Chart Type:", ["Bar Chart", "Line Chart"], horizontal=True)
    if chart_type == "Bar Chart":
        st.bar_chart(st.session_state.df[selected_col])
    else:
        st.line_chart(st.session_state.df[selected_col])
else:
    st.caption("Tip: Put some numbers in the table or upload a CSV with numbers to automatically see interactive charts here.")
