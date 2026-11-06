import streamlit as st
import pandas as pd
import json
import httpx
import os

st.set_page_config(page_title="Ai xlcex", page_icon="📊", layout="wide")

st.title("📊 Ai xlcex - Spreadsheet")
st.caption("AI-assisted smart spreadsheet system")

# Initialize grid in session state
if "df" not in st.session_state:
    st.session_state.df = pd.DataFrame(
        [["" for _ in range(6)] for _ in range(10)],
        columns=[chr(65 + i) for i in range(6)]
    )

# Sidebar for API Key
api_key = st.sidebar.text_input("Gemini API Key", type="password", value=os.getenv("GEMINI_API_KEY", ""))

prompt = st.text_input("Enter AI prompt (e.g. 'Fill Column A with numbers 1 to 10', 'Sum Column A'):")

if st.button("Run AI Prompt"):
    if not prompt:
        st.warning("Please enter a prompt first.")
    else:
        with st.spinner("AI is processing..."):
            if api_key:
                try:
                    url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key={api_key}"
                    system_prompt = (
                        "You are an AI spreadsheet assistant. Given this 2D table as JSON, apply the user request. "
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
                st.info("API key na thakle demo response dekhachhe: Add Gemini key in sidebar.")

# Editable Grid
edited_df = st.data_editor(st.session_state.df, num_rows="dynamic", use_container_width=True)
st.session_state.df = edited_df

# CSV download button
csv = st.session_state.df.to_csv(index=False).encode('utf-8')
st.download_button(
    label="📥 Download as CSV",
    data=csv,
    file_name="ai_xlcex_data.csv",
    mime="text/csv",
)
