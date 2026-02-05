import streamlit as st
import pandas as pd
from io import BytesIO
import base64
import os
import corpnorm_utils as utils

# =========================================================
# 1. Config & Setup
# =========================================================

def load_config(path="config.txt"):
    config = {}
    try:
        with open(path, "r") as f:
            for line in f:
                if "=" in line:
                    k, v = line.strip().split("=", 1)
                    config[k.strip()] = v.strip()
    except FileNotFoundError:
        pass
    return config

def load_rules(path="rules_corpnorm.txt"):
    try:
        with open(path, "r", encoding="utf-8") as f:
            return f.read()
    except: return "You are CorpNorm AI."

CONFIG = load_config()
RULES = load_rules()

# =========================================================
# 1.5 Logo Loading
# =========================================================

# (Logo embedding removed — header will be simple text per user request)

# =========================================================
# 2. Main Streamlit App
# =========================================================

def main():
    st.set_page_config(page_title="CorpNorm AI - Agentic", layout="wide", page_icon="🕵️")
    
    # --- Simple Header with Logo ---
    def _load_logo_datauri(name="corpnorm_logo_applied.svg"):
        p = os.path.join(os.path.dirname(__file__), name)
        if os.path.exists(p):
            with open(p, "rb") as f:
                return "data:image/svg+xml;base64," + base64.b64encode(f.read()).decode()
        return None

    logo_uri = _load_logo_datauri()
    col_logo, col_text = st.columns([1, 6])
    with col_logo:
        if logo_uri:
            st.image(logo_uri, width=120)
    with col_text:
        st.title("CorpNorm AI")
        st.markdown("Agentic Intelligence for Global Data Verification")
    
    # --- Sidebar Configuration ---
    with st.sidebar:
        st.header("⚙️ Configuration")
        
        # Mode Switch
        mode = st.radio("Processing Mode", ["Free (Agentic)", "Premium (AI+SerpAPI)"], index=0)
        
        serpapi_key = ""
        openai_key = ""
        
        if mode == "Premium (AI+SerpAPI)":
            st.info("Premium mode uses SerpAPI (Google) for search and OpenAI for reasoning. High Accuracy.")
            serpapi_key = st.text_input("SerpAPI Key (Google)", value=CONFIG.get("serpapi_key", ""), type="password")
            openai_key = st.text_input("OpenAI API Key", value=CONFIG.get("openai_api_key", ""), type="password")
    
    # Initialize Agent
    agent = utils.CompanyAgent()

    # --- Input Mode Toggle ---
    st.markdown("### 1. Input Mode")
    input_mode = st.radio("Choose input method:", ["Batch (Excel Upload)", "Single Company"], index=0, horizontal=True)
    
    df = None
    df_loaded = False
    
    if input_mode == "Single Company":
        # --- Single Company Input Form ---
        st.markdown("#### Enter Company Details")
        col1, col2 = st.columns(2)
        with col1:
            company_name = st.text_input("Company Name *", placeholder="e.g., Acme Corp", key="single_company")
        with col2:
            country = st.text_input("Country", placeholder="e.g., USA", key="single_country")
        
        col1, col2, col3 = st.columns(3)
        with col1:
            street = st.text_input("Street Address", placeholder="Optional", key="single_street")
        with col2:
            city = st.text_input("City", placeholder="Optional", key="single_city")
        with col3:
            pass  # spacer
        
        if st.button("🔍 Analyze Single Company", type="primary", key="single_analyze"):
            if not company_name.strip():
                st.error("Company name is required!")
            else:
                # Create single-row dataframe from form
                df = pd.DataFrame([{
                    "Raw Company Name": company_name.strip(),
                    "Street Address1": street.strip(),
                    "City Name": city.strip(),
                    "Country Name": country.strip()
                }])
                df_loaded = True
    else:
        # --- Batch Upload Mode ---
        st.markdown("#### Upload Excel File")
        uploaded = st.file_uploader("Upload Excel (.xlsx)", type=["xlsx"], key="batch_upload")
        
        if uploaded:
            try:
                df = pd.read_excel(uploaded)
                st.dataframe(df.head())
                df_loaded = True
            except Exception as e:
                st.error(f"Error reading file: {e}")
                df = None
                df_loaded = False
        
        if uploaded and df_loaded:
            if st.button("🚀 Start Agentic Processing", type="primary", key="batch_process"):
                pass  # Processing logic below will trigger
    
    # --- Processing Logic (works for both modes) ---
    if df is not None and df_loaded:
        results = []
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        total = len(df)
        
        # Create a localized status container
        with st.status("Agent Working...", expanded=True) as status:
            for i, row in df.iterrows():
                raw_name = str(row.get("Raw Company Name", "")).strip()
                if not raw_name: continue
                
                status_text.text(f"Processing {i+1}/{total}: {raw_name}")
                status.write(f"🔍 Analyzing: **{raw_name}**...")
                
                address = {
                    "street1": str(row.get("Street Address1", "")),
                    "city": str(row.get("City Name", "")),
                    "country": str(row.get("Country Name", ""))
                }

                if mode == "Premium (AI+SerpAPI)":
                    if not serpapi_key or not openai_key:
                        status.write("⚠️ Missing Keys! Falling back to Free Agent.")
                        res = agent.process(raw_name, address)
                    else:
                        res = agent.process_premium(raw_name, address, serpapi_key, openai_key, RULES)
                else:
                    # Free Agent using Hybrid Strategy
                    res = agent.process(raw_name, address)

                results.append(res)
                progress_bar.progress((i + 1) / total)
            
            status.update(label="Processing Complete!", state="complete", expanded=False)

        # --- Output Formatting ---
        st.success("Analysis Complete!")
        df_out = pd.DataFrame(results)
        
        # Ensure Column Order matches Output.xlsx (Gold Standard)
        desired_order = [
            "Raw Company Name", 
            "Normalized Company Name", 
            "Website", 
            "Industry", 
            "Third Party Data Source Link",
            "Remark",
            "Confidence Score"
        ]
        
        # Reindex to enforce order
        cols = [c for c in desired_order if c in df_out.columns] + [c for c in df_out.columns if c not in desired_order]
        df_out = df_out[cols]

        st.dataframe(df_out)
        
        # Download
        buffer = BytesIO()
        with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
            df_out.to_excel(writer, index=False)
        buffer.seek(0)
        
        st.download_button(
            "📥 Download Agentic Results",
            data=buffer,
            file_name="CorpNorm_Agentic_Results.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

if __name__ == "__main__":
    main()
