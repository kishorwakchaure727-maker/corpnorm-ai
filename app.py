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

def load_logo_as_datauri():
    """Load SVG and convert to data URI for reliable display in Streamlit"""
    logo_path = os.path.join(os.path.dirname(__file__), "corpnorm_ai_logo.svg")
    if os.path.exists(logo_path):
        with open(logo_path, "rb") as f:
            svg_bytes = f.read()
        svg_base64 = base64.b64encode(svg_bytes).decode()
        return f"data:image/svg+xml;base64,{svg_base64}"
    return None

# =========================================================
# 2. Main Streamlit App
# =========================================================

def main():
    st.set_page_config(page_title="CorpNorm AI - Agentic", layout="wide", page_icon="🕵️")
    
    # --- Custom CSS Styling ---
    st.markdown("""
        <style>
        .header-container {
            display: flex;
            align-items: center;
            gap: 20px;
            margin-bottom: 20px;
        }
        .logo-section {
            display: flex;
            justify-content: center;
        }
        .brand-text {
            color: #1F4B99;
        }
        </style>
    """, unsafe_allow_html=True)
    
    # --- Header with Logo ---
    col1, col2 = st.columns([1, 3])
    
    with col1:
        logo_uri = load_logo_as_datauri()
        if logo_uri:
            st.image(logo_uri, width=200)
        else:
            st.markdown("### 🕵️ CorpNorm AI")
    
    with col2:
        st.markdown("### CorpNorm AI - Agentic Edition")
        st.markdown("**Agentic Intelligence for Global Data Verification**")
    
    st.divider()
    
    st.markdown(
        """
        **Capabilities:**
        - **Verify**: Visits websites to check if they match the company name.
        - **Classify**: Reads page content to determine industry.
        - **Fallback**: Hunts for third-party sources if official site is missing.
        """
    )
    
    # --- Sidebar Configuration ---
    with st.sidebar:
        logo_uri = load_logo_as_datauri()
        if logo_uri:
            st.image(logo_uri, width=180)
        st.header("⚙️ Configuration")
        
        # Mode Switch
        mode = st.radio("Processing Mode", ["Free (Agentic)", "Premium (AI+SerpAPI)"], index=0)
        
        serpapi_key = ""
        openai_key = ""
        
        if mode == "Premium (AI+SerpAPI)":
            st.info("Premium mode uses SerpAPI (Google) for search and OpenAI for reasoning. High Accuracy.")
            serpapi_key = st.text_input("SerpAPI Key (Google)", value=CONFIG.get("serpapi_key", ""), type="password")
            openai_key = st.text_input("OpenAI API Key", value=CONFIG.get("openai_api_key", ""), type="password")
        
        st.divider()
        st.markdown("""
            <div style='text-align: center; color: #999; font-size: 12px; margin-top: 20px;'>
                <p><strong>CorpNorm AI</strong></p>
                <p>Agentic Intelligence for Global Data Verification</p>
            </div>
        """, unsafe_allow_html=True)
    
    # Initialize Agent
    agent = utils.CompanyAgent()

    # --- Main Input ---
    st.markdown("### 1. Upload Input")
    uploaded = st.file_uploader("Upload Excel (.xlsx)", type=["xlsx"])
    
    if uploaded:
        try:
            df = pd.read_excel(uploaded)
            st.dataframe(df.head())
        except Exception as e:
            st.error(f"Error reading file: {e}")
            return

        if st.button("🚀 Start Agentic Processing", type="primary"):
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
