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
    
    # --- Simple Header ---
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
