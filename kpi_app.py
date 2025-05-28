# kpi_app.py
import streamlit as st
import pandas as pd
from kpi import run_kpi_system  # assuming your logic is inside this file

st.set_page_config(page_title="Codixel KPI System", layout="wide")

st.title("📊 Codixel KPI Report System")

if st.button("Generate KPI Report"):
    try:
        file_path = run_kpi_system()  # let this return filename
        st.success("✅ Report generated successfully!")
        with open(file_path, "rb") as f:
            st.download_button("📥 Download Report", f, file_name=file_path, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    except Exception as e:
        st.error(f"❌ Failed: {e}")
