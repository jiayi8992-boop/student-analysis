"""Streamlit hosts the program assets; student files stay in each browser."""
from pathlib import Path
import streamlit as st
from streamlit.components.v1 import declare_component

st.set_page_config(page_title="一卡通月度消费筛选", page_icon="📊", layout="wide", initial_sidebar_state="collapsed")
st.markdown("""<style>
section[data-testid="stMain"] .block-container {max-width:100%;padding:1rem 0 0;}
div[data-testid="stVerticalBlock"] {gap:0;}
</style>""", unsafe_allow_html=True)

# Never use st.file_uploader, setComponentValue, or a server-side student cache.
# This iframe contains the complete WebAssembly program and all of its assets.
screening = declare_component("campus_monthly_screening", path=str(Path(__file__).parent / "web"))
screening(key="campus_monthly_screening_v31", default=None)
