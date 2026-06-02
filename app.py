import streamlit as st

from modules.auth import login
from modules.dashboard import render

st.set_page_config(
    page_title="Onetechmated ERP Pro",
    layout="wide"
)

with open("assets/style.css") as f:
    st.markdown(
        f"<style>{f.read()}</style>",
        unsafe_allow_html=True
    )

if "logged_in" not in st.session_state:
    st.session_state.logged_in = False

if not st.session_state.logged_in:

    login()

else:

    st.sidebar.image(
        "assets/logo.png",
        width=150
    )

    st.sidebar.title("Onetechmated ERP")

    page = st.sidebar.radio(
        "Navigation",
        [
            "Dashboard",
            "Employees",
            "Payroll",
            "Invoices",
            "Finance",
            "Audit",
            "Settings"
        ]
    )

    st.sidebar.write("---")
    st.sidebar.write(
        f"👤 {st.session_state.username}"
    )

    if page == "Dashboard":
        render()