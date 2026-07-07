import time
import os

import streamlit as st

from modules.auth import login
from modules.dashboard import render
from modules.employees import render as employee_render
from modules.settings import render as settings_render
from modules.audit import render as audit_render
from modules.finance import render as finance_render
from modules.invoices import render as invoices_render
from modules.invoice_create import render as invoice_create_render
from modules.payroll import render as payroll_render




st.set_page_config(
    page_title="Onetechmated ERP Pro",
    layout="wide"
)

# Removed automatic refresh to avoid unexpected session reruns that
# can interfere with session_state on some hosting platforms.
st_autorefresh = None


with open("assets/style.css") as f:
    st.markdown(
        f"<style>{f.read()}</style>",
        unsafe_allow_html=True
    )

if "logged_in" not in st.session_state:
    st.session_state.logged_in = False

INACTIVITY_TIMEOUT_SECONDS = 120

if st.session_state.logged_in:
    login_started_at = st.session_state.get("login_started_at")
    if login_started_at and (time.time() - login_started_at) >= INACTIVITY_TIMEOUT_SECONDS:
        st.session_state.logged_in = False
        st.session_state.username = ""
        st.session_state.role = ""
        st.session_state.pop("login_started_at", None)
        st.warning("You were logged out due to inactivity.")
        st.rerun()

if not st.session_state.logged_in:

    login()

else:

    # Load logo in a case-insensitive way (Linux is case-sensitive)
    logo_path = None
    for candidate in ("assets/logo.png", "assets/Logo.png", "assets/logo.PNG"):
        if os.path.exists(candidate):
            logo_path = candidate
            break

    if logo_path:
        st.sidebar.image(logo_path, width=120)

    st.sidebar.markdown(
    """
    ## Onetechmated ERP

    Business Operations Platform
    """
    )


    page = st.sidebar.radio(
        "Navigation",
        [
            "🏠 Dashboard",
            "👥 Employees",
            "💵 Payroll",
            "📄 Invoices",
            "➕ Create Invoice",
            "💰 Finance",
            "📊 Audit",
            "⚙️ Settings"
        ]
    )

    # Record navigation as user activity to prevent inactivity logout
    try:
        st.session_state.login_started_at = time.time()
    except Exception:
        pass

    st.sidebar.write("---")
    st.sidebar.write(
        f"👤 {st.session_state.username}"
    )

    if "Dashboard" in page:
        render()

    elif "Employees" in page:
        employee_render()

    elif "Payroll" in page:
        payroll_render()

    elif "Invoices" in page and "Create" not in page:
        invoices_render()

    elif "Create Invoice" in page:
        invoice_create_render()

    elif "Finance" in page:
        finance_render()

    elif "Audit" in page:
        audit_render()

    elif "Settings" in page:
        settings_render()
        
        
    st.sidebar.divider()

    st.sidebar.markdown(
        f"""
        **User:** {st.session_state.get("username", "Guest")}

        **Role:** {st.session_state.get("role", "-")}
        """
    )

    if st.sidebar.button(
        "🚪 Logout",
        use_container_width=True
    ):

        st.session_state.logged_in = False

        st.rerun()