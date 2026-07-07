import os

import pandas as pd
import streamlit as st
import time

from modules.auth import create_user, delete_user, load_users, update_user

st.markdown("<h1 style='margin-bottom:0'>Dashboard</h1>", unsafe_allow_html=True)

def render():

    employees = 0
    invoices = 0
    revenue = 0

    if os.path.exists("data/employees.csv"):
        employees = len(
            pd.read_csv(
                "data/employees.csv"
            )
        )

    if os.path.exists("data/invoices.csv"):

        inv = pd.read_csv(
            "data/invoices.csv"
        )

        invoices = len(inv)

        for col in inv.columns:

            if "total" in col.lower():

                revenue = pd.to_numeric(
                    inv[col],
                    errors="coerce"
                ).fillna(0).sum()

                break

    st.markdown(
        """
        <h1 style='margin-bottom:0'>
        Executive Dashboard
        </h1>

        <p style='color:#64748b'>
        Welcome to Onetechmated ERP Pro
        </p>
        """,
        unsafe_allow_html=True
    )

    c1, c2, c3, c4 = st.columns(4)

    with c1:

        st.markdown(
            f"""
            <div class="metric-card">
                <h4>👥 Employees</h4>
                <h2>{employees}</h2>
            </div>
            """,
            unsafe_allow_html=True
        )

    with c2:

        st.markdown(
            f"""
            <div class="metric-card">
                <h4>📄 Invoices</h4>
                <h2>{invoices}</h2>
            </div>
            """,
            unsafe_allow_html=True
        )

    with c3:

        st.markdown(
            f"""
            <div class="metric-card">
                <h4>💰 Revenue</h4>
                <h2>${revenue:,.0f}</h2>
            </div>
            """,
            unsafe_allow_html=True
        )

    with c4:

        st.markdown(
            """
            <div class="metric-card">
                <h4>🟢 Status</h4>
                <h2>Online</h2>
            </div>
            """,
            unsafe_allow_html=True
        )

    st.markdown("<br>", unsafe_allow_html=True)

    col1, col2 = st.columns([2, 1])

    with col1:

        st.markdown(
            """
            <div class="metric-card">
                <h3>Business Overview</h3>
                <p>
                Manage employees, invoices,
                finance operations and business
                activities from a single platform.
                </p>
            </div>
            """,
            unsafe_allow_html=True
        )

    with col2:

        st.markdown(
            """
            <div class="metric-card">
                <h3>System Health</h3>
                <p>🟢 All systems operational</p>
            </div>
            """,
            unsafe_allow_html=True
        )

    if st.session_state.get("role") == "Super Admin":
        st.markdown("<br>", unsafe_allow_html=True)
        st.markdown("### 👤 Super Admin User Management")

        with st.expander("Add New User", expanded=True):
            new_username = st.text_input("Username", key="new_user_username")
            new_password = st.text_input("Password", type="password", key="new_user_password")
            new_role = st.selectbox(
                "Role",
                ["Admin", "User", "Super Admin"],
                key="new_user_role",
            )

            if st.button("Create User", use_container_width=True):
                try:
                    create_user(new_username, new_password, new_role)
                    try:
                        st.session_state.login_started_at = time.time()
                    except Exception:
                        pass
                    st.success(f"User '{new_username}' created successfully.")
                    st.rerun()
                except ValueError as exc:
                    st.error(str(exc))

        users_df = load_users()
        users_df["status"] = users_df.get("status", "Active").fillna("Active")

        st.caption("Current users in the login store")
        st.dataframe(users_df[["username", "role", "status"]], use_container_width=True, hide_index=True)

        st.markdown("#### Manage Existing Users")
        manage_user = st.selectbox("Select user to manage", users_df["username"].astype(str).tolist(), key="manage_user_select")

        selected_user = users_df[users_df["username"].astype(str).str.lower() == manage_user.lower()].iloc[0]

        edit_role = st.selectbox(
            "Role",
            ["Admin", "User", "Super Admin"],
            index=["Admin", "User", "Super Admin"].index(selected_user.get("role", "User")) if selected_user.get("role", "User") in ["Admin", "User", "Super Admin"] else 0,
            key="edit_user_role",
        )

        edit_status = st.selectbox(
            "Status",
            ["Active", "Inactive"],
            index=0 if selected_user.get("status", "Active") == "Active" else 1,
            key="edit_user_status",
        )

        new_password = st.text_input("New password (leave blank to keep current)", type="password", key="edit_user_password")

        c1, c2, c3 = st.columns(3)

        with c1:
            if st.button("Save Changes", use_container_width=True):
                try:
                    update_user(manage_user, password=new_password or None, role=edit_role, status=edit_status)
                    try:
                        st.session_state.login_started_at = time.time()
                    except Exception:
                        pass
                    st.success(f"Updated user '{manage_user}'.")
                    st.rerun()
                except ValueError as exc:
                    st.error(str(exc))

        with c2:
            if st.button("Delete User", use_container_width=True):
                try:
                    delete_user(manage_user)
                    try:
                        st.session_state.login_started_at = time.time()
                    except Exception:
                        pass
                    st.warning(f"Deleted user '{manage_user}'.")
                    st.rerun()
                except Exception as exc:
                    st.error(str(exc))

        with c3:
            if st.button("Activate", use_container_width=True):
                try:
                    update_user(manage_user, status="Active")
                    try:
                        st.session_state.login_started_at = time.time()
                    except Exception:
                        pass
                    st.success(f"Activated '{manage_user}'.")
                    st.rerun()
                except ValueError as exc:
                    st.error(str(exc))

            if st.button("Deactivate", use_container_width=True):
                try:
                    update_user(manage_user, status="Inactive")
                    try:
                        st.session_state.login_started_at = time.time()
                    except Exception:
                        pass
                    st.warning(f"Deactivated '{manage_user}'.")
                    st.rerun()
                except ValueError as exc:
                    st.error(str(exc))
