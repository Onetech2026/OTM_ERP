import streamlit as st
import pandas as pd
import os

def render():

    employees = 0
    invoices = 0
    revenue = 0

    if os.path.exists("data/employees.csv"):
        employees = len(pd.read_csv("data/employees.csv"))

    if os.path.exists("data/invoices.csv"):

        inv = pd.read_csv("data/invoices.csv")

        invoices = len(inv)

        for col in inv.columns:

            if "total" in col.lower():
                revenue = inv[col].sum()
                break

    st.markdown(
        '<div class="main-title">Executive Dashboard</div>',
        unsafe_allow_html=True
    )

    c1,c2,c3,c4 = st.columns(4)

    with c1:
        st.metric("Employees", employees)

    with c2:
        st.metric("Invoices", invoices)

    with c3:
        st.metric("Revenue", f"₹{revenue:,.0f}")

    with c4:
        st.metric("Status", "Online")

    st.divider()

    st.subheader("Business Overview")

    st.info(
        "Welcome to Onetechmated ERP Pro."
    )