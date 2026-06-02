import streamlit as st
import pandas as pd
import os

TRANSACTION_FILE = "data/transactions.csv"


def render():

    st.title("💰 Finance Dashboard")

    if not os.path.exists(TRANSACTION_FILE):
        st.warning("transactions.csv not found")
        return

    try:

        df = pd.read_csv(TRANSACTION_FILE)

        if len(df) == 0:
            st.info("No transactions found")
            return

        income = df[
            df["Type"].astype(str).str.lower() == "income"
        ]["Amount"].sum()

        expense = df[
            df["Type"].astype(str).str.lower() == "expense"
        ]["Amount"].sum()

        profit = income - expense

        c1, c2, c3, c4 = st.columns(4)

        with c1:
            st.metric("Income", f"₹{income:,.0f}")

        with c2:
            st.metric("Expense", f"₹{expense:,.0f}")

        with c3:
            st.metric("Net Profit", f"₹{profit:,.0f}")

        with c4:
            st.metric("Transactions", len(df))

        st.divider()

        search = st.text_input(
            "Search Transactions"
        )

        filtered = df.copy()

        if search:

            mask = filtered.astype(str).apply(
                lambda x:
                x.str.contains(
                    search,
                    case=False,
                    na=False
                )
            ).any(axis=1)

            filtered = filtered[mask]

        st.dataframe(
            filtered,
            use_container_width=True
        )

    except Exception as e:

        st.error(str(e))