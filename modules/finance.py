from re import search

import streamlit as st
import pandas as pd
import os
from datetime import datetime
from modules.audit import log_action

TRANSACTION_FILE = "data/transactions.csv"


def get_next_transaction_id(df):

    if len(df) == 0:
        return "TRX0001"

    ids = []

    for tx in df["TransactionID"]:

        try:
            ids.append(
                int(str(tx).replace("TRX", ""))
            )
        except:
            pass

    if len(ids) == 0:
        return "TRX0001"

    return f"TRX{max(ids)+1:04d}"


def save_transaction(transaction):

    if os.path.exists(TRANSACTION_FILE):

        df = pd.read_csv(TRANSACTION_FILE)

        df = pd.concat(
            [df, pd.DataFrame([transaction])],
            ignore_index=True
        )

    else:

        df = pd.DataFrame([transaction])

    df.to_csv(
        TRANSACTION_FILE,
        index=False
    )


def render():

    st.markdown(
        """
        <h1 style='margin-bottom:0'>
        Finance Dashboard
        </h1>

        <p style='color:#64748b'>
        Track income, expenses and profitability
        </p>
        """,
        unsafe_allow_html=True
    )

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

            st.markdown(
                f"""
                <div class="metric-card">
                    <h4>💰 Income</h4>
                    <h2>₹{income:,.0f}</h2>
                </div>
                """,
            unsafe_allow_html=True
            )

        with c2:

            st.markdown(
                f"""
                <div class="metric-card">
                    <h4>💸 Expense</h4>
                    <h2>₹{expense:,.0f}</h2>
                </div>
                """,
            unsafe_allow_html=True
            )

        with c3:

            st.markdown(
                f"""
                <div class="metric-card">
                    <h4>📈 Net Profit</h4>
                    <h2>₹{profit:,.0f}</h2>
                </div>
                """,
            unsafe_allow_html=True
            )

        with c4:

            st.markdown(
                f"""
                <div class="metric-card">
                    <h4>📋 Transactions</h4>
                    <h2>{len(df)}</h2>
                </div>
                """,
            unsafe_allow_html=True
            )

        st.divider()

        st.markdown(
            "## ➕ Add Transaction"
        )

        

        tx_type = st.selectbox(
                "Type",
                ["Income", "Expense"]
            )
        with st.form("transaction_form"):
            if tx_type == "Income":

                category = st.selectbox(
                    "Category",
                    [
                        "Invoice Payment",
                        "Consulting Income",
                        "Training Income",
                        "Other Income"
                    ]
                )
                
                invoice_number = ""

                if category == "Invoice Payment":

                    invoice_df = pd.read_csv(
                      "data/invoices.csv"
                 )

                    invoice_number = st.selectbox(
                     "Select Invoice",
                     invoice_df["Invoice No"].tolist()
                )
                 
            else:

                category = st.selectbox(
                    "Category",
                    [
                        "Salary Payment",
                        "Software & Subscriptions",
                        "Hardware & Equipment",
                        "Office Infrastructure",
                        "Office Supplies",
                        "Travel & Transportation",
                        "Internet & Communication",
                        "Professional Services",
                        "Employee Welfare",
                        "Other"
                    ]
                )

            amount = st.number_input(
                "Amount",
                min_value=0.0,
                step=100.0
            )

            description = st.text_input(
                "Description"
            )

            payment_mode = st.selectbox(
                "Payment Mode",
                [
                    "Bank Transfer",
                    "UPI",
                    "Cash",
                    "Cheque",
                    "Credit Card",
                    "Debit Card"
                ]
            )

            submit = st.form_submit_button(
                "Save Transaction"
            )

            if submit:

                if os.path.exists(TRANSACTION_FILE):
                    existing = pd.read_csv(
                        TRANSACTION_FILE
                    )
                else:
                    existing = pd.DataFrame()

                transaction = {

                    "TransactionID":
                    get_next_transaction_id(existing),

                    "Date":
                    datetime.now().strftime(
                        "%m/%d/%Y"
                    ),

                    "Type":
                    tx_type,

                    "Category":
                    category,

                    "Amount":
                    amount,

                    "Description":
                    description,

                    "PaymentMode":
                    payment_mode,

                    "Status":
                    "Completed",

                    "LinkedTo":
                    invoice_number if category == "Invoice Payment" else "",

                    "Reference":
                    invoice_number if category == "Invoice Payment" else "",

                    "Notes":
                    "",

                    "CreatedAt":
                    datetime.now().strftime(
                        "%m/%d/%Y %H:%M"
                    )
                }

                save_transaction(
                    transaction
                )
                
                if (
                    tx_type == "Income"
                    and category == "Invoice Payment"
                    and invoice_number
                ):

                    tracking_file = (
                        "data/invoice_tracking.csv"
                    )

                    if os.path.exists(
                        tracking_file
                ):

                        tracking_df = pd.read_csv(
                            tracking_file
                        )

                    mask = (
                        tracking_df[
                            "Invoice No"
                        ].astype(str)
                        ==
                        str(invoice_number)
                    )

                    if mask.any():

                        tracking_df.loc[
                            mask,
                            "Status"
                        ] = "Paid"

                        tracking_df.loc[
                            mask,
                            "Received Amount"
                        ] = amount

                        tracking_df.loc[
                            mask,
                            "Payment Date"
                        ] = datetime.now().strftime(
                            "%Y-%m-%d"
                        )

                        tracking_df.to_csv(
                            tracking_file,
                            index=False
                        )

                try:
                    log_action(
                        st.session_state.username,
                        "Added Transaction",
                        f"{tx_type} - {category} - ₹{amount}"
                    )
                except:
                    pass

                st.success(
                    "Transaction Saved Successfully"
                )

                st.rerun()

        st.divider()

        st.markdown(
            "## 📋 Transaction History"
    )

        col1, col2, col3 = st.columns(3)

        with col1:

            search = st.text_input(
                "Search"
            )

        with col2:

            type_filter = st.selectbox(
                "Type",
                [
                    "All",
                    "Income",
                    "Expense"
                ]
            )

        with col3:

            category_filter = st.selectbox(
                "Category",
                ["All"]
                +
                sorted(
                    df["Category"]
                    .dropna()
                    .unique()
                    .tolist()
                )
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

        if type_filter != "All":

            filtered = filtered[
                filtered["Type"]
                ==
                type_filter
            ]

        if category_filter != "All":

            filtered = filtered[
                filtered["Category"]
                ==
            category_filter
            ]

        filtered = filtered.sort_values(
            "CreatedAt",
            ascending=False
        )

        st.dataframe(
            filtered,
            use_container_width=True
        )

        st.divider()

        st.subheader(
            "🗑 Delete Transaction"
        )

        transaction_id = st.selectbox(
            "Transaction ID",
            filtered["TransactionID"]
            .astype(str)
            .tolist()
        )

        if st.button(
            "Delete Transaction"
        ):

            df = df[
                df["TransactionID"]
                .astype(str)
                !=
                str(transaction_id)
            ]

            df.to_csv(
                TRANSACTION_FILE,
                index=False
            )

            st.success(
                "Transaction Deleted"
            )

            st.rerun()
        

    except Exception as e:

        st.error(
            f"Error loading finance data: {e}"
        )