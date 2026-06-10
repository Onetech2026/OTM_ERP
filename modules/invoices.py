import streamlit as st
import pandas as pd
import os


INVOICE_FILE = "data/invoices.csv"

def render():

    st.markdown(
        """
        <h1 style='margin-bottom:0'>
        Invoice Management
        </h1>

        <p style='color:#64748b'>
        Create, track and manage client invoices
        </p>
        """,
        unsafe_allow_html=True
    )
    
    st.info(
        "Use the left menu and select Create Invoice."
    )

    if not os.path.exists(INVOICE_FILE):
        st.warning("invoices.csv not found")
        return

    try:

        df = pd.read_csv(INVOICE_FILE)
        
        tracking_file = "data/invoice_tracking.csv"

        if os.path.exists(tracking_file):

            tracking_df = pd.read_csv(
                tracking_file
            )

            df = df.merge(
                tracking_df,
                on="Invoice No",
                how="left"
            )

        df = df.dropna(
            subset=["Invoice No"]
        )

        if len(df) == 0:
            st.info("No invoices found")
            return

        total_invoices = len(df)

        total_value = pd.to_numeric(
            df["Total"],
            errors="coerce"
        ).fillna(0).sum()
        
        pending_count = 0
        paid_count = 0
        received_total = 0

        if "Status" in df.columns:

            pending_count = len(
                df[df["Status"] == "Pending"]
            )

            paid_count = len(
                df[df["Status"] == "Paid"]
            )

        if "Received Amount" in df.columns:

            received_total = pd.to_numeric(
                df["Received Amount"],
                errors="coerce"
            ).fillna(0).sum()

        outstanding_value = pd.to_numeric(
            df.loc[
                df["Status"] != "Paid",
                "Total"
            ],
            errors="coerce"
        ).fillna(0).sum()

        c1, c2, c3 = st.columns(3)

        with c1:

            st.markdown(
                f"""
                <div class="metric-card">
                    <h4>📄 Total Invoices</h4>
                    <h2>{total_invoices}</h2>
                </div>
                """,
                unsafe_allow_html=True
            )

        with c2:

            st.markdown(
                f"""
                <div class="metric-card">
                    <h4>💵 USD Invoiced</h4>
                    <h2>${total_value:,.0f}</h2>
                </div>
                """,
                unsafe_allow_html=True
            )

        with c3:

            st.markdown(
                f"""
                <div class="metric-card">
                    <h4>₹ INR Received</h4>
                    <h2>₹{received_total:,.0f}</h2>
                </div>
                """,
                unsafe_allow_html=True
            )

        c4, c5 = st.columns(2)

        with c4:

            st.markdown(
                f"""
                <div class="metric-card">
                    <h4>⏳ Pending</h4>
                    <h2>{pending_count}</h2>
                </div>
                """,
                unsafe_allow_html=True
            )

        with c5:

            st.markdown(
                f"""
                <div class="metric-card">
                    <h4>⚠ Outstanding USD</h4>
                    <h2>${outstanding_value:,.0f}</h2>
                </div>
                """,
                unsafe_allow_html=True
            )
        st.divider()

        search = st.text_input(
            "Search Invoice"
        )
        
        status_filter = st.selectbox(
            "Status Filter",
            [
                "All",
                "Pending",
                "Partially Paid",
                "Paid",
                "Cancelled"
            ]
        )

        filtered = df.copy()
        
        if status_filter != "All":

            filtered = filtered[
                filtered["Status"]
                == status_filter
            ]

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

        display_columns = [
            "Invoice No",
            "Invoice Date",
            "Client Name",
            "Client Company Name",
            "Total",
            "Status",
            "Received Amount"
        ]

        existing_cols = [
            col for col in display_columns
            if col in filtered.columns
        ]

        st.dataframe(
            filtered[existing_cols],
            use_container_width=True
        )
        
        st.divider()

        st.subheader(
            "Outstanding Invoices"
        )

        outstanding_df = df[
            df["Status"] != "Paid"
        ]

        outstanding_cols = [
            "Invoice No",
            "Client Company Name",
            "Total",
            "Status"
        ]

        existing_out_cols = [
            col
            for col in outstanding_cols
            if col in outstanding_df.columns
        ]

        st.dataframe(
            outstanding_df[
                existing_out_cols
            ],
            use_container_width=True
        )
        
        st.divider()

        st.subheader(
            "Overdue Invoices"
        )

        today = pd.Timestamp.today()

        overdue_df = df.copy()

        overdue_df["Due Date"] = pd.to_datetime(
            overdue_df["Due Date"],
            errors="coerce"
        )

        overdue_df = overdue_df[
            (overdue_df["Due Date"] < today)
            &
            (overdue_df["Status"] != "Paid")
        ]

        overdue_cols = [
            "Invoice No",
            "Client Company Name",
            "Due Date",
            "Total",
            "Status"
        ]

        existing_overdue_cols = [
            col
            for col in overdue_cols
            if col in overdue_df.columns
        ]

        st.dataframe(
            overdue_df[
                existing_overdue_cols
            ],
            use_container_width=True
        )
        
        st.divider()

        st.subheader("Update Invoice Status")

        invoice_list = df["Invoice No"].astype(str).tolist()

        selected_invoice = st.selectbox(
            "Invoice",
            invoice_list
        )

        status = st.selectbox(
            "Status",
            [
                "Pending",
                "Partially Paid",
                "Paid",
                "Cancelled"
            ]
        )

        received_amount = st.number_input(
            "Received Amount",
            min_value=0.0
        )

        notes = st.text_input(
            "Notes"
        )
        
        st.write(tracking_df.tail())

        if st.button("Update Status"):

            selected_invoice = str(
                selected_invoice
            ).replace(".0", "")

            tracking_df["Invoice No"] = (
                tracking_df["Invoice No"]
                .astype(str)
                .str.replace(
                    ".0",
                    "",
                    regex=False
                )
            )

            row_index = tracking_df[
                tracking_df["Invoice No"]
                == selected_invoice
            ].index

            if len(row_index) > 0:

                tracking_df.loc[
                    row_index,
                    "Status"
                ] = status

                tracking_df.loc[
                    row_index,
                    "Received Amount"
                ] = received_amount

                tracking_df.loc[
                    row_index,
                    "Notes"
                ] = notes

                tracking_df.to_csv(
                    tracking_file,
                    index=False
                )

                st.success(
                    "Status Updated"
                )

                st.rerun()

            else:

                st.error(
                    f"Invoice {selected_invoice} not found"
                )

    except Exception as e:

        st.error(
            f"Invoice Error: {e}"
        )