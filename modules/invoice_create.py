import pandas as pd
import os
import streamlit as st
from datetime import date
import json
from modules.pdf_generator import generate_pdf

INVOICE_FILE = "data/invoices.csv"

def get_next_invoice_number():

    if not os.path.exists(INVOICE_FILE):
        return "202606001"

    df = pd.read_csv(INVOICE_FILE)

    df = df.dropna(
        subset=["Invoice No"]
    )

    if len(df) == 0:
        return "202606001"

    try:

        last_invoice = str(
            df["Invoice No"].iloc[-1]
        )

        return str(
            int(last_invoice) + 1
        )

    except:

        return "202606001"

def load_company_settings():

    try:

        with open(
            "data/settings.json",
            "r"
        ) as f:

            return json.load(f)

    except:

        return {
            "company_name": "",
            "phone": "",
            "email": "",
            "website": "",
            "address": "",
            "gst": ""
        }

def render():

    st.title("📄 Create Invoice")
    
    settings = load_company_settings()
    
    invoice_no = get_next_invoice_number()

    st.info(
        f"Invoice Number: {invoice_no}"
    )

    col1, col2 = st.columns(2)

    with col1:

        invoice_date = st.date_input(
            "Invoice Date",
            value=date.today()
        )

        due_type = st.radio(
            "Due Date Type",
            ["Date", "Custom Terms"],
            horizontal=True
        )

        if due_type == "Date":

            due_date = st.date_input(
                "Due Date"
            )

            due_date_value = due_date.strftime(
                "%Y-%m-%d"
            )

        else:

            due_date_value = st.text_input(
                "Payment Terms",
                value="Immediate"
            )

        customer_id = st.text_input(
            "Customer ID"
        )

    with col2:

        client_name = st.text_input(
            "Client Name"
        )

        company_name = st.text_input(
            "Client Company Name"
        )

        client_phone = st.text_input(
            "Client Phone"
        )

    st.divider()

    st.subheader("Services")

    service_df = pd.DataFrame(
        [
            {
                "Description": "",
                "Unit": "Hour",
                "Rate": 20.0,
                "HoursQty": 0.0
            }
        ]
    )

    edited_df = st.data_editor(
        service_df,
        num_rows="dynamic",
        use_container_width=True,
        key="services_editor"
    )

    for col in ["Rate", "HoursQty", "Amount"]:
        if col in edited_df.columns:
            edited_df[col] = pd.to_numeric(edited_df[col], errors="coerce").fillna(0)

    edited_df["Amount"] = (
        edited_df["Rate"] *
        edited_df["HoursQty"]
    )

    st.dataframe(
        edited_df,
        use_container_width=True
    )

    subtotal = float(edited_df["Amount"].sum()) if not edited_df.empty else 0.0

    st.metric(
        "Invoice Total",
        f"${subtotal:,.2f}"
    )
    
    comments = st.text_area(
    "Comments",
    value="Please let us know if you have any queries."
)
    
    if st.button("💾 Save Invoice"):

        try:

            services = edited_df.to_dict(
                orient="records"
            )

            for service in services:

                service["Currency"] = "$"

            invoice = {

                "Invoice No": invoice_no,
                "Invoice Date": invoice_date,
                "Due Date": due_date_value,
                "Customer ID": customer_id,

                "Organization Name":
                settings.get(
                    "company_name",
                    ""
                ),

                "Organization Address":
                settings.get(
                    "address",
                    ""
                ),

                "Organization City":
                settings.get(
                    "city",
                    ""
                ),  

                "Organization Zip":
                settings.get(
                    "zip_code",
                    ""
                ),

                "Organization Phone":
                settings.get(
                    "phone",
                    ""
                ),

                "Organization Website":
                settings.get(
                    "website",
                    ""
                ),
                
                "Country":
                settings.get(
                    "country",
                    ""
                ),

                "Organization Fax": "",

                "Client Name": client_name,
                "Client Company Name": company_name,
                "Client Address": "",
                "Client City": "",
                "Client Zip": "",
                "Client Phone": client_phone,

                "Services": services,
                "Expenses": [],

                "Subtotal": subtotal,
                "Tax Rate": 0,
                "Tax Due": 0,
                "Total": subtotal,

                "Comments": comments,
                "Currency Symbol": "$"
            }

            existing = pd.read_csv(
                INVOICE_FILE
            )

            existing = pd.concat(
                [
                    existing,
                    pd.DataFrame([invoice])
                ],
                ignore_index=True
            )

            existing.to_csv(
                INVOICE_FILE,
                index=False
            )
            
            tracking_file = "data/invoice_tracking.csv"

            tracking_record = pd.DataFrame([{
                
                "Invoice No": invoice_no,
                "Status": "Pending",
                "Paid Amount": 0,
                "Payment Date": "",
                "Notes": ""
            }])
            
            if os.path.exists(tracking_file):

                try:

                    tracking_df = pd.read_csv(
                        tracking_file,
                        dtype=str
                    )

                    tracking_df = pd.concat(
                        [tracking_df, tracking_record],
                        ignore_index=True
                    )

                except:

                    tracking_df = tracking_record

            else:

                tracking_df = tracking_record

            tracking_df.to_csv(
                tracking_file,
                index=False
            )

            st.success(
                "Tracking record saved"
            )

            st.success(
                f"Invoice {invoice_no} saved successfully"
            )

            pdf_path = generate_pdf(invoice)

            st.success(
                f"PDF Created: {pdf_path}"
            )
            
            st.write("PDF Created Successfully")
            st.write(pdf_path)
            
            with open(pdf_path, "rb") as pdf_file:

                st.download_button(
                    label="📥 Download Invoice PDF",
                    data=pdf_file.read(),
                    file_name=os.path.basename(pdf_path),
                    mime="application/pdf"
                )
        except Exception as e:

            st.error(
                f"Save failed: {e}"
            )