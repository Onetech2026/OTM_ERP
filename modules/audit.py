import streamlit as st
import pandas as pd
from datetime import datetime
import os

AUDIT_FILE = "data/audit_logs.csv"


def log_action(user, action, details=""):

    row = pd.DataFrame([{
        "Timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "User": user,
        "Action": action,
        "Details": details
    }])

    if os.path.exists(AUDIT_FILE):

        try:
            df = pd.read_csv(AUDIT_FILE)
            df = pd.concat([df, row], ignore_index=True)

        except:
            df = row

    else:
        df = row

    df.to_csv(AUDIT_FILE, index=False)


def render():

    st.title("🔍 Audit Logs")

    if not os.path.exists(AUDIT_FILE):
        st.info("No audit logs found")
        return

    try:
        df = pd.read_csv(AUDIT_FILE)

        if len(df) == 0:
            st.info("No audit logs available")
            return

        st.metric("Total Activities", len(df))

        search = st.text_input(
            "Search Logs",
            placeholder="Search user, action, details..."
        )

        if search:

            mask = df.astype(str).apply(
                lambda x: x.str.contains(
                    search,
                    case=False,
                    na=False
                )
            ).any(axis=1)

            df = df[mask]

        st.dataframe(
            df.sort_values(
                "Timestamp",
                ascending=False
            ),
            use_container_width=True
        )

    except Exception as e:
        st.error(str(e))