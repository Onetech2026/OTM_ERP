import streamlit as st
import pandas as pd
import os

EMPLOYEE_FILE = "data/employees.csv"

def render():

    st.title("👥 Employee Management")

    if not os.path.exists(EMPLOYEE_FILE):
        st.warning("Employee database not found")
        return

    df = pd.read_csv(EMPLOYEE_FILE)

    st.subheader("Employee Directory")

    col1, col2 = st.columns([3,1])

    with col1:
        search = st.text_input(
            "Search Employee",
            placeholder="Name, ID, Designation..."
        )

    with col2:
        status = st.selectbox(
            "Status",
            ["All"] + sorted(df["EmployeeStatus"].dropna().unique().tolist())
        )

    filtered = df.copy()

    if search:
        filtered = filtered[
            filtered.astype(str)
            .apply(lambda row: row.str.contains(search, case=False).any(), axis=1)
        ]

    if status != "All":
        filtered = filtered[
            filtered["EmployeeStatus"] == status
        ]

    st.write(f"Total Employees: {len(filtered)}")

    st.divider()

    for _, emp in filtered.iterrows():

        with st.container():

            c1, c2, c3 = st.columns([3,2,2])

            with c1:
                st.markdown(f"### {emp['FullName']}")
                st.write(f"**ID:** {emp['EmployeeID']}")
                st.write(f"**Designation:** {emp['Designation']}")

            with c2:
                st.write(f"**Location:** {emp['WorkLocation']}")
                st.write(f"**Type:** {emp['EmploymentType']}")

            with c3:
                st.write(f"**Status:** {emp['EmployeeStatus']}")
                st.write(f"**CTC:** ₹{emp['CTC']:,.0f}")

            with st.expander("View Employee Profile"):

                tab1, tab2, tab3, tab4 = st.tabs([
                    "Personal",
                    "Employment",
                    "Compensation",
                    "Banking"
                ])

                with tab1:
                    st.write("DOB:", emp["DOB"])
                    st.write("Gender:", emp["Gender"])
                    st.write("Phone:", emp["Phone"])
                    st.write("Email:", emp["Email"])
                    st.write("Current Address:", emp["CurrentAddress"])

                with tab2:
                    st.write("Joining Date:", emp["JoiningDate"])
                    st.write("Manager:", emp["ReportingManager"])
                    st.write("Work Location:", emp["WorkLocation"])
                    st.write("Employment Type:", emp["EmploymentType"])

                with tab3:
                    st.write("CTC:", emp["CTC"])
                    st.write("Basic:", emp["BasicPay"])
                    st.write("HRA:", emp["HRA"])
                    st.write("Travel:", emp["TravelAllowance"])
                    st.write("Medical:", emp["MedicalAllowance"])
                    st.write("Internet:", emp["InternetAllowance"])

                with tab4:
                    st.write("Bank:", emp["BankName"])
                    st.write("Account:", emp["AccountNumber"])
                    st.write("IFSC:", emp["IFSCCode"])
                    st.write("PAN:", emp["PAN"])

            st.divider()