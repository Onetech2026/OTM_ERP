import os

import pandas as pd
import streamlit as st


EMPLOYEE_FILE = "data/employees.csv"
EMPLOYMENT_TYPES = ["Contract", "Full Time", "Full-time", "Intern"]


def load_employees():
    if not os.path.exists(EMPLOYEE_FILE):
        return pd.DataFrame()
    df = pd.read_csv(EMPLOYEE_FILE)

    # Ensure commonly-updated contact and identifier columns are strings
    string_cols = [
        "EmployeeID",
        "Phone",
        "AccountNumber",
        "IFSCCode",
        "PAN",
        "Aadhar",
        "FullName",
        "Email",
    ]

    numeric_cols = [
        "CTC",
        "BasicPay",
        "HRA",
        "TravelAllowance",
        "MedicalAllowance",
        "InternetAllowance",
        "SpecialAllowance",
    ]

    for col in string_cols:
        if col in df.columns:
            df[col] = df[col].fillna("").astype(str)

    for col in numeric_cols:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

    return df


def save_employees(df):
    os.makedirs("data", exist_ok=True)
    df.to_csv(EMPLOYEE_FILE, index=False)


def render():
    st.title("👥 Employee Management")

    df = load_employees()
    # Defensive coercion for in-memory DataFrame to avoid dtype assignment issues
    string_cols = [
        "EmployeeID",
        "Phone",
        "AccountNumber",
        "IFSCCode",
        "PAN",
        "Aadhar",
        "FullName",
        "Email",
        "DOB",
        "Gender",
        "CurrentAddress",
        "PermanentAddress",
        "EmergencyContactName",
        "EmergencyRelation",
        "EmergencyPhone",
        "Designation",
        "JoiningDate",
        "EmploymentType",
        "ReportingManager",
        "WorkLocation",
        "EmployeeStatus",
        "BankName",
        "Branch",
    ]
    numeric_cols = [
        "CTC",
        "BasicPay",
        "HRA",
        "TravelAllowance",
        "MedicalAllowance",
        "InternetAllowance",
        "SpecialAllowance",
        "PFDeduction",
        "ESIDeduction",
        "TDSDeduction",
        "OtherDeductions",
    ]

    for _col in string_cols:
        if _col in df.columns:
            df[_col] = df[_col].fillna("").astype("string")

    for _col in numeric_cols:
        if _col in df.columns:
            df[_col] = pd.to_numeric(df[_col], errors="coerce").fillna(0)

    if df.empty:
        st.warning("Employee database not found or is empty. Add employees first.")
        return

    directory_tab, add_tab, edit_tab = st.tabs(
        ["👥 Employee Directory", "➕ Add Employee", "✏️ Edit Employee"]
    )

    with directory_tab:
        col1, col2 = st.columns([3, 1])

        with col1:
            search = st.text_input(
                "Search Employee",
                placeholder="Name, ID, Designation...",
            )

        with col2:
            status = st.selectbox(
                "Status",
                ["All"] + sorted(df["EmployeeStatus"].dropna().astype(str).unique().tolist()),
            )

        filtered = df.copy()

        if search:
            filtered = filtered[
                filtered.astype(str)
                .apply(lambda row: row.str.contains(search, case=False).any(), axis=1)
            ]

        if status != "All":
            filtered = filtered[filtered["EmployeeStatus"].astype(str) == status]

        st.write(f"Total Employees: {len(filtered)}")
        st.divider()

        for _, emp in filtered.iterrows():
            with st.container():
                c1, c2, c3 = st.columns([3, 2, 2])

                with c1:
                    st.markdown(f"### {emp.get('FullName', 'N/A')}")
                    st.write(f"**ID:** {emp.get('EmployeeID', 'N/A')}")
                    st.write(f"**Designation:** {emp.get('Designation', 'N/A')}")

                with c2:
                    st.write(f"**Location:** {emp.get('WorkLocation', 'N/A')}")
                    st.write(f"**Type:** {emp.get('EmploymentType', 'N/A')}")

                with c3:
                    st.write(f"**Status:** {emp.get('EmployeeStatus', 'N/A')}")
                    st.write(f"**CTC:** ₹{float(emp.get('CTC', 0) or 0):,.0f}")

                with st.expander("View Employee Profile"):
                    personal_tab, employment_tab, compensation_tab, banking_tab = st.tabs(
                        ["Personal", "Employment", "Compensation", "Banking"]
                    )

                    with personal_tab:
                        st.write("DOB:", emp.get("DOB", ""))
                        st.write("Gender:", emp.get("Gender", ""))
                        st.write("Phone:", emp.get("Phone", ""))
                        st.write("Email:", emp.get("Email", ""))
                        st.write("Current Address:", emp.get("CurrentAddress", ""))

                    with employment_tab:
                        st.write("Joining Date:", emp.get("JoiningDate", ""))
                        st.write("Manager:", emp.get("ReportingManager", ""))
                        st.write("Work Location:", emp.get("WorkLocation", ""))
                        st.write("Employment Type:", emp.get("EmploymentType", ""))

                    with compensation_tab:
                        st.write("CTC:", emp.get("CTC", 0))
                        st.write("Basic:", emp.get("BasicPay", 0))
                        st.write("HRA:", emp.get("HRA", 0))
                        st.write("Travel:", emp.get("TravelAllowance", 0))
                        st.write("Medical:", emp.get("MedicalAllowance", 0))
                        st.write("Internet:", emp.get("InternetAllowance", 0))

                    with banking_tab:
                        st.write("Bank:", emp.get("BankName", ""))
                        st.write("Account:", emp.get("AccountNumber", ""))
                        st.write("IFSC:", emp.get("IFSCCode", ""))
                        st.write("PAN:", emp.get("PAN", ""))

                st.divider()

    with add_tab:
        st.subheader("➕ Add Employee")

        c1, c2 = st.columns(2)

        with c1:
            employee_id = st.text_input("Employee ID", placeholder="EMP0001")
            full_name = st.text_input("Full Name")
            dob = st.text_input("DOB")
            gender = st.selectbox("Gender", ["Male", "Female", "Other"])
            phone = st.text_input("Phone")
            email = st.text_input("Email")
            current_address = st.text_area("Current Address")
            permanent_address = st.text_area("Permanent Address")
            emergency_name = st.text_input("Emergency Contact Name")
            emergency_relation = st.text_input("Emergency Relation")
            emergency_phone = st.text_input("Emergency Phone")

        with c2:
            designation = st.text_input("Designation")
            joining_date = st.text_input("Joining Date")
            employment_type = st.selectbox("Employment Type", EMPLOYMENT_TYPES)
            manager = st.text_input("Reporting Manager")
            work_location = st.text_input("Work Location")
            status = st.selectbox("Status", ["Active", "Inactive"])
            ctc = st.number_input("CTC", min_value=0.0)
            basic = st.number_input("Basic Pay", min_value=0.0)
            hra = st.number_input("HRA", min_value=0.0)
            travel = st.number_input("Travel Allowance", min_value=0.0)
            medical = st.number_input("Medical Allowance", min_value=0.0)
            internet = st.number_input("Internet Allowance", min_value=0.0)
            special = st.number_input("Special Allowance", min_value=0.0)

        st.divider()
        st.subheader("Banking Information")

        b1, b2 = st.columns(2)
        with b1:
            bank_name = st.text_input("Bank Name")
            branch = st.text_input("Branch")
            account_number = st.text_input("Account Number")

        with b2:
            ifsc = st.text_input("IFSC Code")
            pan = st.text_input("PAN")
            aadhar = st.text_input("Aadhar")

        if st.button("Save Employee", use_container_width=True):
            try:
                if not full_name.strip():
                    st.error("Please enter the employee name before saving.")
                    return

                if not employee_id.strip():
                    employee_id = f"EMP{len(df) + 1:04d}"

                new_employee = {
                "EmployeeID": employee_id,
                "FullName": full_name,
                "DOB": dob,
                "Gender": gender,
                "Phone": phone,
                "Email": email,
                "CurrentAddress": current_address,
                "PermanentAddress": permanent_address,
                "EmergencyContactName": emergency_name,
                "EmergencyRelation": emergency_relation,
                "EmergencyPhone": emergency_phone,
                "Designation": designation,
                "JoiningDate": joining_date,
                "EmploymentType": employment_type,
                "ReportingManager": manager,
                "WorkLocation": work_location,
                "EmployeeStatus": status,
                "CTC": ctc,
                "BasicPay": basic,
                "HRA": hra,
                "TravelAllowance": travel,
                "MedicalAllowance": medical,
                "InternetAllowance": internet,
                "SpecialAllowance": special,
                "PFDeduction": 0,
                "ESIDeduction": 0,
                "TDSDeduction": 0,
                "OtherDeductions": 0,
                "BankName": bank_name,
                "Branch": branch,
                "AccountNumber": account_number,
                "IFSCCode": ifsc,
                "PAN": pan,
                "Aadhar": aadhar,
            }

                updated_df = pd.concat([df, pd.DataFrame([new_employee])], ignore_index=True)
                save_employees(updated_df)
                # refresh activity timestamp
                try:
                    st.session_state.login_started_at = time.time()
                except Exception:
                    pass
                st.success("Employee Added Successfully")
                st.rerun()
            except Exception as e:
                st.error(f"Failed to add employee: {e}")

    with edit_tab:
        st.subheader("✏️ Edit Employee")

        if df.empty:
            st.info("No employee records available to edit.")
            return

        employee_id = st.selectbox(
            "Select Employee",
            options=df["EmployeeID"].astype(str).tolist(),
            format_func=lambda value: f"{value} - {df.loc[df['EmployeeID'].astype(str) == value, 'FullName'].iloc[0]}",
        )

        selected = df[df["EmployeeID"].astype(str) == employee_id].iloc[0].copy()

        with st.form("edit_employee_form"):
            full_name = st.text_input("Full Name", value=selected.get("FullName", ""))
            phone = st.text_input("Phone", value=selected.get("Phone", ""))
            email = st.text_input("Email", value=selected.get("Email", ""))
            designation = st.text_input("Designation", value=selected.get("Designation", ""))
            selected_employment_type = str(selected.get("EmploymentType", "Full Time") or "Full Time").strip()
            employment_type_index = (
                EMPLOYMENT_TYPES.index(selected_employment_type)
                if selected_employment_type in EMPLOYMENT_TYPES
                else 1
            )
            employment_type = st.selectbox(
                "Employment Type",
                EMPLOYMENT_TYPES,
                index=employment_type_index,
            )
            manager = st.text_input("Reporting Manager", value=selected.get("ReportingManager", ""))
            work_location = st.text_input("Work Location", value=selected.get("WorkLocation", ""))
            status = st.selectbox(
                "Status",
                ["Active", "Inactive"],
                index=["Active", "Inactive"].index(selected.get("EmployeeStatus", "Active")),
            )
            ctc = st.number_input("CTC", min_value=0.0, value=float(selected.get("CTC", 0) or 0))

            submitted = st.form_submit_button("Update Employee")

        if submitted:
            try:
                if not full_name.strip():
                    st.error("Employee name is required.")
                else:
                    mask = df["EmployeeID"].astype(str) == employee_id
                    df.loc[mask, "FullName"] = str(full_name)
                    df.loc[mask, "Phone"] = str(phone)
                    df.loc[mask, "Email"] = str(email)
                    df.loc[mask, "Designation"] = str(designation)
                    df.loc[mask, "EmploymentType"] = str(employment_type)
                    df.loc[mask, "ReportingManager"] = str(manager)
                    df.loc[mask, "WorkLocation"] = str(work_location)
                    df.loc[mask, "EmployeeStatus"] = str(status)
                    df.loc[mask, "CTC"] = float(ctc)

                    save_employees(df)
                    try:
                        st.session_state.login_started_at = time.time()
                    except Exception:
                        pass
                    st.success("Employee updated successfully!")
                    st.rerun()
            except Exception as e:
                st.error(f"Failed to update employee: {e}")