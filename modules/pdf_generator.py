from datetime import datetime, date
import datetime as dt
import streamlit as st
import pandas as pd
import os
from fpdf import FPDF
import json
import win32com.client
import pythoncom
import tempfile
from io import BytesIO
import plotly.express as px

# --- Constants ---
DATA_DIR = "data"
INVOICE_FILE = os.path.join(DATA_DIR, "invoices.csv")
EXPENSE_FILE = os.path.join(DATA_DIR, "expenses.csv")
PAYSLIP_FILE = os.path.join(DATA_DIR, "payslips.csv")
EMPLOYEE_FILE = os.path.join(DATA_DIR, "employees.csv")
PDF_FOLDER = "invoices_pdf"
TRANSACTIONS_FILE = os.path.join(DATA_DIR, "transactions.csv")

# Create required directories
for directory in [DATA_DIR, PDF_FOLDER]:
    if not os.path.exists(directory):
        os.makedirs(directory)

def load_employees():
    if os.path.exists(EMPLOYEE_FILE):
        return pd.read_csv(EMPLOYEE_FILE)
    return pd.DataFrame(columns=[
        "EmployeeID", "FullName", "DOB", "Gender", "Phone", "Email",
        "CurrentAddress", "PermanentAddress", "EmergencyContactName",
        "EmergencyRelation", "EmergencyPhone", "Designation", "JoiningDate",
        "EmploymentType", "ReportingManager", "WorkLocation", "EmployeeStatus",
        "CTC", "BasicPay", "HRA", "TravelAllowance", "MedicalAllowance",
        "InternetAllowance",
        "SpecialAllowance", "PFDeduction", "ESIDeduction", "TDSDeduction",
        "OtherDeductions", "BankName", "Branch", "AccountNumber", "IFSCCode",
        "PAN", "Aadhar"
    ])

def save_employees(df):   # ✅ top-level function
    df.to_csv(EMPLOYEE_FILE, index=False)

def generate_employee_id():
    df = load_employees()
    if df.empty:
        return "EMP0001"
    existing_nums = [int(id.replace("EMP", "")) for id in df["EmployeeID"].values]
    next_num = max(existing_nums) + 1
    return f"EMP{next_num:04d}"

def load_invoices():
    if os.path.exists(INVOICE_FILE):
        return pd.read_csv(INVOICE_FILE)
    return pd.DataFrame(columns=[
        "Invoice No", "Invoice Date", "Due Date", "Customer ID",
        "Organization Name", "Organization Address", "Organization City", "Organization Zip",
        "Organization Phone", "Organization Website", "Organization Fax",
        "Client Name", "Client Company Name", "Client Address", "Client City", "Client Zip", "Client Phone",
        "Services", "Expenses", "Subtotal", "Tax Rate", "Tax Due", "Total", "Comments", "Currency Symbol"
    ])

def save_invoice(invoice):
    invoices = load_invoices()
    invoices = pd.concat([invoices, pd.DataFrame([invoice])], ignore_index=True)
    invoices.to_csv(INVOICE_FILE, index=False)

def send_email_outlook(receiver_email, subject, body, pdf_path):
    try:
        # Initialize COM for Outlook
        pythoncom.CoInitialize()
        
        # Verify PDF exists
        if not os.path.exists(pdf_path):
            st.error(f"PDF file not found at: {pdf_path}")
            return False
            
        # Create Outlook application object
        outlook = win32com.client.Dispatch('Outlook.Application')
        
        # Create a new mail item
        mail = outlook.CreateItem(0)  # 0 = olMailItem
        
        # Set email properties
        mail.To = receiver_email
        mail.Subject = subject
        mail.Body = body
        
        # Add attachment - use absolute path
        absolute_path = os.path.abspath(pdf_path)
        mail.Attachments.Add(absolute_path)
        
        # Send the email
        mail.Send()
        
        # Clean up
        pythoncom.CoUninitialize()
        return True
        
    except Exception as e:
        st.error(f"Failed to send email via Outlook: {str(e)}")
        return False

def wrap_cell(pdf, width, height, text, border=1):
    # Calculate lines that fit in the cell
    lines = []
    words = str(text).split()
    line = ''
    
    # Word wrap logic
    for word in words:
        test_line = f"{line} {word}".strip()
        if pdf.get_string_width(test_line) < width - 4:
            line = test_line
        else:
            lines.append(line)
            line = word
    if line:
        lines.append(line)
    
    # Calculate required height for all lines
    line_height = 5  # height per line
    total_height = max(height, len(lines) * line_height)
    
    # Get starting position
    x_pos = pdf.get_x()
    y_pos = pdf.get_y()
    
    # Print all lines
    for i, line in enumerate(lines):
        pdf.set_xy(x_pos, y_pos + (i * line_height))
        pdf.cell(width, line_height, line.strip(), 0, 0)
    
    # Draw border around entire cell
    pdf.rect(x_pos, y_pos, width, total_height)
    
    # Move to next cell position
    pdf.set_xy(x_pos + width, y_pos)
    
    return total_height

def generate_pdf(invoice):
    pdf = FPDF()
    pdf.add_page()
    pdf.set_auto_page_break(auto=True, margin=15)
    blue = (54, 95, 145)

    # --- Logo and Header ---
    logo_width = 31
    try:
        pdf.image("assets/logo.png", x=12, y=10, w=logo_width)
    except:
        pass

    # Company Address (right aligned)
    pdf.set_xy(12 + logo_width + 10, 12)
    pdf.set_font("Arial", "", 10)
    pdf.set_text_color(50, 50, 50)
    pdf.multi_cell(0, 5,
        f"{invoice.get('Organization Name', '')}\n"
        f"{invoice.get('Organization Address', '')}\n"
        f"{invoice.get('Organization City', '')}, "
        f"{invoice.get('Country', '')}, "
        f"{invoice.get('Organization Zip', '')}\n"
        f"Phone: {invoice.get('Organization Phone', '')}\n"
        f"website: {invoice.get('Organization Website', '')}",
        align="R"
    )

    # Separator line
    pdf.set_draw_color(200, 200, 200)
    pdf.set_line_width(0.3)
    pdf.line(10, 40, 200, 40)

    # INVOICE heading
    pdf.ln(15)
    pdf.set_font("Arial", "B", 20)
    pdf.set_text_color(*blue)
    pdf.cell(0, 12, "INVOICE", ln=1, align="C")

    # Invoice Details
    pdf.set_font("Arial", "", 11)
    pdf.set_text_color(0, 0, 0)
    
    # Invoice info in two columns
    col_width = 95
    pdf.cell(col_width, 8, f"Invoice #: {invoice.get('Invoice No', '')}", border=0)
    pdf.cell(col_width, 8, f"Date: {invoice.get('Invoice Date', '')}", border=0, align="R", ln=1)
    
    pdf.cell(col_width, 8, f"Customer ID: {invoice.get('Customer ID', '')}", border=0)
    pdf.cell(col_width, 8, f"Due Date: {invoice.get('Due Date', '')}", border=0, align="R", ln=1)

    # Bill To section
    pdf.ln(4)
    pdf.set_font("Arial", "B", 12)
    pdf.set_text_color(*blue)
    pdf.cell(0, 8, "Bill To", ln=1)
    pdf.set_font("Arial", "", 11)
    pdf.set_text_color(0, 0, 0)
    pdf.cell(0, 6, invoice.get("Client Company Name", ""), ln=1)
    pdf.cell(0, 6, invoice.get("Client Address", ""), ln=1)
    pdf.cell(0, 6, f"{invoice.get('Client City', '')}, {invoice.get('Client Zip', '')}", ln=1)
    pdf.cell(0, 6, f"Phone: {invoice.get('Client Phone', '')}", ln=1)

    # Services Table
    services = invoice.get("Services", [])
    if services:  # Changed from json.loads()
        pdf.ln(5)
        pdf.set_font("Arial", "B", 12)
        pdf.set_text_color(*blue)
        pdf.cell(0, 8, "Services", ln=1)
        pdf.set_font("Arial", "B", 11)
        pdf.set_text_color(255, 255, 255)
        pdf.set_fill_color(*blue)
        
        # Table headers
        pdf.cell(60, 8, "Description", border=1, fill=True)
        pdf.cell(30, 8, "Unit", border=1, fill=True)
        pdf.cell(30, 8, "Rate", border=1, fill=True)
        pdf.cell(30, 8, "Qty", border=1, fill=True)
        pdf.cell(30, 8, "Amount", border=1, ln=1, fill=True)
        
        pdf.set_text_color(0, 0, 0)
        pdf.set_font("Arial", "", 11)

        # Table rows with word wrap
        for s in services:
            desc = str(s.get("Description", ""))
            current_y = pdf.get_y()
            
            # Get height needed for description
            row_height = wrap_cell(pdf, 60, 8, desc)
            
            # Reset position and print other cells with matched height
            pdf.set_y(current_y)
            pdf.cell(60, row_height, "", border=0)  # Placeholder for description
            pdf.cell(30, row_height, str(s.get("Unit", "")), border=1)
            pdf.cell(30, row_height, f"{s.get('Rate', 0):,.2f}", border=1, align="R")
            pdf.cell(30, row_height, f"{s.get('HoursQty', 0):,.2f}", border=1, align="R")
            pdf.cell(30, row_height, f"{s.get('Amount', 0):,.2f}", border=1, ln=1, align="R")

    # --- Expenses Table (only if expenses exist) ---
    expenses = invoice.get("Expenses", [])  # Changed from json.loads()
    if expenses:
        pdf.ln(5)
        pdf.set_font("Arial", "B", 12)
        pdf.set_text_color(*blue)
        pdf.cell(0, 8, "Expenses", ln=1)
        pdf.set_font("Arial", "B", 11)
        pdf.set_text_color(255, 255, 255)
        pdf.set_fill_color(*blue)
        pdf.cell(60, 8, "Description", border=1, fill=True)
        pdf.cell(30, 8, "Unit", border=1, fill=True)
        pdf.cell(30, 8, "Rate", border=1, fill=True)
        pdf.cell(30, 8, "Qty", border=1, fill=True)
        pdf.cell(30, 8, "Amount", border=1, ln=1, fill=True)
        pdf.set_text_color(0, 0, 0)
        pdf.set_font("Arial", "", 11)

        for e in expenses:
            desc = str(e.get("Description", ""))
            current_y = pdf.get_y()
            
            # Get height needed for description
            row_height = wrap_cell(pdf, 60, 8, desc)
            
            # Reset position and print other cells with matched height
            pdf.set_y(current_y)
            pdf.cell(60, row_height, "", border=0)  # Placeholder for description
            pdf.cell(30, row_height, str(e.get("Unit", "")), border=1)
            pdf.cell(30, row_height, f"{e.get('Rate', 0):,.2f}", border=1, align="R")
            pdf.cell(30, row_height, f"{e.get('Qty', 0):,.2f}", border=1, align="R")
            pdf.cell(30, row_height, f"{e.get('Amount', 0):,.2f}", border=1, ln=1, align="R")

    # --- Totals Section (aligned right) ---
    pdf.ln(5)
    pdf.set_font("Arial", "B", 12)
    total_width = 180  # Total width of the page
    label_width = 140  # Width for labels
    amount_width = 40  # Width for amounts

    # Subtotal
    pdf.cell(label_width, 8, "Subtotal:", border=0, align="R")
    pdf.cell(amount_width, 8, f"{invoice.get('Currency Symbol', '₹')} {invoice.get('Subtotal', 0):,.2f}", 
             border=0, ln=1, align="R")
    
    # Tax Rate
    pdf.cell(label_width, 8, "Tax Rate:", border=0, align="R")
    pdf.cell(amount_width, 8, f"{invoice.get('Tax Rate', 0)}%", 
             border=0, ln=1, align="R")
    
    # Tax Due
    pdf.cell(label_width, 8, "Tax Due:", border=0, align="R")
    pdf.cell(amount_width, 8, f"{invoice.get('Currency Symbol', '₹')} {invoice.get('Tax Due', 0):,.2f}", 
             border=0, ln=1, align="R")
    
    # Total (with line above)
    pdf.set_line_width(0.3)
    pdf.line(label_width - 40, pdf.get_y(), total_width, pdf.get_y())
    pdf.cell(label_width, 10, "Total:", border=0, align="R")
    pdf.cell(amount_width, 10, f"{invoice.get('Currency Symbol', '₹')} {invoice.get('Total', 0):,.2f}", 
             border=0, ln=1, align="R")

    # --- Comments ---
    pdf.ln(8)
    pdf.set_font("Arial", "", 11)
    pdf.multi_cell(0, 8, f"Comments: {invoice.get('Comments', '')}")

    # --- Footer ---
    pdf.ln(10)
    pdf.set_font("Arial", "I", 10)
    pdf.cell(0, 8, "Thank you for your business!", ln=1, align='C')

    pdf_path = os.path.join(PDF_FOLDER, f"Invoice_{invoice.get('Invoice No', 'Unknown')}.pdf")
    pdf.output(pdf_path)
    return pdf_path

def auto_invoice_number():
    today = date.today()
    prefix = today.strftime("%Y%m")
    invoices = load_invoices()
    if invoices.empty:
        return f"{prefix}001"
    current_month_invoices = invoices[invoices["Invoice No"].astype(str).str.startswith(prefix)]
    if len(current_month_invoices) == 0:
        num = 1
    else:
        try:
            # Normalize strings: remove common float suffixes like '.0' produced by CSV/spreadsheet exports
            inv_strs = current_month_invoices["Invoice No"].astype(str).str.replace(r'\.0+$', '', regex=True)

            # Now take the last 3 characters (the sequence number) and parse safely
            seq_part = inv_strs.str[-3:]
            last_nums = pd.to_numeric(seq_part, errors="coerce").fillna(0).astype(int)
            last_num = int(last_nums.max())
            num = last_num + 1
        except Exception:
            # Fallback: if anything goes wrong, start numbering from 1 for the month
            num = 1
    return f"{prefix}{num:03d}"

def load_transactions():
    if os.path.exists(TRANSACTIONS_FILE):
        return pd.read_csv(TRANSACTIONS_FILE)
    return pd.DataFrame(columns=[
        "TransactionID", "Date", "Type", "Category", "Amount", 
        "Description", "PaymentMode", "Status", "LinkedTo", 
        "Reference", "Notes", "CreatedAt"
    ])

def save_transaction(df):
    df.to_csv(TRANSACTIONS_FILE, index=False)

def generate_transaction_id():
    df = load_transactions()
    if df.empty:
        return "TRX0001"
    existing_nums = [int(id.replace("TRX", "")) for id in df["TransactionID"].values]
    next_num = max(existing_nums) + 1
    return f"TRX{next_num:04d}"

# --- Main App ---
def main():
    st.set_page_config(page_title="Business Management System", layout="wide")

    # Add custom CSS styling
    st.markdown("""
        <style>
        /* Main title styling */
        .title {
            color: #1E88E5;
            font-size: 40px;
            font-weight: bold;
            padding: 20px 0;
            text-align: center;
            border-bottom: 2px solid #1E88E5;
            margin-bottom: 30px;
        }
        
        /* Card styling for metrics */
        .metric-card {
            background-color: #f8f9fa;
            padding: 20px;
            border-radius: 10px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        }
        
        /* Success metrics (green) */
        .success-metric {
            color: #28a745;
            font-size: 24px;
            font-weight: bold;
        }
        
        /* Warning metrics (orange) */
        .warning-metric {
            color: #ffc107;
            font-size: 24px;
            font-weight: bold;
        }
        
        /* Danger metrics (red) */
        .danger-metric {
            color: #dc3545;
            font-size: 24px;
            font-weight: bold;
        }
        
        /* Section headers */
        .section-header {
            color: #1E88E5;
            font-size: 24px;
            font-weight: bold;
            padding: 10px 0;
            margin: 20px 0;
        }
        
        /* Sidebar styling */
        .css-1d391kg {
            background-color: #f8f9fa;
        }
        
        /* Button styling */
        .stButton>button {
            background-color: #1E88E5;
            color: white;
            border: none;
            border-radius: 5px;
            padding: 10px 20px;
        }
        
        .stButton>button:hover {
            background-color: #1976D2;
        }
        
        /* Tab styling */
        .stTabs [data-baseweb="tab-list"] {
            gap: 20px;
        }
        
        .stTabs [data-baseweb="tab"] {
            color: #1E88E5;
            border-radius: 5px;
        }
        
        /* DataFrame styling */
        .dataframe {
            border: 1px solid #ddd;
            border-radius: 5px;
            padding: 10px;
        }
        </style>
    """, unsafe_allow_html=True)

    # Use HTML for main title
    st.markdown('<p class="title">Business Management System</p>', unsafe_allow_html=True)

    # Update metrics display
    def display_metric(label, value, type="success"):
        st.markdown(f"""
            <div class="metric-card">
                <p>{label}</p>
                <p class="{type}-metric">{value}</p>
            </div>
        """, unsafe_allow_html=True)

    # Update section headers
    def section_header(title):
        st.markdown(f'<p class="section-header">{title}</p>', unsafe_allow_html=True)

    # Initialize session state
    if 'services' not in st.session_state:
        st.session_state.services = []

    # Sidebar navigation
    menu = st.sidebar.selectbox(
        "Menu",
        ["Employee Database", "Generate Invoice", "View Invoices", 
         "Generate Payslip", "View Payslips", "Income & Expense Tracker"]
    )

    if menu == "Generate Invoice":
        invoice_generator()
    elif menu == "View Invoices":
        view_invoices()
    elif menu == "Employee Database":
        st.header("👥 Employee Database")
        
        tab1, tab2, tab3 = st.tabs(["Add Employee", "View/Edit Employees", "Reports"])
        
        with tab1:
            new_emp_id = generate_employee_id()
            
            with st.form("employee_form", clear_on_submit=True):
                st.subheader("🧑💼 Personal Information")
                col1, col2 = st.columns(2)
                with col1:
                    full_name = st.text_input("Full Name")
                    st.info(f"Employee ID: {new_emp_id}")
                    dob = st.date_input(
                    "Date of Birth",
                    min_value=dt.date(1900, 1, 1),
                    max_value=dt.date.today(),
                    help="Select a valid date of birth (cannot be in the future)"
                )
                    gender = st.selectbox("Gender", ["Male", "Female", "Other"])
                with col2:
                    phone = st.text_input("Phone Number")
                    email = st.text_input("Email")
                
                st.subheader("📍 Address")
                current_address = st.text_area("Current Address")
                permanent_address = st.text_area("Permanent Address")
                
                st.subheader("🆘 Emergency Contact")
                col3, col4 = st.columns(2)
                with col3:
                    emergency_name = st.text_input("Emergency Contact Name")
                    emergency_relation = st.text_input("Relation")
                with col4:
                    emergency_phone = st.text_input("Emergency Contact Phone")
                
                st.subheader("🏢 Employment Details")
                col5, col6 = st.columns(2)
                with col5:
                    designation = st.text_input("Job Title/Designation")
                    joining_date = st.date_input("Date of Joining")
                    emp_type = st.selectbox("Employment Type", 
                        ["Full-time", "Part-time", "Contract", "Intern"])
                with col6:
                    reporting_manager = st.text_input("Reporting Manager")
                    work_location = st.selectbox("Work Location", 
                        ["Office", "Remote", "Hybrid"])
                    emp_status = st.selectbox("Employee Status", 
                        ["Active", "Resigned", "Terminated"])
                
                st.subheader("💵 Payroll & Financial Information")
                col7, col8 = st.columns(2)
                with col7:
                    ctc = st.number_input("CTC (Annual)", min_value=0.0)
                    basic_pay = st.number_input("Basic Pay (Monthly)", min_value=0.0)
                    hra = st.number_input("HRA", min_value=0.0)
                    travel_allowance = st.number_input("Travel Allowance", min_value=0.0)
                    medical_allowance = st.number_input("Medical Allowance", min_value=0.0)
                    special_allowance = st.number_input("Special Allowance", min_value=0.0)
                with col8:
                    pf = st.number_input("PF Deduction", min_value=0.0)
                    esi = st.number_input("ESI Deduction", min_value=0.0)
                    # Calculate TDS automatically as 10% of Basic Pay
                    tds = basic_pay * 0.10
                    st.info(f"TDS Deduction (10% of Basic Pay): ₹{tds:,.2f}")
                    other_deductions = st.number_input("Other Deductions", min_value=0.0)
                
                st.subheader("🏦 Bank Details")
                col9, col10 = st.columns(2)
                with col9:
                    bank_name = st.text_input("Bank Name")
                    branch = st.text_input("Branch")
                    account_no = st.text_input("Account Number")
                with col10:
                    ifsc_code = st.text_input("IFSC Code")
                    pan = st.text_input("PAN Number")
                    aadhar = st.text_input("Aadhar Number")
                
                submitted = st.form_submit_button("Save Employee Data")
                
                if submitted:
                    if not full_name:
                        st.error("Please enter employee name")
                        return
                    
                    emp_data = {
                        "EmployeeID": new_emp_id,
                        "FullName": full_name,
                        "DOB": dob.strftime("%Y-%m-%d"),
                        "Gender": gender,
                        "Phone": phone,
                        "Email": email,
                        "CurrentAddress": current_address,
                        "PermanentAddress": permanent_address,
                        "EmergencyContactName": emergency_name,
                        "EmergencyRelation": emergency_relation,
                        "EmergencyPhone": emergency_phone,
                        "Designation": designation,
                        "JoiningDate": joining_date.strftime("%Y-%m-%d"),
                        "EmploymentType": emp_type,
                        "ReportingManager": reporting_manager,
                        "WorkLocation": work_location,
                        "EmployeeStatus": emp_status,
                        "CTC": ctc,
                        "BasicPay": basic_pay,
                        "HRA": hra,
                        "TravelAllowance": travel_allowance,
                        "MedicalAllowance": medical_allowance,
        "InternetAllowance": st.number_input("Internet Allowance", min_value=0.0),
                        "SpecialAllowance": special_allowance,
                        "PFDeduction": pf,
                        "ESIDeduction": esi,
                        "TDSDeduction": tds,  # Use calculated TDS value
                        "OtherDeductions": other_deductions,
                        "BankName": bank_name,
                        "Branch": branch,
                        "AccountNumber": account_no,
                        "IFSCCode": ifsc_code,
                        "PAN": pan,
                        "Aadhar": aadhar
                    }
                    
                    df = load_employees()
                    df = pd.concat([df, pd.DataFrame([emp_data])], ignore_index=True)
                    save_employees(df)
                    st.success(f"Employee data saved for {full_name}")
        
        with tab2:
            df = load_employees()
            if not df.empty:
                st.subheader("Employee Directory & Edit")
                selected_emp = st.selectbox("Select Employee", df["EmployeeID"].unique())
                emp_data = df[df["EmployeeID"] == selected_emp].iloc[0]
                st.write("### Employee Details")
                st.dataframe(pd.DataFrame([emp_data]))  # Display employee details neatly in a table

                st.write("### Edit Employee")
                with st.form("edit_employee_form"):
                    full_name = st.text_input("Full Name", value=emp_data["FullName"])
                    dob = st.date_input(
                    "Date of Birth",
                    value=pd.to_datetime(emp_data["DOB"]),
                    min_value=dt.date(1900, 1, 1),
                    max_value=dt.date.today(),
                    help="Date of Birth cannot be a future date"
                )
                    gender = st.selectbox("Gender", ["Male", "Female", "Other"], index=["Male", "Female", "Other"].index(emp_data["Gender"]))
                    phone = st.text_input("Phone Number", value=emp_data["Phone"])
                    email = st.text_input("Email", value=emp_data["Email"])
                    current_address = st.text_area("Current Address", value=emp_data["CurrentAddress"])
                    permanent_address = st.text_area("Permanent Address", value=emp_data["PermanentAddress"])
                    emergency_name = st.text_input("Emergency Contact Name", value=emp_data["EmergencyContactName"])
                    emergency_relation = st.text_input("Relation", value=emp_data["EmergencyRelation"])
                    emergency_phone = st.text_input("Emergency Contact Phone", value=emp_data["EmergencyPhone"])
                    designation = st.text_input("Job Title/Designation", value=emp_data["Designation"])
                    joining_date = st.date_input("Date of Joining", value=pd.to_datetime(emp_data["JoiningDate"]))
                    emp_type = st.selectbox("Employment Type", ["Full-time", "Part-time", "Contract", "Intern"], index=["Full-time", "Part-time", "Contract", "Intern"].index(emp_data["EmploymentType"]))
                    reporting_manager = st.text_input("Reporting Manager", value=emp_data["ReportingManager"])
                    work_location = st.selectbox("Work Location", ["Office", "Remote", "Hybrid"], index=["Office", "Remote", "Hybrid"].index(emp_data["WorkLocation"]))
                    emp_status = st.selectbox("Employee Status", ["Active", "Resigned", "Terminated"], index=["Active", "Resigned", "Terminated"].index(emp_data["EmployeeStatus"]))
                    ctc = st.number_input("CTC (Annual)", min_value=0.0, value=float(emp_data["CTC"]))
                    basic_pay = st.number_input("Basic Pay (Monthly)", min_value=0.0, value=float(emp_data["BasicPay"]))
                    hra = st.number_input("HRA", min_value=0.0, value=float(emp_data["HRA"]))
                    travel_allowance = st.number_input("Travel Allowance", min_value=0.0, value=float(emp_data["TravelAllowance"]))
                    medical_allowance = st.number_input("Medical Allowance", min_value=0.0, value=float(emp_data["MedicalAllowance"]))
# Safe float conversion to avoid ValueError
                    def safe_float(value, default=0.0):
                        try:
                            return float(value)
                        except (ValueError, TypeError):
                           return default

                    internet_allowance = st.number_input(
                        "Internet Allowance",
                        min_value=0.0,
                        value=safe_float(emp_data.get("InternetAllowance", 0.0))
                    )
                    special_allowance = st.number_input("Special Allowance", min_value=0.0, value=float(emp_data["SpecialAllowance"]))
                    pf = st.number_input("PF Deduction", min_value=0.0, value=float(emp_data["PFDeduction"]))
                    esi = st.number_input("ESI Deduction", min_value=0.0, value=float(emp_data["ESIDeduction"]))
                    tds = st.number_input("TDS Deduction", min_value=0.0, value=float(emp_data["TDSDeduction"]))
                    other_deductions = st.number_input("Other Deductions", min_value=0.0, value=float(emp_data["OtherDeductions"]))
                    bank_name = st.text_input("Bank Name", value=emp_data["BankName"])
                    branch = st.text_input("Branch", value=emp_data["Branch"])
                    account_no = st.text_input("Account Number", value=emp_data["AccountNumber"])
                    ifsc_code = st.text_input("IFSC Code", value=emp_data["IFSCCode"])
                    pan = st.text_input("PAN Number", value=emp_data["PAN"])
                    aadhar = st.text_input("Aadhar Number", value=emp_data["Aadhar"])
                    edit_submit = st.form_submit_button("Update Employee")
                if edit_submit:
                        # Update only selected employee fields explicitly (safe, field-wise update)
                        df.loc[df["EmployeeID"] == selected_emp, [
                            "FullName", "DOB", "Gender", "Phone", "Email", "CurrentAddress",
                            "PermanentAddress", "EmergencyContactName", "EmergencyRelation", "EmergencyPhone",
                            "Designation", "JoiningDate", "EmploymentType", "ReportingManager", "WorkLocation",
                            "EmployeeStatus", "CTC", "BasicPay", "HRA", "TravelAllowance", "MedicalAllowance",
                            "InternetAllowance", "SpecialAllowance", "PFDeduction", "ESIDeduction",
                            "TDSDeduction", "OtherDeductions", "BankName", "Branch", "AccountNumber",
                            "IFSCCode", "PAN", "Aadhar"
                        ]] = [[
                            full_name, dob.strftime("%Y-%m-%d"), gender, phone, email, current_address,
                            permanent_address, emergency_name, emergency_relation, emergency_phone,
                            designation, joining_date.strftime("%Y-%m-%d"), emp_type, reporting_manager, work_location,
                            emp_status, ctc, basic_pay, hra, travel_allowance, medical_allowance,
                            internet_allowance, special_allowance, pf, esi, tds, other_deductions,
                            bank_name, branch, account_no, ifsc_code, pan, aadhar
                        ]]

                        save_employees(df)
                        st.success("Employee updated successfully!")
                        save_employees(df)
        
        with tab3:
            df = load_employees()
            if not df.empty:
                st.subheader("Employee Reports")
                
                # Employment Status Distribution
                st.write("#### Employee Status Distribution")
                status_count = df["EmployeeStatus"].value_counts()
                st.bar_chart(status_count)
                
                # Department/Designation Distribution
                st.write("#### Designation Distribution")
                designation_count = df["Designation"].value_counts()
                st.bar_chart(designation_count)
                
                # Export Options
                if st.button("Export Employee Data"):
                    df.to_excel("employee_data.xlsx", index=False)
                    st.success("Data exported to employee_data.xlsx")
            else:
                st.info("No data available for reports.")

    elif menu == "Generate Payslip":
        st.header("Payslip Generator")
        df = load_employees()
        if df.empty:
            st.info("No employees found. Add employees first.")
        else:
            emp_id = st.selectbox("Select Employee", df["EmployeeID"].unique())
            emp = df[df["EmployeeID"] == emp_id].iloc[0]
            period_from = st.date_input("Pay Period From")
            period_to = st.date_input("Pay Period To")
            department = st.text_input("Department", "")

            # Earnings
            basic = float(emp["BasicPay"])
            incentives = st.number_input("Incentives", min_value=0.0, value=0.0)
            hra = st.number_input("HRA", min_value=0.0, value=float(emp["HRA"]))
            internet_allowance = st.number_input("Internet Allowance", min_value=0.0, 
                value=float(emp.get("InternetAllowance", 0)))  # Add this line
            total_earnings = basic + incentives + hra + internet_allowance  # Update total

            # Deductions
            pf = st.number_input("PF Deduction", min_value=0.0, value=float(emp["PFDeduction"]))
            professional_tax = st.number_input("Professional Tax", min_value=0.0, value=0.0)
            tds = st.number_input("TDS Deduction", min_value=0.0, value=float(emp["TDSDeduction"]))
            total_deductions = pf + professional_tax + tds

            net_salary = total_earnings - total_deductions

            if st.button("Generate Payslip PDF"):
                pdf = FPDF()
                pdf.add_page()
                
                # Logo and Header
                logo_width = 31
                try:
                    pdf.image("logo.png", x=12, y=10, w=logo_width)
                except:
                    pass

                # Company Address (right aligned)
                pdf.set_xy(12 + logo_width + 10, 12)
                pdf.set_font("Arial", "", 10)
                pdf.set_text_color(50, 50, 50)
                pdf.multi_cell(0, 5,
                    "LIG 11, HUDA Colony\n"
                    "Mehidipatnam, Hyderabad - 500028\n"
                    "Phone: +91 9908272012\n"
                    "Email: hr@onetechmatedsolutions.com",
                    align="R"
                )

                # Separator line
                pdf.set_draw_color(200, 200, 200)
                pdf.set_line_width(0.3)
                pdf.line(10, 40, 200, 40)

                # SALARY SLIP heading
                pdf.ln(15)
                pdf.set_font("Arial", "B", 20)
                pdf.set_text_color(54, 95, 145)
                pdf.cell(0, 12, "SALARY SLIP", ln=1, align="C")

                # Employee Details (single print, no duplicates)
                pdf.set_font("Arial", "", 11)
                pdf.set_text_color(0, 0, 0)
                
                # First row
                pdf.cell(95, 8, f"Employee ID: {emp_id}", border=0)
                pdf.cell(95, 8, f"Pay Period: {period_from.strftime('%d-%m-%Y')} to {period_to.strftime('%d-%m-%Y')}", border=0, align="R", ln=1)
                
                # Additional details
                pdf.cell(0, 8, f"Name: {emp['FullName']}", border=0, ln=1)
                pdf.cell(0, 8, f"Department: {department}", border=0, ln=1)
                pdf.cell(0, 8, f"Designation: {emp['Designation']}", border=0, ln=1)
                pdf.ln(4)

                # --- Earnings and Deductions Table ---
                pdf.set_font("Arial", "B", 12)
                pdf.cell(95, 8, "Earnings", border=1, align="C")
                pdf.cell(95, 8, "Deductions", border=1, align="C", ln=1)
                pdf.set_font("Arial", "", 11)
                pdf.cell(10, 8, "S.No.", border=1, align="C")
                pdf.cell(60, 8, "Salary Head", border=1)
                pdf.cell(25, 8, "Amount", border=1, align="R")
                pdf.cell(10, 8, "S.No.", border=1, align="C")
                pdf.cell(60, 8, "Salary Head", border=1)
                pdf.cell(25, 8, "Amount", border=1, align="R", ln=1)

                # Earnings rows
                earnings = [
                    ("Basic Pay", basic),
                    ("Incentives", incentives),
                    ("HRA", hra),
                    ("Internet Allowance", internet_allowance)  # Add this line
                ]
                for idx, (head, amt) in enumerate(earnings, 1):
                    pdf.cell(10, 8, str(idx), border=1, align="C")
                    pdf.cell(60, 8, head, border=1)
                    pdf.cell(25, 8, f"{amt:.2f}", border=1, align="R")
                    
                    # Print deduction row if exists
                    if idx == 1:
                        pdf.cell(10, 8, "1", border=1, align="C")
                        pdf.cell(60, 8, "PF", border=1)
                        pdf.cell(25, 8, f"{pf:.2f}", border=1, align="R", ln=1)
                    elif idx == 2:
                        pdf.cell(10, 8, "2", border=1, align="C")
                        pdf.cell(60, 8, "Professional Tax", border=1)
                        pdf.cell(25, 8, f"{professional_tax:.2f}", border=1, align="R", ln=1)
                    elif idx == 3:
                        pdf.cell(10, 8, "3", border=1, align="C")
                        pdf.cell(60, 8, "TDS", border=1)
                        pdf.cell(25, 8, f"{tds:.2f}", border=1, align="R", ln=1)
                    elif idx == 4:  # For Internet Allowance row
                        pdf.cell(10, 8, "", border=1)
                        pdf.cell(60, 8, "", border=1)
                        pdf.cell(25, 8, "", border=1, align="R", ln=1)

                # If deductions < earnings, fill empty deduction rows
                for idx in range(len(earnings)+1, 4):
                    pdf.cell(10, 8, "", border=1)
                    pdf.cell(60, 8, "", border=1)
                    pdf.cell(25, 8, "", border=1)
                    pdf.cell(10, 8, str(idx), border=1, align="C")
                    if idx == 1:
                        pdf.cell(60, 8, "PF", border=1)
                        pdf.cell(25, 8, f"{pf:.2f}", border=1, align="R", ln=1)
                    elif idx == 2:
                        pdf.cell(60, 8, "Professional Tax", border=1)
                        pdf.cell(25, 8, f"{professional_tax:.2f}", border=1, align="R", ln=1)
                    elif idx == 3:
                        pdf.cell(60, 8, "TDS", border=1)
                        pdf.cell(25, 8, f"{tds:.2f}", border=1, align="R", ln=1)

                # Totals
                pdf.set_font("Arial", "B", 11)
                pdf.cell(70, 8, "Gross Salary", border=1)
                pdf.cell(25, 8, f"{total_earnings:.2f}", border=1, align="R")
                pdf.cell(70, 8, "Total Deduction", border=1)
                pdf.cell(25, 8, f"{total_deductions:.2f}", border=1, align="R", ln=1)
                pdf.set_font("Arial", "B", 12)
                pdf.cell(70, 10, "Net Salary", border=1)
                pdf.cell(25, 10, f"{net_salary:.2f}", border=1, align="R")
                pdf.cell(70, 10, "", border=1)
                pdf.cell(25, 10, "", border=1, ln=1)
                pdf.ln(5)
                pdf.set_font("Helvetica", "", 10)
                pdf.cell(0, 8, "If you have any questions about this payslip, please contact HR.", ln=1)
                payslip_path = os.path.join(PDF_FOLDER, f"Payslip_{emp_id}_{period_from.strftime('%Y%m%d')}_{period_to.strftime('%Y%m%d')}.pdf")
                pdf.output(payslip_path)
                st.success(f"Payslip generated: {payslip_path}")
                
                # Show download button
                with open(payslip_path, "rb") as f:
                    st.download_button("Download Payslip PDF", f, file_name=os.path.basename(payslip_path))
                
                # Store the path in session state for email section
                st.session_state['current_payslip_path'] = payslip_path
                st.session_state['current_emp_data'] = emp

            # Email section (outside the PDF generation button)
            if 'current_payslip_path' in st.session_state:
                st.markdown("---")
                st.subheader("📧 Send Payslip via Email")
                
                emp = st.session_state['current_emp_data']
                payslip_path = st.session_state['current_payslip_path']
                
                col1, col2 = st.columns(2)
                with col1:
                    receiver_email = st.text_input("Recipient Email", value=emp.get("Email", ""))
                    subject = st.text_input("Email Subject", 
                        value=f"Payslip - {emp['FullName']} - {period_from.strftime('%B %Y')}")
                
                body = st.text_area("Email Body", 
                    value=f"""Dear {emp['FullName'],}

Please find attached your payslip for the period {period_from.strftime('%d-%m-%Y')} to {period_to.strftime('%d-%m-%Y')}.

Regards,
HR Department
Onetechmted Solutions LLP""")

                if st.button("📤 Send Email"):
                    try:
                        with st.spinner("Sending email..."):
                            if os.path.exists(payslip_path):
                                sent = send_email_outlook(receiver_email, subject, body, payslip_path)
                                if sent:
                                    st.success(f"✅ Payslip emailed successfully to {receiver_email}")
                                    # Clear session state after successful send
                                    del st.session_state['current_payslip_path']
                                    del st.session_state['current_emp_data']
                                else:
                                    st.error("❌ Failed to send email. Please check if Outlook is configured properly.")
                            else:
                                st.error("❌ Payslip file not found. Please regenerate the payslip.")
                    except Exception as e:
                        st.error(f"❌ Error sending email: {str(e)}")

    elif menu == "Income & Expense Tracker":
        st.header("💰 Income & Expense Tracker")
        
        # Define tabs within the Income & Expense section
        income_tabs = st.tabs(["Add Transaction", "View Transactions", "Reports"])
        
        # Tab 1: Add Transaction
        with income_tabs[0]:
            with st.form("transaction_form"):
                st.subheader("Add New Transaction")
                
                col1, col2 = st.columns(2)
                with col1:
                    trans_type = st.selectbox("Transaction Type", ["Income", "Expense", "Asset"])
                    category = st.selectbox("Category", 
                        ["Invoice Payment", "Salary Payment", "Rent", "Utilities", 
                         "Office Supplies", "Travel", "Marketing", "Asset Purchase", "Other"])
                    
                    # Show text input for "Other" category
                    if category == "Other":
                        other_category = st.text_input("Specify Other Category")
                        if other_category:
                            category = f"Other: {other_category}"
                    
                    amount = st.number_input("Amount", min_value=0.0, step=100.0)
                    
                with col2:
                    date = st.date_input("Date")
                    payment_mode = st.selectbox("Payment Mode", 
                        ["Cash", "Bank Transfer", "UPI", "Credit Card", "Cheque"])
                    status = st.selectbox("Status", ["Completed", "Pending", "Failed"])
                
                # Link to Invoice/Payslip if applicable
                link_type = st.selectbox("Link to", ["None", "Invoice", "Payslip"])
                if link_type == "Invoice":
                    invoices = load_invoices()
                    if not invoices.empty:
                        linked_to = st.selectbox("Select Invoice", 
                            invoices["Invoice No"].unique())
                    else:
                        st.info("No invoices available")
                        linked_to = "None"
                elif link_type == "Payslip":
                    payslip_files = [f.replace("Payslip_", "").replace(".pdf", "") 
                                   for f in os.listdir(PDF_FOLDER) 
                                   if f.startswith("Payslip_")]
                    if payslip_files:
                        linked_to = st.selectbox("Select Payslip", payslip_files)
                    else:
                        st.info("No payslips available")
                        linked_to = "None"
                else:
                    linked_to = "None"
                
                reference = st.text_input("Reference Number (Optional)")
                description = st.text_area("Description")
                notes = st.text_area("Additional Notes")
                
                submit = st.form_submit_button("Add Transaction")
                
                if submit:
                    df = load_transactions()
                    new_transaction = {
                        "TransactionID": generate_transaction_id(),
                        "Date": date.strftime("%Y-%m-%d"),
                        "Type": trans_type,
                        "Category": category,
                        "Amount": amount,
                        "Description": description,
                        "PaymentMode": payment_mode,
                        "Status": status,
                        "LinkedTo": linked_to,
                        "Reference": reference,
                        "Notes": notes,
                        "CreatedAt": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    }
                    df = pd.concat([df, pd.DataFrame([new_transaction])], ignore_index=True)
                    save_transaction(df)
                    st.success("Transaction added successfully!")

        # Tab 2: View Transactions
        with income_tabs[1]:
            st.subheader("View Transactions")
            df = load_transactions()
            
            if not df.empty:
                # Filters
                col1, col2, col3 = st.columns(3)
                with col1:
                    type_filter = st.multiselect("Transaction Type", 
                        df["Type"].unique().tolist())
                with col2:
                    category_filter = st.multiselect("Category", 
                        df["Category"].unique().tolist())
                with col3:
                    status_filter = st.multiselect("Status", 
                        df["Status"].unique().tolist())
                
                # Apply filters
                filtered_df = df.copy()
                if type_filter:
                    filtered_df = filtered_df[filtered_df["Type"].isin(type_filter)]
                if category_filter:
                    filtered_df = filtered_df[filtered_df["Category"].isin(category_filter)]
                if status_filter:
                    filtered_df = filtered_df[filtered_df["Status"].isin(status_filter)]
                
                # Display transactions
                st.dataframe(filtered_df)
                

                # --- Transaction Edit/Delete Section ---
                if not filtered_df.empty:
                    st.markdown("### ✏️ Edit or 🗑️ Delete Transactions")

                    trans_ids = filtered_df["TransactionID"].tolist()
                    selected_trans = st.selectbox("Select Transaction to Edit/Delete", trans_ids)

                    if selected_trans:
                        trans_row = filtered_df[filtered_df["TransactionID"] == selected_trans].iloc[0]

                        with st.expander("Edit Selected Transaction"):
                            new_type = st.selectbox("Transaction Type", ["Income", "Expense", "Asset"], index=["Income", "Expense", "Asset"].index(trans_row["Type"]))
                            new_category = st.text_input("Category", value=trans_row["Category"])
                            new_amount = st.number_input("Amount", min_value=0.0, value=float(trans_row["Amount"]), step=100.0)
                            new_date = st.date_input("Date", value=pd.to_datetime(trans_row["Date"]))
                            new_payment = st.selectbox("Payment Mode", ["Cash", "Bank Transfer", "UPI", "Credit Card", "Cheque"], index=["Cash", "Bank Transfer", "UPI", "Credit Card", "Cheque"].index(trans_row["PaymentMode"]))
                            new_status = st.selectbox("Status", ["Completed", "Pending", "Failed"], index=["Completed", "Pending", "Failed"].index(trans_row["Status"]))
                            new_description = st.text_area("Description", value=trans_row["Description"])
                            new_notes = st.text_area("Notes", value=trans_row["Notes"])

                            if st.button("💾 Save Changes"):
                                df.loc[df["TransactionID"] == selected_trans, ["Type", "Category", "Amount", "Date", "PaymentMode", "Status", "Description", "Notes"]] = [
                                    new_type, new_category, new_amount, new_date.strftime("%Y-%m-%d"), new_payment, new_status, new_description, new_notes
                                ]
                                save_transaction(df)
                                st.success(f"✅ Transaction {selected_trans} updated successfully!")

                        if st.button("🗑️ Delete Selected Transaction"):
                            df = df[df["TransactionID"] != selected_trans]
                            save_transaction(df)
                            st.warning(f"🗑️ Transaction {selected_trans} deleted successfully!")
                # Export option
                if st.button("Export to Excel"):
                    filtered_df.to_excel("transactions.xlsx", index=False)
                    st.success("Data exported to transactions.xlsx")
            else:
                st.info("No transactions recorded yet.")
        
        # Tab 3: Reports
        with income_tabs[2]:
            st.subheader("Financial Reports")
            df = load_transactions()
            
            if not df.empty:
                # Summary Statistics
                st.write("### Summary")
                total_income = df[df["Type"] == "Income"]["Amount"].sum()
                total_expense = df[df["Type"] == "Expense"]["Amount"].sum()
                net_amount = total_income - total_expense
                
                col1, col2, col3 = st.columns(3)
                with col1:
                    display_metric("Total Income", f"₹ {total_income:,.2f}", "success")
                with col2:
                    display_metric("Total Expenses", f"₹ {total_expense:,.2f}", "warning")
                with col3:
                    display_metric("Net Amount", f"₹ {net_amount:,.2f}", 
                                 "success" if net_amount >= 0 else "danger")
                
                # Monthly Trends
                st.write("### Monthly Trends")
                df["Month"] = pd.to_datetime(df["Date"]).dt.strftime("%Y-%m")
                monthly_data = df.pivot_table(
                    index="Month",
                    columns="Type",
                    values="Amount",
                    aggfunc="sum"
                ).fillna(0)
                st.line_chart(monthly_data)
                
                # Category-wise Analysis
                st.write("### Category-wise Analysis")
                cat_data = df.pivot_table(
                    index="Category",
                    columns="Type",
                    values="Amount",
                    aggfunc="sum"
                ).fillna(0)
                st.bar_chart(cat_data)
                
                # Payment Mode Analysis
                st.write("### Payment Mode Distribution")
                payment_data = df.groupby("PaymentMode")["Amount"].sum()
                
                # Create pie chart using plotly
                fig = px.pie(
                    values=payment_data.values,
                    names=payment_data.index,
                    title="Payment Mode Distribution",
                    hole=0.3
                )
                fig.update_layout(
                    showlegend=True,
                    legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1)
                )
                st.plotly_chart(fig)
                
                # Detailed Reports
                st.write("### Detailed Reports")
                report_type = st.selectbox("Select Report", 
                    ["Monthly Summary", "Category Summary", "Payment Mode Summary"])
                
                if report_type == "Monthly Summary":
                    monthly_summary = df.pivot_table(
                        index="Month",
                        columns="Type",
                        values="Amount",
                        aggfunc="sum"
                    ).reset_index()
                    st.dataframe(monthly_summary)
                    
                elif report_type == "Category Summary":
                    category_summary = df.pivot_table(
                        index="Category",
                        columns="Type",
                        values="Amount",
                        aggfunc="sum"
                    ).reset_index()
                    st.dataframe(category_summary)
                    
                else:
                    payment_summary = df.pivot_table(
                        index="PaymentMode",
                        columns="Type",
                        values="Amount",
                        aggfunc="sum"
                    ).reset_index()
                    st.dataframe(payment_summary)
                
                # Export Reports
                if st.button("Export Reports"):
                    with pd.ExcelWriter("financial_reports.xlsx") as writer:
                        monthly_summary.to_excel(writer, sheet_name="Monthly", index=False)
                        category_summary.to_excel(writer, sheet_name="Category", index=False)
                        payment_summary.to_excel(writer, sheet_name="Payment", index=False)
                    st.success("Reports exported to financial_reports.xlsx")
            else:
                st.info("No transaction data available for reports.")

def invoice_generator():
    st.header("📄 Invoice Generator")

    # ======= FORM SECTION (Data entry only) =======
    with st.form("invoice_form"):
        st.subheader("Company Information")
        org_name = st.text_input("Company Name")
        org_address = st.text_input("Street Address")
        org_city = st.text_input("City")
        org_zip = st.text_input("ZIP Code")
        org_phone = st.text_input("Phone")
        org_website = st.text_input("Website")
        org_fax = st.text_input("Fax")

        st.subheader("Invoice Details")
        invoice_no = auto_invoice_number()
        st.write(f"Invoice Number: {invoice_no}")
        invoice_date = st.date_input("Invoice Date", date.today())
        due_date = st.date_input("Due Date", date.today())
        customer_id = st.text_input("Customer ID")

        st.subheader("Client Information")
        client_name = st.text_input("Client Name")
        client_company = st.text_input("Client Company Name")
        client_address = st.text_input("Client Address")
        client_city = st.text_input("Client City")
        client_zip = st.text_input("Client ZIP")
        client_phone = st.text_input("Client Phone")

        # === Services Section ===
        st.subheader("Services")
        service_desc = st.text_input("Service Description")
        service_unit = st.text_input("Service Unit (e.g., Hour, Day)")
        service_rate = st.number_input("Service Rate", min_value=0.0, step=0.01)
        service_hours = st.number_input("Service Hours/Qty", min_value=0.0, step=0.01)
        service_currency = st.selectbox("Currency Symbol", ["$", "₹"])
        add_service = st.form_submit_button("Add Service")
        if add_service and service_desc:
            if 'services' not in st.session_state:
                st.session_state.services = []
            st.session_state.services.append({
                "Description": service_desc,
                "Unit": service_unit,
                "Rate": service_rate,
                "HoursQty": service_hours,
                "Amount": service_rate * service_hours,
                "Currency": service_currency
            })
            st.success("Service added!")

        # === Expenses Section ===
        st.subheader("Expenses")
        expense_desc = st.text_input("Expense Description")
        expense_unit = st.text_input("Expense Unit (e.g., Unit, Month)")
        expense_rate = st.number_input("Expense Rate", min_value=0.0, step=0.01)
        expense_qty = st.number_input("Expense Qty", min_value=0.0, step=0.01)
        expense_currency = st.selectbox("Expense Currency", ["$", "₹"], key="expense_currency")
        add_expense = st.form_submit_button("Add Expense")
        if add_expense and expense_desc:
            if 'expenses' not in st.session_state:
                st.session_state.expenses = []
            st.session_state.expenses.append({
                "Description": expense_desc,
                "Unit": expense_unit,
                "Rate": expense_rate,
                "Qty": expense_qty,
                "Amount": expense_rate * expense_qty,
                "Currency": expense_currency
            })
            st.success("Expense added!")

        # === Tax Section ===
        st.subheader("💰 Tax Details")
        tax_rate = st.number_input("Tax Rate (%)", min_value=0.0, max_value=100.0, step=0.1, value=0.0)

        # === Additional Information ===
        st.subheader("Additional Information")
        comments = st.text_area("Comments", value="Thank you for your business!")

        # === Generate Invoice Button ===
        generate = st.form_submit_button("Generate Invoice")

        if generate:
            services_total = sum(s["Amount"] for s in st.session_state.get("services", []))
            expenses_total = sum(e["Amount"] for e in st.session_state.get("expenses", []))
            subtotal = services_total + expenses_total
            tax_due = subtotal * (tax_rate / 100)
            total = subtotal + tax_due

            invoice = {
                "Invoice No": invoice_no,
                "Invoice Date": invoice_date.strftime("%Y-%m-%d"),
                "Due Date": due_date.strftime("%Y-%m-%d"),
                "Customer ID": customer_id,
                "Organization Name": org_name,
                "Organization Address": org_address,
                "Organization City": org_city,
                "Organization Zip": org_zip,
                "Organization Phone": org_phone,
                "Organization Website": org_website,
                "Organization Fax": org_fax,
                "Client Name": client_name,
                "Client Company Name": client_company,
                "Client Address": client_address,
                "Client City": client_city,
                "Client Zip": client_zip,
                "Client Phone": client_phone,
                "Services": st.session_state.get("services", []),
                "Expenses": st.session_state.get("expenses", []),
                "Subtotal": subtotal,
                "Tax Rate": tax_rate,
                "Tax Due": tax_due,
                "Total": total,
                "Comments": comments,
                "Currency Symbol": service_currency
            }

            save_invoice(invoice)
            pdf_path = generate_pdf(invoice)
            st.session_state['last_invoice'] = invoice
            st.session_state['last_pdf_path'] = pdf_path
            st.success(f"Invoice #{invoice_no} generated successfully!")

    # ======= OUTSIDE THE FORM =======
    # ✅ You can safely use normal st.button() here now

    if 'services' in st.session_state and st.session_state.services:
        st.subheader("🛠️ Added Services")
        services_df = pd.DataFrame(st.session_state.services)
        st.dataframe(services_df)

        edit_index = st.number_input(
            "Enter service index to edit/delete (starting from 0)",
            min_value=0,
            max_value=len(st.session_state.services) - 1,
            step=1
        )

        col_edit, col_delete = st.columns(2)
        with col_edit:
            if st.button("✏️ Edit Selected Service"):
                service_to_edit = st.session_state.services[int(edit_index)]
                st.info(f"Editing service #{edit_index}: {service_to_edit['Description']}")
                # implement editing logic...

        with col_delete:
            if st.button("🗑️ Delete Selected Service"):
                removed = st.session_state.services.pop(int(edit_index))
                st.warning(f"Removed service: {removed['Description']}")
                
                    # ======= OUTSIDE THE FORM: DOWNLOAD & EMAIL OPTIONS =======
    if 'last_pdf_path' in st.session_state:
        st.subheader("📥 Download or Email Invoice")

        # --- Download Button ---
        with open(st.session_state['last_pdf_path'], "rb") as f:
            st.download_button(
                "⬇️ Download Invoice PDF",
                f,
                file_name=f"Invoice_{st.session_state['last_invoice']['Invoice No']}.pdf",
                mime="application/pdf"
            )

        # --- Email Section ---
        st.subheader("📧 Send Invoice via Email")
        email = st.text_input("Recipient Email")
        subject = st.text_input(
            "Email Subject",
            f"Invoice #{st.session_state['last_invoice']['Invoice No']}"
        )
        body = st.text_area("Email Body", "Please find attached your invoice.")

        if st.button("Send Email"):
            if send_email_outlook(email, subject, body, st.session_state['last_pdf_path']):
                st.success("✅ Email sent successfully!")

                # Clear session data after sending
                st.session_state.pop('services', None)
                st.session_state.pop('expenses', None)
                st.session_state.pop('last_pdf_path', None)
                st.session_state.pop('last_invoice', None)


    
def view_invoices():
    st.header("📋 View Invoices")
    invoices = load_invoices()
    
    if invoices.empty:
        st.info("No invoices found.")
        return
        
    # Filter options
    col1, col2 = st.columns(2)
    with col1:
        client_filter = st.text_input("Filter by Client Name")
    with col2:
        date_filter = st.date_input("Filter by Date", value=None)

    # Apply filters
    filtered_df = invoices.copy()
    if client_filter:
        filtered_df = filtered_df[filtered_df["Client Name"].str.contains(client_filter, case=False, na=False)]
    if date_filter:
        filtered_df = filtered_df[filtered_df["Invoice Date"] == date_filter.strftime("%Y-%m-%d")]

    # Display invoices
    st.dataframe(filtered_df)

    # Select invoice to view details
    selected_invoice = st.selectbox(
        "Select Invoice to View Details",
        filtered_df["Invoice No"].unique()
    )

    if selected_invoice:
        invoice = filtered_df[filtered_df["Invoice No"] == selected_invoice].iloc[0]
        
        col1, col2, col3 = st.columns(3)
        with col1:
            st.info(f"**Invoice #{invoice['Invoice No']}**\n"
                   f"Date: {invoice['Invoice Date']}\n"
                   f"Due: {invoice['Due Date']}")
        with col2:
            st.info(f"**Client**\n"
                   f"{invoice['Client Name']}\n"
                   f"{invoice['Client Company Name']}")
        with col3:
            st.info(f"**Total**\n"
                   f"{invoice['Currency Symbol']} {invoice['Total']:,.2f}")

        # Show services and expenses
        with st.expander("View Services & Expenses"):
            try:
                services = json.loads(invoice["Services"])
                if services:
                    st.write("### Services")
                    st.dataframe(pd.DataFrame(services))
            except:
                st.warning("No services data available")

            try:
                expenses = json.loads(invoice["Expenses"])
                if expenses:
                    st.write("### Expenses")
                    st.dataframe(pd.DataFrame(expenses))
            except:
                st.warning("No expenses data available")

        # Add Comments Editor
        with st.expander("View/Edit Comments"):
            current_comments = invoice['Comments']
            new_comments = st.text_area("Invoice Comments", 
                value=current_comments,
                key="comment_editor")
            
            if st.button("Update Comments"):
                # Update comments in the dataframe
                invoices.loc[invoices["Invoice No"] == selected_invoice, "Comments"] = new_comments
                # Save updated dataframe
                invoices.to_csv(INVOICE_FILE, index=False)
                st.success("Comments updated successfully!")
                
                # Regenerate PDF with new comments
                updated_invoice = invoices[invoices["Invoice No"] == selected_invoice].iloc[0].to_dict()
                pdf_path = generate_pdf(updated_invoice)
                with open(pdf_path, "rb") as f:
                    st.download_button(
                        "Download Updated PDF",
                        f,
                        file_name=f"Invoice_{invoice['Invoice No']}_updated.pdf"
                    )

        # Regenerate PDF
        if st.button("Regenerate PDF"):
            pdf_path = generate_pdf(invoice.to_dict())
            with open(pdf_path, "rb") as f:
                st.download_button(
                    "Download PDF",
                    f,
                    file_name=f"Invoice_{invoice['Invoice No']}.pdf"
                )
if __name__ == "__main__":
    main()