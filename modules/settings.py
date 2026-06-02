import streamlit as st
import json
import os
from modules.audit import log_action

SETTINGS_FILE = "data/settings.json"


def load_settings():

    if os.path.exists(SETTINGS_FILE):

        with open(SETTINGS_FILE, "r") as f:
            return json.load(f)

    return {
        "company_name": "",
        "gst": "",
        "phone": "",
        "email": "",
        "website": "",
        "address": ""
    }


def save_settings(settings):

    with open(SETTINGS_FILE, "w") as f:
        json.dump(settings, f, indent=4)
        
        
def render():

    st.title("⚙ Company Settings")

    settings = load_settings()

    company_name = st.text_input(
        "Company Name",
        settings["company_name"]
    )

    gst = st.text_input(
        "GST Number",
        settings["gst"]
    )

    phone = st.text_input(
        "Phone",
        settings["phone"]
    )

    email = st.text_input(
        "Email",
        settings["email"]
    )

    website = st.text_input(
        "Website",
        settings["website"]
    )

    address = st.text_area(
        "Address",
        settings["address"]
    )

    if st.button("Save Settings"):

        save_settings({
            "company_name": company_name,
            "gst": gst,
            "phone": phone,
            "email": email,
            "website": website,
            "address": address
        })
        log_action(
            st.session_state.username,
            "Updated Settings",
            company_name
        )

        st.success("Settings Saved Successfully")