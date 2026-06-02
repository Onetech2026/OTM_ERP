import streamlit as st
import pandas as pd
import hashlib
import os

USER_FILE = "data/users.csv"

def hash_password(password):
    return hashlib.sha256(password.encode()).hexdigest()

def initialize_users():

    if os.path.exists(USER_FILE):
        return

    users = pd.DataFrame([
        {
            "username":"pruthvi",
            "password":hash_password("Admin@123"),
            "role":"Super Admin"
        },
        {
            "username":"admin",
            "password":hash_password("Admin@123"),
            "role":"Admin"
        }
    ])

    users.to_csv(USER_FILE,index=False)

def login():

    initialize_users()

    st.markdown("# Welcome Back")

    username = st.text_input("Username")
    password = st.text_input("Password", type="password")

    if st.button("Login"):

        users = pd.read_csv(USER_FILE)

        match = users[
            (users["username"] == username)
            &
            (users["password"] == hash_password(password))
        ]

        if len(match) > 0:

            st.session_state.logged_in = True
            st.session_state.username = username
            st.session_state.role = match.iloc[0]["role"]

            st.rerun()

        else:
            st.error("Invalid username or password")