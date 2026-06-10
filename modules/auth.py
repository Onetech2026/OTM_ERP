import os
import time

import hashlib
import pandas as pd
import streamlit as st

USER_FILE = "data/users.csv"


def hash_password(password):
    return hashlib.sha256(password.encode()).hexdigest()


def load_users():
    if not os.path.exists(USER_FILE):
        return pd.DataFrame(columns=["username", "password", "role", "status"])

    users = pd.read_csv(USER_FILE)
    if "status" not in users.columns:
        users["status"] = "Active"
        users.to_csv(USER_FILE, index=False)

    users["status"] = users["status"].fillna("Active").replace({"": "Active"})
    return users


def create_user(username, password, role="User"):
    users = load_users()

    normalized_username = username.strip().lower()
    existing = users["username"].astype(str).str.lower().tolist()

    if not normalized_username:
        raise ValueError("Username is required.")

    if not password:
        raise ValueError("Password is required.")

    if normalized_username in existing:
        raise ValueError("That username already exists.")

    os.makedirs("data", exist_ok=True)

    if "status" not in users.columns:
        users["status"] = "Active"

    users = pd.concat(
        [
            users,
            pd.DataFrame([
                {
                    "username": normalized_username,
                    "password": hash_password(password),
                    "role": role,
                    "status": "Active",
                }
            ]),
        ],
        ignore_index=True,
    )

    users.to_csv(USER_FILE, index=False)
    return {"username": normalized_username, "role": role, "status": "Active"}


def update_user(username, password=None, role=None, status=None):
    users = load_users()
    target = users[users["username"].astype(str).str.lower() == username.strip().lower()]

    if target.empty:
        raise ValueError("User not found.")

    if password:
        users.loc[target.index, "password"] = hash_password(password)

    if role:
        users.loc[target.index, "role"] = role

    if status:
        users.loc[target.index, "status"] = status

    users.to_csv(USER_FILE, index=False)
    return True


def delete_user(username):
    users = load_users()
    users = users[users["username"].astype(str).str.lower() != username.strip().lower()]
    users.to_csv(USER_FILE, index=False)
    return True


def set_user_status(username, status):
    return update_user(username, status=status)


def get_env_credentials():
    env_users = {}

    env_map = {
        "pruthvi": ("ERP_PRUTHVI_PASSWORD", "Super Admin"),
        "admin": ("ERP_ADMIN_PASSWORD", "Admin"),
        "shiva": ("ERP_SHIVA_PASSWORD", "User"),
    }

    for username, (password_env, default_role) in env_map.items():
        password = os.getenv(password_env)
        if password:
            env_users[username] = {
                "password": password,
                "role": os.getenv(f"ERP_{username.upper()}_ROLE", default_role),
            }

    return env_users


def initialize_users():

    if os.path.exists(USER_FILE):
        return

    users = pd.DataFrame([
        {
            "username": "admin",
            "password": hash_password("ChangeMe@123"),
            "role": "Admin"
        }
    ])

    users.to_csv(USER_FILE,index=False)

def login():

    initialize_users()

    left, center, right = st.columns(
        [1, 1.2, 1]
    )

    with center:

        st.markdown(
            """
            <div class="login-card">
            """,
            unsafe_allow_html=True
        )

        try:

            st.image(
                "assets/logo.png",
                width=120
            )

        except:
            pass

        st.markdown(
            """
            <h1 style="text-align:center;margin-bottom:0">
            Onetechmated ERP
            </h1>

            <p style="text-align:center;color:#64748b">
            Business Operations Platform
            </p>
            """,
            unsafe_allow_html=True
        )

        username = st.text_input(
            "Username"
        )

        password = st.text_input(
            "Password",
            type="password"
        )

        login_clicked = st.button(
            "Sign In",
            use_container_width=True
        )

        st.markdown(
            "</div>",
            unsafe_allow_html=True
        )

        if login_clicked:

            env_credentials = get_env_credentials()

            if username in env_credentials and password == env_credentials[username]["password"]:
                st.session_state.logged_in = True
                st.session_state.username = username
                st.session_state.role = env_credentials[username]["role"]
                st.session_state.login_started_at = time.time()
                st.rerun()
                return

            if os.path.exists(USER_FILE):
                users = pd.read_csv(USER_FILE)

                users = users.copy()
                if "status" not in users.columns:
                    users["status"] = "Active"
                users["status"] = users["status"].fillna("Active").replace({"": "Active"})

                match = users[
                    (users["username"] == username)
                    & (users["password"] == hash_password(password))
                    & (users["status"].astype(str).str.lower() != "inactive")
                ]

                if len(match) > 0:
                    st.session_state.logged_in = True
                    st.session_state.username = username
                    st.session_state.role = match.iloc[0]["role"]
                    st.session_state.login_started_at = time.time()
                    st.rerun()
                    return

            st.error("Invalid username or password")