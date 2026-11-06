"""
Run:  pip install streamlit streamlit-aggrid pandas openpyxl
      python manage_users.py add alice        # once per user
      streamlit run app.py
Files:  data/<username>_work.xlsx and data/<username>_life.xlsx
"""
import time

import streamlit as st

import auth
from sheet_page import render

st.set_page_config(page_title="Tracker", layout="wide")


def work_page():
    render("work")


def life_page():
    render("life")


# ---------------------------------------------------------------- login gate
if "user" not in st.session_state:
    _, mid, _ = st.columns([1, 1, 1])
    with mid:
        st.title("Sign in")
        with st.form("login"):
            username = st.text_input("Username").strip()
            password = st.text_input("Password", type="password")
            submitted = st.form_submit_button("Log in", type="primary")
        if submitted:
            if auth.verify(username, password):
                st.session_state["user"] = username
                st.rerun()
            time.sleep(1)                       # slows down password guessing
            st.error("Wrong username or password.")
    st.stop()                                   # nothing below runs until logged in

# ---------------------------------------------------------------- pages (only reached after login)
pg = st.navigation([
    st.Page(work_page, title="Work", icon="💼", url_path="work", default=True),
    st.Page(life_page, title="Life", icon="🏡", url_path="life"),
])
with st.sidebar:
    st.write(f"Signed in as **{st.session_state['user']}**")
    if st.button("Log out"):
        st.session_state.clear()                # also drops the user's loaded data
        st.rerun()
pg.run()
