import streamlit as st

st.title("Test Python Script on Streamlit Cloud")

# Simple input
name = st.text_input("Enter your name:")

# Button to run code
if st.button("Say Hello"):
    st.write(f"Hello, {name}! 🎉")
