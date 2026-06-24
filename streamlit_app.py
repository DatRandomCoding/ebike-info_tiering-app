import streamlit as st

st.title("🚲 E-Bike Tier List & Calculator")
st.write("Welcome to the community e-bike library!")

# Quick test input boxes
name = st.text_input("Bike Name")
price = st.number_input("Price ($)", min_value=0)

if st.button("Calculate"):
    st.success(f"Added {name} to the list at ${price}!")