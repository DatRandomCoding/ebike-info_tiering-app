import streamlit as st
import pandas as pd

st.title("🚲 E-Bike Tier List & Calculator")
st.write("Welcome to the community e-bike library!")

# 1. Read your master list Excel file
# Python uses 'pd.read_excel' to open and read your spreadsheet data
try:
    master_table = pd.read_excel("your_file_name.xlsx")
    
    st.subheader("🏆 Top 10 E-Bikes Leaderboard")
    
    # 2. Sort the table by your rating column and grab the top 10 rows
    # (Replace 'Rating/5:' with the exact column name in your Excel file)
    top_10 = master_table.sort_values(by="Rating/5:", ascending=False).head(10)
    
    # 3. Display it beautifully on your website
    st.dataframe(top_10)

except Exception as e:
    st.error(f"Could not load the Excel file. Make sure the file name matches perfectly! Error: {e}")