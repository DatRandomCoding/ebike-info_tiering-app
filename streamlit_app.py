import os
import streamlit as st
import pandas as pd

st.title("🚲 E-Bike Tier List & Calculator")
st.write("Welcome to the community e-bike library!")

# Default filename to look for in the workspace
DEFAULT_FILENAME = "ebike tier list blank MK1.xlsx"

# Allow the user to upload an Excel file as a fallback
uploaded_file = st.file_uploader("Upload an Excel file (optional)", type=["xlsx", "xls"])

@st.cache_data
def load_excel(file):
    # pandas accepts a file-like object or a path
    return pd.read_excel(file)

master_table = None
if uploaded_file is not None:
    try:
        master_table = load_excel(uploaded_file)
    except Exception as e:
        st.error(f"Error reading uploaded Excel file: {e}")
elif os.path.exists(DEFAULT_FILENAME):
    try:
        master_table = load_excel(DEFAULT_FILENAME)
    except FileNotFoundError:
        st.error(f"File not found: {DEFAULT_FILENAME}")
    except ValueError as e:
        st.error(f"Could not parse the Excel file: {e}")
    except Exception as e:
        st.error(f"Unexpected error loading {DEFAULT_FILENAME}: {e}")
else:
    st.info("No Excel file found in workspace. Upload one using the uploader above.")

if master_table is not None:
    st.subheader("🏆 Top 10 E-Bikes Leaderboard")
    try:
        top_10 = master_table.sort_values(by="Rating/5:", ascending=False).head(10)
        st.dataframe(top_10)
    except KeyError:
        st.error("Expected column 'Rating/5:' not found. Columns: " + ", ".join(master_table.columns.astype(str)))
    except Exception as e:
        st.error(f"Could not process data: {e}")