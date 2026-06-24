import os
import streamlit as st
import pandas as pd

st.title("🚲 E-Bike Tier List & Calculator")
st.write("Welcome to the community e-bike library!")

DEFAULT_FILENAME = "ebike tier list blank MK1.xlsx"

uploaded_file = st.file_uploader("Upload an Excel file", type=["xlsx", "xls"])

def detect_header_row(excel_source, marker="Rating/5:", max_rows=20):
    preview = pd.read_excel(excel_source, header=None, nrows=max_rows)
    for idx, row in preview.iterrows():
        if marker in row.values:
            return idx
    return 0

@st.cache_data
def load_excel(file):
    # Detect the real header row before loading the full dataset.
    if hasattr(file, "seek"):
        file.seek(0)
        header_row = detect_header_row(file)
        file.seek(0)
        return pd.read_excel(file, header=header_row)
    header_row = detect_header_row(file)
    return pd.read_excel(file, header=header_row)

master_table = None
if uploaded_file is not None:
    try:
        master_table = load_excel(uploaded_file)
    except Exception as e:
        st.error(f"Error reading the uploaded Excel file: {e}")
elif os.path.exists(DEFAULT_FILENAME):
    try:
        master_table = load_excel(DEFAULT_FILENAME)
    except Exception as e:
        st.error(f"Error reading {DEFAULT_FILENAME}: {e}")
else:
    st.info(f"No Excel file found. Upload '{DEFAULT_FILENAME}' or add it to the workspace.")

if master_table is not None:
    st.subheader("🏆 Top 10 E-Bikes Leaderboard")
    try:
        top_10 = master_table.sort_values(by="Rating/5:", ascending=False).head(10)
        st.dataframe(top_10)
    except KeyError:
        st.error("Expected column 'Rating/5:' not found. Available columns: " + ", ".join(master_table.columns.astype(str)))
    except Exception as e:
        st.error(f"Could not process data: {e}")