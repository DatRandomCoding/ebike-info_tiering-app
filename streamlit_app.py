import os
import streamlit as st
import pandas as pd
from openpyxl import load_workbook

st.title("🚲 E-Bike Tier List & Calculator")
st.write("Welcome to the community e-bike library!")

DEFAULT_FILENAME = "ebike tier list blank MK1.xlsx"

uploaded_file = st.file_uploader("Upload an Excel file", type=["xlsx", "xls"])

HEADER_LABELS = {
    "Bike name": "Bike name",
    "Price": "Price",
    "Price ($)": "Price",
    "Score (out of 5.01)": "Score",
    "Score": "Score",
    "Unweighted Score": "Score",
}

SUMMARY_COLUMNS = ["Bike name", "Score", "Price"]
SKIP_SCORE_TEXTS = {
    "Speed (mph)",
    "Range (mi)",
    "Price ($)",
    "Budget Max Price: ($)",
}
INVALID_VALUE_PREFIXES = ("#VALUE!",)


def _normalize_header(value):
    if value is None:
        return None
    return HEADER_LABELS.get(str(value).strip(), None)


def find_summary_header(sheet):
    for row in sheet.iter_rows(min_row=1, max_row=20, min_col=1, max_col=30):
        label_map = {}
        for cell in row:
            normalized = _normalize_header(cell.value)
            if normalized:
                label_map[normalized] = cell.column
        if "Bike name" in label_map and ("Score" in label_map or "Price" in label_map):
            return row[0].row, label_map
    return None, {}


def _is_blank_value(value):
    if pd.isna(value):
        return True
    text = str(value).strip()
    return text == "" or any(text.startswith(prefix) for prefix in INVALID_VALUE_PREFIXES)


def _is_summary_row(value_map):
    bike_name = value_map.get("Bike name")
    if bike_name is None or pd.isna(bike_name) or str(bike_name).strip() == "":
        return False

    score_value = value_map.get("Score")
    if score_value is None or pd.isna(score_value):
        return False
    score_text = str(score_value).strip()
    if score_text in SKIP_SCORE_TEXTS or score_text.startswith("Note:"):
        return False
    if any(score_text.startswith(prefix) for prefix in INVALID_VALUE_PREFIXES):
        return False

    return True


def _clean_summary_df(df: pd.DataFrame) -> pd.DataFrame:
    if df is None or df.empty:
        return pd.DataFrame(columns=SUMMARY_COLUMNS)

    df = df.copy()
    for col in SUMMARY_COLUMNS:
        if col not in df.columns:
            df[col] = None

    df = df[SUMMARY_COLUMNS]

    def is_placeholder_row(row):
        bike_name = row["Bike name"]
        score = row["Score"]
        price = row["Price"]

        if _is_blank_value(bike_name) and _is_blank_value(score) and _is_blank_value(price):
            return True

        if bike_name is not None and str(bike_name).strip() == "":
            bike_name = None

        if score is not None:
            score_text = str(score).strip()
            if score_text in SKIP_SCORE_TEXTS or score_text.startswith("Note:") or any(score_text.startswith(prefix) for prefix in INVALID_VALUE_PREFIXES):
                return True

        return False

    cleaned = df[~df.apply(is_placeholder_row, axis=1)]
    return cleaned.reset_index(drop=True)


def extract_summary_table(file):
    try:
        if hasattr(file, "seek"):
            file.seek(0)
        workbook = load_workbook(file, data_only=True)
    except Exception:
        return None

    sheet = workbook[workbook.sheetnames[0]]
    header_row, label_map = find_summary_header(sheet)
    if not label_map:
        return None

    rows = []
    consecutive_empty = 0
    for row_idx in range(header_row + 1, sheet.max_row + 1):
        row_values = {
            label: sheet.cell(row=row_idx, column=label_map.get(label)).value if label_map.get(label) else None
            for label in SUMMARY_COLUMNS
        }

        if all(_is_blank_value(value) for value in row_values.values()):
            consecutive_empty += 1
            if consecutive_empty >= 3:
                break
            continue

        consecutive_empty = 0
        if _is_summary_row(row_values):
            rows.append(row_values)

    if not rows:
        return None

    return _clean_summary_df(pd.DataFrame(rows)[SUMMARY_COLUMNS])


RANKED_SCORE_COLUMNS = ["Score (out of 5.01)", "Unweighted Score", "Score"]

@st.cache_data
def load_data(file):
    if hasattr(file, "seek"):
        file.seek(0)
    summary_table = extract_summary_table(file)
    if summary_table is not None and not summary_table.empty:
        return summary_table

    if hasattr(file, "seek"):
        file.seek(0)
    header_row = detect_header_row(file)
    if hasattr(file, "seek"):
        file.seek(0)
    df = pd.read_excel(file, header=header_row)

    normalized_columns = {}
    for col in df.columns:
        normalized = _normalize_header(col)
        if normalized:
            normalized_columns.setdefault(normalized, []).append(col)

    if normalized_columns:
        result = pd.DataFrame(index=df.index)

        if "Bike name" in normalized_columns:
            result["Bike name"] = df[normalized_columns["Bike name"][0]]

        if "Price" in normalized_columns:
            result["Price"] = df[normalized_columns["Price"][0]]

        score_column = None
        for candidate in RANKED_SCORE_COLUMNS:
            if candidate in df.columns and candidate in normalized_columns.get("Score", []):
                score_column = candidate
                break
        if score_column is None and "Score" in normalized_columns:
            score_column = normalized_columns["Score"][0]
        if score_column is not None:
            result["Score"] = df[score_column]

        if not result.empty:
            return _clean_summary_df(result[SUMMARY_COLUMNS])

    return pd.DataFrame(columns=SUMMARY_COLUMNS)


def detect_header_row(excel_source, marker="Rating/5:", max_rows=20):
    preview = pd.read_excel(excel_source, header=None, nrows=max_rows)
    for idx, row in preview.iterrows():
        if marker in row.values:
            return idx
    return 0


def main():
    master_table = None
    if uploaded_file is not None:
        try:
            master_table = load_data(uploaded_file)
        except Exception as e:
            st.error(f"Error reading the uploaded Excel file: {e}")
    elif os.path.exists(DEFAULT_FILENAME):
        try:
            with open(DEFAULT_FILENAME, "rb") as default_file:
                master_table = load_data(default_file)
        except Exception as e:
            st.error(f"Error reading {DEFAULT_FILENAME}: {e}")
    else:
        st.info(f"No Excel file found. Upload '{DEFAULT_FILENAME}' or add it to the workspace.")

    if master_table is not None:
        st.subheader("🏆 Top 10 E-Bikes Leaderboard")
        if not master_table.empty:
            st.dataframe(master_table.head(10))
        else:
            st.warning("The workbook contains a summary table template, but no bike rows are populated yet.")


if __name__ == "__main__":
    main()
