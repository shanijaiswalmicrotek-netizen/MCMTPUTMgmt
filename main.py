import streamlit as st
import pandas as pd
import requests
import io
import unicodedata

# ---------------- CONFIG ----------------
GOOGLE_SHEET_FILE_ID = "1H_aN66Joy7Tuzx8NjTygOsMA1J2OUjnz"
COL_STUDENT_NAME = "Student Name"
COL_ADMISSION_NO = "Admission No."
COL_FATHER_NAME = "Father Name"   # optional

st.set_page_config(page_title="Marksheet Viewer", layout="wide")


# ---------------- GOOGLE SHEET LOADER ----------------
@st.cache_data(ttl=300)
def download_sheet_xlsx(file_id: str) -> bytes:
    url = f"https://docs.google.com/spreadsheets/d/{file_id}/export?format=xlsx"
    r = requests.get(url, timeout=30)
    r.raise_for_status()
    return r.content


def load_excelfile_from_sheet(file_id: str):
    content = download_sheet_xlsx(file_id)
    return pd.ExcelFile(io.BytesIO(content), engine="openpyxl")


def read_sheet_from_sheet(file_id: str, sheet_name: str):
    content = download_sheet_xlsx(file_id)
    return pd.read_excel(io.BytesIO(content), sheet_name=sheet_name, engine="openpyxl")


# ---------------- HELPERS ----------------
def to_text_one_decimal(val):
    try:
        if pd.isna(val):
            return ""
    except:
        pass

    s = str(val).strip()
    if s == "":
        return ""

    s = unicodedata.normalize("NFKC", s)

    try:
        f = float(s.replace(",", "").replace("%", ""))
        return f"{round(f, 1):.1f}"
    except:
        return s


def parse_sheet(sheetname: str):
    parts = sheetname.rsplit(" ", 1)
    if len(parts) == 2 and parts[1].isdigit():
        return parts[0], parts[1], sheetname
    return sheetname, "", sheetname


# ---------------- MAIN APP ----------------
def main():
    try:
        st.image("header.png", use_container_width=True)
    except:
        pass

    st.markdown(
        "<h2 style='text-align:center;'>Result Sessional Odd Sem 2025</h2>",
        unsafe_allow_html=True
    )

    # Load Google Sheets
    try:
        xls = load_excelfile_from_sheet(GOOGLE_SHEET_FILE_ID)
        sheets = xls.sheet_names
    except Exception as e:
        st.error(f"Unable to load Google Sheet: {e}")
        return

    parsed = [parse_sheet(s) for s in sheets]
    universities = sorted({u for u, sem, orig in parsed})

    uni = st.selectbox("Select University", ["-- choose --"] + universities)
    if uni == "-- choose --":
        return

    sems = sorted({sem for u, sem, orig in parsed if u == uni})
    sem = st.selectbox("Select Semester", ["-- choose --"] + sems)
    if sem == "-- choose --":
        return

    sheet_name = [orig for u, s, orig in parsed if u == uni and s == sem][0]

    try:
        df = read_sheet_from_sheet(GOOGLE_SHEET_FILE_ID, sheet_name)
        df.columns = [str(c).strip() for c in df.columns]
    except Exception as e:
        st.error(f"Sheet load error: {e}")
        return

    if COL_STUDENT_NAME not in df.columns or COL_ADMISSION_NO not in df.columns:
        st.error("Required columns not found.")
        return

    # Normalize student names (IMPORTANT FIX)
    df[COL_STUDENT_NAME] = df[COL_STUDENT_NAME].astype(str).str.strip()

    students = sorted(df[COL_STUDENT_NAME].dropna().unique().tolist())

    student = st.selectbox("Select Student", ["-- choose --"] + students)
    if student == "-- choose --":
        return

    student_row = df.loc[df[COL_STUDENT_NAME] == student]

    if student_row.empty:
        st.error("Student not found.")
        return

    row = student_row.iloc[0]

    # Columns to exclude
    exclude_cols = {COL_STUDENT_NAME, COL_ADMISSION_NO}
    if COL_FATHER_NAME in df.columns:
        exclude_cols.add(COL_FATHER_NAME)

    result_cols = [c for c in df.columns if c not in exclude_cols]

    # ---------------- DISPLAY ----------------
    st.subheader("Student Details")
    st.write(f"**Student Name:** {row[COL_STUDENT_NAME]}")
    st.write(f"**Admission No.:** {row[COL_ADMISSION_NO]}")
    if COL_FATHER_NAME in df.columns:
        st.write(f"**Father Name:** {row[COL_FATHER_NAME]}")

    st.subheader("Subjects / Marks")

    data = []
    for col in result_cols:
        data.append({
            "Subject": col,
            "Marks": to_text_one_decimal(row[col])
        })

    st.dataframe(pd.DataFrame(data), use_container_width=True)


if __name__ == "__main__":
    main()
