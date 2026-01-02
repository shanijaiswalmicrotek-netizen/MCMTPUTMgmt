# main.py
import streamlit as st
import pandas as pd
import requests
import io
import unicodedata

# ---------------- CONFIG ----------------
GOOGLE_SHEET_FILE_ID = "1H_aN66Joy7Tuzx8NjTygOsMA1J2OUjnz"  # from your link
COL_STUDENT_NAME = "Student Name"
COL_ADMISSION_NO = "Admission No."
COL_FATHER_NAME = "Father Name"   # optional

st.set_page_config(page_title="Marksheet Viewer", layout="wide")


# ---------------- GOOGLE SHEETS LOADER ----------------
@st.cache_data(ttl=300)  # cache for 5 minutes; adjust as needed
def download_sheet_xlsx(file_id: str) -> bytes:
    """
    Download a Google Sheet as an XLSX file via the export endpoint.
    The sheet must be shared (Anyone with the link can view) or accessible.
    """
    url = f"https://docs.google.com/spreadsheets/d/{file_id}/export?format=xlsx"
    r = requests.get(url, timeout=30)
    r.raise_for_status()
    return r.content


def load_excelfile_from_sheet(file_id: str):
    """Return a pandas.ExcelFile for the downloaded bytes."""
    content = download_sheet_xlsx(file_id)
    return pd.ExcelFile(io.BytesIO(content), engine="openpyxl")


def read_sheet_from_sheet(file_id: str, sheet_name: str):
    content = download_sheet_xlsx(file_id)
    return pd.read_excel(io.BytesIO(content), sheet_name=sheet_name, engine="openpyxl")


# Clean converter → always returns text
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
    """Extract university + semester from a sheet name like 'MGKVP 1'."""
    parts = sheetname.rsplit(" ", 1)
    if len(parts) == 2 and parts[1].isdigit():
        return parts[0], parts[1], sheetname
    return sheetname, "", sheetname


def main():
    # Header image + Result title (if header missing, show nothing)
    try:
        st.image("header.png", use_container_width=True)
    except Exception:
        pass

    st.markdown(
        "<h2 style='text-align:center; font-weight:700;'>Result Sessional Odd Sem 2025</h2>",
        unsafe_allow_html=True
    )

    # Load sheet names
    try:
        xls = load_excelfile_from_sheet(GOOGLE_SHEET_FILE_ID)
        sheets = xls.sheet_names
    except Exception as e:
        st.error(f"Could not load Google Sheet: {e}")
        return

    parsed = [parse_sheet(s) for s in sheets]
    universities = sorted({u for u, sem, orig in parsed})

    # UI: Select University
    uni = st.selectbox("Select University", ["-- choose --"] + universities)
    if uni == "-- choose --":
        return

    sems = sorted({sem for u, sem, orig in parsed if u == uni})
    sem = st.selectbox("Select Semester", ["-- choose --"] + sems)
    if sem == "-- choose --":
        return

    sheet_name = [orig for u, s, orig in parsed if u == uni and s == sem][0]

    # Load selected sheet
    try:
        df = read_sheet_from_sheet(GOOGLE_SHEET_FILE_ID, sheet_name)
        df.columns = [str(c).strip() for c in df.columns]
    except Exception as e:
        st.error(f"Could not load sheet '{sheet_name}': {e}")
        return

    # Required columns
    if COL_STUDENT_NAME not in df.columns:
        st.error(f"Column '{COL_STUDENT_NAME}' not found.")
        return
    if COL_ADMISSION_NO not in df.columns:
        st.error(f"Column '{COL_ADMISSION_NO}' not found.")
        return

    students = df[COL_STUDENT_NAME].astype(str).dropna().unique().tolist()

    student = st.selectbox("Select Student", ["-- choose --"] + students)
    if student == "-- choose --":
        return

    row = df[df[COL_STUDENT_NAME].astype(str) == student].iloc[0]

    # Columns to exclude
    exclude = {COL_STUDENT_NAME, COL_ADMISSION_NO}
    if COL_FATHER_NAME in df.columns:
        exclude.add(COL_FATHER_NAME)

    attributes = [c for c in df.columns if c not in exclude]

    # ---------------- DISPLAY ----------------
    st.subheader("Student Details")
    st.write(f"**{COL_STUDENT_NAME}:** {row[COL_STUDENT_NAME]}")
    st.write(f"**{COL_ADMISSION_NO}:** {row[COL_ADMISSION_NO]}")
    if COL_FATHER_NAME in df.columns:
        st.write(f"**{COL_FATHER_NAME}:** {row[COL_FATHER_NAME]}")

    st.subheader("Attributes / Subjects")

    attr_list = []
    for col in attributes:
        formatted = to_text_one_decimal(row[col])
        attr_list.append({"Attribute / Subject": col, "Value": formatted})

    st.table(pd.DataFrame(attr_list))


if __name__ == "__main__":
    main()
