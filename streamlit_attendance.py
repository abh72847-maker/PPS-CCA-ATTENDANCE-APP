from datetime import datetime
from pathlib import Path
from io import BytesIO
import zipfile
import xml.etree.ElementTree as ET

import pandas as pd
import streamlit as st


APP_DIR = Path(__file__).resolve().parent
DEFAULT_FILE = APP_DIR / "student list.xlsx"
DAYS = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday"]


@st.cache_data
def load_students_from_excel(file_obj):
    ns = {"a": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
    students = []

    if hasattr(file_obj, "read"):
        content = file_obj.read()
        file_obj = BytesIO(content)

    with zipfile.ZipFile(file_obj) as book:
        shared_strings = []
        if "xl/sharedStrings.xml" in book.namelist():
            root = ET.fromstring(book.read("xl/sharedStrings.xml"))
            for item in root.findall("a:si", ns):
                text = "".join(node.text or "" for node in item.findall(".//a:t", ns))
                shared_strings.append(text.strip())

        sheet = ET.fromstring(book.read("xl/worksheets/sheet1.xml"))
        for row in sheet.findall(".//a:sheetData/a:row", ns):
            values = []
            for cell in row.findall("a:c", ns):
                value = cell.find("a:v", ns)
                if value is None:
                    values.append("")
                elif cell.get("t") == "s":
                    values.append(shared_strings[int(value.text)].strip())
                else:
                    values.append((value.text or "").strip())

            if len(values) >= 3 and values[0].strip().upper() != "SR. NO.":
                students.append({"roll": values[1], "name": values[2]})

    return students


def build_report_rows(students, attendance, current_day):
    rows = []
    day_records = attendance.get(current_day, {})
    for student in students:
        roll = student["roll"]
        record = day_records.get(roll, {})
        rows.append(
            {
                "Day": current_day,
                "Roll No": roll,
                "Student Name": student["name"],
                "Status": record.get("status", "Pending"),
                "Marked Time": record.get("time", ""),
            }
        )
    return rows


def mark_attendance(day, student, status):
    st.session_state.attendance[day][student["roll"]] = {
        "name": student["name"],
        "status": status,
        "time": datetime.now().strftime("%I:%M:%S %p"),
    }


def build_day_dataframe(students, day_records):
    rows = []
    for student in students:
        record = day_records.get(student["roll"], {})
        rows.append(
            {
                "Roll No": student["roll"],
                "Student Name": student["name"],
                "Status": record.get("status", "Pending"),
                "Marked Time": record.get("time", ""),
            }
        )
    return pd.DataFrame(rows)


def save_day_from_editor(current_day, edited_df):
    previous_records = st.session_state.attendance[current_day]
    updated_records = {}

    for _, row in edited_df.iterrows():
        status = row["Status"]
        if status == "Pending":
            continue

        old_record = previous_records.get(row["Roll No"], {})
        updated_records[row["Roll No"]] = {
            "name": row["Student Name"],
            "status": status,
            "time": old_record.get("time") or datetime.now().strftime("%I:%M:%S %p"),
        }

    st.session_state.attendance[current_day] = updated_records


st.set_page_config(page_title="Attendance Dashboard", page_icon="📘", layout="wide")

st.markdown(
    """
    <style>
    .stApp {
        background: linear-gradient(180deg, #eef4ff 0%, #f8fbff 100%);
    }
    .top-card {
        padding: 1rem 1.25rem;
        border-radius: 18px;
        background: #1d4ed8;
        color: white;
        box-shadow: 0 14px 30px rgba(37, 99, 235, 0.18);
        margin-bottom: 1rem;
    }
    .student-row {
        padding: 0.75rem 0.9rem;
        border: 1px solid #dbe6ff;
        border-radius: 14px;
        background: white;
        margin-bottom: 0.55rem;
    }
    .status-pill {
        display: inline-block;
        padding: 0.2rem 0.65rem;
        border-radius: 999px;
        font-size: 0.84rem;
        font-weight: 700;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

st.markdown(
    f"""
    <div class="top-card">
        <h2 style="margin:0;">Student Attendance Dashboard</h2>
        <p style="margin:0.3rem 0 0 0;">{datetime.now().strftime("%A, %d %B %Y")}</p>
    </div>
    """,
    unsafe_allow_html=True,
)

students = []
data_source = None

if DEFAULT_FILE.exists():
    students = load_students_from_excel(DEFAULT_FILE)
    data_source = DEFAULT_FILE.name

if not students:
    st.error("`student list.xlsx` was not found in the project folder.")
    st.stop()

if "students" not in st.session_state or st.session_state.get("data_source") != data_source:
    st.session_state.students = students
    st.session_state.data_source = data_source
    st.session_state.attendance = {day: {} for day in DAYS}

students = st.session_state.students
current_day = st.sidebar.radio("Select attendance day", DAYS, index=0)
st.sidebar.success(f"Using fixed file: {data_source}")

editor_key = f"editor_{current_day}"
if editor_key not in st.session_state:
    st.session_state[editor_key] = build_day_dataframe(
        students, st.session_state.attendance[current_day]
    )

left_col, right_col = st.columns([2.2, 1])

with right_col:
    st.subheader("Summary")
    preview_df = st.session_state[editor_key]
    total = len(preview_df)
    present = int((preview_df["Status"] == "Present").sum())
    absent = int((preview_df["Status"] == "Absent").sum())
    pending = int((preview_df["Status"] == "Pending").sum())

    metric_a, metric_b = st.columns(2)
    metric_c, metric_d = st.columns(2)
    metric_a.metric("Total", total)
    metric_b.metric("Present", present)
    metric_c.metric("Absent", absent)
    metric_d.metric("Pending", pending)

    st.markdown(f"### {current_day} Details")
    detail_rows = preview_df[preview_df["Status"] != "Pending"].copy()

    if not detail_rows.empty:
        st.dataframe(detail_rows, use_container_width=True, hide_index=True)
    else:
        st.info("No attendance marked yet for this day.")

    report_df = pd.DataFrame(build_report_rows(students, st.session_state.attendance, current_day))
    csv_data = report_df.to_csv(index=False).encode("utf-8")
    st.download_button(
        f"Download {current_day} Report",
        data=csv_data,
        file_name=f"attendance_{current_day.lower()}.csv",
        mime="text/csv",
        use_container_width=True,
    )

    if st.button("Reset Current Day", use_container_width=True):
        st.session_state.attendance[current_day] = {}
        st.session_state[editor_key] = build_day_dataframe(students, {})
        st.rerun()

with left_col:
    st.subheader(f"Mark Attendance for {current_day}")
    st.caption(f"Student source: {data_source}")
    edited_df = st.data_editor(
        st.session_state[editor_key],
        use_container_width=True,
        hide_index=True,
        num_rows="fixed",
        column_config={
            "Roll No": st.column_config.TextColumn(disabled=True),
            "Student Name": st.column_config.TextColumn(disabled=True, width="large"),
            "Status": st.column_config.SelectboxColumn(
                "Status",
                options=["Pending", "Present", "Absent"],
                required=True,
                width="medium",
            ),
            "Marked Time": st.column_config.TextColumn(disabled=True),
        },
        key=f"data_editor_{current_day}",
    )
    st.session_state[editor_key] = edited_df

    if st.button(f"Save {current_day} Attendance", type="primary", use_container_width=True):
        save_day_from_editor(current_day, edited_df)
        st.session_state[editor_key] = build_day_dataframe(
            students, st.session_state.attendance[current_day]
        )
        st.success(f"{current_day} attendance saved in the app.")
        st.rerun()
