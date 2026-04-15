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


def build_report_rows(students, attendance):
    rows = []
    for day in DAYS:
        day_records = attendance.get(day, {})
        for student in students:
            roll = student["roll"]
            record = day_records.get(roll, {})
            rows.append(
                {
                    "Day": day,
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
        <p style="margin:0.3rem 0 0 0;">{datetime.now().strftime("%A, %d %B %Y   %I:%M:%S %p")}</p>
    </div>
    """,
    unsafe_allow_html=True,
)

uploaded_file = st.sidebar.file_uploader("Upload student list (.xlsx)", type=["xlsx"])

students = []
data_source = None

if uploaded_file is not None:
    students = load_students_from_excel(uploaded_file)
    data_source = uploaded_file.name
elif DEFAULT_FILE.exists():
    students = load_students_from_excel(DEFAULT_FILE)
    data_source = DEFAULT_FILE.name

if not students:
    st.warning("Upload your `student list.xlsx` file from the sidebar to start the app.")
    st.stop()

if "students" not in st.session_state or st.session_state.get("data_source") != data_source:
    st.session_state.students = students
    st.session_state.data_source = data_source
    st.session_state.attendance = {day: {} for day in DAYS}

students = st.session_state.students
current_day = st.sidebar.radio("Select attendance day", DAYS, index=0)

left_col, right_col = st.columns([2.2, 1])

with right_col:
    st.subheader("Summary")
    day_records = st.session_state.attendance[current_day]
    total = len(students)
    present = sum(1 for item in day_records.values() if item["status"] == "Present")
    absent = sum(1 for item in day_records.values() if item["status"] == "Absent")
    pending = total - len(day_records)

    metric_a, metric_b = st.columns(2)
    metric_c, metric_d = st.columns(2)
    metric_a.metric("Total", total)
    metric_b.metric("Present", present)
    metric_c.metric("Absent", absent)
    metric_d.metric("Pending", pending)

    st.markdown(f"### {current_day} Details")
    detail_rows = []
    for student in students:
        record = day_records.get(student["roll"])
        if record:
            detail_rows.append(
                {
                    "Roll No": student["roll"],
                    "Status": record["status"],
                    "Time": record["time"],
                }
            )

    if detail_rows:
        st.dataframe(pd.DataFrame(detail_rows), use_container_width=True, hide_index=True)
    else:
        st.info("No attendance marked yet for this day.")

    report_df = pd.DataFrame(build_report_rows(students, st.session_state.attendance))
    csv_data = report_df.to_csv(index=False).encode("utf-8")
    st.download_button(
        "Download Attendance Report",
        data=csv_data,
        file_name="attendance_report.csv",
        mime="text/csv",
        use_container_width=True,
    )

    if st.button("Reset Current Day", use_container_width=True):
        st.session_state.attendance[current_day] = {}
        st.rerun()

with left_col:
    st.subheader(f"Mark Attendance for {current_day}")
    st.caption(f"Student source: {data_source}")

    for student in students:
        record = st.session_state.attendance[current_day].get(student["roll"])
        if not record:
            badge = '<span class="status-pill" style="background:#fef3c7;color:#92400e;">Pending</span>'
        elif record["status"] == "Present":
            badge = '<span class="status-pill" style="background:#dcfce7;color:#166534;">Present</span>'
        else:
            badge = '<span class="status-pill" style="background:#fee2e2;color:#991b1b;">Absent</span>'

        st.markdown('<div class="student-row">', unsafe_allow_html=True)
        c1, c2, c3, c4 = st.columns([1.1, 2.8, 1, 1])
        c1.markdown(f"**{student['roll']}**")
        c2.markdown(f"{student['name']}  \n{badge}", unsafe_allow_html=True)
        if c3.button("Present", key=f"p-{current_day}-{student['roll']}", use_container_width=True):
            mark_attendance(current_day, student, "Present")
            st.rerun()
        if c4.button("Absent", key=f"a-{current_day}-{student['roll']}", use_container_width=True):
            mark_attendance(current_day, student, "Absent")
            st.rerun()
        st.markdown("</div>", unsafe_allow_html=True)
