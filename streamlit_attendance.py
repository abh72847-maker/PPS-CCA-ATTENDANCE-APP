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
        background:
            radial-gradient(circle at top left, rgba(56, 189, 248, 0.18), transparent 28%),
            radial-gradient(circle at top right, rgba(14, 165, 233, 0.12), transparent 24%),
            linear-gradient(180deg, #edf4ff 0%, #f8fbff 100%);
    }
    [data-testid="stSidebar"] {
        background: linear-gradient(180deg, #f8fbff 0%, #eef4ff 100%);
        border-right: 1px solid rgba(148, 163, 184, 0.18);
    }
    .block-container {
        padding-top: 2rem;
        padding-bottom: 1.5rem;
    }
    .hero-card {
        padding: 1.5rem 1.6rem;
        border-radius: 26px;
        background: linear-gradient(135deg, #1d4ed8 0%, #0f766e 100%);
        color: white;
        box-shadow: 0 24px 50px rgba(29, 78, 216, 0.20);
        margin-bottom: 1rem;
        position: relative;
        overflow: hidden;
    }
    .hero-card::before {
        content: "";
        position: absolute;
        right: -45px;
        top: -55px;
        width: 190px;
        height: 190px;
        background: rgba(255, 255, 255, 0.08);
        border-radius: 50%;
    }
    .hero-kicker {
        font-size: 0.82rem;
        font-weight: 700;
        letter-spacing: 0.08em;
        text-transform: uppercase;
        opacity: 0.82;
        margin-bottom: 0.4rem;
    }
    .hero-title {
        font-size: 2rem;
        font-weight: 800;
        line-height: 1.1;
        margin: 0;
    }
    .hero-subtitle {
        margin: 0.5rem 0 0 0;
        font-size: 1rem;
        color: rgba(255, 255, 255, 0.92);
        max-width: 42rem;
    }
    .day-chip-row {
        display: flex;
        gap: 0.55rem;
        flex-wrap: wrap;
        margin-top: 1rem;
    }
    .day-chip {
        padding: 0.45rem 0.8rem;
        border-radius: 999px;
        background: rgba(255, 255, 255, 0.16);
        color: rgba(255, 255, 255, 0.88);
        font-size: 0.85rem;
        font-weight: 700;
    }
    .day-chip.active {
        background: white;
        color: #0f766e;
    }
    .info-strip {
        display: flex;
        justify-content: space-between;
        align-items: center;
        gap: 1rem;
        padding: 0.95rem 1rem;
        border-radius: 18px;
        background: rgba(255, 255, 255, 0.82);
        border: 1px solid rgba(191, 219, 254, 0.95);
        margin-bottom: 1rem;
        box-shadow: 0 12px 28px rgba(148, 163, 184, 0.10);
        backdrop-filter: blur(8px);
    }
    .info-main {
        color: #0f172a;
        font-weight: 800;
    }
    .info-sub {
        color: #64748b;
        font-size: 0.93rem;
    }
    .highlight-day {
        color: #0f766e;
    }
    .panel {
        background: rgba(255, 255, 255, 0.82);
        border: 1px solid rgba(191, 219, 254, 0.95);
        border-radius: 22px;
        padding: 1rem 1rem 0.9rem 1rem;
        box-shadow: 0 12px 28px rgba(148, 163, 184, 0.10);
        backdrop-filter: blur(8px);
        margin-bottom: 1rem;
    }
    .section-title {
        font-size: 1.12rem;
        font-weight: 800;
        color: #0f172a;
        margin-bottom: 0.2rem;
    }
    .section-note {
        color: #64748b;
        font-size: 0.92rem;
        margin-bottom: 0.9rem;
    }
    .stat-card {
        border-radius: 18px;
        padding: 0.95rem 1rem;
        color: white;
        margin-bottom: 0.75rem;
        box-shadow: 0 10px 24px rgba(148, 163, 184, 0.14);
    }
    .stat-label {
        font-size: 0.78rem;
        text-transform: uppercase;
        letter-spacing: 0.08em;
        opacity: 0.84;
        margin-bottom: 0.35rem;
    }
    .stat-value {
        font-size: 1.8rem;
        font-weight: 800;
        line-height: 1;
    }
    .total { background: linear-gradient(135deg, #2563eb, #38bdf8); }
    .present { background: linear-gradient(135deg, #15803d, #4ade80); }
    .absent { background: linear-gradient(135deg, #b91c1c, #fb7185); }
    .pending { background: linear-gradient(135deg, #a16207, #fbbf24); }
    div[data-testid="stDataEditor"], div[data-testid="stDataFrame"] {
        background: white;
        border: 1px solid #dbeafe;
        border-radius: 18px;
        box-shadow: 0 8px 22px rgba(148, 163, 184, 0.10);
        overflow: hidden;
    }
    div.stButton > button, div.stDownloadButton > button {
        border-radius: 14px;
        border: none;
        font-weight: 700;
        min-height: 2.8rem;
    }
    </style>
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
today_name = datetime.now().strftime("%A")
default_day_index = DAYS.index(today_name) if today_name in DAYS else 0
current_day = st.sidebar.radio("Select attendance day", DAYS, index=default_day_index)
st.sidebar.success(f"Using fixed file: {data_source}")
st.sidebar.caption("Choose a day, edit the table, then save that day.")

st.markdown(
    f"""
    <div class="hero-card">
        <div class="hero-kicker">Attendance Suite</div>
        <h1 class="hero-title">Student Attendance Dashboard</h1>
        <p class="hero-subtitle">
            Fast day-wise attendance marking with a cleaner classroom dashboard and same-day report export.
        </p>
        <div class="day-chip-row">
            {''.join(f'<span class="day-chip{" active" if day == current_day else ""}">{day}</span>' for day in DAYS)}
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)

st.markdown(
    f"""
    <div class="info-strip">
        <div>
            <div class="info-main">Working on <span class="highlight-day">{current_day}</span></div>
            <div class="info-sub">{datetime.now().strftime("%A, %d %B %Y")} • Source file: {data_source}</div>
        </div>
        <div class="info-sub">Update statuses in the table and save the selected day.</div>
    </div>
    """,
    unsafe_allow_html=True,
)

editor_key = f"editor_{current_day}"
if editor_key not in st.session_state:
    st.session_state[editor_key] = build_day_dataframe(
        students, st.session_state.attendance[current_day]
    )

left_col, right_col = st.columns([2.3, 1])

with right_col:
    st.markdown('<div class="panel">', unsafe_allow_html=True)
    st.markdown('<div class="section-title">Summary</div>', unsafe_allow_html=True)
    st.markdown('<div class="section-note">Quick view of attendance for the selected day.</div>', unsafe_allow_html=True)

    preview_df = st.session_state[editor_key]
    total = len(preview_df)
    present = int((preview_df["Status"] == "Present").sum())
    absent = int((preview_df["Status"] == "Absent").sum())
    pending = int((preview_df["Status"] == "Pending").sum())

    metric_a, metric_b = st.columns(2)
    metric_c, metric_d = st.columns(2)
    metric_a.markdown(
        f'<div class="stat-card total"><div class="stat-label">Total</div><div class="stat-value">{total}</div></div>',
        unsafe_allow_html=True,
    )
    metric_b.markdown(
        f'<div class="stat-card present"><div class="stat-label">Present</div><div class="stat-value">{present}</div></div>',
        unsafe_allow_html=True,
    )
    metric_c.markdown(
        f'<div class="stat-card absent"><div class="stat-label">Absent</div><div class="stat-value">{absent}</div></div>',
        unsafe_allow_html=True,
    )
    metric_d.markdown(
        f'<div class="stat-card pending"><div class="stat-label">Pending</div><div class="stat-value">{pending}</div></div>',
        unsafe_allow_html=True,
    )

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

    st.markdown("</div>", unsafe_allow_html=True)

with left_col:
    st.markdown('<div class="panel">', unsafe_allow_html=True)
    st.markdown(
        f'<div class="section-title">Mark Attendance for {current_day}</div>',
        unsafe_allow_html=True,
    )
    st.markdown(
        '<div class="section-note">Use the status dropdown for each student, then save the day when you are done.</div>',
        unsafe_allow_html=True,
    )

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

    st.markdown("</div>", unsafe_allow_html=True)
