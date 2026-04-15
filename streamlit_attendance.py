from datetime import datetime
from io import BytesIO
from pathlib import Path
import zipfile
import xml.etree.ElementTree as ET

import pandas as pd
import streamlit as st


APP_DIR = Path(__file__).resolve().parent
DEFAULT_FILE = APP_DIR / "student list.xlsx"
DAYS = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday"]
STATUS_OPTIONS = ["Pending", "Present", "Absent"]


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


def build_report_rows(students, attendance, current_day):
    day_records = attendance.get(current_day, {})
    rows = []
    for student in students:
        record = day_records.get(student["roll"], {})
        rows.append(
            {
                "Day": current_day,
                "Roll No": student["roll"],
                "Student Name": student["name"],
                "Status": record.get("status", "Pending"),
                "Marked Time": record.get("time", ""),
            }
        )
    return rows


def calculate_counts(df):
    total = len(df)
    present = int((df["Status"] == "Present").sum())
    absent = int((df["Status"] == "Absent").sum())
    pending = int((df["Status"] == "Pending").sum())
    completion = 0 if total == 0 else int(((present + absent) / total) * 100)
    return total, present, absent, pending, completion


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


def apply_bulk_status(base_df, filtered_df, new_status):
    updated_df = base_df.copy()
    rolls = set(filtered_df["Roll No"].tolist())
    mask = updated_df["Roll No"].isin(rolls)
    updated_df.loc[mask, "Status"] = new_status

    if new_status == "Pending":
        updated_df.loc[mask, "Marked Time"] = ""
    else:
        empty_times = mask & (updated_df["Marked Time"] == "")
        updated_df.loc[empty_times, "Marked Time"] = datetime.now().strftime("%I:%M:%S %p")

    return updated_df


def style_day_overview(day, count, is_current):
    extra = " active" if is_current else ""
    return (
        f'<div class="day-overview{extra}">'
        f'<div class="day-name">{day}</div>'
        f'<div class="day-count">{count}</div>'
        f'<div class="day-meta">marked students</div>'
        f"</div>"
    )


st.set_page_config(page_title="Attendance Dashboard", page_icon="A", layout="wide")

st.markdown(
    """
    <style>
    .stApp {
        background:
            radial-gradient(circle at 10% 0%, rgba(56, 189, 248, 0.20), transparent 26%),
            radial-gradient(circle at 100% 0%, rgba(20, 184, 166, 0.16), transparent 24%),
            linear-gradient(180deg, #eef6ff 0%, #f8fbff 100%);
    }
    [data-testid="stSidebar"] {
        background: linear-gradient(180deg, #f9fbff 0%, #eef5ff 100%);
        border-right: 1px solid rgba(148, 163, 184, 0.18);
    }
    .block-container {
        padding-top: 1.7rem;
        padding-bottom: 1.4rem;
        max-width: 1450px;
    }
    .hero-card {
        padding: 1.6rem 1.7rem;
        border-radius: 28px;
        background:
            radial-gradient(circle at top right, rgba(255,255,255,0.18), transparent 20%),
            linear-gradient(135deg, #1d4ed8 0%, #0f766e 100%);
        color: white;
        box-shadow: 0 24px 55px rgba(29, 78, 216, 0.20);
        margin-bottom: 1rem;
        overflow: hidden;
        position: relative;
    }
    .hero-card::after {
        content: "";
        position: absolute;
        width: 220px;
        height: 220px;
        right: -90px;
        bottom: -90px;
        background: rgba(255, 255, 255, 0.08);
        border-radius: 50%;
    }
    .eyebrow {
        font-size: 0.8rem;
        font-weight: 800;
        letter-spacing: 0.08em;
        text-transform: uppercase;
        opacity: 0.82;
        margin-bottom: 0.35rem;
    }
    .hero-title {
        margin: 0;
        font-size: 2.15rem;
        line-height: 1.05;
        font-weight: 900;
    }
    .hero-subtitle {
        margin: 0.55rem 0 0 0;
        max-width: 46rem;
        color: rgba(255, 255, 255, 0.92);
        font-size: 1rem;
    }
    .hero-row {
        display: flex;
        justify-content: space-between;
        gap: 1rem;
        align-items: flex-end;
        flex-wrap: wrap;
        margin-top: 1rem;
    }
    .hero-badge {
        padding: 0.5rem 0.85rem;
        border-radius: 999px;
        background: rgba(255, 255, 255, 0.16);
        font-size: 0.88rem;
        font-weight: 700;
        color: rgba(255, 255, 255, 0.95);
        margin-right: 0.45rem;
        display: inline-block;
    }
    .hero-badge.active {
        background: white;
        color: #0f766e;
    }
    .info-strip {
        display: flex;
        justify-content: space-between;
        gap: 1rem;
        align-items: center;
        padding: 1rem 1.1rem;
        border-radius: 20px;
        background: rgba(255,255,255,0.82);
        border: 1px solid rgba(191, 219, 254, 0.95);
        box-shadow: 0 12px 30px rgba(148, 163, 184, 0.10);
        backdrop-filter: blur(8px);
        margin-bottom: 1rem;
        flex-wrap: wrap;
    }
    .info-title {
        color: #0f172a;
        font-weight: 800;
    }
    .info-sub {
        color: #64748b;
        font-size: 0.93rem;
    }
    .panel {
        background: rgba(255,255,255,0.84);
        border: 1px solid rgba(191, 219, 254, 0.95);
        border-radius: 24px;
        padding: 1rem 1rem 0.95rem 1rem;
        box-shadow: 0 14px 30px rgba(148, 163, 184, 0.10);
        backdrop-filter: blur(8px);
        margin-bottom: 1rem;
    }
    .section-title {
        font-size: 1.12rem;
        font-weight: 900;
        color: #0f172a;
        margin-bottom: 0.25rem;
    }
    .section-note {
        color: #64748b;
        font-size: 0.93rem;
        margin-bottom: 0.9rem;
    }
    .stat-card {
        border-radius: 18px;
        padding: 1rem;
        color: white;
        margin-bottom: 0.75rem;
        box-shadow: 0 12px 24px rgba(148, 163, 184, 0.15);
    }
    .stat-label {
        font-size: 0.78rem;
        text-transform: uppercase;
        letter-spacing: 0.08em;
        opacity: 0.82;
        margin-bottom: 0.4rem;
    }
    .stat-value {
        font-size: 1.9rem;
        line-height: 1;
        font-weight: 900;
    }
    .stat-total { background: linear-gradient(135deg, #2563eb, #38bdf8); }
    .stat-present { background: linear-gradient(135deg, #15803d, #4ade80); }
    .stat-absent { background: linear-gradient(135deg, #b91c1c, #fb7185); }
    .stat-pending { background: linear-gradient(135deg, #a16207, #fbbf24); }
    .overview-grid {
        display: grid;
        grid-template-columns: repeat(5, minmax(0, 1fr));
        gap: 0.7rem;
        margin-top: 0.7rem;
    }
    .day-overview {
        border-radius: 18px;
        padding: 0.9rem;
        background: linear-gradient(180deg, #f8fbff, #eef6ff);
        border: 1px solid #dbeafe;
    }
    .day-overview.active {
        background: linear-gradient(135deg, #1d4ed8, #0f766e);
        color: white;
        border-color: transparent;
        box-shadow: 0 14px 26px rgba(29, 78, 216, 0.18);
    }
    .day-name {
        font-size: 0.85rem;
        font-weight: 800;
        margin-bottom: 0.3rem;
    }
    .day-count {
        font-size: 1.45rem;
        line-height: 1;
        font-weight: 900;
        margin-bottom: 0.2rem;
    }
    .day-meta {
        font-size: 0.82rem;
        opacity: 0.8;
    }
    div[data-testid="stDataEditor"], div[data-testid="stDataFrame"] {
        background: white;
        border: 1px solid #dbeafe;
        border-radius: 18px;
        box-shadow: 0 8px 24px rgba(148, 163, 184, 0.10);
        overflow: hidden;
    }
    div.stButton > button, div.stDownloadButton > button {
        border-radius: 14px;
        border: none;
        font-weight: 800;
        min-height: 2.8rem;
    }
    .pill {
        display: inline-block;
        padding: 0.32rem 0.72rem;
        border-radius: 999px;
        background: #dcfce7;
        color: #166534;
        font-weight: 800;
        font-size: 0.84rem;
    }
    @media (max-width: 1100px) {
        .overview-grid {
            grid-template-columns: repeat(2, minmax(0, 1fr));
        }
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

st.sidebar.title("Attendance Controls")
current_day = st.sidebar.radio("Select day", DAYS, index=default_day_index)
search_text = st.sidebar.text_input("Search by roll or name")
status_filter = st.sidebar.multiselect(
    "Show statuses",
    STATUS_OPTIONS,
    default=STATUS_OPTIONS,
)
show_only_unmarked = st.sidebar.toggle("Show only pending", value=False)
st.sidebar.markdown("---")
st.sidebar.success(f"Fixed data source: {data_source}")
st.sidebar.caption("Use filters to work faster on large class lists.")

editor_key = f"editor_{current_day}"
if editor_key not in st.session_state:
    st.session_state[editor_key] = build_day_dataframe(
        students, st.session_state.attendance[current_day]
    )

base_df = st.session_state[editor_key].copy()

if search_text:
    query = search_text.strip().lower()
    mask = (
        base_df["Roll No"].str.lower().str.contains(query)
        | base_df["Student Name"].str.lower().str.contains(query)
    )
    base_df = base_df[mask]

if status_filter:
    base_df = base_df[base_df["Status"].isin(status_filter)]

if show_only_unmarked:
    base_df = base_df[base_df["Status"] == "Pending"]

day_mark_counts = {
    day: len(st.session_state.attendance.get(day, {}))
    for day in DAYS
}

full_day_df = st.session_state[editor_key]
total, present, absent, pending, completion = calculate_counts(full_day_df)

st.markdown(
    f"""
    <div class="hero-card">
        <div class="eyebrow">Attendance Suite</div>
        <h1 class="hero-title">Student Attendance Dashboard</h1>
        <p class="hero-subtitle">
            A polished day-wise attendance workspace with fast filtering, quick progress tracking, and ready-to-download same-day reports.
        </p>
        <div class="hero-row">
            <div>
                {''.join(f'<span class="hero-badge{" active" if day == current_day else ""}">{day}</span>' for day in DAYS)}
            </div>
            <div class="hero-badge active">{completion}% complete</div>
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)

st.markdown(
    f"""
    <div class="info-strip">
        <div>
            <div class="info-title">Now working on {current_day}</div>
            <div class="info-sub">{datetime.now().strftime("%A, %d %B %Y")} • {total} students in the class list</div>
        </div>
        <div class="info-sub">Filtered view: {len(base_df)} students • Save after editing to store the day</div>
    </div>
    """,
    unsafe_allow_html=True,
)

overview_panel, top_stats = st.columns([2.3, 1.7])

with overview_panel:
    st.markdown('<div class="panel">', unsafe_allow_html=True)
    st.markdown('<div class="section-title">Weekly Overview</div>', unsafe_allow_html=True)
    st.markdown(
        '<div class="section-note">Each card shows how many students have been marked for that day so far.</div>',
        unsafe_allow_html=True,
    )
    st.markdown(
        '<div class="overview-grid">'
        + "".join(
            style_day_overview(day, day_mark_counts[day], day == current_day)
            for day in DAYS
        )
        + "</div>",
        unsafe_allow_html=True,
    )
    st.markdown("</div>", unsafe_allow_html=True)

with top_stats:
    st.markdown('<div class="panel">', unsafe_allow_html=True)
    st.markdown('<div class="section-title">Live Snapshot</div>', unsafe_allow_html=True)
    st.markdown(
        '<div class="section-note">The summary below always reflects the full selected day.</div>',
        unsafe_allow_html=True,
    )
    a, b = st.columns(2)
    c, d = st.columns(2)
    a.markdown(
        f'<div class="stat-card stat-total"><div class="stat-label">Total</div><div class="stat-value">{total}</div></div>',
        unsafe_allow_html=True,
    )
    b.markdown(
        f'<div class="stat-card stat-present"><div class="stat-label">Present</div><div class="stat-value">{present}</div></div>',
        unsafe_allow_html=True,
    )
    c.markdown(
        f'<div class="stat-card stat-absent"><div class="stat-label">Absent</div><div class="stat-value">{absent}</div></div>',
        unsafe_allow_html=True,
    )
    d.markdown(
        f'<div class="stat-card stat-pending"><div class="stat-label">Pending</div><div class="stat-value">{pending}</div></div>',
        unsafe_allow_html=True,
    )
    st.progress(completion / 100 if completion else 0, text=f"{completion}% of {current_day} marked")
    st.markdown("</div>", unsafe_allow_html=True)

main_col, side_col = st.columns([2.2, 1])

with main_col:
    st.markdown('<div class="panel">', unsafe_allow_html=True)
    st.markdown(
        f'<div class="section-title">Mark Attendance for {current_day}</div>',
        unsafe_allow_html=True,
    )
    st.markdown(
        '<div class="section-note">Filter the list from the sidebar, edit the statuses below, then save the selected day.</div>',
        unsafe_allow_html=True,
    )

    quick_a, quick_b, quick_c = st.columns(3)
    if quick_a.button("Mark Filtered Present", use_container_width=True):
        st.session_state[editor_key] = apply_bulk_status(
            st.session_state[editor_key], base_df, "Present"
        )
        st.rerun()
    if quick_b.button("Mark Filtered Absent", use_container_width=True):
        st.session_state[editor_key] = apply_bulk_status(
            st.session_state[editor_key], base_df, "Absent"
        )
        st.rerun()
    if quick_c.button("Clear Filtered", use_container_width=True):
        st.session_state[editor_key] = apply_bulk_status(
            st.session_state[editor_key], base_df, "Pending"
        )
        st.rerun()

    edited_df = st.data_editor(
        base_df,
        use_container_width=True,
        hide_index=True,
        num_rows="fixed",
        column_config={
            "Roll No": st.column_config.TextColumn(disabled=True, width="small"),
            "Student Name": st.column_config.TextColumn(disabled=True, width="large"),
            "Status": st.column_config.SelectboxColumn(
                "Status",
                options=STATUS_OPTIONS,
                required=True,
                width="medium",
            ),
            "Marked Time": st.column_config.TextColumn(disabled=True, width="medium"),
        },
        key=f"data_editor_{current_day}_{len(base_df)}",
    )

    merged_df = st.session_state[editor_key].copy()
    for _, row in edited_df.iterrows():
        row_mask = merged_df["Roll No"] == row["Roll No"]
        merged_df.loc[row_mask, "Status"] = row["Status"]
        if row["Status"] == "Pending":
            merged_df.loc[row_mask, "Marked Time"] = ""
        elif merged_df.loc[row_mask, "Marked Time"].iloc[0] == "":
            merged_df.loc[row_mask, "Marked Time"] = datetime.now().strftime("%I:%M:%S %p")

    st.session_state[editor_key] = merged_df

    if st.button(f"Save {current_day} Attendance", type="primary", use_container_width=True):
        save_day_from_editor(current_day, st.session_state[editor_key])
        st.session_state[editor_key] = build_day_dataframe(
            students, st.session_state.attendance[current_day]
        )
        st.success(f"{current_day} attendance saved.")
        st.rerun()

    st.markdown("</div>", unsafe_allow_html=True)

with side_col:
    st.markdown('<div class="panel">', unsafe_allow_html=True)
    st.markdown('<div class="section-title">Reports</div>', unsafe_allow_html=True)
    st.markdown(
        '<div class="section-note">Download the selected day only. Pending students remain included in the export.</div>',
        unsafe_allow_html=True,
    )

    marked_df = st.session_state[editor_key]
    marked_rows = marked_df[marked_df["Status"] != "Pending"].copy()
    if not marked_rows.empty:
        st.dataframe(marked_rows, use_container_width=True, hide_index=True)
    else:
        st.info("No saved attendance for this day yet.")

    report_df = pd.DataFrame(
        build_report_rows(students, st.session_state.attendance, current_day)
    )
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
