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
        status = record.get("status", "Pending")
        rows.append(
            {
                "Roll No": student["roll"],
                "Student Name": student["name"],
                "Present": status == "Present",
                "Absent": status == "Absent",
                "Marked Time": record.get("time", ""),
            }
        )
    return pd.DataFrame(rows)


def build_report_rows(students, attendance, current_day):
    rows = []
    day_records = attendance.get(current_day, {})
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


def validate_editor(df):
    invalid = df[df["Present"] & df["Absent"]]
    return invalid["Roll No"].tolist()


def save_day_from_editor(current_day, edited_df):
    invalid_rolls = validate_editor(edited_df)
    if invalid_rolls:
        return invalid_rolls

    previous_records = st.session_state.attendance[current_day]
    updated_records = {}

    for _, row in edited_df.iterrows():
        if row["Present"]:
            status = "Present"
        elif row["Absent"]:
            status = "Absent"
        else:
            status = "Pending"

        if status == "Pending":
            continue

        old_record = previous_records.get(row["Roll No"], {})
        updated_records[row["Roll No"]] = {
            "name": row["Student Name"],
            "status": status,
            "time": old_record.get("time") or row["Marked Time"] or datetime.now().strftime("%I:%M:%S %p"),
        }

    st.session_state.attendance[current_day] = updated_records
    return []


def apply_bulk_choice(base_df, filtered_df, choice):
    updated_df = base_df.copy()
    rolls = set(filtered_df["Roll No"].tolist())
    mask = updated_df["Roll No"].isin(rolls)

    if choice == "Present":
        updated_df.loc[mask, "Present"] = True
        updated_df.loc[mask, "Absent"] = False
    elif choice == "Absent":
        updated_df.loc[mask, "Present"] = False
        updated_df.loc[mask, "Absent"] = True
    else:
        updated_df.loc[mask, "Present"] = False
        updated_df.loc[mask, "Absent"] = False
        updated_df.loc[mask, "Marked Time"] = ""

    if choice in {"Present", "Absent"}:
        empty_times = mask & (updated_df["Marked Time"] == "")
        updated_df.loc[empty_times, "Marked Time"] = datetime.now().strftime("%I:%M:%S %p")

    return updated_df


def calculate_counts(attendance, students, day):
    records = attendance.get(day, {})
    total = len(students)
    present = sum(1 for item in records.values() if item["status"] == "Present")
    absent = sum(1 for item in records.values() if item["status"] == "Absent")
    pending = total - len(records)
    completion = 0 if total == 0 else int(((present + absent) / total) * 100)
    return total, present, absent, pending, completion


def render_stat_card(title, value, class_name):
    return (
        f'<div class="stat-card {class_name}">'
        f'<div class="stat-label">{title}</div>'
        f'<div class="stat-value">{value}</div>'
        f"</div>"
    )


st.set_page_config(page_title="Attendance Dashboard", page_icon="A", layout="wide")

st.markdown(
    """
    <style>
    .stApp {
        background:
            radial-gradient(circle at 10% 0%, rgba(56, 189, 248, 0.20), transparent 26%),
            radial-gradient(circle at 100% 0%, rgba(20, 184, 166, 0.14), transparent 24%),
            linear-gradient(180deg, #eef6ff 0%, #f8fbff 100%);
    }
    .block-container {
        padding-top: 1.5rem;
        padding-bottom: 1.4rem;
        max-width: 1450px;
    }
    [data-testid="stSidebar"] {
        background: linear-gradient(180deg, #f8fbff 0%, #eef4ff 100%);
        border-right: 1px solid rgba(148, 163, 184, 0.18);
    }
    .hero-card {
        padding: 1.65rem 1.7rem;
        border-radius: 28px;
        background:
            radial-gradient(circle at top right, rgba(255,255,255,0.18), transparent 18%),
            linear-gradient(135deg, #1d4ed8 0%, #0f766e 100%);
        color: white;
        box-shadow: 0 24px 55px rgba(29, 78, 216, 0.20);
        margin-bottom: 1rem;
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
        font-size: 2.2rem;
        line-height: 1.05;
        font-weight: 900;
    }
    .hero-subtitle {
        margin: 0.55rem 0 0 0;
        max-width: 48rem;
        color: rgba(255, 255, 255, 0.92);
        font-size: 1rem;
    }
    .hero-meta {
        display: flex;
        gap: 0.7rem;
        flex-wrap: wrap;
        margin-top: 1rem;
    }
    .hero-chip {
        padding: 0.48rem 0.84rem;
        border-radius: 999px;
        background: rgba(255, 255, 255, 0.16);
        font-size: 0.88rem;
        font-weight: 700;
    }
    .panel {
        background: rgba(255, 255, 255, 0.85);
        border: 1px solid rgba(191, 219, 254, 0.95);
        border-radius: 24px;
        padding: 1rem 1rem 0.95rem 1rem;
        box-shadow: 0 14px 30px rgba(148, 163, 184, 0.10);
        backdrop-filter: blur(8px);
        margin-bottom: 1rem;
    }
    .section-title {
        font-size: 1.15rem;
        font-weight: 900;
        color: #0f172a;
        margin-bottom: 0.25rem;
    }
    .section-note {
        color: #64748b;
        font-size: 0.93rem;
        margin-bottom: 0.9rem;
    }
    .day-strip {
        background: rgba(255, 255, 255, 0.85);
        border: 1px solid rgba(191, 219, 254, 0.95);
        border-radius: 22px;
        padding: 0.9rem 1rem 0.5rem 1rem;
        box-shadow: 0 12px 28px rgba(148, 163, 184, 0.10);
        margin-bottom: 1rem;
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
    }
    .overview-card {
        border-radius: 18px;
        padding: 0.95rem;
        background: linear-gradient(180deg, #f8fbff, #eef6ff);
        border: 1px solid #dbeafe;
    }
    .overview-card.active {
        background: linear-gradient(135deg, #1d4ed8, #0f766e);
        color: white;
        border-color: transparent;
        box-shadow: 0 14px 26px rgba(29, 78, 216, 0.18);
    }
    .overview-day {
        font-size: 0.86rem;
        font-weight: 800;
        margin-bottom: 0.3rem;
    }
    .overview-count {
        font-size: 1.5rem;
        line-height: 1;
        font-weight: 900;
    }
    .overview-meta {
        font-size: 0.82rem;
        margin-top: 0.2rem;
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
    div[role="radiogroup"] > label {
        background: white;
        border: 1px solid #dbeafe;
        border-radius: 999px;
        padding: 0.2rem 0.5rem;
        margin-right: 0.35rem;
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
    st.session_state.current_day = DAYS[0]

students = st.session_state.students
today_name = datetime.now().strftime("%A")
default_day = today_name if today_name in DAYS else DAYS[0]
if "current_day" not in st.session_state:
    st.session_state.current_day = default_day

st.sidebar.title("Quick Help")
st.sidebar.info("Use the main page day selector and tabs. The sidebar is now just for guidance.")
st.sidebar.caption(f"Source file: {data_source}")

st.markdown(
    f"""
    <div class="hero-card">
        <div class="eyebrow">Attendance Suite</div>
        <h1 class="hero-title">Teacher Attendance Dashboard</h1>
        <p class="hero-subtitle">
            Faster day selection, easier status marking, and a cleaner teaching workflow designed for real classroom use.
        </p>
        <div class="hero-meta">
            <span class="hero-chip">{datetime.now().strftime("%A, %d %B %Y")}</span>
            <span class="hero-chip">{len(students)} students</span>
            <span class="hero-chip">Fixed source: {data_source}</span>
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)

st.markdown('<div class="day-strip">', unsafe_allow_html=True)
st.markdown('<div class="section-title">Select Day</div>', unsafe_allow_html=True)
st.markdown(
    '<div class="section-note">Choose the working day directly from the main interface.</div>',
    unsafe_allow_html=True,
)
current_day = st.radio(
    "Day",
    DAYS,
    horizontal=True,
    index=DAYS.index(st.session_state.current_day),
    label_visibility="collapsed",
)
st.session_state.current_day = current_day
st.markdown("</div>", unsafe_allow_html=True)

editor_key = f"editor_{current_day}"
if editor_key not in st.session_state:
    st.session_state[editor_key] = build_day_dataframe(
        students, st.session_state.attendance[current_day]
    )

full_day_df = st.session_state[editor_key].copy()

overview_cards = []
for day in DAYS:
    count = len(st.session_state.attendance.get(day, {}))
    extra = " active" if day == current_day else ""
    overview_cards.append(
        f'<div class="overview-card{extra}"><div class="overview-day">{day}</div><div class="overview-count">{count}</div><div class="overview-meta">marked students</div></div>'
    )

overview_tab, attendance_tab, reports_tab, settings_tab = st.tabs(
    ["Overview", "Attendance", "Reports", "Settings"]
)

with overview_tab:
    st.markdown('<div class="panel">', unsafe_allow_html=True)
    st.markdown('<div class="section-title">Weekly Overview</div>', unsafe_allow_html=True)
    st.markdown(
        '<div class="section-note">A quick snapshot of how much attendance has been marked across all five days.</div>',
        unsafe_allow_html=True,
    )
    st.markdown('<div class="overview-grid">' + "".join(overview_cards) + "</div>", unsafe_allow_html=True)
    st.markdown("</div>", unsafe_allow_html=True)

    total, present, absent, pending, completion = calculate_counts(
        st.session_state.attendance, students, current_day
    )
    left_stats, right_stats = st.columns([1.5, 1])
    with left_stats:
        st.markdown('<div class="panel">', unsafe_allow_html=True)
        st.markdown(
            f'<div class="section-title">{current_day} Snapshot</div>',
            unsafe_allow_html=True,
        )
        st.markdown(
            '<div class="section-note">These numbers reflect the currently selected day.</div>',
            unsafe_allow_html=True,
        )
        a, b = st.columns(2)
        c, d = st.columns(2)
        a.markdown(render_stat_card("Total", total, "stat-total"), unsafe_allow_html=True)
        b.markdown(render_stat_card("Present", present, "stat-present"), unsafe_allow_html=True)
        c.markdown(render_stat_card("Absent", absent, "stat-absent"), unsafe_allow_html=True)
        d.markdown(render_stat_card("Pending", pending, "stat-pending"), unsafe_allow_html=True)
        st.progress(completion / 100 if completion else 0, text=f"{completion}% of {current_day} completed")
        st.markdown("</div>", unsafe_allow_html=True)
    with right_stats:
        st.markdown('<div class="panel">', unsafe_allow_html=True)
        st.markdown('<div class="section-title">Teacher Notes</div>', unsafe_allow_html=True)
        st.markdown(
            '<div class="section-note">Use the Attendance tab for quick marking. Use the Reports tab when the day is complete.</div>',
            unsafe_allow_html=True,
        )
        st.info("Tip: The new Present and Absent checkbox columns are easier to click than the old dropdown.")
        st.success(f"Currently working on: {current_day}")
        st.markdown("</div>", unsafe_allow_html=True)

with attendance_tab:
    st.markdown('<div class="panel">', unsafe_allow_html=True)
    st.markdown(
        f'<div class="section-title">Mark Attendance for {current_day}</div>',
        unsafe_allow_html=True,
    )
    st.markdown(
        '<div class="section-note">Use search and quick actions, then click save when you are done. Check only one box per student.</div>',
        unsafe_allow_html=True,
    )

    filter_a, filter_b = st.columns([2, 1])
    search_text = filter_a.text_input("Search by roll number or student name")
    show_only_pending = filter_b.toggle("Show only pending", value=False)

    filtered_df = st.session_state[editor_key].copy()
    if search_text:
        query = search_text.strip().lower()
        mask = (
            filtered_df["Roll No"].str.lower().str.contains(query)
            | filtered_df["Student Name"].str.lower().str.contains(query)
        )
        filtered_df = filtered_df[mask]

    if show_only_pending:
        filtered_df = filtered_df[(~filtered_df["Present"]) & (~filtered_df["Absent"])]

    quick_a, quick_b, quick_c = st.columns(3)
    if quick_a.button("Mark Filtered Present", use_container_width=True):
        st.session_state[editor_key] = apply_bulk_choice(
            st.session_state[editor_key], filtered_df, "Present"
        )
        st.rerun()
    if quick_b.button("Mark Filtered Absent", use_container_width=True):
        st.session_state[editor_key] = apply_bulk_choice(
            st.session_state[editor_key], filtered_df, "Absent"
        )
        st.rerun()
    if quick_c.button("Clear Filtered", use_container_width=True):
        st.session_state[editor_key] = apply_bulk_choice(
            st.session_state[editor_key], filtered_df, "Pending"
        )
        st.rerun()

    edited_df = st.data_editor(
        filtered_df,
        use_container_width=True,
        hide_index=True,
        num_rows="fixed",
        column_config={
            "Roll No": st.column_config.TextColumn(disabled=True, width="small"),
            "Student Name": st.column_config.TextColumn(disabled=True, width="large"),
            "Present": st.column_config.CheckboxColumn("Present", width="small"),
            "Absent": st.column_config.CheckboxColumn("Absent", width="small"),
            "Marked Time": st.column_config.TextColumn(disabled=True, width="medium"),
        },
        key=f"data_editor_{current_day}_{len(filtered_df)}",
    )

    merged_df = st.session_state[editor_key].copy()
    for _, row in edited_df.iterrows():
        row_mask = merged_df["Roll No"] == row["Roll No"]
        merged_df.loc[row_mask, "Present"] = bool(row["Present"])
        merged_df.loc[row_mask, "Absent"] = bool(row["Absent"])
        if row["Present"] or row["Absent"]:
            if merged_df.loc[row_mask, "Marked Time"].iloc[0] == "":
                merged_df.loc[row_mask, "Marked Time"] = datetime.now().strftime("%I:%M:%S %p")
        else:
            merged_df.loc[row_mask, "Marked Time"] = ""

    st.session_state[editor_key] = merged_df

    invalid_rolls = validate_editor(st.session_state[editor_key])
    if invalid_rolls:
        st.error("Only one checkbox can be selected. Fix these roll numbers: " + ", ".join(invalid_rolls[:10]))

    if st.button(f"Save {current_day} Attendance", type="primary", use_container_width=True):
        invalid_rolls = save_day_from_editor(current_day, st.session_state[editor_key])
        if invalid_rolls:
            st.error("Save blocked. Uncheck either Present or Absent for: " + ", ".join(invalid_rolls[:10]))
        else:
            st.session_state[editor_key] = build_day_dataframe(
                students, st.session_state.attendance[current_day]
            )
            st.success(f"{current_day} attendance saved.")
            st.rerun()

    st.markdown("</div>", unsafe_allow_html=True)

with reports_tab:
    st.markdown('<div class="panel">', unsafe_allow_html=True)
    st.markdown('<div class="section-title">Reports</div>', unsafe_allow_html=True)
    st.markdown(
        f'<div class="section-note">Download the report for {current_day}. Pending students are included automatically.</div>',
        unsafe_allow_html=True,
    )

    report_df = pd.DataFrame(
        build_report_rows(students, st.session_state.attendance, current_day)
    )
    marked_rows = report_df[report_df["Status"] != "Pending"].copy()
    if not marked_rows.empty:
        st.dataframe(marked_rows, use_container_width=True, hide_index=True)
    else:
        st.info("No saved attendance for this day yet.")

    csv_data = report_df.to_csv(index=False).encode("utf-8")
    st.download_button(
        f"Download {current_day} Report",
        data=csv_data,
        file_name=f"attendance_{current_day.lower()}.csv",
        mime="text/csv",
        use_container_width=True,
    )
    st.markdown("</div>", unsafe_allow_html=True)

with settings_tab:
    st.markdown('<div class="panel">', unsafe_allow_html=True)
    st.markdown('<div class="section-title">Quick Actions</div>', unsafe_allow_html=True)
    st.markdown(
        '<div class="section-note">Reset only the selected day when you want to start that day again.</div>',
        unsafe_allow_html=True,
    )
    st.warning(f"This will clear only {current_day}. Other days will stay unchanged.")
    if st.button("Reset Current Day", use_container_width=True):
        st.session_state.attendance[current_day] = {}
        st.session_state[editor_key] = build_day_dataframe(students, {})
        st.rerun()
    st.info("The main controls now live on the page itself, so the teacher does not need to work from the sidebar.")
    st.markdown("</div>", unsafe_allow_html=True)
