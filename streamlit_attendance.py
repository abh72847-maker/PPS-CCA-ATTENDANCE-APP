import sqlite3
from datetime import datetime
from io import BytesIO
from pathlib import Path
import zipfile
import xml.etree.ElementTree as ET

import pandas as pd
import streamlit as st

# --- CONFIGURATION ---
APP_DIR = Path(__file__).resolve().parent
DEFAULT_FILE = APP_DIR / "student list.xlsx"
DB_FILE = APP_DIR / "attendance.db"
STATUS_OPTIONS = ["Pending", "Present", "Absent", "Late", "Excused"]

st.set_page_config(page_title="Attendance Pro", page_icon="📝", layout="wide")

# --- DATABASE SETUP ---
def init_db():
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    c.execute('''
        CREATE TABLE IF NOT EXISTS attendance (
            date TEXT,
            roll TEXT,
            name TEXT,
            status TEXT,
            time TEXT,
            UNIQUE(date, roll)
        )
    ''')
    conn.commit()
    return conn

# --- DATA LOADING (Custom XML Reader to avoid extra dependencies) ---
@st.cache_data
def load_students_from_excel(file_obj):
    ns = {"a": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
    students = []
    
    if hasattr(file_obj, "read"):
        content = file_obj.read()
        file_obj = BytesIO(content)

    try:
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
    except Exception as e:
        st.error(f"Error reading Excel file: {e}")
    return students

# --- DATABASE QUERIES ---
def get_attendance_for_date(conn, target_date):
    query = "SELECT roll, status, time FROM attendance WHERE date = ?"
    df = pd.read_sql_query(query, conn, params=(target_date,))
    if df.empty:
        return {}
    return df.set_index('roll').to_dict('index')

def save_attendance_to_db(conn, target_date, final_df):
    records = []
    for _, row in final_df.iterrows():
        if row["Status"] != "Pending":
            records.append((
                target_date, 
                row["Roll No"], 
                row["Student Name"], 
                row["Status"], 
                row["Marked Time"]
            ))
    
    c = conn.cursor()
    # Insert or update based on the UNIQUE(date, roll) constraint
    c.executemany('''
        INSERT OR REPLACE INTO attendance (date, roll, name, status, time)
        VALUES (?, ?, ?, ?, ?)
    ''', records)
    conn.commit()

# --- DATAFRAME BUILDERS ---
def build_day_dataframe(students, db_records):
    rows = []
    for student in students:
        record = db_records.get(student["roll"], {})
        rows.append({
            "Roll No": student["roll"],
            "Student Name": student["name"],
            "Status": record.get("status", "Pending"),
            "Marked Time": record.get("time", "")
        })
    return pd.DataFrame(rows)

def apply_bulk_choice(base_df, filtered_df, choice):
    updated_df = base_df.copy()
    rolls = set(filtered_df["Roll No"].tolist())
    mask = updated_df["Roll No"].isin(rolls)
    
    updated_df.loc[mask, "Status"] = choice
    if choice != "Pending":
        # Only update time if it was previously empty or Pending
        empty_times = mask & (updated_df["Marked Time"] == "")
        updated_df.loc[empty_times, "Marked Time"] = datetime.now().strftime("%I:%M:%S %p")
    else:
        updated_df.loc[mask, "Marked Time"] = ""
        
    return updated_df

# --- MAIN APP ---
def main():
    conn = init_db()

    # 1. Check for Student List
    if not DEFAULT_FILE.exists():
        st.error(f"Could not find `{DEFAULT_FILE.name}` in {APP_DIR}.")
        st.stop()
        
    students = load_students_from_excel(DEFAULT_FILE)
    if not students:
        st.warning("Student list is empty or could not be read.")
        st.stop()

    # 2. Header & Date Selection
    st.title("🎓 Attendance Dashboard Pro")
    
    col1, col2 = st.columns([1, 3])
    with col1:
        selected_date_obj = st.date_input("Select Date", datetime.today())
        selected_date = selected_date_obj.strftime("%Y-%m-%d")
    with col2:
        st.info(f"**{len(students)}** students loaded from `{DEFAULT_FILE.name}`. Data auto-saves to local database.")

    # 3. Load State for Selected Date
    db_records = get_attendance_for_date(conn, selected_date)
    editor_key = f"df_{selected_date}"
    
    if editor_key not in st.session_state:
        st.session_state[editor_key] = build_day_dataframe(students, db_records)

    # 4. Tabs
    tab_attend, tab_dash, tab_report, tab_settings = st.tabs([
        "📝 Mark Attendance", 
        "📊 Dashboard", 
        "📁 Reports & Export", 
        "⚙️ Settings"
    ])

    # --- TAB 1: ATTENDANCE ---
    with tab_attend:
        st.subheader(f"Attendance for {selected_date_obj.strftime('%A, %b %d')}")
        
        # Filters
        f_col1, f_col2 = st.columns([3, 1])
        search_query = f_col1.text_input("🔍 Search Student (Name or Roll No)", key="search")
        show_pending = f_col2.toggle("Show Pending Only", value=False)
        
        # Apply Filters
        current_df = st.session_state[editor_key].copy()
        if search_query:
            query = search_query.lower()
            current_df = current_df[
                current_df["Student Name"].str.lower().str.contains(query) | 
                current_df["Roll No"].str.lower().str.contains(query)
            ]
        if show_pending:
            current_df = current_df[current_df["Status"] == "Pending"]

        # Bulk Actions
        b1, b2, b3, b4 = st.columns(4)
        if b1.button("✔️ Mark Filtered Present", use_container_width=True):
            st.session_state[editor_key] = apply_bulk_choice(st.session_state[editor_key], current_df, "Present")
            st.rerun()
        if b2.button("❌ Mark Filtered Absent", use_container_width=True):
            st.session_state[editor_key] = apply_bulk_choice(st.session_state[editor_key], current_df, "Absent")
            st.rerun()
        if b3.button("⏳ Mark Filtered Late", use_container_width=True):
            st.session_state[editor_key] = apply_bulk_choice(st.session_state[editor_key], current_df, "Late")
            st.rerun()
        if b4.button("🔄 Clear Filtered", use_container_width=True):
            st.session_state[editor_key] = apply_bulk_choice(st.session_state[editor_key], current_df, "Pending")
            st.rerun()

        # Data Editor (Now with Dropdowns and Vectorized processing)
        edited_df = st.data_editor(
            current_df,
            use_container_width=True,
            hide_index=True,
            column_config={
                "Roll No": st.column_config.TextColumn("Roll No", disabled=True),
                "Student Name": st.column_config.TextColumn("Student Name", disabled=True),
                "Status": st.column_config.SelectboxColumn("Status", options=STATUS_OPTIONS, required=True),
                "Marked Time": st.column_config.TextColumn("Time", disabled=True),
            },
            key=f"editor_{selected_date}_{len(current_df)}"
        )

        # FAST Vectorized Time update
        merged_df = st.session_state[editor_key].copy()
        merged_df.update(edited_df)
        
        # If status changed from Pending, add time. If changed to Pending, clear time.
        needs_time = (merged_df["Status"] != "Pending") & (merged_df["Marked Time"] == "")
        if needs_time.any():
            merged_df.loc[needs_time, "Marked Time"] = datetime.now().strftime("%I:%M:%S %p")
        
        clear_time = (merged_df["Status"] == "Pending") & (merged_df["Marked Time"] != "")
        if clear_time.any():
            merged_df.loc[clear_time, "Marked Time"] = ""

        st.session_state[editor_key] = merged_df

        # Save Action
        if st.button("💾 Save Attendance to Database", type="primary", use_container_width=True):
            save_attendance_to_db(conn, selected_date, st.session_state[editor_key])
            st.success(f"Successfully saved {selected_date} to database!")
            st.balloons()


    # --- TAB 2: DASHBOARD ---
    with tab_dash:
        st.subheader("Daily Snapshot")
        dash_df = st.session_state[editor_key]
        
        total = len(dash_df)
        present = len(dash_df[dash_df["Status"] == "Present"])
        absent = len(dash_df[dash_df["Status"] == "Absent"])
        late = len(dash_df[dash_df["Status"] == "Late"])
        excused = len(dash_df[dash_df["Status"] == "Excused"])
        pending = len(dash_df[dash_df["Status"] == "Pending"])
        
        # Native Streamlit metrics (Adapts perfectly to Dark/Light mode)
        m1, m2, m3, m4, m5 = st.columns(5)
        m1.metric("Total Students", total)
        m2.metric("Present", present)
        m3.metric("Absent", absent)
        m4.metric("Late / Excused", late + excused)
        m5.metric("Pending", pending)

        st.divider()
        
        # Visualizations
        st.subheader("Visual Analytics")
        if pending < total:
            status_counts = dash_df["Status"].value_counts().drop("Pending", errors="ignore")
            st.bar_chart(status_counts, color="#3b82f6")
        else:
            st.info("Mark some attendance to see the chart!")


    # --- TAB 3: REPORTS ---
    with tab_report:
        st.subheader("Export Attendance")
        export_df = st.session_state[editor_key].copy()
        
        st.dataframe(export_df, use_container_width=True, hide_index=True)
        
        c1, c2 = st.columns(2)
        
        # CSV Export
        csv = export_df.to_csv(index=False).encode('utf-8')
        c1.download_button(
            label="📄 Download as CSV",
            data=csv,
            file_name=f"Attendance_{selected_date}.csv",
            mime="text/csv",
            use_container_width=True
        )
        
        # Excel Export (Requires openpyxl)
        try:
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                export_df.to_excel(writer, index=False, sheet_name=selected_date)
            excel_data = output.getvalue()
            
            c2.download_button(
                label="📊 Download as Excel",
                data=excel_data,
                file_name=f"Attendance_{selected_date}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
        except ModuleNotFoundError:
            c2.warning("Install `openpyxl` via pip to enable Excel exports.")


    # --- TAB 4: SETTINGS ---
    with tab_settings:
        st.subheader("Database Management")
        st.warning("⚠️ Danger Zone")
        if st.button("Reset Today's Attendance"):
            # Clear from DB
            c = conn.cursor()
            c.execute("DELETE FROM attendance WHERE date = ?", (selected_date,))
            conn.commit()
            # Clear from session state
            del st.session_state[editor_key]
            st.success("Cleared! Reloading...")
            st.rerun()

if __name__ == "__main__":
    main()
