from __future__ import annotations

import os
import tempfile
from io import BytesIO
from datetime import date, datetime, timedelta
from pathlib import Path

import gspread
import pandas as pd
import streamlit as st
from google.oauth2.service_account import Credentials
from gspread.exceptions import SpreadsheetNotFound, WorksheetNotFound
from streamlit.errors import StreamlitSecretNotFoundError
from openpyxl import Workbook, load_workbook
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.worksheet.datavalidation import DataValidation


APP_DIR = Path(__file__).resolve().parent
EXCEL_FILE = APP_DIR / "Homework_Tracker_M1_15.xlsx"
DEFAULT_GOOGLE_SHEET_ID = "1JnxohQh-anm6HFJZudMOLHvnDl8osa-j2V7ynQhYQb0"
STATUSES = ["ยังไม่ส่ง", "ส่งแล้ว", "ส่งช้า", "ยกเว้น"]
STUDENT_COLUMNS = ["เลขที่", "รหัสนักเรียน", "ชื่อ-นามสกุล"]
HOMEWORK_COLUMNS = ["HW_ID", "วิชา", "รายละเอียดงาน", "วันที่สั่ง", "กำหนดส่ง", "หมายเหตุ"]
TRACKING_COLUMNS = ["HW_ID", "เลขที่", "รหัสนักเรียน", "ชื่อ-นามสกุล", "สถานะ", "วันที่ส่ง", "หมายเหตุ"]
SHEETS = {
    "Students": (STUDENT_COLUMNS, 3),
    "Homework": (HOMEWORK_COLUMNS, 6),
    "Tracking": (TRACKING_COLUMNS, 7),
}


def _style_sheet(ws, color: str, widths: list[int]) -> None:
    ws.freeze_panes = "A2"
    ws.sheet_view.showGridLines = False
    fill = PatternFill("solid", fgColor=color)
    for cell in ws[1]:
        cell.fill = fill
        cell.font = Font(color="FFFFFF", bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center")
    for index, width in enumerate(widths, 1):
        ws.column_dimensions[chr(64 + index)].width = width
    ws.auto_filter.ref = ws.dimensions


def _normalize_workbook(path: Path) -> None:
    """Add the required sheets/columns while preserving an existing workbook."""
    if path.exists():
        wb = load_workbook(path)
    else:
        wb = Workbook()
        wb.remove(wb.active)

    definitions = {
        "Students": (STUDENT_COLUMNS, "2563EB", [10, 18, 34]),
        "Homework": (HOMEWORK_COLUMNS, "0F766E", [16, 20, 42, 16, 16, 30]),
        "Tracking": (TRACKING_COLUMNS, "7C3AED", [16, 10, 18, 34, 16, 16, 30]),
    }
    for name, (headers, color, widths) in definitions.items():
        ws = wb[name] if name in wb.sheetnames else wb.create_sheet(name)
        current = [ws.cell(1, i).value for i in range(1, len(headers) + 1)]
        if not any(current):
            for i, header in enumerate(headers, 1):
                ws.cell(1, i, header)
        elif current != headers:
            missing = [h for h in headers if h not in current]
            for header in missing:
                ws.cell(1, ws.max_column + 1, header)
        _style_sheet(ws, color, widths)

    students = wb["Students"]
    if students.max_row <= 1:
        for number in range(1, 41):
            students.append([number, "", f"นักเรียนเลขที่ {number}"])

    homework = wb["Homework"]
    for col in (4, 5):
        for row in range(2, max(homework.max_row, 2) + 1):
            homework.cell(row, col).number_format = "yyyy-mm-dd"
    tracking = wb["Tracking"]
    has_status_validation = any(str(item.sqref) == "E2:E10000" for item in tracking.data_validations.dataValidation)
    if not has_status_validation:
        validation = DataValidation(type="list", formula1='"' + ",".join(STATUSES) + '"', allow_blank=False)
        tracking.add_data_validation(validation)
        validation.add("E2:E10000")
    if len(tracking.conditional_formatting) == 0:
        tracking.conditional_formatting.add("E2:E10000", FormulaRule(formula=['E2="ยังไม่ส่ง"'], fill=PatternFill("solid", fgColor="FEE2E2")))
        tracking.conditional_formatting.add("E2:E10000", FormulaRule(formula=['OR(E2="ส่งแล้ว",E2="ส่งช้า")'], fill=PatternFill("solid", fgColor="DCFCE7")))
    _atomic_save(wb, path)


def _atomic_save(wb, path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    fd, temp_name = tempfile.mkstemp(prefix="homework_", suffix=".xlsx", dir=path.parent)
    os.close(fd)
    try:
        wb.save(temp_name)
        os.replace(temp_name, path)
    finally:
        if os.path.exists(temp_name):
            os.unlink(temp_name)


def load_excel_data() -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    _normalize_workbook(EXCEL_FILE)
    try:
        students = pd.read_excel(EXCEL_FILE, sheet_name="Students", dtype={"รหัสนักเรียน": str})
        homework = pd.read_excel(EXCEL_FILE, sheet_name="Homework", dtype={"HW_ID": str})
        tracking = pd.read_excel(EXCEL_FILE, sheet_name="Tracking", dtype={"HW_ID": str, "รหัสนักเรียน": str})
    except PermissionError as exc:
        raise RuntimeError("กรุณาปิดไฟล์ Excel ก่อนบันทึกหรือรีเฟรชข้อมูล") from exc
    students = students.reindex(columns=STUDENT_COLUMNS).fillna("")
    homework = homework.reindex(columns=HOMEWORK_COLUMNS)
    tracking = tracking.reindex(columns=TRACKING_COLUMNS)
    for frame, columns in ((homework, ["วันที่สั่ง", "กำหนดส่ง"]), (tracking, ["วันที่ส่ง"])):
        for column in columns:
            frame[column] = pd.to_datetime(frame[column], errors="coerce")
    if not tracking.empty:
        tracking["สถานะ"] = tracking["สถานะ"].fillna("ยังไม่ส่ง")
    return students, homework, tracking


def save_excel_frames(students: pd.DataFrame, homework: pd.DataFrame, tracking: pd.DataFrame) -> None:
    _normalize_workbook(EXCEL_FILE)
    wb = load_workbook(EXCEL_FILE)
    for sheet_name, frame, columns in (
        ("Students", students, STUDENT_COLUMNS),
        ("Homework", homework, HOMEWORK_COLUMNS),
        ("Tracking", tracking, TRACKING_COLUMNS),
    ):
        ws = wb[sheet_name]
        if ws.max_row > 1:
            ws.delete_rows(2, ws.max_row - 1)
        clean = frame.reindex(columns=columns).copy()
        clean = clean.where(pd.notna(clean), None)
        for row in clean.itertuples(index=False, name=None):
            ws.append([value.to_pydatetime() if isinstance(value, pd.Timestamp) else value for value in row])
    _style_sheet(wb["Students"], "2563EB", [10, 18, 34])
    _style_sheet(wb["Homework"], "0F766E", [16, 20, 42, 16, 16, 30])
    _style_sheet(wb["Tracking"], "7C3AED", [16, 10, 18, 34, 16, 16, 30])
    _atomic_save(wb, EXCEL_FILE)


def secret_value(key: str, default=None):
    try:
        return st.secrets[key]
    except (KeyError, FileNotFoundError, StreamlitSecretNotFoundError):
        return default


def google_sheets_configured() -> bool:
    return secret_value("gcp_service_account") is not None and bool(secret_value("GOOGLE_SHEET_ID", DEFAULT_GOOGLE_SHEET_ID))


@st.cache_resource(show_spinner=False)
def get_google_spreadsheet():
    info = dict(secret_value("gcp_service_account"))
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    credentials = Credentials.from_service_account_info(info, scopes=scopes)
    client = gspread.authorize(credentials)
    try:
        return client.open_by_key(secret_value("GOOGLE_SHEET_ID", DEFAULT_GOOGLE_SHEET_ID))
    except SpreadsheetNotFound as exc:
        raise RuntimeError("เปิด Google Sheet ไม่ได้ กรุณาตรวจ GOOGLE_SHEET_ID และแชร์ชีตให้อีเมล Service Account") from exc


def _ensure_google_worksheet(spreadsheet, name: str, columns: list[str], min_rows: int):
    try:
        worksheet = spreadsheet.worksheet(name)
    except WorksheetNotFound:
        worksheet = spreadsheet.add_worksheet(title=name, rows=min_rows, cols=len(columns))
    values = worksheet.get_all_values()
    if not values:
        worksheet.update([columns], "A1")
    elif values[0][: len(columns)] != columns:
        worksheet.update([columns], "A1")
    return worksheet


def _google_values_to_frame(values: list[list[str]], columns: list[str]) -> pd.DataFrame:
    if len(values) <= 1:
        return pd.DataFrame(columns=columns)
    rows = [row[: len(columns)] + [""] * max(0, len(columns) - len(row)) for row in values[1:]]
    return pd.DataFrame(rows, columns=columns)


def load_google_data() -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    spreadsheet = get_google_spreadsheet()
    frames = {}
    for name, (columns, column_count) in SHEETS.items():
        min_rows = 200 if name != "Tracking" else 10000
        worksheet = _ensure_google_worksheet(spreadsheet, name, columns, min_rows)
        frames[name] = _google_values_to_frame(worksheet.get_all_values(), columns)

    # On the first cloud launch, migrate the bundled Excel template/data automatically.
    if all(frame.empty for frame in frames.values()) and EXCEL_FILE.exists():
        excel_students, excel_homework, excel_tracking = load_excel_data()
        frames = {"Students": excel_students, "Homework": excel_homework, "Tracking": excel_tracking}
        for name, frame, columns in (
            ("Students", excel_students, STUDENT_COLUMNS),
            ("Homework", excel_homework, HOMEWORK_COLUMNS),
            ("Tracking", excel_tracking, TRACKING_COLUMNS),
        ):
            _write_google_worksheet(spreadsheet.worksheet(name), frame, columns)

    students = frames["Students"]
    if students.empty:
        students = pd.DataFrame([[number, "", f"นักเรียนเลขที่ {number}"] for number in range(1, 41)], columns=STUDENT_COLUMNS)
        _write_google_worksheet(spreadsheet.worksheet("Students"), students, STUDENT_COLUMNS)
    students["เลขที่"] = pd.to_numeric(students["เลขที่"], errors="coerce")
    homework = frames["Homework"]
    tracking = frames["Tracking"]
    for frame, columns in ((homework, ["วันที่สั่ง", "กำหนดส่ง"]), (tracking, ["วันที่ส่ง"])):
        for column in columns:
            frame[column] = pd.to_datetime(frame[column], errors="coerce")
    if not tracking.empty:
        tracking["เลขที่"] = pd.to_numeric(tracking["เลขที่"], errors="coerce")
        tracking["สถานะ"] = tracking["สถานะ"].replace("", "ยังไม่ส่ง").fillna("ยังไม่ส่ง")
    return students.fillna(""), homework, tracking


def _value_for_google(value):
    if value is None or pd.isna(value):
        return ""
    if isinstance(value, (pd.Timestamp, datetime, date)):
        return pd.Timestamp(value).strftime("%Y-%m-%d")
    if hasattr(value, "item"):
        value = value.item()
    return value


def _write_google_worksheet(worksheet, frame: pd.DataFrame, columns: list[str]) -> None:
    clean = frame.reindex(columns=columns)
    values = [columns] + [[_value_for_google(value) for value in row] for row in clean.itertuples(index=False, name=None)]
    worksheet.clear()
    worksheet.update(values, "A1")
    worksheet.freeze(rows=1)
    worksheet.format("1:1", {"backgroundColor": {"red": 0.12, "green": 0.36, "blue": 0.72}, "textFormat": {"foregroundColor": {"red": 1, "green": 1, "blue": 1}, "bold": True}, "horizontalAlignment": "CENTER"})


def save_google_frames(students: pd.DataFrame, homework: pd.DataFrame, tracking: pd.DataFrame) -> None:
    spreadsheet = get_google_spreadsheet()
    for name, frame, columns in (
        ("Students", students, STUDENT_COLUMNS),
        ("Homework", homework, HOMEWORK_COLUMNS),
        ("Tracking", tracking, TRACKING_COLUMNS),
    ):
        min_rows = max(len(frame) + 20, 200 if name != "Tracking" else 10000)
        worksheet = _ensure_google_worksheet(spreadsheet, name, columns, min_rows)
        _write_google_worksheet(worksheet, frame, columns)


def load_data() -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    if google_sheets_configured():
        return load_google_data()
    return load_excel_data()


def save_frames(students: pd.DataFrame, homework: pd.DataFrame, tracking: pd.DataFrame) -> None:
    if google_sheets_configured():
        save_google_frames(students, homework, tracking)
    else:
        save_excel_frames(students, homework, tracking)


def frames_to_excel(students: pd.DataFrame, homework: pd.DataFrame, tracking: pd.DataFrame) -> bytes:
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        students.to_excel(writer, sheet_name="Students", index=False)
        homework.to_excel(writer, sheet_name="Homework", index=False)
        tracking.to_excel(writer, sheet_name="Tracking", index=False)
    return buffer.getvalue()


def require_login() -> None:
    configured_password = str(secret_value("APP_PASSWORD", "")).strip()
    if not configured_password:
        return
    if st.session_state.get("authenticated"):
        return
    st.title("🔐 Homework Tracker ม.1/15")
    with st.form("login_form"):
        password = st.text_input("รหัสผ่าน", type="password")
        login = st.form_submit_button("เข้าสู่ระบบ", type="primary", use_container_width=True)
    if login:
        if password == configured_password:
            st.session_state["authenticated"] = True
            st.rerun()
        else:
            st.error("รหัสผ่านไม่ถูกต้อง")
    st.stop()


def next_hw_id(homework: pd.DataFrame) -> str:
    prefix = datetime.now().strftime("HW-%Y%m%d-")
    used = set(homework["HW_ID"].dropna().astype(str)) if not homework.empty else set()
    number = 1
    while f"{prefix}{number:03d}" in used:
        number += 1
    return f"{prefix}{number:03d}"


def thai_date(value) -> str:
    if pd.isna(value):
        return "-"
    dt = pd.Timestamp(value)
    months = ["ม.ค.", "ก.พ.", "มี.ค.", "เม.ย.", "พ.ค.", "มิ.ย.", "ก.ค.", "ส.ค.", "ก.ย.", "ต.ค.", "พ.ย.", "ธ.ค."]
    return f"{dt.day} {months[dt.month - 1]} {dt.year + 543}"


def homework_label(row: pd.Series) -> str:
    return f"{row['HW_ID']} · {row['วิชา']} · {row['รายละเอียดงาน']}"


st.set_page_config(page_title="Homework Tracker ม.1/15", page_icon="📚", layout="wide", initial_sidebar_state="collapsed")
st.markdown("""
<style>
  .block-container {padding-top: 1.2rem; padding-bottom: 3rem; max-width: 1200px;}
  [data-testid="stMetric"] {background:#f8fafc; border:1px solid #e2e8f0; padding:12px; border-radius:14px;}
  div[data-testid="stForm"] {border:1px solid #e2e8f0; border-radius:16px; padding:1rem;}
  @media (max-width: 640px) {.block-container {padding-left:.75rem; padding-right:.75rem;} h1 {font-size:1.65rem !important;}}
</style>
""", unsafe_allow_html=True)

require_login()
st.title("📚 Homework Tracker ม.1/15")
pages = ["Dashboard", "เพิ่มการบ้าน", "อัปเดตสถานะส่งงาน", "สรุปข้อความส่ง LINE", "รายชื่อนักเรียน"]
page = st.sidebar.radio("เมนู", pages)
storage_name = "Google Sheets ☁️" if google_sheets_configured() else f"Excel: {EXCEL_FILE.name}"
st.sidebar.caption(f"ฐานข้อมูล: {storage_name}")
with st.expander("💡 วิธีใช้งานฉบับย่อ"):
    st.markdown("""
1. **รายชื่อนักเรียน** — แก้ชื่อและรหัสของนักเรียนทั้ง 40 คน แล้วกดบันทึก
2. **เพิ่มการบ้าน** — กรอกวิชา งาน และกำหนดส่ง ระบบจะสร้างรายการติดตามให้นักเรียนทุกคน
3. **อัปเดตสถานะส่งงาน** — เลือกงาน แก้สถานะ/วันที่ส่ง แล้วกดบันทึก (ส่งเกินกำหนดจะเปลี่ยนเป็น “ส่งช้า” อัตโนมัติ)
4. **Dashboard** — ดูจำนวนส่งแล้ว ค้างส่ง และรายชื่อค้างส่งของทุกงาน
5. **สรุปข้อความส่ง LINE** — เลือกงาน คัดลอกข้อความจากกล่อง หรือดาวน์โหลดไฟล์ข้อความ

> ควรปิดไฟล์ Excel ก่อนกดบันทึกในเว็บ เพราะ Windows อาจล็อกไฟล์ไว้
""")

try:
    students_df, homework_df, tracking_df = load_data()
except Exception as exc:
    st.error(f"เปิดข้อมูลไม่ได้: {exc}")
    st.stop()

if google_sheets_configured():
    st.sidebar.download_button(
        "ดาวน์โหลดข้อมูลสำรอง Excel",
        data=frames_to_excel(students_df, homework_df, tracking_df),
        file_name=f"Homework_Tracker_M1_15_backup_{date.today().isoformat()}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

if page == "Dashboard":
    st.subheader("ภาพรวมการส่งงาน")
    today = pd.Timestamp(date.today())
    due = pd.to_datetime(homework_df["กำหนดส่ง"], errors="coerce") if not homework_df.empty else pd.Series(dtype="datetime64[ns]")
    near_due = int(((due >= today) & (due <= today + pd.Timedelta(days=7))).sum())
    c1, c2, c3 = st.columns(3)
    c1.metric("งานทั้งหมด", len(homework_df))
    c2.metric("ใกล้ครบกำหนด (7 วัน)", near_due)
    pending_total = int((tracking_df["สถานะ"] == "ยังไม่ส่ง").sum()) if not tracking_df.empty else 0
    c3.metric("รายการค้างส่ง", pending_total)
    if homework_df.empty:
        st.info("ยังไม่มีการบ้าน เลือกเมนู “เพิ่มการบ้าน” เพื่อเริ่มต้น")
    else:
        summary_rows = []
        for _, hw in homework_df.sort_values("กำหนดส่ง", na_position="last").iterrows():
            rows = tracking_df[tracking_df["HW_ID"].astype(str) == str(hw["HW_ID"])]
            sent = int(rows["สถานะ"].isin(["ส่งแล้ว", "ส่งช้า"]).sum())
            pending = rows[rows["สถานะ"] == "ยังไม่ส่ง"]
            names = ", ".join(f"{int(r['เลขที่'])}. {r['ชื่อ-นามสกุล']}" for _, r in pending.iterrows())
            summary_rows.append({"HW_ID": hw["HW_ID"], "วิชา": hw["วิชา"], "งาน": hw["รายละเอียดงาน"], "กำหนดส่ง": hw["กำหนดส่ง"], "ส่งแล้ว": sent, "ค้างส่ง": len(pending), "รายชื่อค้างส่ง": names or "-"})
        st.dataframe(pd.DataFrame(summary_rows), use_container_width=True, hide_index=True, column_config={"กำหนดส่ง": st.column_config.DateColumn(format="DD/MM/YYYY"), "รายชื่อค้างส่ง": st.column_config.TextColumn(width="large")})

elif page == "เพิ่มการบ้าน":
    st.subheader("เพิ่มการบ้านใหม่")
    with st.form("add_homework", clear_on_submit=True):
        hw_id = st.text_input("HW_ID", value=next_hw_id(homework_df))
        subject = st.text_input("วิชา *")
        detail = st.text_area("รายละเอียดงาน *")
        c1, c2 = st.columns(2)
        assigned = c1.date_input("วันที่สั่ง", value=date.today())
        due_date = c2.date_input("กำหนดส่ง", value=date.today() + timedelta(days=7))
        note = st.text_area("หมายเหตุ")
        submitted = st.form_submit_button("เพิ่มการบ้านและสร้างรายการ 40 คน", type="primary", use_container_width=True)
    if submitted:
        errors = []
        if not hw_id.strip() or not subject.strip() or not detail.strip(): errors.append("กรุณากรอก HW_ID วิชา และรายละเอียดงาน")
        if hw_id.strip() in set(homework_df["HW_ID"].astype(str)): errors.append("HW_ID นี้มีอยู่แล้ว")
        if due_date < assigned: errors.append("กำหนดส่งต้องไม่ก่อนวันที่สั่ง")
        if errors:
            st.error(" · ".join(errors))
        else:
            new_hw = pd.DataFrame([[hw_id.strip(), subject.strip(), detail.strip(), pd.Timestamp(assigned), pd.Timestamp(due_date), note.strip()]], columns=HOMEWORK_COLUMNS)
            homework_df = pd.concat([homework_df, new_hw], ignore_index=True)
            active_students = students_df[students_df["ชื่อ-นามสกุล"].astype(str).str.strip().ne("")].copy()
            new_tracking = pd.DataFrame({"HW_ID": hw_id.strip(), "เลขที่": active_students["เลขที่"], "รหัสนักเรียน": active_students["รหัสนักเรียน"], "ชื่อ-นามสกุล": active_students["ชื่อ-นามสกุล"], "สถานะ": "ยังไม่ส่ง", "วันที่ส่ง": pd.NaT, "หมายเหตุ": ""})
            try:
                save_frames(students_df, homework_df, pd.concat([tracking_df, new_tracking], ignore_index=True))
                st.success(f"เพิ่ม {hw_id} และสร้างรายการติดตาม {len(new_tracking)} คนแล้ว")
                st.rerun()
            except PermissionError:
                st.error("บันทึกไม่ได้ กรุณาปิดไฟล์ Excel แล้วลองใหม่")

elif page == "อัปเดตสถานะส่งงาน":
    st.subheader("อัปเดตสถานะส่งงาน")
    if homework_df.empty:
        st.info("ยังไม่มีการบ้าน")
    else:
        options = {homework_label(row): str(row["HW_ID"]) for _, row in homework_df.iloc[::-1].iterrows()}
        selected = st.selectbox("เลือกการบ้าน", list(options))
        hw_id = options[selected]
        mask = tracking_df["HW_ID"].astype(str) == hw_id
        subset = tracking_df.loc[mask, TRACKING_COLUMNS].sort_values("เลขที่").reset_index(drop=True)
        edited = st.data_editor(subset, use_container_width=True, hide_index=True, disabled=["HW_ID", "เลขที่", "รหัสนักเรียน", "ชื่อ-นามสกุล"], column_config={"สถานะ": st.column_config.SelectboxColumn(options=STATUSES, required=True), "วันที่ส่ง": st.column_config.DateColumn(format="DD/MM/YYYY"), "หมายเหตุ": st.column_config.TextColumn(width="medium")}, key=f"editor-{hw_id}")
        if st.button("บันทึกสถานะ", type="primary", use_container_width=True):
            due_value = homework_df.loc[homework_df["HW_ID"].astype(str) == hw_id, "กำหนดส่ง"].iloc[0]
            edited["วันที่ส่ง"] = pd.to_datetime(edited["วันที่ส่ง"], errors="coerce")
            auto_late = (edited["สถานะ"] == "ส่งแล้ว") & edited["วันที่ส่ง"].notna() & (edited["วันที่ส่ง"] > pd.Timestamp(due_value))
            edited.loc[auto_late, "สถานะ"] = "ส่งช้า"
            tracking_df = pd.concat([tracking_df.loc[~mask], edited], ignore_index=True)
            save_frames(students_df, homework_df, tracking_df)
            st.success("บันทึกสถานะแล้ว")
            st.rerun()

elif page == "สรุปข้อความส่ง LINE":
    st.subheader("สร้างข้อความสรุปสำหรับ LINE")
    if homework_df.empty:
        st.info("ยังไม่มีการบ้าน")
    else:
        options = {homework_label(row): str(row["HW_ID"]) for _, row in homework_df.iloc[::-1].iterrows()}
        selected = st.selectbox("เลือกการบ้าน", list(options), key="line_hw")
        hw_id = options[selected]
        hw = homework_df[homework_df["HW_ID"].astype(str) == hw_id].iloc[0]
        rows = tracking_df[tracking_df["HW_ID"].astype(str) == hw_id].sort_values("เลขที่")
        sent = int(rows["สถานะ"].isin(["ส่งแล้ว", "ส่งช้า"]).sum())
        pending = rows[rows["สถานะ"] == "ยังไม่ส่ง"]
        pending_lines = [f"{i}. เลขที่ {int(row['เลขที่'])} {row['ชื่อ-นามสกุล']}" for i, (_, row) in enumerate(pending.iterrows(), 1)]
        message = "\n".join(["สรุปการบ้าน ม.1/15", f"วิชา: {hw['วิชา']}", f"งาน: {hw['รายละเอียดงาน']}", f"กำหนดส่ง: {thai_date(hw['กำหนดส่ง'])}", "", f"ส่งแล้ว: {sent} คน", f"ค้างส่ง: {len(pending)} คน", "รายชื่อค้างส่ง:", *(pending_lines or ["ไม่มี 🎉"])])
        st.text_area("ข้อความพร้อมคัดลอก", value=message, height=360)
        st.download_button("ดาวน์โหลดเป็นไฟล์ข้อความ", message.encode("utf-8-sig"), file_name=f"LINE_{hw_id}.txt", mime="text/plain", use_container_width=True)

else:
    st.subheader("รายชื่อนักเรียน 40 คน")
    st.caption("แก้ชื่อและรหัสนักเรียนได้โดยตรง เลขที่ต้องไม่ซ้ำกัน")
    edited_students = st.data_editor(students_df, use_container_width=True, hide_index=True, num_rows="fixed", column_config={"เลขที่": st.column_config.NumberColumn(min_value=1, max_value=40, step=1, required=True), "รหัสนักเรียน": st.column_config.TextColumn(), "ชื่อ-นามสกุล": st.column_config.TextColumn(required=True)}, key="students_editor")
    if st.button("บันทึกรายชื่อนักเรียน", type="primary", use_container_width=True):
        numbers = pd.to_numeric(edited_students["เลขที่"], errors="coerce")
        if len(edited_students) != 40 or numbers.isna().any() or numbers.duplicated().any():
            st.error("ต้องมีนักเรียน 40 คน และเลขที่ 1–40 ต้องไม่ซ้ำกัน")
        else:
            edited_students["เลขที่"] = numbers.astype(int)
            # Keep names in existing tracking rows synchronized by student number.
            student_map = edited_students.set_index("เลขที่")
            if not tracking_df.empty:
                for idx, row in tracking_df.iterrows():
                    number = int(row["เลขที่"])
                    if number in student_map.index:
                        tracking_df.at[idx, "รหัสนักเรียน"] = student_map.at[number, "รหัสนักเรียน"]
                        tracking_df.at[idx, "ชื่อ-นามสกุล"] = student_map.at[number, "ชื่อ-นามสกุล"]
            save_frames(edited_students.sort_values("เลขที่"), homework_df, tracking_df)
            st.success("บันทึกรายชื่อแล้ว")
            st.rerun()
