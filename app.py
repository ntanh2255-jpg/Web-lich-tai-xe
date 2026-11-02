
import datetime
import re
import textwrap
from pathlib import Path

import numpy as np
import pandas as pd
import requests
import streamlit as st


# ============================================================
# 1. CẤU HÌNH FILE GA
# ============================================================

BASE_DIR = Path(__file__).resolve().parent
SCHEDULE_FILE = BASE_DIR / "Lich_tai_xe_GA_30_ngay_can_bang.xlsx"
TIME_FILE = BASE_DIR / "Giờ đến trường GA.xlsx"

# Giữ kết nối Google Sheets để đồng bộ trạng thái thu gom giữa các thiết bị.
SCRIPT_URL = "https://script.google.com/macros/s/AKfycbzmPXp8-KTeiM9uCDMMshtKucCz1LoM269Dpc5hm-rWNFE_nKXV0uKgrkXQ020WmYCM/exec"
REQUEST_TIMEOUT = 15
BREAK_THRESHOLD = 20

st.set_page_config(page_title="Lịch làm việc tài xế - GA", page_icon="🚚", layout="wide")


# ============================================================
# 2. GIAO DIỆN
# ============================================================

st.markdown("""
<style>
:root{--blue:#2563eb;--blue-bg:#eff6ff;--orange:#d97706;--orange-bg:#fffbeb;--red:#dc2626;--red-bg:#fff1f2;--green:#15803d;--green-bg:#f0fdf4}
.block-container{max-width:1100px;padding-top:2.5rem!important;padding-bottom:3rem}
.stApp h1{line-height:1.4!important;padding-top:.15rem!important;overflow:visible!important}
[data-testid="stSidebar"]{min-width:260px;background:#f8fafc}
div[data-testid="stMetric"]{background:#fff;border:1px solid #e5e7eb;border-radius:14px;padding:11px 13px}
.route-header{padding:17px 20px;background:linear-gradient(135deg,#f8fafc,#eef2ff);border:1px solid #e2e8f0;border-radius:16px;margin:10px 0 14px}
.route-title{font-size:1.25rem;font-weight:800;color:#0f172a}
.route-sub{margin-top:6px;font-size:.92rem;color:#64748b}
.driver-label{display:inline-block;padding:7px 12px;border-radius:999px;margin:3px 6px 8px 0;font-size:.88rem;font-weight:700}
.blue-label{background:#dbeafe;color:#1d4ed8}
.orange-label{background:#fef3c7;color:#b45309}
.route-card{border:1px solid #e5e7eb;border-left:6px solid;border-radius:14px;padding:14px 16px;margin:8px 0 5px}
.driver-blue{border-left-color:var(--blue);background:var(--blue-bg)}
.driver-orange{border-left-color:var(--orange);background:var(--orange-bg)}
.start-card{border-left-color:#475569;background:#f8fafc}
.done-card{border-left-color:#94a3b8!important;background:#f1f5f9!important;opacity:.7}
.time{font-size:1.05rem;font-weight:800;color:#0f172a}
.place{margin-top:4px;font-size:1rem;font-weight:650;color:#334155}
.meta{margin-top:6px;font-size:.86rem;color:#64748b}
.event-card{border-left:6px solid;border-radius:13px;padding:12px 15px;margin:10px 0}
.switch-card{border-left-color:var(--red);background:var(--red-bg)}
.break-card{border-left-color:var(--green);background:var(--green-bg)}
.event-title{font-size:.96rem;font-weight:800}
.event-detail{margin-top:4px;font-size:.86rem;color:#64748b}
.progress-box{padding:11px 14px;background:#f8fafc;border:1px solid #e2e8f0;border-radius:13px;margin:9px 0}
div.stButton>button{border-radius:10px!important;min-height:38px!important;font-weight:700!important}
@media(max-width:700px){.block-container{padding:1.4rem .8rem 2.5rem}.route-header{padding:14px}.route-card{padding:13px}}
</style>
""", unsafe_allow_html=True)


def html(content):
    st.markdown(textwrap.dedent(content).strip(), unsafe_allow_html=True)


# ============================================================
# 3. ĐỌC VÀ KIỂM TRA DỮ LIỆU
# ============================================================

@st.cache_data
def load_data():
    schedule_sheets = pd.ExcelFile(SCHEDULE_FILE).sheet_names
    time_sheets = pd.ExcelFile(TIME_FILE).sheet_names
    schedule_data = {name: pd.read_excel(SCHEDULE_FILE, sheet_name=name) for name in schedule_sheets}
    time_data = {name: pd.read_excel(TIME_FILE, sheet_name=name) for name in time_sheets}
    return schedule_data, time_data


if not SCHEDULE_FILE.exists() or not TIME_FILE.exists():
    missing_files = [p.name for p in [SCHEDULE_FILE, TIME_FILE] if not p.exists()]
    st.error("Chưa tìm thấy file Excel: " + ", ".join(missing_files))
    st.info("Hãy tải đủ 2 file GA vào mục Files của Google Colab, đặt cùng thư mục với app.py và giữ nguyên tên file.")
    st.stop()

try:
    schedule_data, time_data = load_data()
except Exception as exc:
    st.error(f"Không đọc được file Excel. Chi tiết: {exc}")
    st.stop()

required_sheets = ["LICH_PHAN_CONG", "CHI_TIET_DOI_LAI", "TONG_HOP_TAI_XE"]
missing_sheets = [name for name in required_sheets if name not in schedule_data]
if missing_sheets:
    st.error("File lịch tài xế GA thiếu sheet: " + ", ".join(missing_sheets))
    st.stop()

schedule = schedule_data["LICH_PHAN_CONG"].copy()
drive_detail = schedule_data["CHI_TIET_DOI_LAI"].copy()
driver_summary = schedule_data["TONG_HOP_TAI_XE"].copy()

if "Chi_tiet_thoi_gian" in time_data:
    route_time = time_data["Chi_tiet_thoi_gian"].copy()
elif "Gio_den_truong" in time_data:
    route_time = time_data["Gio_den_truong"].copy()
else:
    st.error("File giờ đến trường GA thiếu sheet Chi_tiet_thoi_gian hoặc Gio_den_truong.")
    st.stop()

required_schedule_cols = ["Ngày", "Xe", "Tuyến", "Tài xế 1", "Tài xế 2", "Người lái đầu"]
required_drive_cols = ["Ngày", "Xe", "Tuyến", "Thứ tự chặng", "Tài xế lái", "Đổi lái trước chặng"]
required_time_cols = ["Tuyến", "Thứ tự chặng", "Tên điểm đi", "Tên điểm đến", "Điểm đến", "Giờ xuất phát", "Giờ đến"]
for frame, required, label in [
    (schedule, required_schedule_cols, "LICH_PHAN_CONG"),
    (drive_detail, required_drive_cols, "CHI_TIET_DOI_LAI"),
    (route_time, required_time_cols, "file giờ đến trường")
]:
    missing = [col for col in required if col not in frame.columns]
    if missing:
        st.error(f"Sheet {label} thiếu cột: " + ", ".join(missing))
        st.stop()

route_col, order_col = "Tuyến", "Thứ tự chặng"
from_name_col, to_name_col, to_code_col = "Tên điểm đi", "Tên điểm đến", "Điểm đến"
depart_col, arrival_col = "Giờ xuất phát", "Giờ đến"
service_col = next((c for c in ["Thời gian phục vụ (phút)", "Thời gian phục vụ"] if c in route_time.columns), None)
driver_col, switch_col = "Tài xế lái", "Đổi lái trước chặng"
reason_col = "Lý do đổi lái" if "Lý do đổi lái" in drive_detail.columns else None
first_driver_col = "Người lái đầu"

drivers = sorted(schedule[["Tài xế 1", "Tài xế 2"]].astype(str).stack().unique().tolist())
routes = sorted(route_time[route_col].dropna().astype(str).unique().tolist())


# ============================================================
# 4. HÀM XỬ LÝ DỮ LIỆU
# ============================================================

def normalize(value):
    try:
        number = float(value)
        return str(int(number)) if number.is_integer() else str(number)
    except Exception:
        return str(value).strip()


def time_to_minutes(value):
    if pd.isna(value):
        return None
    if hasattr(value, "hour"):
        return value.hour * 60 + value.minute + getattr(value, "second", 0) / 60
    if isinstance(value, (int, float, np.number)):
        number = float(value)
        return number * 1440 if 0 <= number < 2 else number
    try:
        parts = str(value).strip().split(":")
        return int(parts[0]) * 60 + int(parts[1]) + (float(parts[2]) / 60 if len(parts) > 2 else 0)
    except Exception:
        return None


def format_time(value):
    minute = time_to_minutes(value)
    if minute is None:
        return ""
    minute = int(round(minute)) % 1440
    return f"{minute // 60:02d}:{minute % 60:02d}"


def format_minutes(minute):
    minute = int(round(minute)) % 1440
    return f"{minute // 60:02d}:{minute % 60:02d}"


def is_depot(name, code):
    return "kho" in str(name).lower() or "kho" in str(code).lower()


def get_route_driver_maps(day, vehicle, route):
    detail = drive_detail[
        (pd.to_numeric(drive_detail["Ngày"], errors="coerce") == day)
        & (drive_detail["Xe"].astype(str) == vehicle)
        & (drive_detail["Tuyến"].astype(str) == route)
    ]
    driver_map, switch_map = {}, {}
    for _, row in detail.iterrows():
        segment = normalize(row[order_col])
        driver_map[segment] = str(row[driver_col])
        if str(row[switch_col]).strip().lower() in ["có", "co", "yes", "true"]:
            switch_map[segment] = row
    return driver_map, switch_map


def break_before(previous_row, current_row):
    if previous_row is None:
        return 0
    previous_arrival = time_to_minutes(previous_row[arrival_col])
    current_departure = time_to_minutes(current_row[depart_col])
    if previous_arrival is None or current_departure is None:
        return 0
    service = float(previous_row[service_col]) if service_col and pd.notna(previous_row[service_col]) else 0
    return max(0, current_departure - previous_arrival - service)


# ============================================================
# 5. ĐỒNG BỘ TRẠNG THÁI THU GOM
# ============================================================

def fetch_status():
    try:
        response = requests.get(SCRIPT_URL, timeout=REQUEST_TIMEOUT)
        response.raise_for_status()
        data = response.json()
        return (data if isinstance(data, list) else []), None
    except Exception as exc:
        return [], str(exc)


def send_status(payload):
    try:
        response = requests.post(SCRIPT_URL, json=payload, timeout=REQUEST_TIMEOUT)
        response.raise_for_status()
        return True, None
    except Exception as exc:
        return False, str(exc)


def completed_points(status_rows, year, month, day, route):
    completed = set()
    for row in status_rows:
        same_date = (
            normalize(row.get("Nam", "")) == str(year)
            and normalize(row.get("Thang", "")) == str(month)
            and normalize(row.get("Ngay", "")) == str(day)
        )
        if same_date and normalize(row.get("Tuyen", "")) == str(route):
            completed.add(str(row.get("Ma_diem", "")).strip())
    return completed


# ============================================================
# 6. TIÊU ĐỀ VÀ BỘ LỌC
# ============================================================

st.title("🚚 LỊCH LÀM VIỆC TÀI XẾ")
st.caption("Phương án GA · Tra cứu lịch thu gom, phân công lái xe và tiến độ tại từng điểm trường")

today = datetime.datetime.now()
years = list(range(today.year - 1, today.year + 5))

with st.sidebar:
    st.header("TRA CỨU LỊCH")
    selected_driver = st.selectbox("Tài xế", drivers)
    col_month, col_year = st.columns(2)
    selected_month = col_month.selectbox("Tháng", list(range(1, 13)), index=today.month - 1)
    selected_year = col_year.selectbox("Năm", years, index=years.index(today.year))
    selected_day = st.selectbox("Ngày trong chu kỳ 30 ngày", list(range(1, 31)), index=min(today.day, 30) - 1)
    st.divider()
    st.caption("Lịch GA được tổ chức theo chu kỳ 30 ngày. Ngày không có phân công là ngày nghỉ.")

status_rows, status_error = fetch_status()
if status_error:
    st.warning("Chưa kết nối được trạng thái thu gom. Bạn vẫn có thể xem lịch và giờ đến trường.")

tab_day, tab_month, tab_route = st.tabs(["📍 Lịch theo ngày", "📅 Lịch tháng", "🛣️ Chi tiết tuyến"])


# ============================================================
# 7. LỊCH THEO NGÀY
# ============================================================

with tab_day:
    st.subheader(f"Lịch của {selected_driver} · {selected_day:02d}/{selected_month:02d}/{selected_year}")
    daily = schedule[
        (pd.to_numeric(schedule["Ngày"], errors="coerce") == selected_day)
        & (
            (schedule["Tài xế 1"].astype(str) == selected_driver)
            | (schedule["Tài xế 2"].astype(str) == selected_driver)
        )
    ].copy().sort_values(["Xe", "Tuyến"])

    if daily.empty:
        st.info("Không có tuyến thu gom được phân công cho tài xế này trong ngày.")
    for _, assignment in daily.iterrows():
        route, vehicle = str(assignment["Tuyến"]), str(assignment["Xe"])
        driver_1, driver_2 = str(assignment["Tài xế 1"]), str(assignment["Tài xế 2"])
        companion = driver_2 if selected_driver == driver_1 else driver_1
        route_data = route_time[route_time[route_col].astype(str) == route].copy().sort_values(order_col).reset_index(drop=True)
        if route_data.empty:
            st.warning(f"Không tìm thấy dữ liệu giờ đến cho tuyến {route}.")
            continue

        driver_map, switch_map = get_route_driver_maps(selected_day, vehicle, route)
        done_points = completed_points(status_rows, selected_year, selected_month, selected_day, route)
        first_driver = str(assignment[first_driver_col])
        start_time, end_time = format_time(route_data.iloc[0][depart_col]), format_time(route_data.iloc[-1][arrival_col])
        school_rows = [row for _, row in route_data.iterrows() if not is_depot(row[to_name_col], row[to_code_col])]
        done_count = sum(
            f"{normalize(row[to_code_col])}|{normalize(row[order_col])}" in done_points
            for row in school_rows
        )

        html(f"""<div class="route-header"><div class="route-title">{route} · {vehicle}</div><div class="route-sub">Đi cùng: <b>{companion}</b> &nbsp;•&nbsp; Người lái đầu: <b>{first_driver}</b></div></div>""")
        c1, c2, c3 = st.columns(3)
        c1.metric("Xuất phát", start_time)
        c2.metric("Về kho", end_time)
        c3.metric("Số điểm trường", len(school_rows))
        html(f"""<span class="driver-label blue-label">● {driver_1}</span><span class="driver-label orange-label">● {driver_2}</span><div class="progress-box"><b>Tiến độ thu gom: {done_count}/{len(school_rows)} điểm</b></div>""")
        st.progress(done_count / len(school_rows) if school_rows else 0)

        first_row = route_data.iloc[0]
        start_driver = driver_map.get(normalize(first_row[order_col]), first_driver)
        html(f"""<div class="route-card start-card"><div class="time">{start_time} · XUẤT PHÁT</div><div class="place">{first_row[from_name_col]}</div><div class="meta">Người lái đầu: {start_driver}</div></div>""")
        previous_row = None

        for _, row in route_data.iterrows():
            segment = normalize(row[order_col])
            name, code = str(row[to_name_col]), str(row[to_code_col])
            current_driver = driver_map.get(segment, first_driver)
            driver_class = "driver-blue" if current_driver == driver_1 else "driver-orange"
            gap = break_before(previous_row, row)

            if previous_row is not None and gap >= BREAK_THRESHOLD:
                previous_arrival = time_to_minutes(previous_row[arrival_col])
                service = float(previous_row[service_col]) if service_col and pd.notna(previous_row[service_col]) else 0
                rest_start = previous_arrival + service
                rest_end = time_to_minutes(row[depart_col])
                html(f"""<div class="event-card break-card"><div class="event-title">NGHỈ GIỮA CHẶNG · {format_minutes(rest_start)}–{format_minutes(rest_end)}</div><div class="event-detail">Thời gian nghỉ: {int(round(gap))} phút</div></div>""")

            if segment in switch_map:
                switch_row = switch_map[segment]
                switch_place = str(switch_row["Tên điểm đi"]) if "Tên điểm đi" in switch_row.index else str(row[from_name_col])
                reason = str(switch_row[reason_col]) if reason_col and pd.notna(switch_row[reason_col]) else "Phân công đổi lái"
                html(f"""<div class="event-card switch-card"><div class="event-title">ĐỔI LÁI · {format_time(row[depart_col])}</div><div class="event-detail">Tại: <b>{switch_place}</b><br>Người lái tiếp theo: <b>{current_driver}</b><br>{reason}</div></div>""")

            if is_depot(name, code):
                html(f"""<div class="route-card {driver_class}"><div class="time">{format_time(row[arrival_col])} · VỀ KHO</div><div class="place">{name}</div><div class="meta">Người lái chặng cuối: {current_driver}</div></div>""")
            else:
                point_id = f"{normalize(code)}|{segment}"
                is_done = point_id in done_points
                done_class = "done-card" if is_done else ""
                done_text = "<div class='meta'><b>✓ Đã thu gom</b></div>" if is_done else ""
                html(f"""<div class="route-card {driver_class} {done_class}"><div class="time">{'✓ ' if is_done else ''}{format_time(row[arrival_col])}</div><div class="place">{name}</div><div class="meta">Người lái chặng này: {current_driver}</div>{done_text}</div>""")
                button_label = "↩ Hoàn tác" if is_done else "✓ Đã thu gom"
                action = "undo" if is_done else "complete"
                if st.button(button_label, key=f"point_{selected_year}_{selected_month}_{selected_day}_{route}_{point_id}", use_container_width=False):
                    ok, error = send_status({
                        "action": action, "Nam": selected_year, "Thang": selected_month, "Ngay": selected_day,
                        "Tuyen": route, "Ma_diem": point_id, "Ten_diem": name, "Tai_xe": selected_driver
                    })
                    if ok:
                        st.rerun()
                    else:
                        st.error("Không lưu được trạng thái thu gom: " + str(error))
            previous_row = row

        reset_key = f"reset_{selected_year}_{selected_month}_{selected_day}_{route}"
        if st.button("↻ Reset trạng thái tuyến này", key="ask_" + reset_key):
            st.session_state["confirm_" + reset_key] = True
        if st.session_state.get("confirm_" + reset_key, False):
            st.warning("Xác nhận reset sẽ đưa toàn bộ điểm của tuyến về trạng thái chưa thu gom.")
            yes, no = st.columns(2)
            if yes.button("Xác nhận reset", key="yes_" + reset_key, type="primary", use_container_width=True):
                ok, error = send_status({
                    "action": "reset_route", "Nam": selected_year, "Thang": selected_month,
                    "Ngay": selected_day, "Tuyen": route
                })
                if ok:
                    st.session_state["confirm_" + reset_key] = False
                    st.rerun()
                else:
                    st.error("Reset không thành công: " + str(error))
            if no.button("Hủy", key="no_" + reset_key, use_container_width=True):
                st.session_state["confirm_" + reset_key] = False
                st.rerun()
        st.divider()


# ============================================================
# 8. LỊCH THÁNG
# ============================================================

with tab_month:
    st.subheader(f"Lịch tháng {selected_month:02d}/{selected_year} của {selected_driver}")
    month_rows = []
    driver_schedule = schedule[
        (schedule["Tài xế 1"].astype(str) == selected_driver)
        | (schedule["Tài xế 2"].astype(str) == selected_driver)
    ].copy()
    for _, assignment in driver_schedule.iterrows():
        route = str(assignment["Tuyến"])
        data = route_time[route_time[route_col].astype(str) == route].copy().sort_values(order_col)
        companion = assignment["Tài xế 2"] if str(assignment["Tài xế 1"]) == selected_driver else assignment["Tài xế 1"]
        month_rows.append({
            "Ngày": int(assignment["Ngày"]), "Tuyến": route, "Xe": assignment["Xe"],
            "Xuất phát": format_time(data.iloc[0][depart_col]) if not data.empty else "",
            "Về kho": format_time(data.iloc[-1][arrival_col]) if not data.empty else "",
            "Đi cùng": companion, "Người lái đầu": assignment[first_driver_col]
        })
    if month_rows:
        month_df = pd.DataFrame(month_rows).sort_values(["Ngày", "Xe"])
        st.dataframe(month_df, use_container_width=True, hide_index=True)
    else:
        st.info("Tài xế không có lịch phân công trong chu kỳ này.")


# ============================================================
# 9. CHI TIẾT TUYẾN
# ============================================================

with tab_route:
    st.subheader("Chi tiết giờ đến từng điểm trường")
    selected_route = st.selectbox("Chọn tuyến", routes)
    data = route_time[route_time[route_col].astype(str) == selected_route].copy().sort_values(order_col)
    if not data.empty:
        schools = []
        for _, row in data.iterrows():
            if not is_depot(row[to_name_col], row[to_code_col]):
                schools.append({
                    "STT": normalize(row[order_col]),
                    "Giờ đến": format_time(row[arrival_col]),
                    "Điểm trường": str(row[to_name_col]),
                })
        c1, c2, c3 = st.columns(3)
        c1.metric("Xuất phát", format_time(data.iloc[0][depart_col]))
        c2.metric("Về kho", format_time(data.iloc[-1][arrival_col]))
        c3.metric("Số điểm trường", len(schools))
        st.dataframe(pd.DataFrame(schools), use_container_width=True, hide_index=True)

st.divider()
st.caption("Lịch được xây dựng theo phương án GA. Thời gian đến trường là thời gian kế hoạch; trạng thái thu gom được đồng bộ qua Google Sheets.")

