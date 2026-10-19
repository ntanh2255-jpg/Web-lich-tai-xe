
import datetime
import textwrap
from pathlib import Path

import numpy as np
import pandas as pd
import requests
import streamlit as st


# ============================================================
# 1. CẤU HÌNH
# ============================================================

BASE_DIR = Path(__file__).resolve().parent
SCHEDULE_FILE = BASE_DIR / "Lich_tai_xe_ALNS_30_ngay_can_bang.xlsx"
TIME_FILE = BASE_DIR / "Giờ đến trường ALNS.xlsx"

SCRIPT_URL = "https://script.google.com/macros/s/AKfycbzmPXp8-KTeiM9uCDMMshtKucCz1LoM269Dpc5hm-rWNFE_nKXV0uKgrkXQ020WmYCM/exec"

LUNCH_BREAK_THRESHOLD = 20
GET_TIMEOUT = 15
POST_TIMEOUT = 20

st.set_page_config(page_title="Lịch làm việc tài xế", page_icon="🚚", layout="wide", initial_sidebar_state="expanded")


# ============================================================
# 2. GIAO DIỆN
# ============================================================

st.markdown("""
<style>
:root{--blue:#2563eb;--blue-bg:#eff6ff;--orange:#f59e0b;--orange-bg:#fffbeb;--red:#dc2626;--red-bg:#fff1f2;--green:#16a34a;--green-bg:#f0fdf4;--gray:#64748b}
.block-container{max-width:1000px;padding-top:2rem;padding-bottom:4rem}
[data-testid="stSidebar"]{min-width:260px;background:#f8fafc}
div[data-testid="stMetric"]{background:#fff;border:1px solid #e5e7eb;border-radius:14px;padding:11px 13px}
.route-header{padding:16px 18px;background:linear-gradient(135deg,#f8fafc,#eef2ff);border:1px solid #e2e8f0;border-radius:16px;margin:8px 0 14px}
.route-title{font-size:1.25rem;font-weight:800;color:#0f172a}
.route-sub{margin-top:4px;font-size:.9rem;color:#64748b}
.driver-label{display:inline-block;padding:7px 12px;border-radius:999px;margin:3px 6px 8px 0;font-size:.88rem;font-weight:700}
.blue-label{background:#dbeafe;color:#1d4ed8}
.orange-label{background:#fef3c7;color:#b45309}
.route-card{border:1px solid #e5e7eb;border-left:6px solid;border-radius:14px;padding:14px 16px;margin:7px 0 4px}
.driver-blue{border-left-color:var(--blue);background:var(--blue-bg)}
.driver-orange{border-left-color:var(--orange);background:var(--orange-bg)}
.start-card{border-left-color:#475569;background:#f8fafc}
.done-card{border-left-color:#94a3b8!important;background:#f1f5f9!important;opacity:.65}
.time{font-size:1.05rem;font-weight:800;color:#0f172a}
.place{margin-top:3px;font-size:1rem;font-weight:600;color:#334155}
.meta{margin-top:6px;font-size:.84rem;color:#64748b}
.event-card{border-left:6px solid;border-radius:13px;padding:12px 15px;margin:9px 0}
.switch-card{border-left-color:var(--red);background:var(--red-bg)}
.lunch-card{border-left-color:var(--green);background:var(--green-bg)}
.event-title{font-size:.96rem;font-weight:800}
.event-detail{margin-top:4px;font-size:.86rem;color:#64748b}
.progress-box{padding:11px 14px;background:#f8fafc;border:1px solid #e2e8f0;border-radius:13px;margin:8px 0 10px}
div.stButton>button{border-radius:10px!important;min-height:38px!important;font-weight:700!important}
@media(max-width:700px){.block-container{padding:1rem .8rem 3rem}.route-header{padding:14px}.route-card{padding:13px}}
</style>
""", unsafe_allow_html=True)


def html(content):
    """Render HTML sạch, tránh Streamlit hiểu nhầm thành code block."""
    st.markdown(textwrap.dedent(content).strip(), unsafe_allow_html=True)


# ============================================================
# 3. ĐỌC DỮ LIỆU
# ============================================================

@st.cache_data
def load_data():
    schedule_xls = pd.ExcelFile(SCHEDULE_FILE)
    time_xls = pd.ExcelFile(TIME_FILE)
    schedule_data = {s: pd.read_excel(SCHEDULE_FILE, sheet_name=s) for s in schedule_xls.sheet_names}
    time_data = {s: pd.read_excel(TIME_FILE, sheet_name=s) for s in time_xls.sheet_names}
    return schedule_data, time_data


def find_col(df, names):
    for name in names:
        if name in df.columns:
            return name
    return None


try:
    schedule_data, time_data = load_data()
except FileNotFoundError as e:
    st.error(f"Không tìm thấy file dữ liệu: {e}")
    st.stop()

required_sheets = ["LICH_PHAN_CONG", "CHI_TIET_DOI_LAI", "TONG_HOP_TAI_XE"]
missing = [s for s in required_sheets if s not in schedule_data]

if missing:
    st.error("File lịch tài xế thiếu sheet: " + ", ".join(missing))
    st.stop()

schedule = schedule_data["LICH_PHAN_CONG"].copy()
drive_detail = schedule_data["CHI_TIET_DOI_LAI"].copy()
driver_summary = schedule_data["TONG_HOP_TAI_XE"].copy()

if "Chi_tiet_thoi_gian" in time_data:
    route_time = time_data["Chi_tiet_thoi_gian"].copy()
elif "Gio_den_truong" in time_data:
    route_time = time_data["Gio_den_truong"].copy()
else:
    st.error("File giờ đến trường thiếu sheet Chi_tiet_thoi_gian hoặc Gio_den_truong.")
    st.stop()


# ============================================================
# 4. XÁC ĐỊNH CỘT
# ============================================================

route_col = find_col(route_time, ["Tuyến"])
order_col = find_col(route_time, ["Thứ tự chặng"])
from_name_col = find_col(route_time, ["Tên điểm đi"])
to_name_col = find_col(route_time, ["Tên điểm đến"])
to_code_col = find_col(route_time, ["Điểm đến"])
depart_col = find_col(route_time, ["Giờ xuất phát"])
arrival_col = find_col(route_time, ["Giờ đến"])
service_col = find_col(route_time, ["Thời gian phục vụ (phút)", "Thời gian phục vụ"])

driver_col = find_col(drive_detail, ["Tài xế lái"])
switch_col = find_col(drive_detail, ["Đổi lái trước chặng", "Đổi lái"])
reason_col = find_col(drive_detail, ["Lý do đổi lái"])
drive_order_col = find_col(drive_detail, ["Thứ tự chặng"])
first_driver_col = find_col(schedule, ["Người lái đầu", "Tài xế lái đầu"])

drivers = sorted(driver_summary["Tài xế"].dropna().astype(str).unique().tolist())


# ============================================================
# 5. HÀM CHUẨN HÓA
# ============================================================

def normalize_number(value):
    try:
        number = float(value)
        return str(int(number)) if number.is_integer() else str(number)
    except Exception:
        return str(value).strip()


def segment_key(value):
    return normalize_number(value)


def time_to_minutes(value):
    if pd.isna(value):
        return None

    if hasattr(value, "hour"):
        return value.hour * 60 + value.minute + getattr(value, "second", 0) / 60

    if isinstance(value, (int, float, np.number)):
        value = float(value)
        return value * 1440 if 0 <= value < 2 else value

    try:
        parts = str(value).strip().split(":")
        hour, minute = int(parts[0]), int(parts[1])
        second = float(parts[2]) if len(parts) > 2 else 0
        return hour * 60 + minute + second / 60
    except Exception:
        return None


def format_time(value):
    minute = time_to_minutes(value)

    if minute is None:
        return ""

    minute = int(round(minute)) % 1440
    return f"{minute // 60:02d}:{minute % 60:02d}"


def format_minute_value(minute):
    minute = int(round(minute)) % 1440
    return f"{minute // 60:02d}:{minute % 60:02d}"


def is_depot(name, code):
    return "kho" in str(name).lower() or "kho" in str(code).lower()


# ============================================================
# 6. GOOGLE SHEETS - TRẠNG THÁI THU GOM
# ============================================================

def fetch_status():
    try:
        response = requests.get(SCRIPT_URL, timeout=GET_TIMEOUT)
        response.raise_for_status()
        data = response.json()
        return data if isinstance(data, list) else [], None
    except Exception as e:
        return [], str(e)


def send_status(payload):
    try:
        response = requests.post(SCRIPT_URL, json=payload, timeout=POST_TIMEOUT)
        response.raise_for_status()
        return True, None
    except Exception as e:
        return False, str(e)


def completed_points(status_rows, year, month, day, route):
    result = set()

    for row in status_rows:
        same_date = (
            normalize_number(row.get("Nam", "")) == str(year)
            and normalize_number(row.get("Thang", "")) == str(month)
            and normalize_number(row.get("Ngay", "")) == str(day)
        )

        same_route = normalize_number(row.get("Tuyen", "")) == str(route)

        if same_date and same_route:
            result.add(normalize_number(row.get("Ma_diem", "")))

    return result


# ============================================================
# 7. XỬ LÝ PHÂN CÔNG LÁI
# ============================================================

def route_driver_maps(day, vehicle, route):
    detail = drive_detail[
        (pd.to_numeric(drive_detail["Ngày"], errors="coerce") == day)
        & (drive_detail["Xe"].astype(str) == vehicle)
        & (drive_detail["Tuyến"].astype(str) == route)
    ].copy()

    driver_map, switch_map = {}, {}

    for _, row in detail.iterrows():
        segment = segment_key(row[drive_order_col])
        driver_map[segment] = str(row[driver_col])

        if switch_col and str(row[switch_col]).strip().lower() == "có":
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
    return current_departure - previous_arrival - service


# ============================================================
# 8. TIÊU ĐỀ
# ============================================================

st.title("🚚 LỊCH LÀM VIỆC TÀI XẾ")
st.caption("Theo dõi tuyến thu gom, phân công lái xe và tiến độ tại từng điểm trường")

# ============================================================
# 9. BỘ LỌC
# ============================================================

today = datetime.datetime.now()
years = list(range(today.year - 1, today.year + 5))

with st.sidebar:
    st.header("TRA CỨU LỊCH")

    selected_driver = st.selectbox("Tài xế", drivers)

    col_month, col_year = st.columns(2)
    selected_month = col_month.selectbox("Tháng", list(range(1, 13)), index=today.month - 1)
    selected_year = col_year.selectbox("Năm", years, index=years.index(today.year))

    selected_day = st.selectbox("Ngày", list(range(1, 31)), index=min(today.day, 30) - 1)

    st.divider()
    st.caption("Trạng thái thu gom được lưu tập trung và đồng bộ giữa các thiết bị.")


status_rows, status_error = fetch_status()

if status_error:
    st.warning("Không kết nối được trạng thái thu gom. Lịch vẫn có thể xem bình thường.")


# ============================================================
# 10. CÁC TAB
# ============================================================

tab_day, tab_month, tab_route = st.tabs(["📍 Lịch theo ngày", "📅 Lịch tháng", "🛣️ Chi tiết tuyến"])


# ============================================================
# 11. LỊCH THEO NGÀY
# ============================================================

with tab_day:
    st.subheader(f"Lịch của {selected_driver} · {selected_day:02d}/{selected_month:02d}/{selected_year}")

    daily_schedule = schedule[
        (pd.to_numeric(schedule["Ngày"], errors="coerce") == selected_day)
        & (
            (schedule["Tài xế 1"].astype(str) == selected_driver)
            | (schedule["Tài xế 2"].astype(str) == selected_driver)
        )
    ].copy()

    if daily_schedule.empty:
        st.info("Không có tuyến thu gom trong ngày này.")

    for _, assignment in daily_schedule.iterrows():
        route = str(assignment["Tuyến"])
        vehicle = str(assignment["Xe"])
        driver_1 = str(assignment["Tài xế 1"])
        driver_2 = str(assignment["Tài xế 2"])
        companion = driver_2 if selected_driver == driver_1 else driver_1

        route_data = route_time[route_time[route_col].astype(str) == route].copy()
        route_data = route_data.sort_values(order_col).reset_index(drop=True)

        if route_data.empty:
            continue

        driver_map, switch_map = route_driver_maps(selected_day, vehicle, route)
        done_points = completed_points(status_rows, selected_year, selected_month, selected_day, route)

        first_driver = str(assignment[first_driver_col]) if first_driver_col else driver_1
        start_time = format_time(route_data.iloc[0][depart_col])
        end_time = format_time(route_data.iloc[-1][arrival_col])

        html(f"""
        <div class="route-header">
            <div class="route-title">{route} · {vehicle}</div>
            <div class="route-sub">Đi cùng: <b>{companion}</b> &nbsp;•&nbsp; Người lái đầu: <b>{first_driver}</b></div>
        </div>
        """)

        m1, m2, m3 = st.columns(3)
        m1.metric("Xuất phát", start_time)
        m2.metric("Về kho", end_time)
        m3.metric("Đi cùng", companion)

        html(f"""
        <span class="driver-label blue-label">● {driver_1}</span>
        <span class="driver-label orange-label">● {driver_2}</span>
        """)

        school_rows = [row for _, row in route_data.iterrows() if not is_depot(row[to_name_col], row[to_code_col])]
        total_points = len(school_rows)

        done_count = 0
        for row in school_rows:
            point_id = f"{normalize_number(row[to_code_col])}|{segment_key(row[order_col])}"
            if point_id in done_points:
                done_count += 1

        html(f'<div class="progress-box"><b>Tiến độ thu gom: {done_count}/{total_points} điểm</b></div>')
        st.progress(done_count / total_points if total_points else 0)

        st.markdown("### Lộ trình")

        first_row = route_data.iloc[0]
        start_driver = driver_map.get(segment_key(first_row[order_col]), first_driver)

        html(f"""
        <div class="route-card start-card">
            <div class="time">{start_time} · Xuất phát</div>
            <div class="place">{first_row[from_name_col]}</div>
            <div class="meta">Người lái: {start_driver}</div>
        </div>
        """)

        previous_row = None

        for _, row in route_data.iterrows():
            segment = segment_key(row[order_col])
            name = str(row[to_name_col])
            code = str(row[to_code_col])
            current_driver = driver_map.get(segment, first_driver)
            driver_class = "driver-blue" if current_driver == driver_1 else "driver-orange"

            gap = break_before(previous_row, row)

            if previous_row is not None and gap >= LUNCH_BREAK_THRESHOLD:
                previous_arrival = time_to_minutes(previous_row[arrival_col])
                service = float(previous_row[service_col]) if service_col and pd.notna(previous_row[service_col]) else 0
                lunch_start = previous_arrival + service
                lunch_end = time_to_minutes(row[depart_col])

                html(f"""
                <div class="event-card lunch-card">
                    <div class="event-title">🍱 {format_minute_value(lunch_start)}–{format_minute_value(lunch_end)} · NGHỈ TRƯA</div>
                    <div class="event-detail">Thời gian nghỉ: {int(round(gap))} phút</div>
                </div>
                """)

            if segment in switch_map:
                switch_row = switch_map[segment]
                place = str(switch_row["Tên điểm đi"]) if "Tên điểm đi" in switch_row.index else ""
                reason = str(switch_row[reason_col]) if reason_col and pd.notna(switch_row[reason_col]) else ""

                html(f"""
                <div class="event-card switch-card">
                    <div class="event-title">🔄 {format_time(row[depart_col])} · ĐỔI LÁI</div>
                    <div class="event-detail">Tại: {place}<br>Người lái tiếp theo: <b>{current_driver}</b>{"<br>" + reason if reason else ""}</div>
                </div>
                """)

            if is_depot(name, code):
                html(f"""
                <div class="route-card {driver_class}">
                    <div class="time">{format_time(row[arrival_col])} · Về kho</div>
                    <div class="place">{name}</div>
                    <div class="meta">Người lái chặng này: {current_driver}</div>
                </div>
                """)

            else:
                point_id = f"{normalize_number(code)}|{segment}"
                is_done = point_id in done_points
                done_class = "done-card" if is_done else ""

                html(f"""
                <div class="route-card {driver_class} {done_class}">
                    <div class="time">{"✓ " if is_done else ""}{format_time(row[arrival_col])}</div>
                    <div class="place">{name}</div>
                    <div class="meta">Người lái chặng này: {current_driver}</div>
                    {"<div class='meta'><b>✓ Đã thu gom</b></div>" if is_done else ""}
                </div>
                """)

                button_col, _ = st.columns([1, 2.6])
                button_label = "↩ Hoàn tác" if is_done else "✓ Đã thu gom"
                action = "undo" if is_done else "complete"

                if button_col.button(
                    button_label,
                    key=f"point_{selected_year}_{selected_month}_{selected_day}_{route}_{point_id}",
                    use_container_width=True
                ):
                    ok, error = send_status({
                        "action": action,
                        "Nam": selected_year,
                        "Thang": selected_month,
                        "Ngay": selected_day,
                        "Tuyen": route,
                        "Ma_diem": point_id,
                        "Ten_diem": name,
                        "Tai_xe": selected_driver,
                    })

                    if ok:
                        st.rerun()
                    else:
                        st.error("Không lưu được trạng thái: " + str(error))

            previous_row = row

        st.divider()

        reset_key = f"reset_{selected_year}_{selected_month}_{selected_day}_{route}"

        if not st.session_state.get(reset_key, False):
            if st.button("↻ Reset tuyến hôm nay", key="ask_" + reset_key):
                st.session_state[reset_key] = True
                st.rerun()

        else:
            st.warning("Reset sẽ đưa toàn bộ điểm của tuyến này về trạng thái chưa thu gom.")

            confirm_col, cancel_col = st.columns(2)

            if confirm_col.button("Xác nhận reset", key="confirm_" + reset_key, type="primary", use_container_width=True):
                ok, error = send_status({
                    "action": "reset_route",
                    "Nam": selected_year,
                    "Thang": selected_month,
                    "Ngay": selected_day,
                    "Tuyen": route,
                })

                if ok:
                    st.session_state[reset_key] = False
                    st.rerun()
                else:
                    st.error("Reset không thành công: " + str(error))

            if cancel_col.button("Hủy", key="cancel_" + reset_key, use_container_width=True):
                st.session_state[reset_key] = False
                st.rerun()


# ============================================================
# 12. LỊCH THÁNG
# ============================================================

with tab_month:
    st.subheader(f"Lịch tháng {selected_month:02d}/{selected_year} của {selected_driver}")

    month_schedule = schedule[
        (schedule["Tài xế 1"].astype(str) == selected_driver)
        | (schedule["Tài xế 2"].astype(str) == selected_driver)
    ].copy()

    rows = []

    for _, assignment in month_schedule.iterrows():
        route = str(assignment["Tuyến"])
        data = route_time[route_time[route_col].astype(str) == route].copy().sort_values(order_col)
        companion = assignment["Tài xế 2"] if str(assignment["Tài xế 1"]) == selected_driver else assignment["Tài xế 1"]

        rows.append({
            "Ngày": int(assignment["Ngày"]),
            "Tuyến": route,
            "Xe": assignment["Xe"],
            "Xuất phát": format_time(data.iloc[0][depart_col]) if not data.empty else "",
            "Về kho": format_time(data.iloc[-1][arrival_col]) if not data.empty else "",
            "Đi cùng": companion,
        })

    if rows:
        st.dataframe(pd.DataFrame(rows).sort_values("Ngày"), use_container_width=True, hide_index=True)
    else:
        st.info("Không có lịch trong tháng.")


# ============================================================
# 13. CHI TIẾT TUYẾN
# ============================================================

with tab_route:
    st.subheader("Chi tiết tuyến")

    route_list = sorted(route_time[route_col].dropna().astype(str).unique().tolist())
    selected_route = st.selectbox("Chọn tuyến", route_list)

    data = route_time[route_time[route_col].astype(str) == selected_route].copy().sort_values(order_col)

    if not data.empty:
        schools = []

        for _, row in data.iterrows():
            if is_depot(row[to_name_col], row[to_code_col]):
                continue

            schools.append({
                "STT": segment_key(row[order_col]),
                "Giờ đến": format_time(row[arrival_col]),
                "Điểm trường": str(row[to_name_col]),
            })

        c1, c2, c3 = st.columns(3)
        c1.metric("Xuất phát", format_time(data.iloc[0][depart_col]))
        c2.metric("Về kho", format_time(data.iloc[-1][arrival_col]))
        c3.metric("Số điểm trường", len(schools))

        st.dataframe(pd.DataFrame(schools), use_container_width=True, hide_index=True)


st.divider()
st.caption("Hệ thống hỗ trợ tài xế theo dõi lịch thu gom, phân công lái và tiến độ thực hiện tại các điểm trường.")
