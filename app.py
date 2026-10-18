
import streamlit as st
import pandas as pd
import os, unicodedata

SCHEDULE_FILE = "Lich_tai_xe_ALNS_30_ngay_can_bang.xlsx"
TIME_FILE = "Giờ đến trường ALNS.xlsx"

st.set_page_config(page_title="Lịch làm việc tài xế", page_icon="🚚", layout="wide", initial_sidebar_state="expanded")

st.markdown("""
<style>
.block-container{padding-top:1.4rem;padding-bottom:3rem;max-width:1150px}
[data-testid="stSidebar"]{min-width:270px}
div[data-testid="stMetric"]{border:1px solid #e5e7eb;padding:12px;border-radius:12px}
.timeline{padding:11px 15px;border-left:4px solid #2563eb;background:#f8fafc;border-radius:8px;margin-bottom:8px}
.switch{padding:12px 15px;border-left:4px solid #dc2626;background:#fff7f7;border-radius:8px;margin:12px 0}
.note{font-size:14px;color:#6b7280}
</style>
""", unsafe_allow_html=True)

st.title("🚚 LỊCH LÀM VIỆC TÀI XẾ")
st.caption("Tra cứu tuyến, giờ xuất phát, giờ đến trường và thời điểm đổi lái")

@st.cache_data
def load_data():
    sx = pd.ExcelFile(SCHEDULE_FILE); tx = pd.ExcelFile(TIME_FILE)
    sdata = {s:pd.read_excel(SCHEDULE_FILE,sheet_name=s) for s in sx.sheet_names}
    tdata = {s:pd.read_excel(TIME_FILE,sheet_name=s) for s in tx.sheet_names}
    return sdata,tdata

sdata,tdata = load_data()

required = ["LICH_PHAN_CONG","CHI_TIET_DOI_LAI","TONG_HOP_TAI_XE"]
missing = [s for s in required if s not in sdata]
if missing:
    st.error("File lịch tài xế thiếu sheet: "+", ".join(missing)); st.stop()

schedule = sdata["LICH_PHAN_CONG"].copy()
drive_detail = sdata["CHI_TIET_DOI_LAI"].copy()
summary = sdata["TONG_HOP_TAI_XE"].copy()

if "Chi_tiet_thoi_gian" in tdata: route_time = tdata["Chi_tiet_thoi_gian"].copy()
elif "Gio_den_truong" in tdata: route_time = tdata["Gio_den_truong"].copy()
else:
    st.error("File giờ đến trường không có sheet Chi_tiet_thoi_gian hoặc Gio_den_truong."); st.stop()

def col(df,names):
    for x in names:
        if x in df.columns: return x
    return None

def gio(x):
    if pd.isna(x): return ""
    if isinstance(x,str):
        x=x.strip()
        if len(x)>=5 and ":" in x: return x[:5]
        return x
    if hasattr(x,"strftime"):
        try: return x.strftime("%H:%M")
        except: pass
    try: return pd.to_datetime(str(x)).strftime("%H:%M")
    except: return str(x)

first_driver_col = col(schedule,["Người lái đầu","Tài xế lái đầu"])
switch_col = col(drive_detail,["Đổi lái trước chặng","Đổi lái"])
driver_col = col(drive_detail,["Tài xế lái"])
segment_col = col(drive_detail,["Thứ tự chặng"])
route_col = col(route_time,["Tuyến"])
order_col = col(route_time,["Thứ tự chặng"])
from_name_col = col(route_time,["Tên điểm đi"])
to_name_col = col(route_time,["Tên điểm đến"])
to_code_col = col(route_time,["Điểm đến"])
depart_col = col(route_time,["Giờ xuất phát"])
arrival_col = col(route_time,["Giờ đến"])

drivers = sorted(summary["Tài xế"].dropna().astype(str).unique().tolist())

with st.sidebar:
    st.header("TRA CỨU LỊCH")
    selected_driver = st.selectbox("Tài xế",drivers)
    selected_day = st.selectbox("Ngày trong tháng",list(range(1,31)))
    st.divider()
    st.caption("Chọn tài xế và ngày cần xem.")

tab1,tab2,tab3 = st.tabs(["📍 Lịch theo ngày","📅 Lịch tháng","🛣️ Chi tiết tuyến"])

# ============================================================
# TAB 1
# ============================================================

with tab1:
    st.subheader(f"Lịch của {selected_driver} - Ngày {selected_day}")

    day_rows = schedule[(schedule["Ngày"]==selected_day)&((schedule["Tài xế 1"].astype(str)==selected_driver)|(schedule["Tài xế 2"].astype(str)==selected_driver))].copy()

    if day_rows.empty:
        st.info("Không có tuyến thu gom trong ngày này.")
    else:
        for _,r in day_rows.iterrows():
            route=str(r["Tuyến"]); vehicle=str(r["Xe"]); mate=r["Tài xế 2"] if str(r["Tài xế 1"])==selected_driver else r["Tài xế 1"]
            rt=route_time[route_time[route_col].astype(str)==route].copy()
            if order_col: rt=rt.sort_values(order_col)

            start_time=gio(rt.iloc[0][depart_col]) if not rt.empty and depart_col else "-"
            end_time=gio(rt.iloc[-1][arrival_col]) if not rt.empty and arrival_col else "-"

            st.markdown(f"### {route} · {vehicle}")
            c1,c2,c3,c4=st.columns(4)
            c1.metric("Xuất phát",start_time); c2.metric("Về kho",end_time); c3.metric("Đi cùng",str(mate)); c4.metric("Xe",vehicle)

            if first_driver_col:
                fd=str(r[first_driver_col])
                if fd==selected_driver: st.success(f"Bạn ({selected_driver}) là người lái đầu.")
                else: st.info(f"{fd} là người lái đầu.")

            st.markdown("### Lộ trình")

            if not rt.empty:
                first=rt.iloc[0]
                st.markdown(f'<div class="timeline"><b>{start_time} — Xuất phát</b><br>{first[from_name_col] if from_name_col else "Kho"}</div>',unsafe_allow_html=True)

                for _,x in rt.iterrows():
                    name=str(x[to_name_col]) if to_name_col else ""
                    code=str(x[to_code_col]) if to_code_col else ""
                    arr=gio(x[arrival_col]) if arrival_col else ""

                    if "kho" in code.lower() or "kho" in name.lower():
                        st.markdown(f'<div class="timeline"><b>{arr} — Về kho</b><br>{name}</div>',unsafe_allow_html=True)
                    else:
                        st.markdown(f'<div class="timeline"><b>{arr}</b><br>{name}</div>',unsafe_allow_html=True)

            if switch_col and driver_col:
                dd=drive_detail[(drive_detail["Ngày"]==selected_day)&(drive_detail["Xe"].astype(str)==vehicle)&(drive_detail["Tuyến"].astype(str)==route)].copy()
                switches=dd[dd[switch_col].astype(str).str.strip().str.lower()=="có"].copy()

                if not switches.empty:
                    st.markdown("### Điểm đổi lái")

                    for _,sw in switches.iterrows():
                        seg=sw[segment_col] if segment_col else None
                        place=sw["Tên điểm đi"] if "Tên điểm đi" in sw.index else ""
                        new_driver=sw[driver_col]
                        sw_time=""

                        if seg is not None and order_col and depart_col:
                            temp=rt[rt[order_col]==seg]
                            if not temp.empty: sw_time=gio(temp.iloc[0][depart_col])

                        st.markdown(f'<div class="switch"><b>🔄 {sw_time} — ĐỔI LÁI</b><br>Tại: {place}<br>Người lái tiếp theo: <b>{new_driver}</b></div>',unsafe_allow_html=True)

            st.markdown('<div class="note">Thời gian hiển thị là thời gian dự kiến theo lịch trình.</div>',unsafe_allow_html=True)

# ============================================================
# TAB 2
# ============================================================

with tab2:
    st.subheader(f"Lịch tháng của {selected_driver}")

    month=schedule[(schedule["Tài xế 1"].astype(str)==selected_driver)|(schedule["Tài xế 2"].astype(str)==selected_driver)].copy()
    rows=[]

    for _,r in month.iterrows():
        route=str(r["Tuyến"]); rt=route_time[route_time[route_col].astype(str)==route].copy()
        if order_col: rt=rt.sort_values(order_col)

        start=gio(rt.iloc[0][depart_col]) if not rt.empty and depart_col else ""
        end=gio(rt.iloc[-1][arrival_col]) if not rt.empty and arrival_col else ""
        mate=r["Tài xế 2"] if str(r["Tài xế 1"])==selected_driver else r["Tài xế 1"]

        rows.append({"Ngày":int(r["Ngày"]),"Tuyến":route,"Xe":r["Xe"],"Xuất phát":start,"Về kho":end,"Đi cùng":mate})

    mv=pd.DataFrame(rows)

    if mv.empty: st.info("Không có lịch trong tháng.")
    else: st.dataframe(mv.sort_values("Ngày"),use_container_width=True,hide_index=True)

# ============================================================
# TAB 3
# ============================================================

with tab3:
    st.subheader("Chi tiết tuyến")

    route_list=sorted(route_time[route_col].dropna().astype(str).unique().tolist())
    selected_route=st.selectbox("Chọn tuyến",route_list,key="route")

    rt=route_time[route_time[route_col].astype(str)==selected_route].copy()
    if order_col: rt=rt.sort_values(order_col)

    if not rt.empty:
        start=gio(rt.iloc[0][depart_col]) if depart_col else ""
        end=gio(rt.iloc[-1][arrival_col]) if arrival_col else ""

        schools=0
        if to_name_col:
            for x in rt[to_name_col].astype(str):
                if "kho" not in x.lower(): schools+=1

        c1,c2,c3=st.columns(3)
        c1.metric("Xuất phát",start); c2.metric("Về kho",end); c3.metric("Số điểm trường",schools)

        rows=[]
        for _,x in rt.iterrows():
            name=str(x[to_name_col]) if to_name_col else ""
            code=str(x[to_code_col]) if to_code_col else ""

            if "kho" in name.lower() or "kho" in code.lower(): continue

            rows.append({"STT":int(x[order_col]) if order_col else len(rows)+1,"Giờ đến":gio(x[arrival_col]) if arrival_col else "","Điểm trường":name})

        st.dataframe(pd.DataFrame(rows),use_container_width=True,hide_index=True)

st.divider()
st.caption("Hệ thống hỗ trợ tài xế theo dõi lịch thu gom và thời gian dự kiến tại các điểm trường.")
