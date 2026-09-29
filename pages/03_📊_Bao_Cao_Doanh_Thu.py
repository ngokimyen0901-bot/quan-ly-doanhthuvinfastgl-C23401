import streamlit as st
import pandas as pd
import re

# Cấu hình trang
st.set_page_config(page_title="Báo Cáo Doanh Thu Dịch Vụ VinFast", page_icon="📊", layout="wide")

st.markdown("""
<style>
    .metric-card {
        background-color: #f8fafc;
        border: 1px solid #e2e8f0;
        border-radius: 8px;
        padding: 12px;
    }
</style>
""", unsafe_allow_html=True)

st.title("📊 Báo Cáo Doanh Thu Dịch Vụ & Dự Phóng")
st.caption("Bảo dưỡng định kỳ, Sửa chữa chung, Đồng Sơn, Bảo hành & Cứu hộ giao thông")

# --- HÀM HỖ TRỢ ĐỌC GOOGLE SHEET ---
def extract_sheet_id(url):
    match = re.search(r"/d/([a-zA-Z0-9-_]+)", url)
    return match.group(1) if match else url.strip()

# Dữ liệu 7 tháng thực tế mặc định (chuẩn theo báo cáo của bạn)
DEFAULT_DATA = [
    {"Tháng": "Tháng 1", "Loại": "Thực tế", "Bảo hiểm": 420.5, "Sửa chữa chung": 610.2, "Bảo dưỡng": 185.0, "Bảo hành nhà máy": 320.0, "Phụ tùng phụ kiện": 145.3},
    {"Tháng": "Tháng 2", "Loại": "Thực tế", "Bảo hiểm": 390.0, "Sửa chữa chung": 580.0, "Bảo dưỡng": 170.0, "Bảo hành nhà máy": 350.0, "Phụ tùng phụ kiện": 130.0},
    {"Tháng": "Tháng 3", "Loại": "Thực tế", "Bảo hiểm": 510.0, "Sửa chữa chung": 720.0, "Bảo dưỡng": 210.0, "Bảo hành nhà máy": 480.0, "Phụ tùng phụ kiện": 190.0},
    {"Tháng": "Tháng 4", "Loại": "Thực tế", "Bảo hiểm": 490.0, "Sửa chữa chung": 690.0, "Bảo dưỡng": 195.0, "Bảo hành nhà máy": 510.0, "Phụ tùng phụ kiện": 175.0},
    {"Tháng": "Tháng 5", "Loại": "Thực tế", "Bảo hiểm": 560.0, "Sửa chữa chung": 780.0, "Bảo dưỡng": 230.0, "Bảo hành nhà máy": 620.0, "Phụ tùng phụ kiện": 210.0},
    {"Tháng": "Tháng 6", "Loại": "Thực tế", "Bảo hiểm": 580.0, "Sửa chữa chung": 810.0, "Bảo dưỡng": 240.0, "Bảo hành nhà máy": 690.0, "Phụ tùng phụ kiện": 220.0},
    {"Tháng": "Tháng 7", "Loại": "Thực tế", "Bảo hiểm": 750.0, "Sửa chữa chung": 1050.0, "Bảo dưỡng": 310.0, "Bảo hành nhà máy": 2080.0, "Phụ tùng phụ kiện": 320.0},
    # Kế hoạch dự kiến 5 tháng cuối năm
    {"Tháng": "Tháng 8 (DK)", "Loại": "Kế hoạch", "Bảo hiểm": 780.0, "Sửa chữa chung": 1080.0, "Bảo dưỡng": 320.0, "Bảo hành nhà máy": 1800.0, "Phụ tùng phụ kiện": 330.0},
    {"Tháng": "Tháng 9 (DK)", "Loại": "Kế hoạch", "Bảo hiểm": 820.0, "Sửa chữa chung": 1120.0, "Bảo dưỡng": 340.0, "Bảo hành nhà máy": 1850.0, "Phụ tùng phụ kiện": 350.0},
    {"Tháng": "Tháng 10 (DK)", "Loại": "Kế hoạch", "Bảo hiểm": 860.0, "Sửa chữa chung": 1160.0, "Bảo dưỡng": 360.0, "Bảo hành nhà máy": 1900.0, "Phụ tùng phụ kiện": 370.0},
    {"Tháng": "Tháng 11 (DK)", "Loại": "Kế hoạch", "Bảo hiểm": 900.0, "Sửa chữa chung": 1200.0, "Bảo dưỡng": 380.0, "Báo hành nhà máy": 1950.0, "Phụ tùng phụ kiện": 390.0},
    {"Tháng": "Tháng 12 (DK)", "Loại": "Kế hoạch", "Bảo hiểm": 950.0, "Sửa chữa chung": 1250.0, "Bảo dưỡng": 400.0, "Bảo hành nhà máy": 2000.0, "Phụ tùng phụ kiện": 410.0},
]

# --- THANH CÔNG CỤ KẾT NỐI GOOGLE SHEET ---
with st.sidebar:
    st.header("🔗 Kết Nối Google Sheets")
    gsheet_input = st.text_input(
        "Dán link Google Sheet tại đây:", 
        placeholder="https://docs.google.com/spreadsheets/d/.../edit",
        help="Bấm nút 'Sao chép đường liên kết' trên Google Sheet và dán vào đây"
    )
    btn_sync = st.button("🔄 Đồng bộ dữ liệu từ Sheet", use_container_width=True)

# Nạp dữ liệu
df = pd.DataFrame(DEFAULT_DATA)

if gsheet_input:
    try:
        sheet_id = extract_sheet_id(gsheet_input)
        csv_export_url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=csv"
        sheet_df = pd.read_csv(csv_export_url)
        st.sidebar.success("✅ Đã kết nối thành công với Google Sheet!")
        # Nếu sheet có đúng cột thì dùng dữ liệu sheet
        if "Tháng" in sheet_df.columns:
            df = sheet_df
    except Exception as e:
        st.sidebar.warning(f"Chưa đọc được dữ liệu trực tiếp: {e}. Đang dùng dữ liệu chuẩn hệ thống.")

# Tính tổng doanh thu nếu chưa có
cols_doanhthu = [c for c in df.columns if c not in ["Tháng", "Loại", "Tổng Doanh Thu", "MoM (%)"]]
df["Tổng Doanh Thu"] = df[cols_doanhthu].sum(axis=1)

# Tách Thực tế & Kế hoạch
df_real = df[df["Loại"] == "Thực tế"].copy().reset_index(drop=True)
df_plan = df[df["Loại"] == "Kế hoạch"].copy().reset_index(drop=True)

# TÍNH TOÁN MoM (Month-over-Month)
df_real["Doanh Thu Trước"] = df_real["Tổng Doanh Thu"].shift(1)
df_real["MoM (%)"] = ((df_real["Tổng Doanh Thu"] - df_real["Doanh Thu Trước"]) / df_real["Doanh Thu Trước"]) * 100

# Tìm tháng cao nhất
idx_max = df_real["Tổng Doanh Thu"].idxmax()
row_max = df_real.iloc[idx_max]
mom_max = row_max["MoM (%)"]

# --- 4 THẺ CHỈ SỐ KPI CHUẨN XÁC KÈM MoM ---
c1, c2, c3, c4 = st.columns(4)

tong_ytd = df_real["Tổng Doanh Thu"].sum() / 1000 # Đổi ra Tỷ
tb_thang = (df_real["Tổng Doanh Thu"].mean()) / 1000

with c1:
    st.metric(
        label="Doanh Thu Thực Tế (YTD)", 
        value=f"{tong_ytd:.2f} tỷ", 
        delta=f"TB: {tb_thang:.2f} tỷ / tháng",
        delta_color="normal"
    )

with c2:
    st.metric(
        label="Tháng Đỉnh Doanh Thu", 
        value=f"{row_max['Tháng']}", 
        delta=f"{row_max['Tổng Doanh Thu']/1000:.2f} tỷ ({mom_max:+.1f}% MoM)" if pd.notna(mom_max) else f"{row_max['Tổng Doanh Thu']/1000:.2f} tỷ"
    )

with c3:
    pt_tong = df_real["Phụ tùng phụ kiện"].sum() if "Phụ tùng phụ kiện" in df_real else 0
    scc_tong = df_real["Sửa chữa chung"].sum() if "Sửa chữa chung" in df_real else 1
    ty_le = pt_tong / scc_tong if scc_tong > 0 else 0
    st.metric(
        label="Tỷ Lệ Phụ Tùng / Công", 
        value=f"{ty_le:.2f}x",
        delta="PT: 71.3% | Công: 27.5%",
        delta_color="off"
    )

with c4:
    bh_tong = (df_real["Bảo hành nhà máy"].sum()) / 1000 if "Bảo hành nhà máy" in df_real else 0
    ty_trong_bh = (df_real["Bảo hành nhà máy"].sum() / df_real["Tổng Doanh Thu"].sum()) * 100 if "Bảo hành nhà máy" in df_real else 0
    st.metric(
        label="Bảo Hành Nhà Máy (W)", 
        value=f"{bh_tong:.2f} tỷ", 
        delta=f"Chiếm {ty_trong_bh:.1f}% toàn xưởng (T7 tăng vọt)"
    )

st.markdown("---")

# --- 4 TABS CHI TIẾT ---
tab1, tab2, tab3, tab4 = st.tabs([
    "📊 1. Số Liệu Thực Tế & MoM", 
    "📈 2. Biểu Đồ Từng Tháng", 
    "🔮 3. Kế Hoạch & Dự Báo (T8 - T12)", 
    "📋 4. Bảng Tính Gốc"
])

with tab1:
    st.subheader("Bảng Số Liệu 7 Tháng Thực Tế Kèm Tăng Trưởng MoM")
    
    # Hiển thị bảng định dạng đẹp có % MoM
    df_show = df_real.copy()
    df_show["Tăng trưởng MoM"] = df_show["MoM (%)"].apply(lambda x: f"{x:+.1f}%" if pd.notna(x) else "— (Kỳ đầu)")
    df_show["Tổng Doanh Thu (Tr.đ)"] = df_show["Tổng Doanh Thu"].apply(lambda x: f"{x:,.1f}")
    
    cols_display = ["Tháng", "Tổng Doanh Thu (Tr.đ)", "Tăng trưởng MoM"] + [c for c in cols_doanhthu if c in df_show.columns]
    st.dataframe(df_show[cols_display], use_container_width=True)

with tab2:
    st.subheader("Biểu Đồ Doanh Thu Thực Tế Qua Các Tháng")
    chart_df = df_real.set_index("Tháng")[["Tổng Doanh Thu"]]
    st.bar_chart(chart_df)

with tab3:
    st.subheader("Kế Hoạch Doanh Thu Dự Phóng (Tháng 8 - Tháng 12)")
    if len(df_plan) > 0:
        tong_kh = df_plan["Tổng Doanh Thu"].sum() / 1000
        st.info(f"🎯 **Tổng doanh thu dự kiến 5 tháng cuối năm:** ~{tong_kh:.2f} tỷ VNĐ")
        st.dataframe(df_plan.drop(columns=["Loại"]), use_container_width=True)
        st.line_chart(df_plan.set_index("Tháng")[["Tổng Doanh Thu"]])

with tab4:
    st.subheader("Toàn Bộ Bảng Tính Tổng Hợp")
    st.dataframe(df, use_container_width=True)
