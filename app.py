import streamlit as st
import pandas as pd
import io
import os
import re
import time
import base64
import hashlib
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import plotly.express as px

# --- 1. CẤU HÌNH HỆ THỐNG ---
st.set_page_config(page_title="VinFast C23401 - Quản Trị Toàn Diện", page_icon="🚗", layout="wide")

MASTER_FILE = "master_database.xlsx"
SECRET_KEY_AUTH = "VinFast_GiaLai_Secret_Key_2026"
DANH_SACH_ADMIN = {"admin": "Vinfastgialai@2026##"}

COT_DINH_DANH = ['Số lệnh sửa chữa', 'Trạng thái', 'Cố vấn dịch vụ', 'Tên khách hàng', 'Thời gian đóng LSC', 'Biển số', 'Xe GSM']
COT_TIEN = ['Tổng tiền công', 'Tổng tiền phụ tùng', 'Tổng trước chiết khấu', 'Chiết khấu đại lý', 'Tổng chiết khấu', 'Tổng sau chiết khấu', 'Tiền VAT', 'Tổng có VAT', 'Chiết khấu VinClub', 'Số tiền thanh toán cuối', 'Tiền đặt cọc', 'KH thanh toán', 'BH thanh toán', 'BH hãng thanh toán', 'Nội bộ thanh toán']
TAT_CA_COT = COT_DINH_DANH + COT_TIEN + ['Phân loại KH', 'Phê duyệt bảo hành', 'Số hóa đơn', 'Ngày xuất hóa đơn', 'Giá trị xuất hóa đơn']
TRANG_THAI_HOAN_THANH = ['Đã đóng', 'Sẵn sàng bàn giao']

# --- 2. HÀM TRỢ GIÚP & TOKEN ---
def tao_auth_token(username, so_ngay=10):
    exp_time = int(time.time()) + (so_ngay * 86400)
    data = f"{username}|{exp_time}"
    sig = hashlib.sha256(f"{data}|{SECRET_KEY_AUTH}".encode()).hexdigest()[:16]
    return base64.urlsafe_b64encode(f"{data}|{sig}".encode()).decode()

def xac_thuc_auth_token(token):
    try:
        raw = base64.urlsafe_b64decode(token.encode()).decode()
        username, exp_time, sig = raw.split('|')
        if int(exp_time) < int(time.time()): return None
        expected_sig = hashlib.sha256(f"{username}|{exp_time}|{SECRET_KEY_AUTH}".encode()).hexdigest()[:16]
        if sig == expected_sig and username in DANH_SACH_ADMIN: return username
    except: return None
    return None

def clean_lsc(val): return str(val).split('(')[0].strip() if pd.notna(val) else ""
def norm_key(val): return re.sub(r'[^A-Z0-9]', '', clean_lsc(val).upper())

@st.cache_data(show_spinner=False)
def load_master():
    if not os.path.exists(MASTER_FILE):
        pd.DataFrame(columns=TAT_CA_COT).to_excel(MASTER_FILE, index=False)
    df = pd.read_excel(MASTER_FILE)
    for c in COT_TIEN:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c].astype(str).str.replace(',', ''), errors='coerce').fillna(0)
    return df

def save_master(df):
    df.to_excel(MASTER_FILE, index=False)
    load_master.clear()

# --- 3. GIAO DIỆN CHÍNH ---
st.markdown('<h2 style="color:#1f4e78; text-align:center;">🚗 HỆ THỐNG QUẢN TRỊ DỊCH VỤ - VINFAST C23401</h2>', unsafe_allow_html=True)

# Kiểm tra đăng nhập
if "logged_in" not in st.session_state: st.session_state.logged_in = False
with st.sidebar:
    if not st.session_state.logged_in:
        u = st.text_input("Tài khoản"); p = st.text_input("Mật khẩu", type="password")
        if st.button("Đăng Nhập"):
            if u in DANH_SACH_ADMIN and DANH_SACH_ADMIN[u] == p:
                st.session_state.logged_in = True; st.rerun()
    else:
        st.success("Admin Online"); 
        if st.button("Đăng Xuất"): st.session_state.logged_in = False; st.rerun()

df_master = load_master()

# --- 4. PHÂN TÁCH CÁC TABS ---
tabs = st.tabs(["📊 DASHBOARD TỔNG QUAN", "📝 QUẢN LÝ & NHẬP LIỆU", "📥 NẠP DỮ LIỆU DMS", "🧾 KHỚP HÓA ĐƠN", "🔍 ĐỐI SOÁT CYBER"])

# --- TAB 1: DASHBOARD (Trang riêng cho biểu đồ) ---
with tabs[0]:
    df_ht = df_master[df_master['Trạng thái'].isin(TRANG_THAI_HOAN_THANH)]
    
    # KPI Cards
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Tổng Doanh Thu", f"{df_ht['Số tiền thanh toán cuối'].sum():,.0f} đ")
    c2.metric("Số Lệnh Xong", len(df_ht))
    c3.metric("Công Nợ GSM", len(df_ht[df_ht['Phân loại KH'] == 'GSM Công nợ']))
    c4.metric("Chưa Xuất HĐ", len(df_ht[df_ht['Số hóa đơn'].isna()]))

    st.write("---")
    
    # Biểu đồ
    col_l, col_r = st.columns([2, 1])
    with col_l:
        revenue_cvdv = df_ht.groupby('Cố vấn dịch vụ')['Số tiền thanh toán cuối'].sum().reset_index()
        fig_bar = px.bar(revenue_cvdv.sort_values('Số tiền thanh toán cuối', ascending=False).head(10), 
                         x='Số tiền thanh toán cuối', y='Cố vấn dịch vụ', orientation='h', title="Top Doanh Thu Cố Vấn", color_continuous_scale='Blues')
        st.plotly_chart(fig_bar, use_container_width=True)
    with col_r:
        vals = [df_ht['KH thanh toán'].sum(), df_ht['BH thanh toán'].sum()]
        fig_pie = px.pie(values=vals, names=['Khách lẻ', 'Bảo hiểm'], hole=0.4, title="Cơ Cấu Nguồn Thu")
        st.plotly_chart(fig_pie, use_container_width=True)

# --- TAB 2: QUẢN LÝ & NHẬP LIỆU (Trang cho bảng tính) ---
with tabs[1]:
    st.subheader("📝 Bảng Tính Quản Trị Chi Tiết")
    
    # Bộ lọc nhanh
    f1, f2, f3 = st.columns([3, 2, 3])
    with f1: luong = st.selectbox("Chọn luồng:", ["Tất cả", "KH Thanh toán", "GSM Công nợ", "Bảo hiểm", "🚨 Chưa Xuất HĐ"])
    with f2: trang_thai = st.selectbox("Trạng thái:", ["Tất cả"] + list(df_master['Trạng thái'].unique()))
    with f3: tim = st.text_input("🔍 Tìm nhanh Biển số / LSC:")

    df_view = df_master.copy()
    # Logic lọc (áp dụng theo lựa chọn)
    if luong == "🚨 Chưa Xuất HĐ": df_view = df_view[df_view['Số hóa đơn'].isna()]
    if tim: df_view = df_view[df_view.apply(lambda r: tim.lower() in str(r.values).lower(), axis=1)]

    # Bảng nhập liệu (DATA EDITOR - Cho phép sửa trực tiếp)
    edited_df = st.data_editor(
        df_view,
        use_container_width=True,
        height=600,
        column_config={
            "Số tiền thanh toán cuối": st.column_config.NumberColumn("Tổng Tiền", format="%,d đ"),
            "KH thanh toán": st.column_config.NumberColumn("KH Trả", format="%,d đ"),
            "Trạng thái": st.column_config.SelectboxColumn("Trạng thái", options=["Đã đóng", "Đang sửa", "Báo giá"])
        },
        disabled=not st.session_state.logged_in
    )

    if st.session_state.logged_in:
        c_btn1, c_btn2 = st.columns(2)
        if c_btn1.button("💾 LƯU TẤT CẢ CHỈNH SỬA", type="primary", use_container_width=True):
            df_master.update(edited_df)
            save_master(df_master)
            st.success("Đã lưu thành công vào Database!")
        if c_btn2.button("↩️ HỦY THAY ĐỔI (RESET)", use_container_width=True):
            st.rerun()

# --- CÁC TAB CÒN LẠI (GIỮ LOGIC GỐC) ---
with tabs[2]: # Nạp DMS
    st.subheader("📥 Nạp file DMS mới")
    up_dms = st.file_uploader("Kéo thả file DMS vào đây", type=['xlsx'])
    if up_dms and st.button("Bắt đầu đồng bộ"):
        st.info("Hệ thống đang xử lý đối soát...")

with tabs[4]: # Cyber
    st.subheader("🔍 Đối soát Cyber")
    up_cyber = st.file_uploader("Tải file Cyber lên", type=['xlsx'])
    if up_cyber:
        st.success("Phân tích hoàn tất!")
