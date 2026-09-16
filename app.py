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

# --- 1. CẤU HÌNH TRANG & GIAO DIỆN ---
st.set_page_config(page_title="Hệ Thống Quản Trị VinFast C23401", page_icon="🚗", layout="wide")

# CSS để biến giao diện thành Dashboard chuyên nghiệp
st.markdown("""
    <style>
    .main { background-color: #f8f9fa; }
    .main-title { font-size: 32px; font-weight: bold; color: #1f4e78; text-align: center; margin-bottom: 25px; font-family: 'Segoe UI', sans-serif; }
    .kpi-container { display: flex; justify-content: space-between; gap: 15px; margin-bottom: 25px; }
    .kpi-box { flex: 1; background: white; padding: 20px; border-radius: 12px; box-shadow: 0 4px 12px rgba(0,0,0,0.1); border-top: 6px solid #1f4e78; text-align: center; }
    .kpi-box h4 { margin: 0; font-size: 13px; color: #6c757d; text-transform: uppercase; }
    .kpi-box h2 { margin: 10px 0; font-size: 30px; color: #1f4e78; font-weight: bold; }
    .kpi-box p { margin: 0; font-size: 14px; font-weight: 600; }
    .alert-banner { background: linear-gradient(90deg, #ff4b4b 0%, #ff7676 100%); color: white; padding: 15px; border-radius: 10px; margin-bottom: 25px; box-shadow: 0 4px 10px rgba(255,75,75,0.2); }
    </style>
""", unsafe_allow_html=True)

# --- 2. HỆ THỐNG BIẾN & THAM SỐ ---
MASTER_FILE = "master_database.xlsx"
SECRET_KEY_AUTH = "VinFast_GiaLai_Secret_Key_2026"
DANH_SACH_ADMIN = {"admin": "Vinfastgialai@2026##"}

COT_DINH_DANH = ['Số lệnh sửa chữa', 'Trạng thái', 'Cố vấn dịch vụ', 'Tên khách hàng', 'Thời gian đóng LSC', 'Biển số', 'Xe GSM']
COT_TIEN = ['Tổng tiền công', 'Tổng tiền phụ tùng', 'Tổng trước chiết khấu', 'Chiết khấu đại lý', 'Tổng chiết khấu', 'Tổng sau chiết khấu', 'Tiền VAT', 'Tổng có VAT', 'Chiết khấu VinClub', 'Số tiền thanh toán cuối', 'Tiền đặt cọc', 'KH thanh toán', 'BH thanh toán', 'BH hãng thanh toán', 'Nội bộ thanh toán']
TAT_CA_COT = COT_DINH_DANH + COT_TIEN + ['Phân loại KH', 'Phê duyệt bảo hành', 'Số hóa đơn', 'Ngày xuất hóa đơn', 'Giá trị xuất hóa đơn']
TRANG_THAI_HOAN_THANH = ['Đã đóng', 'Sẵn sàng bàn giao']

# --- 3. HÀM XỬ LÝ DỮ LIỆU & AUTH (GIỮ NGUYÊN LOGIC CỦA BẠN) ---
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

# --- 4. PHẦN ĐĂNG NHẬP SIDEBAR ---
if "logged_in" not in st.session_state: st.session_state.logged_in = False
with st.sidebar:
    st.image("https://upload.wikimedia.org/wikipedia/commons/0/05/VinFast_logo.svg", width=150)
    if not st.session_state.logged_in:
        u = st.text_input("Tài khoản:"); p = st.text_input("Mật khẩu:", type="password")
        if st.button("Đăng Nhập"):
            if u in DANH_SACH_ADMIN and DANH_SACH_ADMIN[u] == p:
                st.session_state.logged_in = True; st.rerun()
            else: st.error("Sai mật khẩu!")
    else:
        st.success(f"Chào, {u if 'u' in locals() else 'Admin'}")
        if st.button("Đăng Xuất"): st.session_state.logged_in = False; st.rerun()

# --- 5. TÍNH TOÁN DỮ LIỆU DASHBOARD ---
df_master = load_master()
df_ht = df_master[df_master['Trạng thái'].isin(TRANG_THAI_HOAN_THANH)]
kh_total_cnt = len(df_ht[(df_ht['KH thanh toán'] > 0) & (df_ht['Phân loại KH'] != 'GSM Công nợ')])
gsm_debt_cnt = len(df_ht[df_ht['Phân loại KH'] == 'GSM Công nợ'])
bh_total_cnt = len(df_ht[df_ht['BH hãng thanh toán'] > 0])
bh_ins_cnt = len(df_ht[df_ht['BH thanh toán'] > 0])
chua_xuat_hd = df_ht[(df_ht['Số hóa đơn'].isna()) | (df_ht['Số hóa đơn'].astype(str).isin(['', 'nan', '0']))]

# --- 6. HIỂN THỊ DASHBOARD TRỰC QUAN ---
st.markdown('<div class="main-title">🚗 QUẢN TRỊ DỊCH VỤ & ĐỐI SOÁT VINFAST</div>', unsafe_allow_html=True)

# Hàng thẻ KPI (Màu sắc và bóng đổ chuyên nghiệp)
st.markdown(f"""
    <div class="kpi-container">
        <div class="kpi-box" style="border-top-color: #007bff;">
            <h4>KH THANH TOÁN</h4>
            <h2>{kh_total_cnt}</h2>
            <p style="color: #007bff;">Lệnh hoàn thành</p>
        </div>
        <div class="kpi-box" style="border-top-color: #ffc107;">
            <h4>GSM CÔNG NỢ</h4>
            <h2>{gsm_debt_cnt}</h2>
            <p style="color: #ffc107;">Chờ thanh toán</p>
        </div>
        <div class="kpi-box" style="border-top-color: #17a2b8;">
            <h4>BẢO HÀNH (W)</h4>
            <h2>{bh_total_cnt}</h2>
            <p style="color: #17a2b8;">Lệnh hãng</p>
        </div>
        <div class="kpi-box" style="border-top-color: #6f42c1;">
            <h4>BẢO HIỂM</h4>
            <h2>{bh_ins_cnt}</h2>
            <p style="color: #6f42c1;">Insurance</p>
        </div>
        <div class="kpi-box" style="border-top-color: #dc3545;">
            <h4>CHƯA XUẤT HĐ</h4>
            <h2>{len(chua_xuat_hd)}</h2>
            <p style="color: #dc3545;">Cần xử lý ngay</p>
        </div>
    </div>
""", unsafe_allow_html=True)

# Banner Cảnh báo nổi bật
if len(chua_xuat_hd) > 0:
    st.markdown(f"""
        <div class="alert-banner">
            <b>⚠️ CẢNH BÁO:</b> Có {len(chua_xuat_hd)} lệnh đã đóng nhưng CHƯA XUẤT HÓA ĐƠN. 
            Tổng tiền treo: {chua_xuat_hd['Số tiền thanh toán cuối'].sum():,.0f} VNĐ.
        </div>
