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

st.markdown("""
    <style>
    .main { background-color: #f8f9fa; }
    .kpi-card {
        background: white; padding: 20px; border-radius: 12px;
        box-shadow: 0 4px 12px rgba(0,0,0,0.1);
        border-top: 5px solid #1f4e78; text-align: center;
    }
    .kpi-value { font-size: 28px; font-weight: bold; color: #1f4e78; margin: 5px 0; }
    .kpi-label { font-size: 13px; color: #666; font-weight: bold; text-transform: uppercase; }
    </style>
""", unsafe_allow_html=True)

# --- 2. HÀM XỬ LÝ DỮ LIỆU GỐC CỦA BẠN ---
MASTER_FILE = "master_database.xlsx"
TRANG_THAI_HOAN_THANH = ['Đã đóng', 'Sẵn sàng bàn giao']

def clean_lsc(val): return str(val).split('(')[0].strip() if pd.notna(val) else ""

@st.cache_data(show_spinner=False)
def load_master():
    if not os.path.exists(MASTER_FILE):
        return pd.DataFrame()
    df = pd.read_excel(MASTER_FILE)
    # Convert tiền
    for col in df.columns:
        if 'tiền' in col.lower() or 'thanh toán' in col.lower() or 'Giá trị' in col:
            df[col] = pd.to_numeric(df[col].astype(str).str.replace(',', ''), errors='coerce').fillna(0)
    return df

# --- 3. ĐỌC DỮ LIỆU THẬT ---
df_master = load_master()

if df_master.empty:
    st.error("❌ Không tìm thấy file dữ liệu master_database.xlsx. Vui lòng nạp dữ liệu ở Tab 2.")
else:
    # Tính toán số liệu thật
    df_ht = df_master[df_master['Trạng thái'].isin(TRANG_THAI_HOAN_THANH)]
    kh_total = len(df_ht[(df_ht['KH thanh toán'] > 0) & (df_ht['Phân loại KH'] != 'GSM Công nợ')])
    gsm_debt = len(df_ht[df_ht['Phân loại KH'] == 'GSM Công nợ'])
    bh_hang = len(df_ht[df_ht['BH hãng thanh toán'] > 0])
    bh_ins = len(df_ht[df_ht['BH thanh toán'] > 0])
    dang_lam = len(df_master[~df_master['Trạng thái'].isin(TRANG_THAI_HOAN_THANH + ['Đã hủy'])])

    # --- 4. HIỂN THỊ KPI CARD VỚI DATA THẬT ---
    st.markdown('<h2 style="text-align:center; color:#1f4e78;">🚗 QUẢN TRỊ DỊCH VỤ - VINFAST C23401</h2>', unsafe_allow_html=True)
    
    c1, c2, c3, c4, c5 = st.columns(5)
    with c1: st.markdown(f'<div class="kpi-card"><h4>🚗 KHÁCH LẺ</h4><div class="kpi-value">{kh_total}</div></div>', unsafe_allow_html=True)
    with c2: st.markdown(f'<div class="kpi-card" style="border-top-color:#ffc107"><h4>🚕 GSM NỢ</h4><div class="kpi-value">{gsm_debt}</div></div>', unsafe_allow_html=True)
    with c3: st.markdown(f'<div class="kpi-card" style="border-top-color:#17a2b8"><h4>🛡️ BẢO HÀNH</h4><div class="kpi-value">{bh_hang}</div></div>', unsafe_allow_html=True)
    with c4: st.markdown(f'<div class="kpi-card" style="border-top-color:#6f42c1"><h4>🏢 BẢO HIỂM</h4><div class="kpi-value">{bh_ins}</div></div>', unsafe_allow_html=True)
    with c5: st.markdown(f'<div class="kpi-card" style="border-top-color:#dc3545"><h4>⏳ ĐANG LÀM</h4><div class="kpi-value">{dang_lam}</div></div>', unsafe_allow_html=True)

    st.write("---")

    # --- 5. BIỂU ĐỒ VỚI DATA THẬT ---
    col_l, col_r = st.columns([2, 1])
    with col_l:
        st.subheader("📊 Doanh thu theo Cố vấn dịch vụ")
        revenue_cvdv = df_ht.groupby('Cố vấn dịch vụ')['Số tiền thanh toán cuối'].sum().reset_index()
        fig_bar = px.bar(revenue_cvdv.sort_values('Số tiền thanh toán cuối', ascending=False).head(10), 
                         x='Số tiền thanh toán cuối', y='Cố vấn dịch vụ', orientation='h',
                         color='Số tiền thanh toán cuối', color_continuous_scale='Blues')
        st.plotly_chart(fig_bar, use_container_width=True)
    
    with col_r:
        st.subheader("💰 Cơ cấu doanh thu")
        vals = [df_ht['KH thanh toán'].sum(), df_ht['BH thanh toán'].sum(), df_ht['BH hãng thanh toán'].sum()]
        names = ['Khách lẻ', 'Bảo hiểm', 'Bảo hành']
        fig_pie = px.pie(values=vals, names=names, hole=0.4)
        st.plotly_chart(fig_pie, use_container_width=True)

    # --- 6. CÁC TÍNH NĂNG QUẢN TRỊ (TABS) ---
    tabs = st.tabs(["📊 Bảng Tính Chi Tiết", "📥 Nạp DMS", "🧾 Khớp HĐ", "🔍 Cyber Đối Soát"])
    
    with tabs[0]:
        st.subheader("📝 Danh sách lệnh sửa chữa (Dữ liệu thật)")
        st.data_editor(df_master, use_container_width=True, height=500)
    
    with tabs[1]:
        st.subheader("📥 Nạp dữ liệu DMS mới")
        up_file = st.file_uploader("Chọn file Excel từ DMS", type=['xlsx'])
        if up_file:
            st.success("Đã nhận file. Bạn có thể viết tiếp logic lưu file tại đây.")

    with tabs[3]:
        st.subheader("🔍 Đối soát Cyber")
        st.info("Tính năng đối soát Cyber đang sẵn sàng.")
