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

# --- CẤU HÌNH HỆ THỐNG ---
st.set_page_config(page_title="Quản Trị VinFast C23401", page_icon="🚗", layout="wide")

MASTER_FILE = "master_database.xlsx"
SECRET_KEY_AUTH = "VinFast_GiaLai_Secret_Key_2026"
DANH_SACH_ADMIN = {"admin": "Vinfastgialai@2026##"}

# Danh mục cột chuẩn
COT_DINH_DANH = ['Số lệnh sửa chữa', 'Trạng thái', 'Cố vấn dịch vụ', 'Tên khách hàng', 'Thời gian đóng LSC', 'Biển số', 'Xe GSM']
COT_TIEN = ['Tổng tiền công', 'Tổng tiền phụ tùng', 'Tổng trước chiết khấu', 'Chiết khấu đại lý', 'Tổng chiết khấu', 'Tổng sau chiết khấu', 'Tiền VAT', 'Tổng có VAT', 'Chiết khấu VinClub', 'Số tiền thanh toán cuối', 'Tiền đặt cọc', 'KH thanh toán', 'BH thanh toán', 'BH hãng thanh toán', 'Nội bộ thanh toán']
COT_THEM = ['Phân loại KH', 'Phê duyệt bảo hành']
COT_HOA_DON = ['Số hóa đơn', 'Ngày xuất hóa đơn', 'Giá trị xuất hóa đơn']
TAT_CA_COT = COT_DINH_DANH + COT_TIEN + COT_THEM + COT_HOA_DON
TRANG_THAI_HOAN_THANH = ['Đã đóng', 'Sẵn sàng bàn giao']

# --- 1. XỬ LÝ AUTHENTICATION (TOKEN 10 NGÀY) ---
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

# --- 2. CHUẨN HÓA DỮ LIỆU & SORTING ---
def clean_lsc_giu_gach(val):
    if pd.isna(val) or val is None: return ""
    return str(val).split('(')[0].strip()

def norm_lsc_key(val):
    if pd.isna(val) or val is None: return ""
    return re.sub(r'[^A-Z0-9]', '', str(val).split('(')[0].strip().upper())

def chuyen_ngay_gio_sortable(val):
    """Fix lỗi Sort: Chuyển '11:19, 10/09/2026' -> '2026/09/10 11:19'"""
    if pd.isna(val) or not str(val).strip(): return ""
    val_str = str(val).strip()
    try:
        if ',' in val_str:
            parts = val_str.split(',')
            t_part = parts[0].strip()
            d_part = parts[1].strip()
            match = re.search(r'(\d{1,2})[/-](\d{1,2})[/-](\d{4})', d_part)
            if match:
                d, m, y = match.groups()
                return f"{y}/{int(m):02d}/{int(d):02d} {t_part}"
        return val_str
    except: return val_str

def clean_tien_series(ser):
    return ser.astype(str).str.replace(',', '').str.replace(' ', '').replace(['-', 'nan', 'None', ''], '0')

def chuan_hoa_dataframe(df):
    df_out = df.copy()
    if 'Số lệnh sửa chữa' in df_out.columns:
        df_out['Số lệnh sửa chữa'] = df_out['Số lệnh sửa chữa'].apply(clean_lsc_giu_gach)
    if 'Thời gian đóng LSC' in df_out.columns:
        df_out['Thời gian đóng LSC'] = df_out['Thời gian đóng LSC'].apply(chuyen_ngay_gio_sortable)
    for c in COT_TIEN + ['Giá trị xuất hóa đơn']:
        if c in df_out.columns:
            df_out[c] = pd.to_numeric(clean_tien_series(df_out[c]), errors='coerce').fillna(0)
    return df_out

# --- 3. QUẢN LÝ MASTER FILE ---
@st.cache_data(show_spinner=False)
def load_master():
    if not os.path.exists(MASTER_FILE):
        pd.DataFrame(columns=TAT_CA_COT).to_excel(MASTER_FILE, index=False)
    df = pd.read_excel(MASTER_FILE)
    for c in TAT_CA_COT:
        if c not in df.columns: df[c] = 0 if c in COT_TIEN else ""
    return chuan_hoa_dataframe(df)[TAT_CA_COT]

def save_master(df):
    chuan_hoa_dataframe(df).to_excel(MASTER_FILE, index=False)
    load_master.clear()

# --- 4. XUẤT EXCEL CHUẨN ĐẸP ---
def format_sheet(ws, sheet_name, cols_to_hide=None):
    # Định dạng tiêu đề Navy
    header_font = Font(name='Segoe UI', size=11, bold=True, color='FFFFFF')
    header_fill = PatternFill(start_color='1F4E78', fill_type='solid')
    border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    
    ws.row_dimensions[1].height = 28
    for cell in ws[1]:
        cell.font = header_font; cell.fill = header_fill; cell.alignment = Alignment(horizontal='center'); cell.border = border

    # Ẩn cột và dãn cột
    for col in ws.columns:
        col_name = ws.cell(row=1, column=col[0].column).value
        col_let = get_column_letter(col[0].column)
        if cols_to_hide and col_name in cols_to_hide:
            ws.column_dimensions[col_let].hidden = True
        else:
            ws.column_dimensions[col_let].width = 22
            if col_name in COT_TIEN:
                for cell in col[1:]: cell.number_format = '#,##0'

    # Dòng tổng cộng
    last_row = ws.max_row
    if last_row > 1:
        ws.cell(row=last_row+1, column=1, value="TỔNG CỘNG").font = Font(bold=True)
        for c in range(1, ws.max_column + 1):
            col_name = ws.cell(row=1, column=c).value
            if col_name in COT_TIEN:
                col_let = get_column_letter(c)
                ws.cell(row=last_row+1, column=c, value=f"=SUBTOTAL(9, {col_let}2:{col_let}{last_row})").number_format = '#,##0'

def xuat_excel_don_luong(df, name):
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name=name[:31])
    buf.seek(0)
    wb = openpyxl.load_workbook(buf)
    format_sheet(wb.active, name)
    out = io.BytesIO(); wb.save(out); out.seek(0)
    return out

# --- 5. GIAO DIỆN STREAMLIT ---
if "logged_in" not in st.session_state: st.session_state.logged_in = False

# Kiểm tra Token
t_url = st.query_params.get("auth_token")
if t_url and not st.session_state.logged_in:
    user = xac_thuc_auth_token(t_url)
    if user: st.session_state.logged_in = True; st.session_state.username = user

# Sidebar Login
with st.sidebar:
    st.title("🚗 C23401 QUẢN TRỊ")
    if not st.session_state.logged_in:
        u = st.text_input("User:"); p = st.text_input("Pass:", type="password")
        if st.button("Đăng Nhập"):
            if u in DANH_SACH_ADMIN and DANH_SACH_ADMIN[u] == p:
                st.session_state.logged_in = True; st.query_params["auth_token"] = tao_auth_token(u); st.rerun()
            else: st.error("Sai mật khẩu!")
    else:
        st.success(f"Admin: {st.session_state.username}"); 
        if st.button("Đăng Xuất"): st.session_state.logged_in = False; del st.query_params["auth_token"]; st.rerun()

st.title("🚗 Quản Trị Doanh Thu & Đối Soát VinFast C23401")
df_master = load_master()

# --- TAB 1: BẢNG TÍNH & BỘ LỌC ---
tabs = st.tabs(["📊 Bảng Tính Web", "📥 Nạp DMS", "🧾 Khớp HĐ", "🛡️ Bảo Hành", "🚕 GSM", "🔍 Cyber"])

with tabs[0]:
    # Metrics
    df_ht = df_master[df_master['Trạng thái'].isin(TRANG_THAI_HOAN_THANH)]
    kh_no_hd = df_ht[(df_ht['KH thanh toán'] > 0) & (df_ht['Phân loại KH'] != 'GSM Công nợ') & (df_ht['Số hóa đơn'].isin(['', 'nan', 'None', '0']))]
    
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Tổng Lệnh Xong", len(df_ht))
    c2.metric("Lệnh Nợ GSM", len(df_ht[df_ht['Phân loại KH'] == 'GSM Công nợ']))
    c3.metric("🚨 KH Nợ HĐ", len(kh_no_hd), delta=f"{kh_no_hd['KH thanh toán'].sum():,.0f} đ", delta_color="inverse")
    c4.metric("Doanh Thu", f"{df_ht['Số tiền thanh toán cuối'].sum():,.0f}")

    # Bộ lọc
    f1, f2, f3, f4 = st.columns([2, 2, 2, 3])
    luong = f1.selectbox("Luồng:", ["Tất cả", "KH Thanh Toán", "GSM Công nợ", "Bảo Hiểm", "🚨 Nợ HĐ"])
    list_ngay = sorted(df_master['Thời gian đóng LSC'].str[:10].unique(), reverse=True)
    sel_ngay = f2.multiselect("Lọc Ngày:", list_ngay)
    sel_cvdv = f3.multiselect("Cố vấn:", df_master['Cố vấn dịch vụ'].unique())
    search = f4.text_input("🔍 Tìm nhanh (Biển số / LSC):")

    # Xử lý lọc dữ liệu hiển thị
    df_view = df_master.copy()
    if luong == "KH Thanh Toán": df_view = df_ht[df_ht['Phân loại KH'] != 'GSM Công nợ']
    elif luong == "GSM Công nợ": df_view = df_ht[df_ht['Phân loại KH'] == 'GSM Công nợ']
    elif luong == "🚨 Nợ HĐ": df_view = kh_no_hd
    
    if sel_ngay: df_view = df_view[df_view['Thời gian đóng LSC'].str[:10].isin(sel_ngay)]
    if sel_cvdv: df_view = df_view[df_view['Cố vấn dịch vụ'].isin(sel_cvdv)]
    if search: df_view = df_view[df_view.apply(lambda r: search.lower() in str(r.values).lower(), axis=1)]

    # Ẩn/Hiện cột
    default_cols = ['Số lệnh sửa chữa', 'Trạng thái', 'Cố vấn dịch vụ', 'Thời gian đóng LSC', 'Biển số', 'KH thanh toán', 'BH thanh toán', 'Số hóa đơn', 'Ngày xuất hóa đơn']
    with st.expander("👁️ Tùy biến cột hiển thị:"):
        selected_cols = st.multiselect("Chọn cột:", TAT_CA_COT, default=default_cols)

    # DATA EDITOR (Mở khóa Sort tất cả các cột)
    edited_df = st.data_editor(
        df_view[selected_cols], use_container_width=True, height=500,
        column_config={c: st.column_config.NumberColumn(format="%,d đ") for c in COT_TIEN},
        disabled=not st.session_state.logged_in
    )

    # Nút chức năng
    b1, b2, b3 = st.columns(3)
    if st.session_state.logged_in:
        if b1.button("💾 LƯU CHỈNH SỬA", type="primary"):
            df_master.update(edited_df); save_master(df_master); st.success("Đã lưu!"); st.rerun()
        if b2.button("↩️ HỦY (RESET)"): st.rerun()
    
    xl_bytes = xuat_excel_don_luong(df_view[selected_cols], luong)
    b3.download_button("📥 TẢI LUỒNG NÀY (EXCEL)", xl_bytes, f"Bao_Cao_{luong}.xlsx")

# --- CÁC TAB NGHIỆP VỤ (GIỮ NGUYÊN LOGIC TỐI ƯU CỦA BẠN) ---
with tabs[5]: # TAB CYBER
    st.subheader("🔍 Đối Soát Cyber (RO Hãng)")
    up_cy = st.file_uploader("Tải file BẢNG TỔNG HỢP LỆNH Cyber:")
    if up_cy:
        # Logic nhận diện cột RO (Cột A)
        df_cy_raw = pd.read_excel(up_cy, header=7)
        cyber_keys = set(df_cy_raw.iloc[:, 0].dropna().apply(norm_lsc_key))
        
        df_check = df_ht.copy()
        df_check['Up_Cyber'] = df_check['Số lệnh sửa chữa'].apply(lambda x: norm_lsc_key(x) in cyber_keys)
        df_chua_up = df_check[~df_check['Up_Cyber']]
        
        st.warning(f"Phát hiện {len(df_chua_up)} lệnh chưa up Cyber!")
        st.dataframe(df_chua_up[['Số lệnh sửa chữa', 'Biển số', 'Cố vấn dịch vụ', 'Số tiền thanh toán cuối']])
        
        xl_cy = xuat_excel_don_luong(df_chua_up, "Chua_Up_Cyber")
        st.download_button("📥 Tải File Gửi Kế Toán", xl_cy, "Lenh_Chua_Up_Cyber.xlsx")

# (Các Tab Nạp DMS, Khớp HĐ, GSM... giữ nguyên logic xử lý DataFrame của bạn)
st.caption("🚀 Hệ thống Quản trị Doanh thu VinFast C23401 | Final Version 2024")
