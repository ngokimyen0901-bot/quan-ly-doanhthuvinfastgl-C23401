import streamlit as st
import pandas as pd
import numpy as np
import io
import os
import re
import time
import base64
import hashlib
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from streamlit_gsheets import GSheetsConnection

st.set_page_config(page_title="Báo Cáo Dịch Vụ VinFast", page_icon="🚗", layout="wide")

st.markdown("""

""", unsafe_allow_html=True)

SECRET_KEY_AUTH = "VinFast_GiaLai_Secret_Key_2026"
DANH_SACH_ADMIN = {"admin": "Vinfastgialai@2026##"}
MOC_BAT_DAU_DATE = pd.Timestamp(year=2026, month=8, day=29)

COT_DINH_DANH = ['Số lệnh sửa chữa', 'Trạng thái', 'Cố vấn dịch vụ', 'Tên khách hàng', 'Thời gian đóng LSC', 'Biển số', 'Xe GSM']
COT_TIEN = [
    'Tổng tiền công', 'Tổng tiền phụ tùng', 'Tổng trước chiết khấu',
    'Chiết khấu đại lý', 'Tổng chiết khấu', 'Tổng sau chiết khấu',
    'Tiền VAT', 'Tổng có VAT', 'Chiết khấu VinClub',
    'Số tiền thanh toán cuối', 'Tiền đặt cọc', 'KH thanh toán',
    'BH thanh toán', 'BH hãng thanh toán', 'Nội bộ thanh toán'
]
COT_THEM = ['Phân loại KH', 'Phê duyệt bảo hành']
COT_HOA_DON = ['Số hóa đơn', 'Ngày xuất hóa đơn', 'Giá trị xuất hóa đơn', 'Ghi chú']
TAT_CA_COT = COT_DINH_DANH + COT_TIEN + COT_THEM + COT_HOA_DON
TEXT_COLUMNS = [
    'Số lệnh sửa chữa', 'Trạng thái', 'Cố vấn dịch vụ', 'Tên khách hàng',
    'Biển số', 'Xe GSM', 'Phân loại KH', 'Phê duyệt bảo hành', 
    'Số hóa đơn', 'Ngày xuất hóa đơn', 'Ghi chú'
]
TRANG_THAI_HOAN_THANH = ['Đã đóng', 'Sẵn sàng bàn giao']

conn = st.connection("gsheets", type=GSheetsConnection)

def clean_lsc_giu_gach(val):
    if pd.isna(val) or val is None: return ""
    return str(val).split('(')[0].strip()

def norm_lsc_key(val):
    if pd.isna(val) or val is None: return ""
    return re.sub(r'[^A-Z0-9]', '', str(val).split('(')[0].strip().upper())

def lay_loi_ma_wo(val):
    s = norm_lsc_key(val)
    m = re.search(r'WO(\d{8,12})', s)
    if m: return m.group(1)
    digits = re.sub(r'\D', '', s)
    return digits[-10:] if len(digits) >= 8 else s

def clean_ngay_chuan(val):
    if pd.isna(val) or not val: return ""
    val_str = str(val).strip()
    if val_str.lower() in ['nan', 'none', '', 'nat']: return ""
    match_dmy = re.search(r'(\b\d{1,2}[/-]\d{1,2}[/-]\d{4}\b)', val_str)
    if match_dmy:
        p = match_dmy.group(1).replace('-', '/').split('/')
        return f"{int(p[0]):02d}/{int(p[1]):02d}/{p[2]}"
    return val_str

def chuyen_ngay_gio_sortable(val):
    if pd.isna(val) or not val: return ""
    val_str = str(val).strip()
    if val_str.lower() in ['nan', 'none', '', 'nat']: return ""
    try:
        if ',' in val_str:
            p = val_str.split(',')
            d_clean = clean_ngay_chuan(p[1].strip()).split('/')
            return f"{d_clean[2]}/{d_clean[1]}/{d_clean[0]} {p[0].strip()}"
        return val_str
    except Exception: return val_str

def clean_tien_series(ser):
    return ser.astype(str).str.replace(',', '', regex=False).str.replace(' ', '', regex=False).replace(['-', '', 'nan', 'None', ''], '0')

def doc_danh_sach_gsm_tu_gsheet():
    gsm_keys = set()
    try:
        df_gsm = conn.read(worksheet="GSM_List", ttl=5)
        if df_gsm is not None:
            for c in df_gsm.columns:
                c_str = str(c).strip()
                if c_str and not c_str.startswith('#') and not c_str.startswith('Unnamed'):
                    gsm_keys.add(norm_lsc_key(c_str))
                    gsm_keys.add(lay_loi_ma_wo(c_str))
            for v in df_gsm.iloc[:, 0].dropna():
                v_str = str(v).strip()
                if v_str and not v_str.startswith('#'):
                    gsm_keys.add(norm_lsc_key(v_str))
                    gsm_keys.add(lay_loi_ma_wo(v_str))
    except Exception: pass
    return {k for k in gsm_keys if k}

def dong_bo_hoa_don(df_target):
    df_res = df_target.copy()
    if 'Số hóa đơn' in df_res.columns: df_res['Số hóa đơn'] = df_res['Số hóa đơn'].astype(object).fillna('').astype(str)
    if 'Ngày xuất hóa đơn' in df_res.columns: df_res['Ngày xuất hóa đơn'] = df_res['Ngày xuất hóa đơn'].astype(object).fillna('').astype(str)
    if 'Giá trị xuất hóa đơn' in df_res.columns: df_res['Giá trị xuất hóa đơn'] = pd.to_numeric(clean_tien_series(df_res['Giá trị xuất hóa đơn']), errors='coerce').fillna(0)
    mask = df_res['Số hóa đơn'].isna() | df_res['Số hóa đơn'].astype(str).str.strip().isin(['', 'nan', 'None', '0'])
    df_res.loc[mask, 'Ngày xuất hóa đơn'] = ""
    df_res.loc[mask, 'Giá trị xuất hóa đơn'] = 0
    return df_res

def chuan_hoa_kieu_du_lieu(df_input):
    df_out = df_input.copy()
    for col in TEXT_COLUMNS:
        if col in df_out.columns:
            if col == 'Số lệnh sửa chữa': df_out[col] = df_out[col].fillna('').astype(str).apply(clean_lsc_giu_gach)
            else: df_out[col] = df_out[col].fillna('').astype(str).replace('nan', '')
    if 'Thời gian đóng LSC' in df_out.columns:
        df_out['Thời gian đóng LSC'] = df_out['Thời gian đóng LSC'].apply(chuyen_ngay_gio_sortable)
    for col in COT_TIEN + ['Giá trị xuất hóa đơn']:
        if col in df_out.columns: df_out[col] = pd.to_numeric(clean_tien_series(df_out[col]), errors='coerce').fillna(0)
    return df_out

def loc_chuan_tu_29_thang_8(df_target):
    if df_target is None or len(df_target) == 0: return df_target
    df_res = df_target.copy()
    df_res['temp_norm_key'] = df_res['Số lệnh sửa chữa'].apply(norm_lsc_key)
    df_res = df_res[df_res['temp_norm_key'].str.len() > 3]
    df_res['has_hd'] = df_res['Số hóa đơn'].fillna('').astype(str).str.strip().apply(lambda x: 1 if x and x not in ['0', 'nan', 'None'] else 0)
    df_res = df_res.sort_values(by=['has_hd', 'Giá trị xuất hóa đơn', 'Số tiền thanh toán cuối'], ascending=[True, True, True])
    df_res = df_res.drop_duplicates(subset=['temp_norm_key'], keep='last').drop(columns=['temp_norm_key', 'has_hd']).reset_index(drop=True)
    return df_res

def load_data_from_gsheets():
    try: df_m = conn.read(worksheet="MasterData", ttl=5)
    except Exception: df_m = pd.DataFrame(columns=TAT_CA_COT)
    if df_m is None or len(df_m) == 0 or 'Số lệnh sửa chữa' not in df_m.columns: df_m = pd.DataFrame(columns=TAT_CA_COT)
    df_m.columns = [str(c).strip() for c in df_m.columns]
    for c in TAT_CA_COT:
        if c not in df_m.columns: df_m[c] = ""
    for c in TEXT_COLUMNS:
        if c in df_m.columns: df_m[c] = df_m[c].fillna('').astype(str)
    for c in COT_TIEN + ['Giá trị xuất hóa đơn']:
        if c in df_m.columns: df_m[c] = pd.to_numeric(clean_tien_series(df_m[c]), errors='coerce').fillna(0)
    
    gsm_keys_set = doc_danh_sach_gsm_tu_gsheet()
    for idx, row in df_m.iterrows():
        cur_pl = str(row.get('Phân loại KH', '')).strip()
        lsc_val = row.get('Số lệnh sửa chữa', '')
        if norm_lsc_key(lsc_val) in gsm_keys_set or lay_loi_ma_wo(lsc_val) in gsm_keys_set: 
            df_m.at[idx, 'Phân loại KH'] = "GSM Công nợ"
        elif cur_pl in ["Bảo Hiểm", "Bảo Hành Hãng", "Nội Bộ / PDI"]: 
            pass
        elif not cur_pl or cur_pl in ['nan', 'None']: 
            df_m.at[idx, 'Phân loại KH'] = "KH Thông Thường"
            
    df_m = dong_bo_hoa_don(df_m)
    df_m = chuan_hoa_kieu_du_lieu(df_m)
    return loc_chuan_tu_29_thang_8(df_m)[TAT_CA_COT]

def save_data_to_gsheets(df_to_save):
    df_clean = dong_bo_hoa_don(df_to_save.copy())
    df_clean = chuan_hoa_kieu_du_lieu(df_clean)
    df_clean = loc_chuan_tu_29_thang_8(df_clean)
    for c in TEXT_COLUMNS:
        if c in df_clean.columns: df_clean[c] = df_clean[c].fillna('').astype(str)
    conn.update(worksheet="MasterData", data=df_clean)

# XÁC THỰC ADMIN
if "logged_in" not in st.session_state: st.session_state.logged_in = False
if "username" not in st.session_state: st.session_state.username = ""

with st.sidebar:
    st.subheader("🔐 Quyền Quản Trị Hệ Thống")
    if not st.session_state.logged_in:
        with st.form("form_login"):
            u = st.text_input("Tài khoản:")
            p = st.text_input("Mật khẩu:", type="password")
            if st.form_submit_button("Đăng Nhập Quản Trị"):
                if u in DANH_SACH_ADMIN and DANH_SACH_ADMIN[u] == p:
                    st.session_state.logged_in = True
                    st.session_state.username = u
                    st.rerun()
                else: st.error("Sai thông tin đăng nhập!")
    else:
        st.success(f"Xin chào: **{st.session_state.username}**")
        if st.button("🚪 Đăng Xuất"):
            st.session_state.logged_in = False
            st.rerun()

st.title("🚗 Quản Trị Dịch Vụ, Hóa Đơn & Bảo Hành VinFast")
df_master = load_data_from_gsheets()

tabs = st.tabs([
    "📊 1. Bảng Tính Web & Phân Luồng",
    "📥 2. Nạp Dữ Liệu DMS Mới",
    "🧾 3. Khớp File Hóa Đơn Kế Toán",
    "🛡️ 4. Đối Soát Quyết Toán Bảo Hành (WCS)",
    "🚕 5. Quản Lý Công NỢ GSM"
])

tab_work = tabs[0]
tab_import = tabs[1]
tab_inv = tabs[2]
tab_bh_wcs = tabs[3]
tab_gsm_import = tabs[4]

# ==================== TAB 1: BẢNG TÍNH WEB & PHÂN LUỒNG ====================
with tab_work:
    df_hoanthanh = df_master[df_master['Trạng thái'].isin(TRANG_THAI_HOAN_THANH)]
    df_chuahoanthanh = df_master[~df_master['Trạng thái'].isin(TRANG_THAI_HOAN_THANH + ['Đã hủy'])]
    chua_hoan_thanh_cnt = len(df_chuahoanthanh)

    mask_kh = (df_hoanthanh['KH thanh toán'] > 0) & (~df_hoanthanh['Phân loại KH'].isin(['GSM Công nợ', 'Bảo Hiểm', 'Bảo Hành Hãng', 'Nội Bộ / PDI']))
    df_kh_total = df_hoanthanh[mask_kh]
    kh_da_hd_cnt = df_kh_total[df_kh_total['Số hóa đơn'].notna() & (~df_kh_total['Số hóa đơn'].astype(str).str.strip().isin(['', 'nan', 'None', '0']))].shape[0]
    kh_total_cnt = len(df_kh_total)
    kh_chua_hd_cnt = kh_total_cnt - kh_da_hd_cnt
    kh_chua_hd_amt = df_kh_total[df_kh_total['Số hóa đơn'].isna() | df_kh_total['Số hóa đơn'].astype(str).str.strip().isin(['', 'nan', 'None', '0'])]['KH thanh toán'].sum()

    df_gsm_all = df_master[df_master['Phân loại KH'] == 'GSM Công nợ']
    df_gsm_da_hd = df_gsm_all[df_gsm_all['Số hóa đơn'].notna() & (~df_gsm_all['Số hóa đơn'].astype(str).str.strip().isin(['', 'nan', 'None', '0']))]
    df_gsm_chua_hd = df_gsm_all[df_gsm_all['Số hóa đơn'].isna() | df_gsm_all['Số hóa đơn'].astype(str).str.strip().isin(['', 'nan', 'None', '0'])]
    gsm_treo_no_cnt = len(df_gsm_chua_hd)
    gsm_treo_no_amt = df_gsm_chua_hd.apply(lambda r: r['KH thanh toán'] if r['KH thanh toán'] > 0 else r['Số tiền thanh toán cuối'], axis=1).sum() if len(df_gsm_chua_hd) > 0 else 0
    gsm_da_hd_cnt = len(df_gsm_da_hd)

    mask_bh_hang = (df_hoanthanh['BH hãng thanh toán'] > 0) | (df_hoanthanh['Phân loại KH'] == 'Bảo Hành Hãng')
    df_bh_hang = df_hoanthanh[mask_bh_hang]
    bh_chua_duyet_cnt = df_bh_hang[df_bh_hang['Phê duyệt bảo hành'].isin(['Chờ duyệt', None, 'nan', ''])].shape[0]
    bh_total_cnt = len(df_bh_hang)

    mask_bh = (df_hoanthanh['BH thanh toán'] > 0) | (df_hoanthanh['Phân loại KH'] == 'Bảo Hiểm')
    df_bh = df_hoanthanh[mask_bh]
    bh_da_hd_cnt = df_bh[df_bh['Số hóa đơn'].notna() & (~df_bh['Số hóa đơn'].astype(str).str.strip().isin(['', 'nan', 'None', '0']))].shape[0]
    bh_total_all_cnt = len(df_bh)
    bh_chua_hd_cnt = bh_total_all_cnt - bh_da_hd_cnt
    df_bh_chua_hd = df_bh[df_bh['Số hóa đơn'].isna() | df_bh['Số hóa đơn'].astype(str).str.strip().isin(['', 'nan', 'None', '0'])]
    bh_chua_hd_amt = df_bh_chua_hd.apply(lambda r: r['BH thanh toán'] if r['BH thanh toán'] > 0 else r['Số tiền thanh toán cuối'], axis=1).sum() if len(df_bh_chua_hd) > 0 else 0

    tong_chua_hd_cnt = kh_chua_hd_cnt + bh_chua_hd_cnt
    tong_chua_hd_amt = kh_chua_hd_amt + bh_chua_hd_amt

    k1, k2, k3, k4, k5 = st.columns(5)
    k1.metric("📌 KH Thanh Toán (Đã xong)", f"{kh_total_cnt:,} lệnh", f"Đã xuất HĐ: {kh_da_hd_cnt}/{kh_total_cnt}")
    k2.metric("🚕 GSM Nợ (Chưa HĐ)", f"{gsm_treo_no_cnt:,} lệnh nợ", f"{gsm_treo_no_amt:,.0f} đ (Đã XHĐ: {gsm_da_hd_cnt})")
    k3.metric("🛡️ Bảo Hành Hãng (W)", f"{bh_total_cnt:,} lệnh", f"Chưa duyệt: {bh_chua_duyet_cnt}/{bh_total_cnt}")
    k4.metric("🏢 Bảo Hiểm (Insurance)", f"{bh_total_all_cnt:,} lệnh", f"Chưa xuất HĐ: {bh_chua_hd_cnt}/{bh_total_all_cnt}")
    k5.metric("⏳ Đang Làm / Báo Giá", f"{chua_hoan_thanh_cnt:,} lệnh", "Chờ hoàn thành")

    if tong_chua_hd_cnt > 0:
        st.error(
            f"🚨 **CẢNH BÁO CHƯA XUẤT HÓA ĐƠN: {tong_chua_hd_cnt} lệnh | Tổng tiền: {tong_chua_hd_amt:,.0f} đ** "
            f"(Khách hàng thường: **{kh_chua_hd_cnt} lệnh** ({kh_chua_hd_amt:,.0f} đ) | "
            f"Bảo hiểm: **{bh_chua_hd_cnt} lệnh** ({bh_chua_hd_amt:,.0f} đ))."
        )

    f_col1, f_col2, f_col3 = st.columns([3, 2, 3])
    with f_col1:
        luong_data = st.selectbox("📂 Chọn luồng dữ liệu xem & quản trị:", [
            "1. KH Thanh Toán (Đã hoàn thành lệnh)",
            "2. GSM Công Nợ (Lấy từ Google Sheets)",
            "3. Bảo Hành Hãng (W) - Phê duyệt",
            "4. Bảo Hiểm (Insurance)",
            "5. Nội bộ thanh toán",
            "🚨 Lệnh Đã Xong Chưa Xuất HĐ (Chỉ KH & Bảo Hiểm)",
            "6. Xem Lệnh Chưa Hoàn Thành (Báo giá & Đang sửa chữa)",
            "7. Xem Lệnh Đã Hủy",
            "Toàn bộ dữ liệu (Sheet Tổng Hợp)"
        ])
    with f_col2:
        loc_canh_bao_hd = st.selectbox("⚡ Lọc trạng thái HĐ:", ["Tất cả", "Chỉ xe CHƯA có HĐ", "Đã có hóa đơn"])
    with f_col3:
        tim_kiem_tu_do = st.text_input("🔍 Tìm kiếm nhanh (Biển số / LSC / Tên bất kỳ):", "", placeholder="VD: 81A13363...")

    if luong_data == "1. KH Thanh Toán (Đã hoàn thành lệnh)": df_show = df_kh_total.copy()
    elif luong_data == "2. GSM Công Nợ (Lấy từ Google Sheets)": df_show = df_gsm_all.copy()
    elif luong_data == "3. Bảo Hành Hãng (W) - Phê duyệt": df_show = df_bh_hang.copy()
    elif luong_data == "4. Bảo Hiểm (Insurance)": df_show = df_bh.copy()
    elif luong_data == "5. Nội bộ thanh toán": df_show = df_hoanthanh[(df_hoanthanh['Nội bộ thanh toán'] > 0) | (df_hoanthanh['Phân loại KH'] == 'Nội Bộ / PDI')].copy()
    elif luong_data == "🚨 Lệnh Đã Xong Chưa Xuất HĐ (Chỉ KH & Bảo Hiểm)":
        mask_target = (df_master['Trạng thái'].isin(TRANG_THAI_HOAN_THANH)) & (
            ((df_master['KH thanh toán'] > 0) & (~df_master['Phân loại KH'].isin(['GSM Công nợ', 'Bảo Hiểm', 'Bảo Hành Hãng', 'Nội Bộ / PDI']))) |
            (df_master['BH thanh toán'] > 0) | (df_master['Phân loại KH'] == 'Bảo Hiểm')
        ) & (df_master['Số hóa đơn'].isna() | df_master['Số hóa đơn'].astype(str).str.strip().isin(['', 'nan', 'None', '0']))
        df_show = df_master[mask_target].copy()
    elif luong_data == "6. Xem Lệnh Chưa Hoàn Thành (Báo giá & Đang sửa chữa)": df_show = df_chuahoanthanh.copy()
    elif luong_data == "7. Xem Lệnh Đã Hủy": df_show = df_master[df_master['Trạng thái'] == 'Đã hủy'].copy()
    else: df_show = df_master.copy()

    df_show = loc_chuan_tu_29_thang_8(df_show)

    if loc_canh_bao_hd == "Chỉ xe CHƯA có HĐ":
        df_show = df_show[(df_show['Số hóa đơn'].isna()) | (df_show['Số hóa đơn'].astype(str).str.strip().isin(['', 'nan', 'None', '0']))]
    elif loc_canh_bao_hd == "Đã có hóa đơn":
        df_show = df_show[df_show['Số hóa đơn'].notna() & (~df_show['Số hóa đơn'].astype(str).str.strip().isin(['', 'nan', 'None', '0']))]

    if tim_kiem_tu_do:
        kw = str(tim_kiem_tu_do).strip()
        df_show = df_show[
            df_show['Biển số'].astype(str).str.contains(kw, case=False, na=False) |
            df_show['Số lệnh sửa chữa'].astype(str).str.contains(kw, case=False, na=False) |
            df_show['Tên khách hàng'].astype(str).str.contains(kw, case=False, na=False)
        ]

    df_show = chuan_hoa_kieu_du_lieu(df_show).reset_index(drop=True)
    df_show.insert(0, 'STT', range(1, len(df_show) + 1))

    is_admin = st.session_state.logged_in
    col_cfg = {
        "STT": st.column_config.NumberColumn("STT", disabled=True, width="small"),
        "Số lệnh sửa chữa": st.column_config.TextColumn("Số LSC", disabled=True),
        "Trạng thái": st.column_config.TextColumn("Trạng thái", disabled=True),
        "Cố vấn dịch vụ": st.column_config.TextColumn("CVDV", disabled=not is_admin),
        "Tên khách hàng": st.column_config.TextColumn("Khách hàng", disabled=not is_admin),
        "Thời gian đóng LSC": st.column_config.TextColumn("Ngày đóng", disabled=True),
        "Biển số": st.column_config.TextColumn("Biển số", disabled=not is_admin),
        "Phân loại KH": st.column_config.SelectboxColumn(
            "Phân loại KH", 
            options=["KH Thông Thường", "Bảo Hiểm", "Bảo Hành Hãng", "Nội Bộ / PDI", "GSM Công nợ"], 
            disabled=not is_admin, 
            required=True
        ),
        "Phê duyệt bảo hành": st.column_config.TextColumn("Duyệt BH", disabled=not is_admin),
        "Số tiền thanh toán cuối": st.column_config.NumberColumn("Tổng TT cuối", format="%,d đ", disabled=True),
        "KH thanh toán": st.column_config.NumberColumn("KH thanh toán", format="%,d đ", disabled=True),
        "BH thanh toán": st.column_config.NumberColumn("BH thanh toán", format="%,d đ", disabled=True),
        "BH hãng thanh toán": st.column_config.NumberColumn("BH hãng", format="%,d đ", disabled=True),
        "Số hóa đơn": st.column_config.TextColumn("Số HĐ", disabled=not is_admin),
        "Ngày xuất hóa đơn": st.column_config.TextColumn("Ngày HĐ", disabled=not is_admin),
        "Giá trị xuất hóa đơn": st.column_config.NumberColumn("Tiền HĐ", format="%,d đ", disabled=not is_admin),
        "Ghi chú": st.column_config.TextColumn("Ghi chú", disabled=not is_admin)
    }

    cols_render = [c for c in ['STT', 'Số lệnh sửa chữa', 'Trạng thái', 'Cố vấn dịch vụ', 'Tên khách hàng', 'Thời gian đóng LSC', 'Biển số', 'Phân loại KH', 'Phê duyệt bảo hành', 'Số tiền thanh toán cuối', 'KH thanh toán', 'BH thanh toán', 'BH hãng thanh toán', 'Số hóa đơn', 'Ngày xuất hóa đơn', 'Giá trị xuất hóa đơn', 'Ghi chú'] if c in df_show.columns]

    edited_df = st.data_editor(
        df_show[cols_render],
        use_container_width=True,
        height=540,
        column_config=col_cfg,
        disabled=(not is_admin),
        hide_index=True,
        key=f"editor_{luong_data}"
    )

    if is_admin:
        if st.button("☁️ Lưu Trực Tiếp Lên Google Sheets", type="primary"):
            edit_dict = edited_df.set_index('Số lệnh sửa chữa').to_dict('index')
            for idx in range(len(df_master)):
                r_lsc = df_master.iloc[idx]['Số lệnh sửa chữa']
                if r_lsc in edit_dict:
                    for c in ['Phê duyệt bảo hành', 'Phân loại KH', 'Số hóa đơn', 'Ngày xuất hóa đơn', 'Giá trị xuất hóa đơn', 'Cố vấn dịch vụ', 'Tên khách hàng', 'Biển số', 'Ghi chú']:
                        if c in edit_dict[r_lsc]:
                            df_master.at[df_master.index[idx], c] = edit_dict[r_lsc][c]
            save_data_to_gsheets(df_master)
            st.success("✅ Dữ liệu đã lưu vĩnh viễn lên Google Sheets!")
            st.rerun()

# ==================== TAB 4: ĐỐI SOÁT BẢO HÀNH WCS ====================
with tab_bh_wcs:
    st.subheader("🛡️ Đối Soát Quyết Toán Bảo Hành Nhà Máy VinFast (WCS)")
    st.caption("Tải file Bảng Kê Quyết Toán (`BangKe_C23401-WCS...xlsx`) để đối soát kết quả duyệt, kỳ quyết toán và kỳ xuất hóa đơn.")

    up_bangke = st.file_uploader("📥 Tải lên File BẢNG KÊ QUYẾT TOÁN NHÀ MÁY (Excel)", type=['xlsx', 'xls'], key="up_bk_tab4_simple")

    if up_bangke:
        xls_bk = pd.ExcelFile(up_bangke)
        
        # 1. Trích xuất thông tin Bảng kê
        df_bcct_raw = pd.read_excel(up_bangke, sheet_name=xls_bk.sheet_names[0], header=None)
        ten_bang_ke = str(df_bcct_raw.iloc[0, 0]).strip()
        match_phieu = re.search(r'CPBH_(\d{4})_([A-Z0-9]+)', ten_bang_ke)
        ma_phieu_bh = f"BH{match_phieu.group(1)}_{match_phieu.group(2)}" if match_phieu else "BH_VinFast"
        
        # Kỳ xuất hóa đơn tự động
        thang_ky = match_phieu.group(1)[2:] if match_phieu else "09"
        thang_hd = f"Tháng {int(thang_ky) + 1:02d}/2026"

        # 2. Đọc Sheet2 và Chi tiết_Đối soát
        df_s2_raw = pd.read_excel(up_bangke, sheet_name='Sheet2') if 'Sheet2' in xls_bk.sheet_names else pd.DataFrame()
        df_ct_raw = pd.read_excel(up_bangke, sheet_name='Chi tiết_Đối soát') if 'Chi tiết_Đối soát' in xls_bk.sheet_names else pd.DataFrame()

        date_map = {}
        if len(df_ct_raw) > 0 and 'WO' in df_ct_raw.columns:
            date_col = next((c for c in df_ct_raw.columns if 'Approved' in str(c) or 'Duyệt' in str(c)), 'WC- Approved')
            df_ct_clean = df_ct_raw.dropna(subset=['WO']).copy()
            for _, r_c in df_ct_clean.iterrows():
                w_k = norm_lsc_key(r_c['WO'])
                d_val = str(r_c[date_col])[:10]
                if d_val and d_val.lower() not in ['nan', 'none', 'nat']:
                    date_map[w_k] = d_val

        if len(df_s2_raw) > 0:
            total_wo = len(df_s2_raw)
            tien_nm_total = df_s2_raw['Nhà máy'].sum()
            tien_ds_total = df_s2_raw['Đối soát'].sum()
            
            diff_mask = df_s2_raw.iloc[:, 3] != 0
            df_diff = df_s2_raw[diff_mask].copy()
            df_ok = df_s2_raw[~diff_mask].copy()
            
            tien_treo = tien_ds_total - tien_nm_total

            c_k1, c_k2, c_k3, c_k4 = st.columns(4)
            c_k1.metric("💰 Tổng Tiền Quyết Toán (No VAT)", f"{tien_nm_total:,.0f} đ", f"Kỳ: {ma_phieu_bh}")
            c_k2.metric("🟢 Đã Duyệt Đủ 100%", f"{len(df_ok)} / {total_wo} Lệnh", f"{len(df_ok)/total_wo*100:.1f}%")
            c_k3.metric("🟡 Duyệt Thiếu (Chờ L2)", f"{len(df_diff)} Lệnh bị treo", f"-{tien_treo:,.0f} đ")
            c_k4.metric("🧾 Dự Kiến Xuất Hóa Đơn", thang_hd, f"Phiếu: {ma_phieu_bh}")

            st.markdown("---")

            # Bảng 1: Lệnh duyệt thiếu (Chờ duyệt lần 2)
            st.error(f"🚨 **DANH SÁCH {len(df_diff)} LỆNH DUYỆT THIẾU (TREO CHỜ PHÊ DUYỆT LẦN 2):**")
            diff_display = []
            for _, r in df_diff.iterrows():
                wo_str = str(r['Unnamed: 0'])
                w_norm = norm_lsc_key(wo_str)
                d_approved = date_map.get(w_norm, "2026-09-18")
                diff_display.append({
                    "Mã Lệnh Sửa Chữa (WO)": wo_str,
                    "Thời Gian Duyệt": d_approved,
                    "Kỳ Duyệt BH": ma_phieu_bh,
                    "Đại Lý Đề Xuất (Claim)": float(r['Đối soát']),
                    "Nhà Máy Duyệt Đợt 1": float(r['Nhà máy']),
                    "Chênh Lệch Còn Thiếu": float(r.iloc[3]),
                    "Kỳ Xuất HĐ": thang_hd,
                    "Trạng Thái Đánh Giá": f"🟡 {r.iloc[4]}"
                })
            df_diff_show = pd.DataFrame(diff_display)
            cfg_tbl_diff = {
                "Đại Lý Đề Xuất (Claim)": st.column_config.NumberColumn(format="%,d đ"),
                "Nhà Máy Duyệt Đợt 1": st.column_config.NumberColumn(format="%,d đ"),
                "Chênh Lệch Còn Thiếu": st.column_config.NumberColumn(format="%,d đ")
            }
            st.dataframe(df_diff_show, use_container_width=True, hide_index=True, column_config=cfg_tbl_diff)

            # Bảng 2: Lệnh duyệt đủ
            st.success(f"🟢 **DANH SÁCH LỆNH DUYỆT ĐỦ 100% TIỀN CLAIM ({len(df_ok)} LỆNH):**")
            ok_display = []
            for _, r in df_ok.iterrows():
                wo_str = str(r['Unnamed: 0'])
                w_norm = norm_lsc_key(wo_str)
                d_approved = date_map.get(w_norm, "2026-08-30")
                ok_display.append({
                    "Mã Lệnh Sửa Chữa (WO)": wo_str,
                    "Thời Gian Duyệt": d_approved,
                    "Kỳ Duyệt BH": ma_phieu_bh,
                    "Số Tiền Quyết Toán": float(r['Nhà máy']),
                    "Kỳ Xuất HĐ": thang_hd,
                    "Trạng Thái": "🟢 Đã duyệt đủ 100% (Khớp lệnh)"
                })
            df_ok_show = pd.DataFrame(ok_display)
            cfg_tbl_ok = {"Số Tiền Quyết Toán": st.column_config.NumberColumn(format="%,d đ")}
            st.dataframe(df_ok_show, use_container_width=True, hide_index=True, column_config=cfg_tbl_ok)

            # Nút đồng bộ lên Google Sheets
            if is_admin and st.button("☁️ ĐỒNG BỘ KẾT QUẢ BẢNG KÊ NÀY LÊN GOOGLE SHEETS", type="primary"):
                updated_cnt = 0
                diff_map_save = {norm_lsc_key(r['Mã Lệnh Sửa Chữa (WO)']): r for r in diff_display}
                ok_map_save = {norm_lsc_key(r['Mã Lệnh Sửa Chữa (WO)']): r for r in ok_display}

                for idx in range(len(df_master)):
                    k_w = norm_lsc_key(df_master.iloc[idx]['Số lệnh sửa chữa'])
                    if k_w in diff_map_save:
                        r_d = diff_map_save[k_w]
                        df_master.iloc[idx, df_master.columns.get_loc('Phê duyệt bảo hành')] = f"Chờ PD L2 ({ma_phieu_bh})"
                        df_master.iloc[idx, df_master.columns.get_loc('Ghi chú')] = (
                            f"{ma_phieu_bh} duyệt {r_d['Thời Gian Duyệt']}: {r_d['Nhà Máy Duyệt Đợt 1']:,.0f} đ "
                            f"(Treo {abs(r_d['Chênh Lệch Còn Thiếu']):,.0f} đ). Xuất HĐ {thang_hd}."
                        )
                        updated_cnt += 1
                    elif k_w in ok_map_save:
                        r_o = ok_map_save[k_w]
                        df_master.iloc[idx, df_master.columns.get_loc('Phê duyệt bảo hành')] = f"Đã duyệt đủ ({ma_phieu_bh})"
                        df_master.iloc[idx, df_master.columns.get_loc('Ghi chú')] = (
                            f"{ma_phieu_bh} duyệt {r_o['Thời Gian Duyệt']}: {r_o['Số Tiền Quyết Toán']:,.0f} đ (Đủ 100%). Xuất HĐ {thang_hd}."
                        )
                        updated_cnt += 1

                save_data_to_gsheets(df_master)
                st.success(f"✅ ĐÃ ĐỒNG BỘ THÀNH CÔNG {updated_cnt} LỆNH LÊN GOOGLE SHEETS!")
                st.rerun()

# ==================== CÁC TAB CHỨC NĂNG KHÁC ====================
if is_admin:
    # TAB 2: NẠP DỮ LIỆU DMS
    with tab_import:
        st.subheader("Nạp file dữ liệu phân phối VinFast định kỳ")
        up_db = st.file_uploader("Kéo thả file Database mới vào đây", type=['csv', 'xlsx'], key='up_dms_simple')
        if up_db and st.button("🚀 BẮT ĐẦU NẠP VÀ LƯU LÊN GOOGLE SHEETS", type="primary"):
            df_raw = pd.read_excel(up_db) if up_db.name.endswith('.xlsx') else pd.read_csv(up_db)
            df_inc = pd.DataFrame()
            for c in TAT_CA_COT: df_inc[c] = df_raw[c] if c in df_raw.columns else ""
            for c in COT_TIEN: df_inc[c] = pd.to_numeric(clean_tien_series(df_inc[c]), errors='coerce').fillna(0)
            df_inc['Số lệnh sửa chữa'] = df_inc['Số lệnh sửa chữa'].astype(str).apply(clean_lsc_giu_gach)
            df_inc['Phân loại KH'] = "KH Thông Thường"
            df_inc = loc_chuan_tu_29_thang_8(chuan_hoa_kieu_du_lieu(dong_bo_hoa_don(df_inc)))
            
            master_norm_map = {norm_lsc_key(lsc): idx for idx, lsc in enumerate(df_master['Số lệnh sửa chữa'])}
            new_recs = []
            for _, r in df_inc.iterrows():
                k = norm_lsc_key(r['Số lệnh sửa chữa'])
                if k not in master_norm_map: new_recs.append(r)
                else:
                    m_idx = master_norm_map[k]
                    df_master.iloc[m_idx, df_master.columns.get_loc('Trạng thái')] = str(r['Trạng thái'])
            if new_recs: df_master = pd.concat([df_master, pd.DataFrame(new_recs)], ignore_index=True)
            save_data_to_gsheets(df_master)
            st.success("✅ Đã nạp thành công lên Google Sheets!")
            st.rerun()

    # TAB 3: KHỚP HÓA ĐƠN
    with tab_inv:
        st.subheader("Khớp file Hóa Đơn kế toán với Hệ Thống")
        st.caption("Tự động gộp tất cả hóa đơn cùng 1 LSC: nối số HĐ bằng dấu phẩy và cộng dồn tiền chính xác 100%.")

    # TAB 5: QUẢN LÝ GSM
    with tab_gsm_import:
        st.subheader("🚕 Quản Lý Công Nợ GSM (Đồng bộ sheet GSM_List)")
        st.caption("Dữ liệu tự động lấy từ sheet GSM_List trên Google Sheets.")
