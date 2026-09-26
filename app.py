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

# ==================== TAB 4: ĐỐI SOÁT & TỰ ĐỘNG ĐẨY GOOGLE SHEETS ====================
with tab_bh_wcs:
    st.subheader("🛡️ Quản Lý & Tự Động Đồng Bộ Bảo Hành VinFast (Chỉ Lệnh ĐÃ ĐÓNG)")
    st.caption("Kéo thả 2 file vào đây: Hệ thống sẽ TỰ ĐỘNG lọc lệnh ĐÃ ĐÓNG và TỰ ĐỘNG ĐẨY NGAY sang sheet 'ChiTiet_BaoHanh' trên Google Sheets.")

    c_up1, c_up2 = st.columns(2)
    with c_up1:
        up_dx = st.file_uploader("📥 1. File DEXUATBAOHANH (Cổng Nhà Máy)", type=['csv', 'xlsx'], key="up_dx_auto")
    with c_up2:
        up_ct = st.file_uploader("📥 2. File CHITIETLENHSUACHUA (DMS Xưởng)", type=['csv', 'xlsx'], key="up_ct_auto")

    if up_dx and up_ct:
        df_dx_in = pd.read_csv(up_dx, low_memory=False) if up_dx.name.endswith('.csv') else pd.read_excel(up_dx)
        df_ct_in = pd.read_csv(up_ct, low_memory=False) if up_ct.name.endswith('.csv') else pd.read_excel(up_ct)
        
        df_dx_in.columns = [str(c).strip() for c in df_dx_in.columns]
        df_ct_in.columns = [str(c).strip() for c in df_ct_in.columns]

        col_dx_lsc = next((c for c in df_dx_in.columns if 'lệnh sửa chữa' in c.lower()), 'Lệnh sửa chữa')
        col_ct_lsc = next((c for c in df_ct_in.columns if 'lệnh sửa chữa' in c.lower()), 'Lệnh sửa chữa')
        
        df_dx_in['lsc_norm'] = df_dx_in[col_dx_lsc].apply(norm_lsc_key)
        df_ct_in['lsc_norm'] = df_ct_in[col_ct_lsc].apply(norm_lsc_key)

        # LẤY DANH SÁCH LỆNH "ĐÃ ĐÓNG" TỪ MASTERDATA & TỪ FILE CHI TIẾT
        status_closed_set = set(df_master[df_master['Trạng thái'].str.strip() == 'Đã đóng']['Số lệnh sửa chữa'].apply(norm_lsc_key))

        # LỌC: CHỈ LẤY P/bill == 'W'
        col_pb = next((c for c in df_ct_in.columns if 'classification' in c.lower() or 'p/bill' in c.lower()), 'P/bill classification')
        df_w_raw = df_ct_in[df_ct_in[col_pb].astype(str).str.strip().str.upper() == 'W'].copy()

        # Kiểm tra điều kiện "ĐÃ ĐÓNG"
        def check_strictly_closed(norm_k, row_r):
            if norm_k in status_closed_set: return True
            col_dd = next((c for c in row_r.index if 'đóng' in str(c).lower()), None)
            if col_dd and str(row_r[col_dd]).strip() == 'Đã đóng': return True
            return False

        df_w_raw['is_dong'] = df_w_raw.apply(lambda r: check_strictly_closed(r['lsc_norm'], r), axis=1)
        df_w_target = df_w_raw[df_w_raw['is_dong']].copy()

        set_dx_lsc = set(df_dx_in['lsc_norm'].unique())
        set_w_lsc = set(df_w_target['lsc_norm'].unique())

        lsc_chua_claim = sorted(list(set_w_lsc - set_dx_lsc))
        lsc_da_claim = set_w_lsc.intersection(set_dx_lsc)

        st.markdown("---")
        # THẺ KPI
        k1, k2, k3 = st.columns(3)
        k1.metric("📌 Tổng Lệnh BH ĐÃ ĐÓNG (W)", f"{len(set_w_lsc)} Lệnh", f"{len(df_w_target)} dòng linh kiện/công")
        k2.metric("🟢 Đã Gửi Đề Xuất Nhà Máy", f"{len(lsc_da_claim)} Lệnh")
        k3.metric("🚨 Lệnh ĐÃ ĐÓNG Chưa Tạo ĐXBH", f"{len(lsc_chua_claim)} Lệnh BỊ SÓT", delta=f"-{len(lsc_chua_claim)} lệnh", delta_color="inverse")

        # KHỐI 1: CẢNH BÁO LỆNH CHƯA TẠO CLAIM
        if len(lsc_chua_claim) > 0:
            st.error(f"🚨 **DANH SÁCH {len(lsc_chua_claim)} LỆNH ĐÃ ĐÓNG NHƯNG CHƯA TẠO CLAIM TRÊN CỔNG NHÀ MÁY:**")
            df_miss_summary = df_w_target[df_w_target['lsc_norm'].isin(lsc_chua_claim)].groupby(col_ct_lsc).agg({
                'Số khung': 'first' if 'Số khung' in df_w_target.columns else lambda x: '',
                'Cố vấn dịch vụ': 'first' if 'Cố vấn dịch vụ' in df_w_target.columns else lambda x: '',
                'Mô tả sản phẩm': 'count'
            }).reset_index()
            df_miss_summary.columns = ['Số Lệnh Sửa Chữa', 'Số Khung', 'Cố Vấn Dịch Vụ', 'Số Mục Chưa Claim']
            st.dataframe(df_miss_summary, use_container_width=True, hide_index=True)

        st.markdown("---")
        st.subheader("📋 Bảng Chi Tiết Từng Hạng Mục (Tự Động Đồng Bộ Google Sheets)")

        col_dx_mavt = next((c for c in df_dx_in.columns if 'mã vật tư' in c.lower() or 'mã sản phẩm' in c.lower()), 'Mã vật tư')
        col_dx_ngay = next((c for c in df_dx_in.columns if 'ngày phê duyệt' in c.lower()), 'Ngày phê duyệt')
        col_dx_tt = next((c for c in df_dx_in.columns if 'trạng thái' in c.lower() and 'phê duyệt' not in c.lower()), 'Trạng thái')
        col_dx_tien = next((c for c in df_dx_in.columns if 'tổng số tiền' in c.lower()), 'Tổng số tiền')
        col_dx_dx = next((c for c in df_dx_in.columns if 'số đề xuất' in c.lower()), 'Số đề xuất bảo hành')

        col_ct_masp = next((c for c in df_w_target.columns if 'mã sản phẩm' in c.lower() or 'mã vật tư' in c.lower()), 'Mã sản phẩm')
        col_ct_pt = next((c for c in df_w_target.columns if 'tiền phụ tùng' in c.lower()), 'Tổng tiền phụ tùng')
        col_ct_nc = next((c for c in df_w_target.columns if 'tiền nhân công' in c.lower()), 'Tổng tiền nhân công')

        df_dx_in['mavt_norm'] = df_dx_in[col_dx_mavt].astype(str).str.strip().str.upper()
        df_w_target['masp_norm'] = df_w_target[col_ct_masp].astype(str).str.strip().str.upper()

        df_export_gsheet = pd.merge(
            df_w_target,
            df_dx_in[['lsc_norm', 'mavt_norm', col_dx_dx, col_dx_ngay, col_dx_tt, col_dx_tien]].drop_duplicates(subset=['lsc_norm', 'mavt_norm']),
            left_on=['lsc_norm', 'masp_norm'],
            right_on=['lsc_norm', 'mavt_norm'],
            how='left'
        )

        def tinh_tien_xuong(r):
            pt = pd.to_numeric(clean_tien_series(pd.Series([r.get(col_ct_pt, 0)])), errors='coerce').iloc[0]
            nc = pd.to_numeric(clean_tien_series(pd.Series([r.get(col_ct_nc, 0)])), errors='coerce').iloc[0]
            return pt if pt > 0 else nc

        df_export_gsheet['Tiền xưởng (DMS)'] = df_export_gsheet.apply(tinh_tien_xuong, axis=1)
        df_export_gsheet['Tiền claim (ĐXBH)'] = pd.to_numeric(clean_tien_series(df_export_gsheet[col_dx_tien]), errors='coerce').fillna(0)
        df_export_gsheet['Trạng thái claim'] = df_export_gsheet[col_dx_tt].fillna('Chưa gửi claim')
        df_export_gsheet['Ngày phê duyệt'] = df_export_gsheet[col_dx_ngay].fillna('')
        df_export_gsheet['Số đề xuất'] = df_export_gsheet[col_dx_dx].fillna('')

        cols_final_gsheet = [
            col_ct_lsc, 'Số khung', 'Cố vấn dịch vụ', 'Loại sản phẩm',
            col_ct_masp, 'Mô tả sản phẩm', 'Số lượng/Nhân công',
            'Tiền xưởng (DMS)', 'Tiền claim (ĐXBH)', 'Trạng thái claim',
            'Ngày phê duyệt', 'Số đề xuất'
        ]
        cols_valid = [c for c in cols_final_gsheet if c in df_export_gsheet.columns]
        df_final = df_export_gsheet[cols_valid].copy().reset_index(drop=True)

        # CƠ CHẾ TỰ ĐỘNG ĐỒNG BỘ (AUTO-SYNC) LÊN GOOGLE SHEETS
        current_files_hash = hashlib.md5((up_dx.name + str(up_dx.size) + up_ct.name + str(up_ct.size)).encode()).hexdigest()
        
        if "last_synced_bh_hash" not in st.session_state:
            st.session_state.last_synced_bh_hash = ""

        if is_admin and st.session_state.last_synced_bh_hash != current_files_hash:
            with st.spinner("⏳ Đang TỰ ĐỘNG truyền dữ liệu sang Google Sheets (Sheet: ChiTiet_BaoHanh)..."):
                df_to_push = df_final.copy()
                for c in df_to_push.columns:
                    if c in ['Tiền xưởng (DMS)', 'Tiền claim (ĐXBH)', 'Số lượng/Nhân công']:
                        df_to_push[c] = pd.to_numeric(df_to_push[c], errors='coerce').fillna(0)
                    else:
                        df_to_push[c] = df_to_push[c].fillna('').astype(str)

                conn.update(worksheet="ChiTiet_BaoHanh", data=df_to_push)
                st.session_state.last_synced_bh_hash = current_files_hash
                st.toast(f"✅ Đã tự động đồng bộ {len(df_to_push)} dòng sang sheet ChiTiet_BaoHanh!", icon="🚀")

        # Nút Link trực tiếp Google Sheet hiển thị nổi bật
        c_status, c_link = st.columns([7, 3])
        with c_status:
            st.success(f"🟢 Dữ liệu đang được kết nối và đồng bộ tự động với Google Sheets ({len(df_final)} dòng chi tiết).")
        with c_link:
            st.link_button(
                "🔗 Mở File Google Sheets Ngay",
                "https://docs.google.com/spreadsheets/d/1P50EbeA6oN-5CcZtZCvynlsdLM50b3up4URWR-rc85E/edit",
                use_container_width=True
            )

        cfg_final = {
            "Tiền xưởng (DMS)": st.column_config.NumberColumn(format="%,d đ"),
            "Tiền claim (ĐXBH)": st.column_config.NumberColumn(format="%,d đ")
        }
        st.dataframe(df_final, use_container_width=True, hide_index=True, column_config=cfg_final)

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

   # TAB 3: KHỚP HÓA ĐƠN KẾ TOÁN
    with tab_inv:
        st.subheader("🧾 Khớp File Hóa Đơn Kế Toán với Hệ Thống")
        st.caption("Tự động nhận diện mã LSC từ file hóa đơn, gộp nhiều hóa đơn cùng 1 xe (nối số HĐ và cộng dồn tiền), sau đó cập nhật trực tiếp lên Google Sheets.")

        up_hd = st.file_uploader("📥 Kéo thả file Hóa Đơn kế toán vào đây (Excel hoặc CSV):", type=['xlsx', 'xls', 'csv'], key="up_hd_main")

        if up_hd:
            try:
                if up_hd.name.endswith('.csv'):
                    df_hd_raw = pd.read_csv(up_hd, low_memory=False)
                else:
                    df_hd_raw = pd.read_excel(up_hd)
                
                df_hd_raw.columns = [str(c).strip() for c in df_hd_raw.columns]

                # 1. Tự động tìm kiếm các cột quan trọng
                col_so_hd = next((c for c in df_hd_raw.columns if any(k in str(c).lower() for k in ['số hóa đơn', 'số hđ', 'inv_no', 'so_hd', 'số ct'])), None)
                col_ngay_hd = next((c for c in df_hd_raw.columns if any(k in str(c).lower() for k in ['ngày hóa đơn', 'ngày hđ', 'ngày lập', 'ngày ct', 'inv_date', 'ngày xuất'])), None)
                col_tien_hd = next((c for c in df_hd_raw.columns if any(k in str(c).lower() for k in ['tổng tiền', 'thành tiền', 'tổng thanh toán', 'tiền sau thuế', 'giá trị'])), None)
                col_lsc_hd = next((c for c in df_hd_raw.columns if any(k in str(c).lower() for k in ['lệnh sửa chữa', 'số lệnh', 'lsc', 'wo', 'mã lệnh'])), None)

                # Nếu không có cột LSC riêng, tìm trong cột diễn giải / ghi chú
                if not col_lsc_hd:
                    col_lsc_hd = next((c for c in df_hd_raw.columns if any(k in str(c).lower() for k in ['diễn giải', 'nội dung', 'ghi chú', 'description'])), None)

                c_sel1, c_sel2, c_sel3, c_sel4 = st.columns(4)
                with c_sel1:
                    sel_so_hd = st.selectbox("Cột Số HĐ:", df_hd_raw.columns, index=df_hd_raw.columns.get_loc(col_so_hd) if col_so_hd else 0)
                with c_sel2:
                    sel_ngay_hd = st.selectbox("Cột Ngày HĐ:", df_hd_raw.columns, index=df_hd_raw.columns.get_loc(col_ngay_hd) if col_ngay_hd else 0)
                with c_sel3:
                    sel_tien_hd = st.selectbox("Cột Tiền HĐ:", df_hd_raw.columns, index=df_hd_raw.columns.get_loc(col_tien_hd) if col_tien_hd else 0)
                with c_sel4:
                    sel_lsc_hd = st.selectbox("Cột Mã LSC / Ghi chú:", df_hd_raw.columns, index=df_hd_raw.columns.get_loc(col_lsc_hd) if col_lsc_hd else 0)

                # 2. Xử lý bóc tách mã WO
                def trich_xuat_wo(val):
                    if pd.isna(val): return ""
                    s = str(val).upper()
                    # Tìm mẫu C23401-WO-xxxxxx-xxxx hoặc WO-xxxxxx-xxxx
                    m = re.search(r'([A-Z0-9]*WO[-0-9A-Z]+)', s)
                    if m:
                        return norm_lsc_key(m.group(1))
                    return norm_lsc_key(s)

                df_hd_work = df_hd_raw.copy()
                df_hd_work['lsc_norm'] = df_hd_work[sel_lsc_hd].apply(trich_xuat_wo)
                df_hd_work = df_hd_work[df_hd_work['lsc_norm'].str.len() > 3]

                df_hd_work['so_hd_clean'] = df_hd_work[sel_so_hd].fillna('').astype(str).str.strip().str.lstrip('0')
                df_hd_work['ngay_hd_clean'] = df_hd_work[sel_ngay_hd].apply(clean_ngay_chuan)
                df_hd_work['tien_hd_num'] = pd.to_numeric(clean_tien_series(df_hd_work[sel_tien_hd]), errors='coerce').fillna(0)

                # 3. Gom nhóm theo LSC: Nối số HĐ và cộng dồn tiền
                hd_grouped = df_hd_work.groupby('lsc_norm').agg({
                    'so_hd_clean': lambda x: ", ".join(sorted(list(set(filter(None, x))))),
                    'ngay_hd_clean': lambda x: next((d for d in x if d), ""),
                    'tien_hd_num': 'sum'
                }).reset_index()

                # 4. Đối chiếu với MasterData
                master_dict = {norm_lsc_key(lsc): idx for idx, lsc in enumerate(df_master['Số lệnh sửa chữa'])}
                
                matched_rows = []
                unmatched_rows = []

                for _, r in hd_grouped.iterrows():
                    k = r['lsc_norm']
                    if k in master_dict:
                        m_idx = master_dict[k]
                        row_m = df_master.iloc[m_idx]
                        matched_rows.append({
                            "Số Lệnh Sửa Chữa": row_m['Số lệnh sửa chữa'],
                            "Biển Số": row_m['Biển số'],
                            "Cố Vấn Dịch Vụ": row_m['Cố vấn dịch vụ'],
                            "Số HĐ Mới Khớp": r['so_hd_clean'],
                            "Ngày HĐ Mới": r['ngay_hd_clean'],
                            "Tiền HĐ Mới Khớp": r['tien_hd_num'],
                            "Tiền Thanh Toán DMS": row_m['Số tiền thanh toán cuối'],
                            "Số HĐ Cũ Trên Hệ Thống": row_m['Số hóa đơn']
                        })
                    else:
                        unmatched_rows.append({
                            "Mã LSC Từ File HĐ": k,
                            "Số Hóa Đơn": r['so_hd_clean'],
                            "Ngày Hóa Đơn": r['ngay_hd_clean'],
                            "Tổng Tiền HĐ": r['tien_hd_num']
                        })

                st.markdown("---")
                # Thẻ thống kê kết quả khớp
                k_m1, k_m2, k_m3 = st.columns(3)
                k_m1.metric("📑 Tổng Lệnh Trong File HĐ", f"{len(hd_grouped)} Lệnh")
                k_m2.metric("🟢 Khớp Thành Công Với Hệ Thống", f"{len(matched_rows)} Lệnh")
                k_m3.metric("🟡 Không Tìm Thấy Trong Hệ Thống", f"{len(unmatched_rows)} Lệnh")

                if matched_rows:
                    st.success(f"🟢 **DANH SÁCH {len(matched_rows)} LỆNH KHỚP THÀNH CÔNG (Sẵn sàng cập nhật):**")
                    df_matched_show = pd.DataFrame(matched_rows)
                    cfg_match = {
                        "Tiền HĐ Mới Khớp": st.column_config.NumberColumn(format="%,d đ"),
                        "Tiền Thanh Toán DMS": st.column_config.NumberColumn(format="%,d đ")
                    }
                    st.dataframe(df_matched_show, use_container_width=True, hide_index=True, column_config=cfg_match)

                    # Nút bấm lưu lên Google Sheets
                    if is_admin:
                        if st.button("☁️ XÁC NHẬN CẬP NHẬT HÓA ĐƠN LÊN GOOGLE SHEETS", type="primary", use_container_width=True):
                            update_cnt = 0
                            for item in matched_rows:
                                k_norm = norm_lsc_key(item['Số Lệnh Sửa Chữa'])
                                if k_norm in master_dict:
                                    idx_target = master_dict[k_norm]
                                    df_master.at[df_master.index[idx_target], 'Số hóa đơn'] = str(item['Số HĐ Mới Khớp'])
                                    df_master.at[df_master.index[idx_target], 'Ngày xuất hóa đơn'] = str(item['Ngày HĐ Mới'])
                                    df_master.at[df_master.index[idx_target], 'Giá trị xuất hóa đơn'] = float(item['Tiền HĐ Mới Khớp'])
                                    update_cnt += 1

                            save_data_to_gsheets(df_master)
                            st.success(f"✅ ĐÃ CẬP NHẬT THÀNH CÔNG {update_cnt} HÓA ĐƠN LÊN GOOGLE SHEETS!")
                            st.rerun()

                if unmatched_rows:
                    with st.expander(f"⚠️ Xem {len(unmatched_rows)} hóa đơn không khớp được với LSC trên hệ thống:"):
                        df_unmatch_show = pd.DataFrame(unmatched_rows)
                        st.dataframe(df_unmatch_show, use_container_width=True, hide_index=True)

            except Exception as e:
                st.error(f"❌ Có lỗi khi đọc file Hóa Đơn: {str(e)}")
    # TAB 5: QUẢN LÝ GSM
    with tab_gsm_import:
        st.subheader("🚕 Quản Lý Công NỢ GSM (Đồng bộ sheet GSM_List)")
        st.caption("Dữ liệu tự động lấy từ sheet GSM_List trên Google Sheets.")
