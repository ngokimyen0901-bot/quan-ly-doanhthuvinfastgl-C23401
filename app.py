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
from streamlit_gsheets import GSheetsConnection

st.set_page_config(page_title="Báo Cáo Dịch Vụ VinFast", page_icon="🚗", layout="wide")

st.markdown("""

""", unsafe_allow_html=True)

SECRET_KEY_AUTH = "VinFast_GiaLai_Secret_Key_2026"

DANH_SACH_ADMIN = {
    "admin": "Vinfastgialai@2026##"
}

MOC_BAT_DAU_DATE = pd.Timestamp(year=2026, month=8, day=29)

COT_DINH_DANH = [
    'Số lệnh sửa chữa', 'Trạng thái', 'Cố vấn dịch vụ', 'Tên khách hàng', 
    'Thời gian đóng LSC', 'Biển số', 'Xe GSM'
]

COT_TIEN = [
    'Tổng tiền công', 'Tổng tiền phụ tùng', 'Tổng trước chiết khấu',
    'Chiết khấu đại lý', 'Tổng chiết khấu', 'Tổng sau chiết khấu',
    'Tiền VAT', 'Tổng có VAT', 'Chiết khấu VinClub',
    'Số tiền thanh toán cuối', 'Tiền đặt cọc', 'KH thanh toán',
    'BH thanh toán', 'BH hãng thanh toán', 'Nội bộ thanh toán'
]

COT_THEM = ['Phân loại KH', 'Phê duyệt bảo hành']
COT_HOA_DON = ['Số hóa đơn', 'Ngày xuất hóa đơn', 'Giá trị xuất hóa đơn']

TAT_CA_COT = COT_DINH_DANH + COT_TIEN + COT_THEM + COT_HOA_DON

TEXT_COLUMNS = [
    'Số lệnh sửa chữa', 'Trạng thái', 'Cố vấn dịch vụ', 'Tên khách hàng',
    'Biển số', 'Xe GSM', 'Phân loại KH', 
    'Phê duyệt bảo hành', 'Số hóa đơn', 'Ngày xuất hóa đơn'
]

TRANG_THAI_HOAN_THANH = ['Đã đóng', 'Sẵn sàng bàn giao']

# KHỞI TẠO KẾT NỐI GOOGLE SHEETS
conn = st.connection("gsheets", type=GSheetsConnection)

def tao_auth_token(username, so_ngay=10):
    exp_time = int(time.time()) + (so_ngay * 86400)
    data = f"{username}|{exp_time}"
    sig = hashlib.sha256(f"{data}|{SECRET_KEY_AUTH}".encode()).hexdigest()[:16]
    token = base64.urlsafe_b64encode(f"{data}|{sig}".encode()).decode()
    return token

def xac_thuc_auth_token(token):
    try:
        raw = base64.urlsafe_b64decode(token.encode()).decode()
        username, exp_time, sig = raw.split('|')
        if int(exp_time) < int(time.time()):
            return None
        expected_sig = hashlib.sha256(f"{username}|{exp_time}|{SECRET_KEY_AUTH}".encode()).hexdigest()[:16]
        if sig == expected_sig and username in DANH_SACH_ADMIN:
            return username
    except Exception:
        return None
    return None

def clean_lsc_giu_gach(val):
    if pd.isna(val) or val is None:
        return ""
    return str(val).split('(')[0].strip()

def norm_lsc_key(val):
    if pd.isna(val) or val is None:
        return ""
    s = str(val).split('(')[0].strip().upper()
    return re.sub(r'[^A-Z0-9]', '', s)

def lay_loi_ma_wo(val):
    s = norm_lsc_key(val)
    m = re.search(r'WO(\d{8,12})', s)
    if m:
        return m.group(1)
    digits = re.sub(r'\D', '', s)
    if len(digits) >= 8:
        return digits[-10:]
    return s

def clean_ngay_chuan(val):
    if pd.isna(val) or val is None:
        return ""
    val_str = str(val).strip()
    if val_str.lower() in ['nan', 'none', '', 'nat']:
        return ""
    match_dmy = re.search(r'(\b\d{1,2}[/-]\d{1,2}[/-]\d{4}\b)', val_str)
    if match_dmy:
        d_str = match_dmy.group(1).replace('-', '/')
        parts = d_str.split('/')
        return f"{int(parts[0]):02d}/{int(parts[1]):02d}/{parts[2]}"
    match_ymd = re.search(r'(\b\d{4}[/-]\d{1,2}[/-]\d{1,2}\b)', val_str)
    if match_ymd:
        parts = match_ymd.group(1).replace('/', '-').split('-')
        return f"{int(parts[2]):02d}/{int(parts[1]):02d}/{parts[0]}"
    return val_str

def chuyen_ngay_gio_sortable(val):
    if pd.isna(val) or not val:
        return ""
    val_str = str(val).strip()
    if val_str.lower() in ['nan', 'none', '', 'nat']:
        return ""
    try:
        if ',' in val_str:
            parts = val_str.split(',')
            t_part = parts[0].strip()
            d_part = parts[1].strip()
            d_clean = clean_ngay_chuan(d_part)
            d_split = d_clean.split('/')
            return f"{d_split[2]}/{d_split[1]}/{d_split[0]} {t_part}"
        return val_str
    except Exception:
        return val_str

def lay_ngay_lsc(lsc_str, tg_str=""):
    m1 = re.search(r'WO-(\d{2})-(\d{2})-(\d{2})', str(lsc_str), re.IGNORECASE)
    if m1:
        yy, mm, dd = int(m1.group(1)), int(m1.group(2)), int(m1.group(3))
        return pd.Timestamp(year=2000+yy, month=mm, day=dd)
    m2 = re.search(r'WO-(\d{2})(\d{2})(\d{2})', str(lsc_str), re.IGNORECASE)
    if m2:
        yy, mm, dd = int(m2.group(1)), int(m2.group(2)), int(m2.group(3))
        return pd.Timestamp(year=2000+yy, month=mm, day=dd)
    m3 = re.search(r'(\d{4})[/-](\d{1,2})[/-](\d{1,2})', str(tg_str))
    if m3:
        return pd.Timestamp(year=int(m3.group(1)), month=int(m3.group(2)), day=int(m3.group(3)))
    m4 = re.search(r'(\d{1,2})[/-](\d{1,2})[/-](\d{4})', str(tg_str))
    if m4:
        return pd.Timestamp(year=int(m4.group(3)), month=int(m4.group(2)), day=int(m4.group(1)))
    return pd.NaT

def clean_tien_series(ser):
    return (
        ser.astype(str)
        .str.replace(',', '', regex=False)
        .str.replace(' ', '', regex=False)
        .replace(['-', '', 'nan', 'None', ''], '0')
    )

def doc_danh_sach_gsm_tu_gsheet():
    gsm_keys = set()
    try:
        df_gsm = conn.read(worksheet="GSM_List", ttl=5)
        if df_gsm is not None and len(df_gsm) > 0:
            for val in df_gsm.iloc[:, 0].dropna():
                val_str = str(val).strip()
                if val_str and not val_str.startswith('#'):
                    k_norm = norm_lsc_key(val_str)
                    k_core = lay_loi_ma_wo(val_str)
                    if k_norm: gsm_keys.add(k_norm)
                    if k_core: gsm_keys.add(k_core)
    except Exception:
        pass
    return gsm_keys

def doc_file_hoa_don_chuan(file_obj):
    if file_obj.name.endswith('.csv'):
        df_raw = pd.read_csv(file_obj, header=None)
    else:
        df_raw = pd.read_excel(file_obj, header=None)

    h_idx = 0
    for idx, row in df_raw.head(25).iterrows():
        row_str = " ".join([str(v).lower() for v in row.values if pd.notna(v)])
        if ('số ro' in row_str or 'ro hãng' in row_str) and ('chứng từ' in row_str or 'khách hàng' in row_str or 'loại tt' in row_str):
            h_idx = idx
            break

    row_top = df_raw.iloc[h_idx].fillna('').astype(str).tolist()
    row_sub = df_raw.iloc[h_idx + 1].fillna('').astype(str).tolist() if (h_idx + 1) < len(df_raw) else [''] * len(row_top)
    
    df_data = df_raw.iloc[h_idx + 2:].copy().reset_index(drop=True)
    df_clean_data = []
    for _, r in df_data.iterrows():
        r_str = " ".join([str(v) for v in r.values if pd.notna(v)])
        if 'tổng số lượt' in r_str.lower() or 'tổng cộng' in r_str.lower():
            continue
        if len(r_str.strip()) > 0:
            df_clean_data.append(r.values)

    df_result = pd.DataFrame(df_clean_data)
    col_labels = []
    for c_idx in range(len(row_top)):
        t = row_top[c_idx].strip() if row_top[c_idx].lower() not in ['none', 'nan'] else ''
        s = row_sub[c_idx].strip() if row_sub[c_idx].lower() not in ['none', 'nan'] else ''
        sample_val = ""
        if len(df_result) > 0 and c_idx < df_result.shape[1]:
            raw_s = df_result.iloc[0, c_idx]
            if pd.notna(raw_s) and str(raw_s).lower() not in ['nan', 'none']:
                sample_val = str(raw_s).strip()
                if len(sample_val) > 20:
                    sample_val = sample_val[:17] + "..."
        parts = [p for p in [t, s] if p]
        base_name = " - ".join(parts) if parts else f"Cột {c_idx + 1}"
        preview = f" [Mẫu: {sample_val}]" if sample_val else ""
        col_labels.append(f"{base_name}{preview}")

    return df_result, col_labels

def doc_file_cyber(file_obj):
    if file_obj.name.endswith('.csv'):
        df_raw = pd.read_csv(file_obj, header=None)
    else:
        df_raw = pd.read_excel(file_obj, header=None)

    h_idx = -1
    col_lsc_idx = 0
    for idx, row in df_raw.head(20).iterrows():
        for c_i, v in enumerate(row.values):
            val_str = str(v).lower()
            if 'ro hãng' in val_str or 'ro hang' in val_str or 'lệnh hãng' in val_str:
                h_idx = idx
                col_lsc_idx = c_i
                break
        if h_idx != -1:
            break

    if h_idx == -1:
        h_idx = 7
        col_lsc_idx = 0

    cyber_series = df_raw.iloc[h_idx + 1:, col_lsc_idx].dropna().astype(str)
    cyber_keys = set()
    for x in cyber_series:
        val_clean = clean_lsc_giu_gach(x)
        if len(val_clean) > 5 and 'tổng cộng' not in val_clean.lower():
            cyber_keys.add(val_clean)
            cyber_keys.add(norm_lsc_key(val_clean))
    return cyber_keys

def doc_file_db(file_obj):
    if file_obj.name.endswith('.csv'):
        df = pd.read_csv(file_obj)
    else:
        df = pd.read_excel(file_obj)
    df.columns = [str(c).strip() for c in df.columns]
    return df

# ĐÃ KHẮC PHỤC TRIỆT ĐỂ LỖI LOSSYSETITEM TRÊN PYTHON 3.14
def dong_bo_hoa_don(df_target):
    df_res = df_target.copy()
    if 'Số hóa đơn' in df_res.columns:
        df_res['Số hóa đơn'] = df_res['Số hóa đơn'].astype(object).fillna('').astype(str)
    if 'Ngày xuất hóa đơn' in df_res.columns:
        df_res['Ngày xuất hóa đơn'] = df_res['Ngày xuất hóa đơn'].astype(object).fillna('').astype(str)
    if 'Giá trị xuất hóa đơn' in df_res.columns:
        df_res['Giá trị xuất hóa đơn'] = pd.to_numeric(clean_tien_series(df_res['Giá trị xuất hóa đơn']), errors='coerce').fillna(0)

    mask_chua_hd = df_res['Số hóa đơn'].isna() | df_res['Số hóa đơn'].astype(str).str.strip().isin(['', 'nan', 'None', '0'])
    df_res.loc[mask_chua_hd, 'Ngày xuất hóa đơn'] = ""
    df_res.loc[mask_chua_hd, 'Giá trị xuất hóa đơn'] = 0
    return df_res

def chuan_hoa_kieu_du_lieu(df_input):
    df_out = df_input.copy()
    for col in TEXT_COLUMNS:
        if col in df_out.columns:
            if col == 'Số lệnh sửa chữa':
                df_out[col] = df_out[col].fillna('').astype(str).apply(clean_lsc_giu_gach)
            else:
                df_out[col] = df_out[col].fillna('').astype(str).replace('nan', '')
    
    if 'Thời gian đóng LSC' in df_out.columns:
        df_out['Thời gian đóng LSC'] = df_out['Thời gian đóng LSC'].apply(chuyen_ngay_gio_sortable)

    for col in COT_TIEN + ['Giá trị xuất hóa đơn']:
        if col in df_out.columns:
            df_out[col] = pd.to_numeric(clean_tien_series(df_out[col]), errors='coerce').fillna(0)
    return df_out

def loc_chuan_tu_29_thang_8(df_target):
    if df_target is None or len(df_target) == 0:
        return df_target
    df_res = df_target.copy()
    df_res['temp_norm_key'] = df_res['Số lệnh sửa chữa'].apply(norm_lsc_key)
    df_res = df_res[df_res['temp_norm_key'].str.len() > 3]

    def check_ngay_hop_le(row):
        d = lay_ngay_lsc(row.get('Số lệnh sửa chữa', ''), row.get('Thời gian đóng LSC', ''))
        if pd.isna(d):
            return True
        return d >= MOC_BAT_DAU_DATE

    mask_hop_le = df_res.apply(check_ngay_hop_le, axis=1)
    df_res = df_res[mask_hop_le]

    df_res['has_hd'] = df_res['Số hóa đơn'].fillna('').astype(str).str.strip().apply(lambda x: 1 if x and x not in ['0', 'nan', 'None'] else 0)
    df_res = df_res.sort_values(by=['has_hd', 'Giá trị xuất hóa đơn', 'Số tiền thanh toán cuối'], ascending=[True, True, True])
    df_res = df_res.drop_duplicates(subset=['temp_norm_key'], keep='last')
    df_res = df_res.drop(columns=['temp_norm_key', 'has_hd'])
    df_res = df_res.reset_index(drop=True)
    return df_res

# ĐỌC VÀ LƯU TRỰC TIẾP VÀO GOOGLE SHEETS
def load_data_from_gsheets():
    try:
        df_m = conn.read(worksheet="MasterData", ttl=5)
    except Exception:
        df_m = pd.DataFrame(columns=TAT_CA_COT)

    if df_m is None or len(df_m) == 0 or 'Số lệnh sửa chữa' not in df_m.columns:
        df_m = pd.DataFrame(columns=TAT_CA_COT)
    
    df_m.columns = [str(c).strip() for c in df_m.columns]
    
    cols_to_drop = [c for c in ['Tổng bằng chữ', 'Phân loại bảo hành', 'Loại BH'] if c in df_m.columns]
    if cols_to_drop:
        df_m = df_m.drop(columns=cols_to_drop)

    for c in TAT_CA_COT:
        if c not in df_m.columns:
            df_m[c] = ""

    for c in TEXT_COLUMNS:
        if c in df_m.columns:
            df_m[c] = df_m[c].fillna('').astype(str)

    for c in COT_TIEN + ['Giá trị xuất hóa đơn']:
        if c in df_m.columns:
            df_m[c] = pd.to_numeric(clean_tien_series(df_m[c]), errors='coerce').fillna(0)

    # ĐỐI CHIẾU GSM TỪ SHEET GSM_List
    gsm_keys_set = doc_danh_sach_gsm_tu_gsheet()
    if gsm_keys_set:
        df_m['Phân loại KH'] = df_m['Số lệnh sửa chữa'].apply(
            lambda x: "GSM Công nợ" if norm_lsc_key(x) in gsm_keys_set or lay_loi_ma_wo(x) in gsm_keys_set else "KH Thông Thường"
        )
    else:
        if 'Phân loại KH' not in df_m.columns:
            df_m['Phân loại KH'] = "KH Thông Thường"
        df_m['Phân loại KH'] = df_m['Phân loại KH'].replace('', 'KH Thông Thường').fillna("KH Thông Thường")

    df_m = dong_bo_hoa_don(df_m)
    df_m = chuan_hoa_kieu_du_lieu(df_m)
    df_m = loc_chuan_tu_29_thang_8(df_m)
    return df_m[TAT_CA_COT]

def save_data_to_gsheets(df_to_save):
    df_clean = dong_bo_hoa_don(df_to_save.copy())
    df_clean = chuan_hoa_kieu_du_lieu(df_clean)
    df_clean = loc_chuan_tu_29_thang_8(df_clean)
    for c in TEXT_COLUMNS:
        if c in df_clean.columns:
            df_clean[c] = df_clean[c].fillna('').astype(str)
    conn.update(worksheet="MasterData", data=df_clean)

def format_sheet_in_workbook(ws, sheet_name, cols_to_hide=None):
    font_header = Font(name='Segoe UI', size=11, bold=True, color='FFFFFF')
    fill_header = PatternFill(start_color='1F4E78', fill_type='solid')
    fill_even = PatternFill(start_color='F4F7FA', fill_type='solid')
    fill_total = PatternFill(start_color='D9E1F2', fill_type='solid')
    fill_alert = PatternFill(start_color='FCE4D6', fill_type='solid')
    font_total = Font(name='Segoe UI', size=11, bold=True, color='000000')
    border_thin = Side(border_style='thin', color='D3D3D3')
    border_all = Border(left=border_thin, right=border_thin, top=border_thin, bottom=border_thin)
    border_double = Border(top=Side(border_style='thin', color='000000'), bottom=Side(border_style='double', color='000000'))

    ws.row_dimensions[1].height = 28
    for cell in ws[1]:
        cell.font = font_header
        cell.fill = fill_header
        cell.alignment = Alignment(horizontal='center', vertical='center')
        cell.border = border_all

    last_row = ws.max_row
    col_shd_idx = -1
    col_tt_idx = -1
    col_kh_idx = -1
    col_bh_idx = -1
    col_pl_idx = -1

    for c in range(1, ws.max_column + 1):
        c_val = ws.cell(row=1, column=c).value
        if c_val == 'Số hóa đơn': col_shd_idx = c
        elif c_val == 'Trạng thái': col_tt_idx = c
        elif c_val == 'KH thanh toán': col_kh_idx = c
        elif c_val == 'BH thanh toán': col_bh_idx = c
        elif c_val == 'Phân loại KH': col_pl_idx = c

    for r in range(2, last_row + 1):
        ws.row_dimensions[r].height = 22
        is_even = (r % 2 == 0)
        
        is_chua_xuat_hd = False
        if col_shd_idx != -1:
            shd_val = str(ws.cell(row=r, column=col_shd_idx).value or '').strip()
            tt_val = str(ws.cell(row=r, column=col_tt_idx).value or '').strip() if col_tt_idx != -1 else ''
            kh_val = float(ws.cell(row=r, column=col_kh_idx).value or 0) if col_kh_idx != -1 else 0
            bh_val = float(ws.cell(row=r, column=col_bh_idx).value or 0) if col_bh_idx != -1 else 0
            pl_val = str(ws.cell(row=r, column=col_pl_idx).value or '').strip() if col_pl_idx != -1 else ''

            if tt_val in TRANG_THAI_HOAN_THANH and (shd_val in ['', 'None', 'nan', '0']):
                if (kh_val > 0 and pl_val != 'GSM Công nợ') or (bh_val > 0):
                    is_chua_xuat_hd = True

        for c in range(1, ws.max_column + 1):
            cell = ws.cell(row=r, column=c)
            col_name = ws.cell(row=1, column=c).value
            cell.font = Font(name='Segoe UI', size=10)
            cell.border = border_all
            
            if is_chua_xuat_hd and col_name in ['Số lệnh sửa chữa', 'Biển số', 'Số tiền thanh toán cuối', 'Số hóa đơn']:
                cell.fill = fill_alert
            elif is_even:
                cell.fill = fill_even

            if col_name in COT_TIEN or col_name == 'Giá trị xuất hóa đơn':
                cell.number_format = '#,##0'
                cell.alignment = Alignment(horizontal='right', vertical='center')
            elif col_name in ['Số lệnh sửa chữa', 'Trạng thái', 'Biển số', 'Thời gian đóng LSC', 'Ngày xuất hóa đơn', 'Số hóa đơn', 'Phê duyệt bảo hành', 'Xe GSM', 'Phân loại KH']:
                cell.alignment = Alignment(horizontal='center', vertical='center')
            else:
                cell.alignment = Alignment(horizontal='left', vertical='center')

    if last_row >= 2 and 'Lệnh Đã Hủy' not in sheet_name:
        tot_row = last_row + 1
        ws.row_dimensions[tot_row].height = 25
        ws.cell(row=tot_row, column=1, value="TỔNG CỘNG").font = font_total
        ws.cell(row=tot_row, column=1).alignment = Alignment(horizontal='center', vertical='center')
        ws.cell(row=tot_row, column=1).fill = fill_total

        for c in range(1, ws.max_column + 1):
            cell = ws.cell(row=tot_row, column=c)
            col_name = ws.cell(row=1, column=c).value
            col_let = get_column_letter(c)
            cell.border = border_double
            cell.fill = fill_total
            if col_name in COT_TIEN or col_name == 'Giá trị xuất hóa đơn':
                cell.value = f"=SUBTOTAL(9, {col_let}2:{col_let}{last_row})"
                cell.font = font_total
                cell.number_format = '#,##0'
                cell.alignment = Alignment(horizontal='right', vertical='center')

    ws.auto_filter.ref = f"A1:{get_column_letter(ws.max_column)}{last_row}"
    ws.freeze_panes = 'G2'

    for col in ws.columns:
        col_header = ws.cell(row=1, column=col[0].column).value
        col_let = get_column_letter(col[0].column)
        if cols_to_hide and col_header in cols_to_hide:
            ws.column_dimensions[col_let].hidden = True
        else:
            max_l = max(len(str(c.value or '')) for c in col[:25])
            ws.column_dimensions[col_let].width = max(max_l + 6, 15)

def xuat_excel_don_luong(df_source, ten_sheet, cols_to_hide=None):
    buffer = io.BytesIO()
    df_clean = dong_bo_hoa_don(df_source.copy())
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        df_clean.to_excel(writer, index=False, sheet_name=ten_sheet[:31])
    buffer.seek(0)
    wb = openpyxl.load_workbook(buffer)
    format_sheet_in_workbook(wb.active, ten_sheet, cols_to_hide=cols_to_hide)
    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return out

def xuat_excel_da_sheet_with_progress(df_full, p_bar=None, status_txt=None):
    buffer = io.BytesIO()
    df_clean = dong_bo_hoa_don(df_full.copy())
    
    if status_txt: status_txt.write("⏳ Đang phân tích và tách các nguồn dữ liệu... (20%)")
    if p_bar: p_bar.progress(20)

    mask_chua_hd = (df_clean['Trạng thái'].isin(TRANG_THAI_HOAN_THANH)) & (
        df_clean['Số hóa đơn'].isna() | df_clean['Số hóa đơn'].astype(str).str.strip().isin(['', 'nan', 'None', '0'])
    ) & (
        ((df_clean['KH thanh toán'] > 0) & (df_clean['Phân loại KH'] != 'GSM Công nợ')) |
        (df_clean['BH thanh toán'] > 0)
    )
    df_canh_bao = df_clean[mask_chua_hd].copy()

    df_kh_all = df_clean[(df_clean['Trạng thái'].isin(TRANG_THAI_HOAN_THANH)) & (df_clean['KH thanh toán'] > 0) & (df_clean['Phân loại KH'] != 'GSM Công nợ')].copy()
    df_gsm_debt = df_clean[(df_clean['Trạng thái'].isin(TRANG_THAI_HOAN_THANH)) & (df_clean['KH thanh toán'] > 0) & (df_clean['Phân loại KH'] == 'GSM Công nợ')].copy()
    df_bh_hang = df_clean[(df_clean['Trạng thái'].isin(TRANG_THAI_HOAN_THANH)) & (df_clean['BH hãng thanh toán'] > 0)].copy()
    df_bh = df_clean[(df_clean['Trạng thái'].isin(TRANG_THAI_HOAN_THANH)) & (df_clean['BH thanh toán'] > 0)].copy()
    df_noi_bo = df_clean[(df_clean['Trạng thái'].isin(TRANG_THAI_HOAN_THANH)) & (df_clean['Nội bộ thanh toán'] > 0)].copy()
    df_huy = df_clean[df_clean['Trạng thái'] == 'Đã hủy'].copy()

    if status_txt: status_txt.write("⏳ Đang ghi dữ liệu vào 8 sheet Excel... (50%)")
    if p_bar: p_bar.progress(50)

    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        df_canh_bao.to_excel(writer, index=False, sheet_name='🚨 Chưa Xuất HĐ (KH & BH)')
        df_kh_all.to_excel(writer, index=False, sheet_name='1. KH Thanh Toán (Đã Đóng)')
        df_gsm_debt.to_excel(writer, index=False, sheet_name='2. GSM Công NỢ (Tách Riêng)')
        df_bh_hang.to_excel(writer, index=False, sheet_name='3. Bảo Hành Hãng (W)')
        df_bh.to_excel(writer, index=False, sheet_name='4. Bảo Hiểm (Insurance)')
        df_noi_bo.to_excel(writer, index=False, sheet_name='5. Nội Bộ Thanh Toán')
        df_huy.to_excel(writer, index=False, sheet_name='6. Lệnh Đã Hủy')
        df_clean.to_excel(writer, index=False, sheet_name='Tổng Hợp Toàn Bộ (Gốc)')

    buffer.seek(0)
    wb = openpyxl.load_workbook(buffer)
    total_sheets = len(wb.sheetnames)
    
    for s_i, s_name in enumerate(wb.sheetnames):
        pct = int(50 + ((s_i + 1) / total_sheets) * 45)
        if status_txt: status_txt.write(f"🎨 Đang định dạng chuẩn đẹp Sheet [{s_name}]... ({pct}%)")
        if p_bar: p_bar.progress(pct)
        format_sheet_in_workbook(wb[s_name], s_name)

    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    if p_bar: p_bar.progress(100)
    if status_txt: status_txt.write("✅ ĐÃ CHẠY XONG CHU TRÌNH TẠO BÁO CÁO! (100%)")
    return out

# --- PHÂN QUYỀN ĐĂNG NHẬP & GHI NHỚ 10 NGÀY ---
if "logged_in" not in st.session_state:
    st.session_state.logged_in = False
if "username" not in st.session_state:
    st.session_state.username = ""

token_url = st.query_params.get("auth_token")
if token_url and not st.session_state.logged_in:
    valid_user = xac_thuc_auth_token(token_url)
    if valid_user:
        st.session_state.logged_in = True
        st.session_state.username = valid_user

with st.sidebar:
    st.subheader("🔐 Quyền Truy Cập Hệ Thống")
    if not st.session_state.logged_in:
        st.info("Chế độ hiện tại: **Chỉ xem & Tải báo cáo**")
        with st.form("form_login"):
            u = st.text_input("Tài khoản:", value="")
            p = st.text_input("Mật khẩu:", type="password", value="")
            remember_me = st.checkbox("Ghi nhớ đăng nhập (10 ngày)", value=True)
            btn_login = st.form_submit_button("Đăng Nhập Quản Trị")
            if btn_login:
                if u in DANH_SACH_ADMIN and DANH_SACH_ADMIN[u] == p:
                    st.session_state.logged_in = True
                    st.session_state.username = u
                    if remember_me:
                        new_token = tao_auth_token(u, so_ngay=10)
                        st.query_params["auth_token"] = new_token
                    st.success("✅ Đăng nhập thành công!")
                    st.rerun()
                else:
                    st.error("❌ Sai tài khoản hoặc mật khẩu!")
    else:
        st.success(f"Xin chào: **{st.session_state.username}** (Admin)")
        st.caption("🟢 Trạng thái: Đã ghi nhớ đăng nhập 10 ngày")
        if st.button("🚪 Đăng Xuất"):
            st.session_state.logged_in = False
            st.session_state.username = ""
            if "auth_token" in st.query_params:
                del st.query_params["auth_token"]
            st.rerun()

# --- GIAO DIỆN CHÍNH ---
st.title("🚗 Quản Trị Dịch Vụ, Hóa Đơn & Đối Soát Cyber")
st.caption("☁️ **Lưu trữ đám mây:** Kết nối trực tiếp Google Sheets. Hoạt động từ **29/08/2026**.")

df_master = load_data_from_gsheets()

if st.session_state.logged_in:
    tabs = st.tabs([
        "📊 1. Bảng Tính Web & Phân Luồng",
        "📥 2. Nạp Dữ Liệu DMS Mới",
        "🧾 3. Khớp File Hóa Đơn Kế Toán",
        "🛡️ 4. Import Phê Duyệt Bảo Hành (Excel)",
        "🚕 5. Quản Lý Công NỢ GSM (Google Sheets)",
        "🔍 6. Đối Soát Upload Lên Cyber"
    ])
    tab_work = tabs[0]
    tab_import = tabs[1]
    tab_inv = tabs[2]
    tab_bh_import = tabs[3]
    tab_gsm_import = tabs[4]
    tab_cyber = tabs[5]
else:
    tabs = st.tabs(["📊 1. Bảng Tính Tra Cứu & Báo Cáo"])
    tab_work = tabs[0]

# TAB 1: BẢNG TÍNH WEB
with tab_work:
    df_hoanthanh = df_master[df_master['Trạng thái'].isin(TRANG_THAI_HOAN_THANH)]
    df_chuahoanthanh = df_master[~df_master['Trạng thái'].isin(TRANG_THAI_HOAN_THANH + ['Đã hủy'])]
    chua_hoan_thanh_cnt = len(df_chuahoanthanh)

    df_kh_total = df_hoanthanh[(df_hoanthanh['KH thanh toán'] > 0) & (df_hoanthanh['Phân loại KH'] != 'GSM Công nợ')]
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

    df_bh_hang = df_hoanthanh[df_hoanthanh['BH hãng thanh toán'] > 0]
    bh_chua_duyet_cnt = df_bh_hang[df_bh_hang['Phê duyệt bảo hành'].isin(['Chờ duyệt', None, 'nan', ''])].shape[0]
    bh_total_cnt = len(df_bh_hang)

    df_bh = df_hoanthanh[df_hoanthanh['BH thanh toán'] > 0]
    bh_da_hd_cnt = df_bh[df_bh['Số hóa đơn'].notna() & (~df_bh['Số hóa đơn'].astype(str).str.strip().isin(['', 'nan', 'None', '0']))].shape[0]
    bh_total_all_cnt = len(df_bh)
    bh_chua_hd_cnt = bh_total_all_cnt - bh_da_hd_cnt
    bh_chua_hd_amt = df_bh[df_bh['Số hóa đơn'].isna() | df_bh['Số hóa đơn'].astype(str).str.strip().isin(['', 'nan', 'None', '0'])]['BH thanh toán'].sum()

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
            f"🚨 **CẢNH BÁO CHƯA XUẤT HÓA ĐƠN (CHỈ LỆNH ĐÃ HOÀN THÀNH - KH & BẢO HIỂM): {tong_chua_hd_cnt} lệnh | Tổng tiền: {tong_chua_hd_amt:,.0f} đ** "
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
        tim_kiem_tu_do = st.text_input("🔍 Tìm kiếm nhanh (Biển số / LSC / Tên bất kỳ):", "", placeholder="VD: 81A13363, C23401-WO-260909-0014...")

    if luong_data == "1. KH Thanh Toán (Đã hoàn thành lệnh)":
        df_show = df_kh_total.copy()
        sheet_file_name = "1_KH_Thanh_Toan"
    elif luong_data == "2. GSM Công Nợ (Lấy từ Google Sheets)":
        df_show = df_gsm_all.copy()
        sheet_file_name = "2_GSM_Cong_No"
    elif luong_data == "3. Bảo Hành Hãng (W) - Phê duyệt":
        df_show = df_bh_hang.copy()
        sheet_file_name = "3_Bao_Hanh_Hang"
    elif luong_data == "4. Bảo Hiểm (Insurance)":
        df_show = df_bh.copy()
        sheet_file_name = "4_Bao_Hiem"
    elif luong_data == "5. Nội bộ thanh toán":
        df_show = df_hoanthanh[df_hoanthanh['Nội bộ thanh toán'] > 0].copy()
        sheet_file_name = "5_Noi_Bo"
    elif luong_data == "🚨 Lệnh Đã Xong Chưa Xuất HĐ (Chỉ KH & Bảo Hiểm)":
        mask_target = (df_master['Trạng thái'].isin(TRANG_THAI_HOAN_THANH)) & (
            ((df_master['KH thanh toán'] > 0) & (df_master['Phân loại KH'] != 'GSM Công nợ')) |
            (df_master['BH thanh toán'] > 0)
        ) & (df_master['Số hóa đơn'].isna() | df_master['Số hóa đơn'].astype(str).str.strip().isin(['', 'nan', 'None', '0']))
        df_show = df_master[mask_target].copy()
        sheet_file_name = "Canh_Bao_Chua_Xuat_HD"
    elif luong_data == "6. Xem Lệnh Chưa Hoàn Thành (Báo giá & Đang sửa chữa)":
        df_show = df_chuahoanthanh.copy()
        sheet_file_name = "6_Lenh_Chua_Xong"
    elif luong_data == "7. Xem Lệnh Đã Hủy":
        df_show = df_master[df_master['Trạng thái'] == 'Đã hủy'].copy()
        sheet_file_name = "7_Lenh_Da_Huy"
    else:
        df_show = df_master.copy()
        sheet_file_name = "Tong_Hop_Toan_Bo"

    df_show = loc_chuan_tu_29_thang_8(df_show)

    st.markdown("##### 📌 Lọc Chi Tiết Theo Tiêu Đề Cột:")
    f_box1, f_box2, f_box3 = st.columns(3)

    df_show['ngay_dong_lsc_loc'] = df_show['Thời gian đóng LSC'].astype(str).str.extract(r'(\b\d{4}[/-]\d{1,2}[/-]\d{1,2}\b)')[0].fillna('')

    with f_box1:
        list_all_days = sorted([x for x in df_show['ngay_dong_lsc_loc'].unique() if x], reverse=True)
        sel_ngay_dong = st.multiselect("📅 Lọc Ngày đóng LSC:", options=list_all_days, default=[], placeholder="Chọn ngày đóng (VD: 2026/09/09)...")

    with f_box2:
        list_cvdv = sorted([x for x in df_show['Cố vấn dịch vụ'].dropna().unique() if str(x).strip()])
        sel_cvdv = st.multiselect("👨‍🔧 Lọc Cố vấn dịch vụ:", options=list_cvdv, default=[], placeholder="Chọn CVDV...")

    with f_box3:
        list_trang_thai = sorted([x for x in df_show['Trạng thái'].dropna().unique() if str(x).strip()])
        sel_trang_thai = st.multiselect("📋 Lọc Trạng thái:", options=list_trang_thai, default=[], placeholder="Đã đóng / Sẵn sàng...")

    if loc_canh_bao_hd == "Chỉ xe CHƯA có HĐ":
        df_show = df_show[(df_show['Số hóa đơn'].isna()) | (df_show['Số hóa đơn'].astype(str).str.strip().isin(['', 'nan', 'None', '0']))]
    elif loc_canh_bao_hd == "Đã có hóa đơn":
        df_show = df_show[df_show['Số hóa đơn'].notna() & (~df_show['Số hóa đơn'].astype(str).str.strip().isin(['', 'nan', 'None', '0']))]

    if tim_kiem_tu_do:
        kw = str(tim_kiem_tu_do).strip()
        kw_norm = norm_lsc_key(kw)
        mask = (
            df_show['Biển số'].astype(str).str.contains(kw, case=False, na=False) |
            df_show['Số lệnh sửa chữa'].astype(str).str.contains(kw, case=False, na=False) |
            df_show['Số lệnh sửa chữa'].apply(norm_lsc_key).str.contains(kw_norm, case=False, na=False) |
            df_show['Tên khách hàng'].astype(str).str.contains(kw, case=False, na=False) |
            df_show['Thời gian đóng LSC'].astype(str).str.contains(kw, case=False, na=False) |
            df_show['Số hóa đơn'].astype(str).str.contains(kw, case=False, na=False)
        )
        df_show = df_show[mask]

    if sel_ngay_dong:
        df_show = df_show[df_show['ngay_dong_lsc_loc'].isin(sel_ngay_dong)]

    if sel_cvdv:
        df_show = df_show[df_show['Cố vấn dịch vụ'].isin(sel_cvdv)]

    if sel_trang_thai:
        df_show = df_show[df_show['Trạng thái'].isin(sel_trang_thai)]

    df_show = df_show.drop(columns=['ngay_dong_lsc_loc'])

    tat_ca_cot_bang = [
        'Số lệnh sửa chữa', 'Trạng thái', 'Cố vấn dịch vụ', 'Tên khách hàng', 
        'Thời gian đóng LSC', 'Biển số', 'Xe GSM', 'Phân loại KH', 'Phê duyệt bảo hành',
        'Tổng tiền công', 'Tổng tiền phụ tùng', 'Tổng trước chiết khấu',
        'Chiết khấu đại lý', 'Tổng chiết khấu', 'Tổng sau chiết khấu',
        'Tiền VAT', 'Tổng có VAT', 'Chiết khấu VinClub',
        'Số tiền thanh toán cuối', 'Tiền đặt cọc',
        'KH thanh toán', 'BH thanh toán', 'BH hãng thanh toán', 'Nội bộ thanh toán',
        'Số hóa đơn', 'Ngày xuất hóa đơn', 'Giá trị xuất hóa đơn'
    ]

    cols_base = ['Số lệnh sửa chữa', 'Trạng thái', 'Cố vấn dịch vụ', 'Tên khách hàng', 'Thời gian đóng LSC', 'Biển số']
    cols_hd = ['Số hóa đơn', 'Ngày xuất hóa đơn', 'Giá trị xuất hóa đơn']
    cols_4_thanh_toan = ['KH thanh toán', 'BH thanh toán', 'BH hãng thanh toán', 'Nội bộ thanh toán']

    if luong_data == "1. KH Thanh Toán (Đã hoàn thành lệnh)":
        default_cols = cols_base + ['Số tiền thanh toán cuối', 'KH thanh toán'] + cols_hd
    elif luong_data == "2. GSM Công Nợ (Lấy từ Google Sheets)":
        default_cols = cols_base + ['Phân loại KH', 'Số tiền thanh toán cuối', 'KH thanh toán'] + cols_hd
    elif luong_data == "3. Bảo Hành Hãng (W) - Phê duyệt":
        default_cols = cols_base + ['BH hãng thanh toán', 'Phê duyệt bảo hành'] + cols_hd
    elif luong_data == "4. Bảo Hiểm (Insurance)":
        default_cols = cols_base + ['Số tiền thanh toán cuối', 'BH thanh toán'] + cols_hd
    elif luong_data == "5. Nội bộ thanh toán":
        default_cols = cols_base + ['Số tiền thanh toán cuối', 'Nội bộ thanh toán'] + cols_hd
    elif luong_data == "🚨 Lệnh Đã Xong Chưa Xuất HĐ (Chỉ KH & Bảo Hiểm)":
        default_cols = cols_base + ['Số tiền thanh toán cuối', 'KH thanh toán', 'BH thanh toán'] + cols_hd
    elif luong_data == "6. Xem Lệnh Chưa Hoàn Thành (Báo giá & Đang sửa chữa)":
        default_cols = cols_base + ['Số tiền thanh toán cuối'] + cols_4_thanh_toan
    elif luong_data == "7. Xem Lệnh Đã Hủy":
        default_cols = cols_base + ['Số tiền thanh toán cuối']
    else:
        default_cols = cols_base + ['Phân loại KH', 'Phê duyệt bảo hành', 'Số tiền thanh toán cuối'] + cols_4_thanh_toan + cols_hd

    with st.expander("👁️ Tùy biến Cột hiển thị", expanded=False):
        selected_cols = st.multiselect(
            "Chọn các cột bạn muốn xem trên bảng:",
            options=tat_ca_cot_bang,
            default=[c for c in default_cols if c in tat_ca_cot_bang],
            key=f"col_filter_{luong_data}"
        )

    df_show = chuan_hoa_kieu_du_lieu(df_show)
    df_show = df_show.reset_index(drop=True)
    df_show.insert(0, 'STT', range(1, len(df_show) + 1))

    actual_cols = ['STT'] + [c for c in tat_ca_cot_bang if c in selected_cols and c in df_show.columns]
    df_render = df_show[actual_cols].copy()
    cols_to_hide_in_excel = [c for c in tat_ca_cot_bang if c not in selected_cols]

    is_admin = st.session_state.logged_in

    col_cfg = {
        "STT": st.column_config.NumberColumn("STT", disabled=True, width="small"),
        "Số lệnh sửa chữa": st.column_config.TextColumn("Số LSC", disabled=True),
        "Trạng thái": st.column_config.TextColumn("Trạng thái", disabled=True),
        "Cố vấn dịch vụ": st.column_config.TextColumn("Cố vấn dịch vụ", disabled=not is_admin),
        "Tên khách hàng": st.column_config.TextColumn("Tên khách hàng", disabled=not is_admin),
        "Thời gian đóng LSC": st.column_config.TextColumn("Thời gian đóng LSC", disabled=True),
        "Biển số": st.column_config.TextColumn("Biển số", disabled=not is_admin),
        "Phân loại KH": st.column_config.SelectboxColumn("Phân loại KH", options=["KH Thông Thường", "GSM Công nợ"], disabled=not is_admin, required=True),
        "Xe GSM": st.column_config.TextColumn("Xe GSM", disabled=True),
        "Phê duyệt bảo hành": st.column_config.SelectboxColumn(
            "Phê duyệt BH",
            options=["Chờ duyệt", "Đã duyệt", "Từ chối", "Không duyệt"],
            disabled=not is_admin,
            required=False
        ),
        "Tổng tiền công": st.column_config.NumberColumn("Tổng tiền công", format="%,d đ", disabled=True),
        "Tổng tiền phụ tùng": st.column_config.NumberColumn("Tổng phụ tùng", format="%,d đ", disabled=True),
        "Tổng trước chiết khấu": st.column_config.NumberColumn("Trước CK", format="%,d đ", disabled=True),
        "Chiết khấu đại lý": st.column_config.NumberColumn("CK đại lý", format="%,d đ", disabled=True),
        "Tổng chiết khấu": st.column_config.NumberColumn("Tổng CK", format="%,d đ", disabled=True),
        "Tổng sau chiết khấu": st.column_config.NumberColumn("Sau CK", format="%,d đ", disabled=True),
        "Tiền VAT": st.column_config.NumberColumn("Tiền VAT", format="%,d đ", disabled=True),
        "Tổng có VAT": st.column_config.NumberColumn("Tổng có VAT", format="%,d đ", disabled=True),
        "Chiết khấu VinClub": st.column_config.NumberColumn("CK VinClub", format="%,d đ", disabled=True),
        "Số tiền thanh toán cuối": st.column_config.NumberColumn("Tổng TT cuối", format="%,d đ", disabled=True),
        "Tiền đặt cọc": st.column_config.NumberColumn("Tiền đặt cọc", format="%,d đ", disabled=True),
        "KH thanh toán": st.column_config.NumberColumn("KH thanh toán", format="%,d đ", disabled=True),
        "BH thanh toán": st.column_config.NumberColumn("BH thanh toán", format="%,d đ", disabled=True),
        "BH hãng thanh toán": st.column_config.NumberColumn("BH hãng thanh toán", format="%,d đ", disabled=True),
        "Nội bộ thanh toán": st.column_config.NumberColumn("Nội bộ thanh toán", format="%,d đ", disabled=True),
        "Số hóa đơn": st.column_config.TextColumn("Số HĐ", disabled=not is_admin),
        "Ngày xuất hóa đơn": st.column_config.TextColumn("Ngày HĐ", disabled=not is_admin),
        "Giá trị xuất hóa đơn": st.column_config.NumberColumn("Tiền HĐ", format="%,d đ", disabled=not is_admin),
    }

    edited_df = st.data_editor(
        df_render,
        use_container_width=True,
        height=560,
        column_config=col_cfg,
        disabled=(not is_admin),
        num_rows="fixed",
        hide_index=True,
        key=f"data_editor_table_gsheets_{luong_data}"
    )

    st.markdown("---")
    c_btn1, c_btn1_undo, c_btn_clean, c_btn2, c_btn3 = st.columns([2.5, 1.8, 2.5, 2.2, 3.0])

    with c_btn1:
        if is_admin:
            if st.button("☁️ Lưu Trực Tiếp Lên Google Sheets", type="primary", use_container_width=True):
                prog = st.progress(0)
                edit_dict = edited_df.set_index('Số lệnh sửa chữa').to_dict('index')
                total_rows = len(df_master)
                for idx in range(total_rows):
                    r_lsc = df_master.iloc[idx]['Số lệnh sửa chữa']
                    if r_lsc in edit_dict:
                        for c in ['Phê duyệt bảo hành', 'Phân loại KH', 'Số hóa đơn', 'Ngày xuất hóa đơn', 'Giá trị xuất hóa đơn', 'Cố vấn dịch vụ', 'Tên khách hàng', 'Biển số']:
                            if c in edit_dict[r_lsc]:
                                df_master.at[df_master.index[idx], c] = edit_dict[r_lsc][c]
                    if idx % 100 == 0:
                        prog.progress(int((idx / total_rows) * 80))
                        
                save_data_to_gsheets(df_master)
                prog.progress(100)
                st.success("✅ Dữ liệu đã lưu vĩnh viễn lên Google Sheets!")

    with c_btn1_undo:
        if is_admin:
            if st.button("↩️ Hủy Thay Đổi", use_container_width=True):
                st.rerun()

    with c_btn_clean:
        if is_admin:
            if st.button("🧹 Lọc Chặn Chuẩn (Từ 29/08)", use_container_width=True):
                truoc_do = len(df_master)
                df_master = loc_chuan_tu_29_thang_8(df_master)
                save_data_to_gsheets(df_master)
                da_xoa = truoc_do - len(df_master)
                st.success(f"✅ Đã dọn dẹp sạch {da_xoa} dòng cũ trên Google Sheets.")
                st.rerun()

    with c_btn2:
        df_export_single = df_show.drop(columns=['STT']) if 'STT' in df_show.columns else df_show
        excel_single_bytes = xuat_excel_don_luong(df_export_single, luong_data, cols_to_hide=cols_to_hide_in_excel)
        st.download_button(
            label=f"⬇️ Tải Riêng Luồng Này",
            data=excel_single_bytes,
            file_name=f"Bao_Cao_{sheet_file_name}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

    with c_btn3:
        with st.expander("📦 Xuất File Excel Tổng Hợp", expanded=False):
            if st.button("🚀 Khởi tạo toàn bộ 8 Sheet", use_container_width=True):
                p_bar_excel = st.progress(0)
                st_txt_excel = st.empty()
                excel_bytes = xuat_excel_da_sheet_with_progress(df_master, p_bar_excel, st_txt_excel)
                st.download_button(
                    label="⬇️ TẢI BÁO CÁO 8 SHEET VỀ MÁY",
                    data=excel_bytes,
                    file_name="Bao_Cao_VinFast_Tu_29_08_Den_Nay.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary",
                    use_container_width=True
                )

# CÁC TAB CHỨC NĂNG KHI ĐÃ ĐĂNG NHẬP
if is_admin:
    # TAB 2: NẠP DỮ LIỆU DMS
    with tab_import:
        st.subheader("Nạp file dữ liệu phân phối VinFast định kỳ")
        st.info("💡 Dữ liệu nạp mới sẽ được đồng bộ và lưu vĩnh viễn lên Google Sheets.")
        up_db = st.file_uploader("Kéo thả file Database mới vào đây", type=['csv', 'xlsx'], key='up_dms')

        if up_db:
            df_raw = doc_file_db(up_db)
            st.write("🔎 Xem trước file vừa tải lên:")
            st.dataframe(df_raw.head(3), use_container_width=True)

            if st.button("🚀 BẮT ĐẦU NẠP VÀ LƯU LÊN GOOGLE SHEETS", type="primary", use_container_width=True):
                df_inc = pd.DataFrame()
                for c in TAT_CA_COT:
                    df_inc[c] = df_raw[c] if c in df_raw.columns else ""

                for c in COT_TIEN:
                    df_inc[c] = pd.to_numeric(clean_tien_series(df_inc[c]), errors='coerce').fillna(0)

                df_inc['Số lệnh sửa chữa'] = df_inc['Số lệnh sửa chữa'].astype(str).apply(clean_lsc_giu_gach)
                df_inc['Phân loại KH'] = "KH Thông Thường"
                df_inc.loc[(df_inc['BH hãng thanh toán'] > 0) & (df_inc['Phê duyệt bảo hành'].isna()), 'Phê duyệt bảo hành'] = "Chờ duyệt"
                df_inc = dong_bo_hoa_don(df_inc)
                df_inc = chuan_hoa_kieu_du_lieu(df_inc)
                df_inc = loc_chuan_tu_29_thang_8(df_inc)

                p_bar_dms = st.progress(0)
                txt_dms = st.empty()
                txt_dms.write("⏳ Đang đối soát và cập nhật dữ liệu... (10%)")

                df_master = df_master.reset_index(drop=True)
                master_norm_map = {norm_lsc_key(lsc): idx for idx, lsc in enumerate(df_master['Số lệnh sửa chữa'])}
                new_records = []
                status_updated_cnt = 0
                total_inc = len(df_inc)

                for i, (_, row) in enumerate(df_inc.iterrows()):
                    lsc = row['Số lệnh sửa chữa']
                    k_norm = norm_lsc_key(lsc)
                    if not k_norm:
                        continue
                    if k_norm not in master_norm_map:
                        new_records.append(row)
                        master_norm_map[k_norm] = len(df_master) + len(new_records) - 1
                    else:
                        m_idx = master_norm_map[k_norm]
                        old_status = str(df_master.iloc[m_idx]['Trạng thái'])
                        new_status = str(row['Trạng thái'])
                        if old_status != new_status:
                            df_master.iloc[m_idx, df_master.columns.get_loc('Trạng thái')] = new_status
                            status_updated_cnt += 1
                        for col in COT_TIEN + ['Thời gian đóng LSC', 'Xe GSM']:
                            if col in df_master.columns and col in row:
                                df_master.iloc[m_idx, df_master.columns.get_loc(col)] = row[col]
                    
                    if i % 30 == 0:
                        pct = int(10 + (i / total_inc) * 80)
                        p_bar_dms.progress(pct)

                if new_records:
                    df_master = pd.concat([df_master, pd.DataFrame(new_records)], ignore_index=True)

                save_data_to_gsheets(df_master)
                p_bar_dms.progress(100)
                txt_dms.empty()
                st.success(f"✅ ĐÃ ĐỒNG BỘ LÊN GOOGLE SHEETS! Thêm **{len(new_records)}** lệnh mới, cập nhật **{status_updated_cnt}** lệnh.")
                st.rerun()

    # TAB 3: KHỚP HÓA ĐƠN
    with tab_inv:
        st.subheader("Khớp file Hóa Đơn kế toán với Hệ Thống")
        st.caption("Tự động gộp tất cả hóa đơn cùng 1 LSC: nối số HĐ bằng dấu phẩy và cộng dồn tiền chính xác 100%.")
        up_inv = st.file_uploader("Tải file Bảng Kê Hóa Đơn", type=['csv', 'xlsx'], key='up_inv_tab')

        if up_inv:
            df_inv, col_labels = doc_file_hoa_don_chuan(up_inv)
            
            def tim_chinh_xac_cot(uu_tien_list, default_idx=0):
                for kw in uu_tien_list:
                    for i, label in enumerate(col_labels):
                        if kw in label.lower():
                            return i
                return default_idx

            idx_lsc = tim_chinh_xac_cot(['ro hãng', 'ro_hãng', 'lệnh hãng', 'ro hang', 'số ro hãng', 'số ro'], 0)
            idx_shd = tim_chinh_xac_cot(['hóa đơn đt', 'hóa đơn điện tử', 'số hóa đơn đt', 'số hđ đt', 'số hóa đơn', 'số hđ'], 1 if len(col_labels) > 1 else 0)
            idx_nhd = tim_chinh_xac_cot(['chứng từ - ngày', 'chứng từ ngày', 'ngày chứng từ', 'ngày hđ', 'ngày'], 0)
            idx_gt  = tim_chinh_xac_cot(['tổng thanh toán', 'thanh toán', 'tổng tiền', 'thành tiền'], 0)

            c1, c2, c3, c4 = st.columns(4)
            opt_indices = list(range(len(col_labels)))
            sel_lsc_idx = c1.selectbox("📌 Cột Số Lệnh SC / RO:", opt_indices, format_func=lambda i: col_labels[i], index=idx_lsc)
            sel_shd_idx = c2.selectbox("🧾 Cột Số Hóa Đơn:", opt_indices, format_func=lambda i: col_labels[i], index=idx_shd)
            
            opt_with_none = [-1] + opt_indices
            sel_nhd_idx = c3.selectbox("📅 Cột Ngày Hóa Đơn:", opt_with_none, format_func=lambda i: "Bỏ qua" if i == -1 else col_labels[i], index=idx_nhd + 1)
            sel_gt_idx  = c4.selectbox("💰 Cột Giá Trị Hóa Đơn:", opt_with_none, format_func=lambda i: "Bỏ qua" if i == -1 else col_labels[i], index=idx_gt + 1)

            if st.button("🚀 BẮT ĐẦU KHỚP HÓA ĐƠN", type="primary", use_container_width=True):
                p_bar_inv = st.progress(0)
                txt_inv = st.empty()
                txt_inv.write("⏳ Đang chuẩn hóa và gom các hóa đơn theo LSC... (15%)")
                p_bar_inv.progress(15)

                df_master = df_master.reset_index(drop=True)
                
                norm_map = {}
                for idx, lsc in enumerate(df_master['Số lệnh sửa chữa']):
                    norm_map[norm_lsc_key(lsc)] = idx
                    c_key = lay_loi_ma_wo(lsc)
                    if c_key:
                        norm_map[c_key] = idx
                
                inv_aggregated = {}
                total_inv_rows = len(df_inv)

                for r_idx in range(total_inv_rows):
                    raw_lsc = str(df_inv.iloc[r_idx, sel_lsc_idx])
                    key_norm = norm_lsc_key(raw_lsc)
                    key_core = lay_loi_ma_wo(raw_lsc)

                    shd_val = str(df_inv.iloc[r_idx, sel_shd_idx]).strip() if pd.notna(df_inv.iloc[r_idx, sel_shd_idx]) else ""
                    if shd_val.endswith('.0'): shd_val = shd_val[:-2]

                    nhd_val = ""
                    if sel_nhd_idx != -1 and pd.notna(df_inv.iloc[r_idx, sel_nhd_idx]):
                        nhd_val = clean_ngay_chuan(df_inv.iloc[r_idx, sel_nhd_idx])

                    gt_val = 0
                    if sel_gt_idx != -1 and pd.notna(df_inv.iloc[r_idx, sel_gt_idx]):
                        try:
                            gt_val = float(str(df_inv.iloc[r_idx, sel_gt_idx]).replace(',', '').replace(' ', ''))
                        except ValueError:
                            gt_val = 0

                    match_key = key_norm if key_norm in norm_map else (key_core if key_core in norm_map else key_norm)

                    if not match_key or len(match_key) < 4:
                        continue

                    if match_key not in inv_aggregated:
                        inv_aggregated[match_key] = {
                            'so_hd': [shd_val] if shd_val and shd_val not in ['', 'nan', 'None'] else [],
                            'ngay_hd': [nhd_val] if nhd_val else [],
                            'tong_tien': gt_val
                        }
                    else:
                        if shd_val and shd_val not in ['', 'nan', 'None'] and shd_val not in inv_aggregated[match_key]['so_hd']:
                            inv_aggregated[match_key]['so_hd'].append(shd_val)
                        if nhd_val and nhd_val not in inv_aggregated[match_key]['ngay_hd']:
                            inv_aggregated[match_key]['ngay_hd'].append(nhd_val)
                        inv_aggregated[match_key]['tong_tien'] += gt_val

                matched_records = []
                total_keys = len(inv_aggregated)

                for k_i, (k_target, val_dict) in enumerate(inv_aggregated.items()):
                    m_idx = norm_map.get(k_target)

                    if m_idx is not None:
                        if str(df_master.iloc[m_idx]['Trạng thái']) == 'Đã hủy':
                            continue

                        shd_str = ", ".join(val_dict['so_hd'])
                        nhd_str = ", ".join(val_dict['ngay_hd'])
                        gt_tong = val_dict['tong_tien']

                        df_master.iloc[m_idx, df_master.columns.get_loc('Số hóa đơn')] = shd_str
                        df_master.iloc[m_idx, df_master.columns.get_loc('Ngày xuất hóa đơn')] = nhd_str
                        df_master.iloc[m_idx, df_master.columns.get_loc('Giá trị xuất hóa đơn')] = gt_tong

                        matched_records.append({
                            'Số LSC': df_master.iloc[m_idx]['Số lệnh sửa chữa'],
                            'Số Hóa Đơn': shd_str,
                            'Ngày HĐ': nhd_str,
                            'Giá Trị HĐ': f"{gt_tong:,.0f}" if gt_tong else "0"
                        })

                    if k_i % 30 == 0:
                        pct = int(15 + (k_i / max(total_keys, 1)) * 75)
                        p_bar_inv.progress(pct)

                save_data_to_gsheets(df_master)
                p_bar_inv.progress(100)
                txt_inv.empty()
                st.success(f"✅ ĐÃ KHỚP & LƯU LÊN GOOGLE SHEETS! Khớp và gộp {len(matched_records)} lệnh.")

    # TAB 4: IMPORT BẢO HÀNH
    with tab_bh_import:
        st.subheader("🛡️ Import Danh Sách Phê Duyệt Bảo Hành Hãng Tự Động")
        up_bh_file = st.file_uploader("Tải file duyệt bảo hành (XLSX / CSV)", type=['csv', 'xlsx'], key='up_bh_file')
        
        if up_bh_file:
            df_bh_input = doc_file_db(up_bh_file)
            bh_cols = list(df_bh_input.columns)
            b1, b2, b3, b4 = st.columns(4)
            
            idx_bh_lsc = next((i for i, c in enumerate(bh_cols) if any(k in c.lower() for k in ['lệnh', 'lsc', 'ro'])), 0)
            idx_bh_tt = next((i for i, c in enumerate(bh_cols) if any(k in c.lower() for k in ['duyệt', 'trạng thái', 'status'])), 0)
            idx_bh_shd = next((i for i, c in enumerate(bh_cols) if any(k in c.lower() for k in ['hóa đơn', 'số hđ'])), 0)
            idx_bh_nhd = next((i for i, c in enumerate(bh_cols) if any(k in c.lower() for k in ['ngày'])), 0)

            s_bh_lsc = b1.selectbox("Cột Số LSC:", bh_cols, index=idx_bh_lsc)
            s_bh_tt  = b2.selectbox("Cột Trạng Thái Duyệt:", bh_cols, index=idx_bh_tt)
            s_bh_shd = b3.selectbox("Cột Số Hóa Đơn:", [None] + bh_cols, index=idx_bh_shd + 1 if idx_bh_shd != 0 else 0)
            s_bh_nhd = b4.selectbox("Cột Ngày Hóa Đơn:", [None] + bh_cols, index=idx_bh_nhd + 1 if idx_bh_nhd != 0 else 0)

            if st.button("🚀 CẬP NHẬT PHÊ DUYỆT LÊN GOOGLE SHEETS", type="primary", use_container_width=True):
                p_bar_bh = st.progress(0)
                txt_bh = st.empty()
                txt_bh.write("⏳ Đang đối soát danh sách bảo hành...")

                df_master = df_master.reset_index(drop=True)
                norm_map = {norm_lsc_key(lsc): idx for idx, lsc in enumerate(df_master['Số lệnh sửa chữa'])}
                
                bh_updated = 0
                total_bh_rows = len(df_bh_input)

                for r_i, (_, r) in enumerate(df_bh_input.iterrows()):
                    key_norm = norm_lsc_key(r[s_bh_lsc])
                    if not key_norm:
                        continue
                    m_idx = norm_map.get(key_norm)

                    if m_idx is not None:
                        val_tt = str(r[s_bh_tt]).strip()
                        if any(w in val_tt.lower() for w in ['đã duyệt', 'duyệt', 'approved', 'pass', 'ok']):
                            norm_tt = "Đã duyệt"
                        elif any(w in val_tt.lower() for w in ['từ chối', 'reject', 'không duyệt']):
                            norm_tt = "Từ chối"
                        else:
                            norm_tt = "Chờ duyệt"
                        
                        df_master.iloc[m_idx, df_master.columns.get_loc('Phê duyệt bảo hành')] = norm_tt
                        
                        if s_bh_shd and pd.notna(r[s_bh_shd]):
                            df_master.iloc[m_idx, df_master.columns.get_loc('Số hóa đơn')] = str(r[s_bh_shd]).strip()
                        if s_bh_nhd and pd.notna(r[s_bh_nhd]):
                            df_master.iloc[m_idx, df_master.columns.get_loc('Ngày xuất hóa đơn')] = clean_ngay_chuan(r[s_bh_nhd])
                        bh_updated += 1
                        
                    if r_i % 30 == 0:
                        pct = int((r_i / total_bh_rows) * 90)
                        p_bar_bh.progress(pct)

                save_data_to_gsheets(df_master)
                p_bar_bh.progress(100)
                txt_bh.empty()
                st.success(f"✅ ĐÃ CẬP NHẬT LÊN GOOGLE SHEETS cho **{bh_updated}** lệnh bảo hành.")

    # TAB 5: QUẢN LÝ CÔNG NỢ GSM (ĐỌC TỪ SHEET GSM_List)
    with tab_gsm_import:
        st.subheader("🚕 Quản Lý Công Nợ GSM (Đồng Bộ Google Sheets)")
        st.info("📌 **Quy tắc:** Dữ liệu tự động lấy từ sheet `GSM_List` trên file Google Sheets của bạn. Bạn có thể mở file Google Sheets dán thêm mã bất cứ lúc nào.")

        gsm_keys_set = doc_danh_sach_gsm_tu_gsheet()
        st.write(f"📊 Hiện tại hệ thống đang nhận diện: **{len(gsm_keys_set)}** mã WO GSM.")

        if st.button("🔄 Nạp Lại Dữ Liệu Từ Google Sheets", type="primary", use_container_width=True):
            st.rerun()

        st.markdown("---")
        df_gsm_view = df_master[(df_master['Phân loại KH'] == 'GSM Công nợ') & (df_master['Trạng thái'].isin(TRANG_THAI_HOAN_THANH))].copy()
        
        c_f_g1, c_f_g2 = st.columns([3, 5])
        with c_f_g1:
            loc_tinh_trang_gsm = st.selectbox("Xem tình trạng nợ:", ["Tất cả", "🔴 Chỉ lệnh ĐANG TREO NỢ (Chưa có HĐ)", "🟢 Chỉ lệnh ĐÃ XUẤT HĐ (Đã quyết toán)"])

        if len(df_gsm_view) > 0:
            df_gsm_view['Tình trạng HĐ'] = df_gsm_view['Số hóa đơn'].apply(
                lambda x: "🟢 Đã XHĐ (Đã quyết toán)" if str(x).strip() not in ['', 'nan', 'None', '0'] else "🔴 Đang treo nợ"
            )

            if "ĐANG TREO NỢ" in loc_tinh_trang_gsm:
                df_gsm_view = df_gsm_view[df_gsm_view['Tình trạng HĐ'] == "🔴 Đang treo nợ"]
            elif "ĐÃ XUẤT HĐ" in loc_tinh_trang_gsm:
                df_gsm_view = df_gsm_view[df_gsm_view['Tình trạng HĐ'] == "🟢 Đã XHĐ (Đã quyết toán)"]

            cols_gsm_show = ['Số lệnh sửa chữa', 'Trạng thái', 'Biển số', 'Tên khách hàng', 'KH thanh toán', 'Số tiền thanh toán cuối', 'Số hóa đơn', 'Ngày xuất hóa đơn', 'Tình trạng HĐ']
            
            cfg_gsm = {
                "KH thanh toán": st.column_config.NumberColumn("KH thanh toán", format="%,d đ"),
                "Số tiền thanh toán cuối": st.column_config.NumberColumn("Tổng TT cuối", format="%,d đ")
            }
            st.dataframe(df_gsm_view[cols_gsm_show], use_container_width=True, hide_index=True, column_config=cfg_gsm)
        else:
            st.warning("⚠️ Chưa tìm thấy mã WO GSM nào. Hãy kiểm tra sheet `GSM_List` trên Google Sheets của bạn.")

    # TAB 6: ĐỐI SOÁT CYBER
    with tab_cyber:
        st.subheader("🔍 Đối Soát Lệnh Đã Hoàn Thành Chưa Up Lên Phần Mềm Cyber")
        up_cyber_file = st.file_uploader("Tải lên file BẢNG TỔNG HỢP LỆNH SỬA CHỮA từ Cyber (Excel)", type=['xlsx', 'xls', 'csv'], key='up_cyber_file')
        
        if up_cyber_file:
            p_bar_cy = st.progress(0)
            txt_cy = st.empty()
            txt_cy.write("⏳ Đang quét dữ liệu Cyber... (20%)")
            p_bar_cy.progress(20)

            cyber_keys = doc_file_cyber(up_cyber_file)
            df_comp = df_master[df_master['Trạng thái'].isin(TRANG_THAI_HOAN_THANH)].copy()
            df_comp['da_up_cyber'] = df_comp['Số lệnh sửa chữa'].apply(
                lambda x: clean_lsc_giu_gach(x) in cyber_keys or norm_lsc_key(x) in cyber_keys
            )
            
            df_chua_up = df_comp[~df_comp['da_up_cyber']].copy()
            df_da_up = df_comp[df_comp['da_up_cyber']].copy()
            p_bar_cy.progress(100)
            txt_cy.empty()
            
            c1, c2, c3 = st.columns(3)
            c1.metric("📌 Tổng Lệnh Đã Hoàn Thành", f"{len(df_comp):,} lệnh")
            c2.metric("🟢 Đã Up Lên Cyber", f"{len(df_da_up):,} lệnh")
            c3.metric("🚨 CHƯA UP LÊN CYBER", f"{len(df_chua_up):,} lệnh", delta=f"-{len(df_chua_up)} lệnh", delta_color="inverse")
            
            if len(df_chua_up) > 0:
                st.error(f"⚠️ Có {len(df_chua_up)} lệnh hoàn thành chưa được đẩy lên Cyber.")
                cols_display = ['Số lệnh sửa chữa', 'Trạng thái', 'Biển số', 'Cố vấn dịch vụ', 'Tên khách hàng', 'Số tiền thanh toán cuối', 'KH thanh toán', 'BH thanh toán', 'BH hãng thanh toán']
                cols_valid = [c for c in cols_display if c in df_chua_up.columns]
                st.dataframe(df_chua_up[cols_valid], use_container_width=True)
