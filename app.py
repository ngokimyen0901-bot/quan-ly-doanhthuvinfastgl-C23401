import streamlit as st
import pandas as pd
import io
import os
import re
import time
import base64
import hashlib
from datetime import datetime
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Báo Cáo Dịch Vụ VinFast", page_icon="🚗", layout="wide")

MASTER_FILE = "master_database.xlsx"
INVOICE_DB_FILE = "hoa_don_database.xlsx"
ACC_511_FILE = "tk511_database.xlsx"
SECRET_KEY_AUTH = "VinFast_GiaLai_Secret_Key_2026"

DANH_SACH_ADMIN = {
    "admin": "Vinfastgialai@2026##"
}

COT_DINH_DANH = [
    'Số lệnh sửa chữa', 'Trạng thái', 'Cố vấn dịch vụ', 'Tên khách hàng', 
    'Thời gian đóng LSC', 'Biển số', 'Xe GSM'
]

# Đầy đủ tất cả các cột chi tiết tiền từ DMS gốc
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
    'Thời gian đóng LSC', 'Biển số', 'Xe GSM', 'Phân loại KH', 
    'Phê duyệt bảo hành', 'Số hóa đơn', 'Ngày xuất hóa đơn'
]

TRANG_THAI_HOAN_THANH = ['Đã đóng', 'Sẵn sàng bàn giao']

# --- AUTH TOKEN ---
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

# --- CHUẨN HÓA DỮ LIỆU ---
def clean_lsc_giu_gach(val):
    if pd.isna(val) or val is None:
        return ""
    return str(val).split('(')[0].strip()

def norm_lsc_key(val):
    if pd.isna(val) or val is None:
        return ""
    s = str(val).split('(')[0].strip().upper()
    return re.sub(r'[^A-Z0-9]', '', s)

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

def clean_tien_series(ser):
    return (
        ser.astype(str)
        .str.replace(',', '', regex=False)
        .str.replace(' ', '', regex=False)
        .replace(['-', '', 'nan', 'None', ''], '0')
    )

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

def doc_file_so_511(file_obj):
    if file_obj.name.endswith('.csv'):
        df = pd.read_csv(file_obj, header=None)
    else:
        df = pd.read_excel(file_obj, header=None)
    
    # Tìm dòng chứa 'Ngày C.từ' hoặc 'Số hóa đơn' hoặc 'Phát sinh' (thường là dòng 8 hoặc 9)
    h_idx = -1
    for idx, r in df.head(15).iterrows():
        r_txt = " ".join([str(v).lower() for v in r.values if pd.notna(v)])
        if 'ngày c.từ' in r_txt or ('ngày' in r_txt and 'chứng từ' in r_txt) or ('phát sinh' in r_txt and 'có' in r_txt):
            h_idx = idx
            break
            
    if h_idx == -1:
        h_idx = 7

    row_header = df.iloc[h_idx].fillna('').astype(str).tolist()
    if h_idx + 1 < len(df):
        row_sub = df.iloc[h_idx + 1].fillna('').astype(str).tolist()
    else:
        row_sub = [''] * len(row_header)

    col_names = []
    for c_i in range(len(row_header)):
        t = row_header[c_i].strip()
        s = row_sub[c_i].strip()
        name = f"{t} {s}".strip() if s and s.lower() not in t.lower() else (t if t else s)
        col_names.append(name if name else f"Col_{c_i}")

    df_data = df.iloc[h_idx + 2:].copy().reset_index(drop=True)
    df_data.columns = [f"{col_names[i]}_{i}" if col_names.count(col_names[i]) > 1 else col_names[i] for i in range(len(col_names))]
    return df_data, col_names

def doc_file_cyber(file_obj):
    if file_obj.name.endswith('.csv'):
        df_raw = pd.read_csv(file_obj, header=None)
    else:
        df_raw = pd.read_excel(file_obj, header=None)

    h_idx = -1
    col_lsc_idx = 1
    for idx, row in df_raw.head(20).iterrows():
        for c_i, v in enumerate(row.values):
            if 'lệnh hãng' in str(v).lower():
                h_idx = idx
                col_lsc_idx = c_i
                break
        if h_idx != -1:
            break

    if h_idx == -1:
        h_idx = 7
        col_lsc_idx = 1

    cyber_series = df_raw.iloc[h_idx + 1:, col_lsc_idx].dropna().astype(str)
    cyber_keys = set()
    for x in cyber_series:
        val_clean = clean_lsc_giu_gach(x)
        if len(val_clean) > 5:
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

def dong_bo_hoa_don(df_target):
    mask_chua_hd = df_target['Số hóa đơn'].isna() | df_target['Số hóa đơn'].astype(str).str.strip().isin(['', 'nan', 'None', '0'])
    df_target.loc[mask_chua_hd, 'Ngày xuất hóa đơn'] = ""
    df_target.loc[mask_chua_hd, 'Giá trị xuất hóa đơn'] = 0
    return df_target

def chuan_hoa_kieu_du_lieu(df_input):
    df_out = df_input.copy()
    for col in TEXT_COLUMNS:
        if col in df_out.columns:
            if col == 'Số lệnh sửa chữa':
                df_out[col] = df_out[col].fillna('').astype(str).apply(clean_lsc_giu_gach)
            else:
                df_out[col] = df_out[col].fillna('').astype(str).replace('nan', '')
    for col in COT_TIEN + ['Giá trị xuất hóa đơn']:
        if col in df_out.columns:
            df_out[col] = pd.to_numeric(df_out[col], errors='coerce').fillna(0)
    return df_out

@st.cache_data(show_spinner=False)
def load_cached_master():
    if not os.path.exists(MASTER_FILE):
        df_empty = pd.DataFrame(columns=TAT_CA_COT)
        df_empty.to_excel(MASTER_FILE, index=False)
    df_m = pd.read_excel(MASTER_FILE, dtype={'Số lệnh sửa chữa': str, 'Số hóa đơn': str, 'Ngày xuất hóa đơn': str, 'Xe GSM': str})
    df_m.columns = [str(c).strip() for c in df_m.columns]
    
    cols_to_drop = [c for c in ['Tổng bằng chữ', 'Phân loại bảo hành', 'Loại BH'] if c in df_m.columns]
    if cols_to_drop:
        df_m = df_m.drop(columns=cols_to_drop)

    for c in TAT_CA_COT:
        if c not in df_m.columns:
            df_m[c] = None

    for c in COT_TIEN + ['Giá trị xuất hóa đơn']:
        if c in df_m.columns:
            df_m[c] = pd.to_numeric(clean_tien_series(df_m[c]), errors='coerce').fillna(0)
            
    if 'Phân loại KH' in df_m.columns:
        df_m['Phân loại KH'] = df_m['Phân loại KH'].fillna("KH Thông Thường")
        df_m.loc[df_m['Phân loại KH'] != 'GSM Công nợ', 'Phân loại KH'] = "KH Thông Thường"

    df_m = dong_bo_hoa_don(df_m)
    df_m = chuan_hoa_kieu_du_lieu(df_m)
    return df_m[TAT_CA_COT]

def save_master(df_to_save):
    df_clean = dong_bo_hoa_don(df_to_save.copy())
    df_clean = chuan_hoa_kieu_du_lieu(df_clean)
    df_clean.to_excel(MASTER_FILE, index=False)
    load_cached_master.clear()

@st.cache_data(show_spinner=False)
def load_cached_invoices():
    if os.path.exists(INVOICE_DB_FILE):
        try:
            return pd.read_excel(INVOICE_DB_FILE)
        except Exception:
            return pd.DataFrame()
    return pd.DataFrame()

def save_invoice_db(df_inv_clean):
    df_inv_clean.to_excel(INVOICE_DB_FILE, index=False)
    load_cached_invoices.clear()

@st.cache_data(show_spinner=False)
def load_cached_511():
    if os.path.exists(ACC_511_FILE):
        try:
            return pd.read_excel(ACC_511_FILE)
        except Exception:
            return pd.DataFrame()
    return pd.DataFrame()

def save_511_db(df_511):
    df_511.to_excel(ACC_511_FILE, index=False)
    load_cached_511.clear()

def tao_file_mau_gsm():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Mau_Danh_Sach_No_GSM"
    headers = ["STT", "Số lệnh sửa chữa (Bắt buộc)", "Biển số xe (Tùy chọn)", "Ghi chú công nợ (Tùy chọn)"]
    ws.row_dimensions[1].height = 26
    font_h = Font(name='Segoe UI', size=11, bold=True, color='FFFFFF')
    fill_h = PatternFill(start_color='1F4E78', fill_type='solid')
    border_t = Side(border_style='thin', color='D3D3D3')
    border_all = Border(left=border_t, right=border_t, top=border_t, bottom=border_t)
    
    for col_idx, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col_idx, value=h)
        cell.font = font_h
        cell.fill = fill_h
        cell.alignment = Alignment(horizontal='center', vertical='center')
        cell.border = border_all
        
    sample_data = [
        [1, "C23401-WO-26-08-24-027", "81G00056", "Nợ GSM đợt 1"],
        [2, "C23401-WO-26-08-30-001", "81G00056", "Nợ GSM đợt 1"],
        [3, "C23401-WO-26-08-31-001", "81H05070", "Chờ GSM thanh toán"]
    ]
    for r_idx, row in enumerate(sample_data, 2):
        ws.row_dimensions[r_idx].height = 20
        for c_idx, val in enumerate(row, 1):
            cell = ws.cell(row=r_idx, column=c_idx, value=val)
            cell.font = Font(name='Segoe UI', size=10)
            cell.border = border_all
            cell.alignment = Alignment(horizontal='center' if c_idx <= 3 else 'left', vertical='center')
            
    ws.column_dimensions['A'].width = 8
    ws.column_dimensions['B'].width = 30
    ws.column_dimensions['C'].width = 20
    ws.column_dimensions['D'].width = 35
    
    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return out

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

df_master = load_cached_master()
df_invoice_db = load_cached_invoices()
df_511_db = load_cached_511()

if st.session_state.logged_in:
    tabs = st.tabs([
        "📈 0. Dashboard Doanh Thu & Đối Soát TK 511",
        "📊 1. Bảng Tính Web & Phân Luồng",
        "📥 2. Nạp Dữ Liệu DMS Mới",
        "🧾 3. Khớp File Hóa Đơn Kế Toán",
        "🛡️ 4. Import Phê Duyệt Bảo Hành (Excel)",
        "🚕 5. Quản Lý Công Nợ GSM (Import & Nhập Tay)",
        "🔍 6. Đối Soát Upload Lên Cyber"
    ])
    tab_dash = tabs[0]
    tab_work = tabs[1]
    tab_import = tabs[2]
    tab_inv = tabs[3]
    tab_bh_import = tabs[4]
    tab_gsm_import = tabs[5]
    tab_cyber = tabs[6]
else:
    tabs = st.tabs(["📈 0. Dashboard Doanh Thu & Đối Soát TK 511", "📊 1. Bảng Tính Tra Cứu & Báo Cáo"])
    tab_dash = tabs[0]
    tab_work = tabs[1]

# TAB 0: DASHBOARD LẤY DỮ LIỆU TỪ BẢNG KÊ (CHỈ LẤY 01.S, LOẠI BỎ 01.B) & ĐỐI SOÁT TK 511 (5112 + 5114)
with tab_dash:
    st.subheader("📈 Phân Tích Doanh Thu Xuất HĐ (Chỉ tính 01.S) & Đối Soát TK 511 (5112 + 5114)")
    st.caption("Tự động loại bỏ mã 01.B (Phụ kiện bán lẻ), cộng lũy kế từng ngày và double-check với Sổ cái 511.")

    col_up_hd, col_up_511 = st.columns(2)
    with col_up_hd:
        with st.expander("📂 1. Nạp Bảng Kê Hóa Đơn (Tự động lọc 01.S)", expanded=(len(df_invoice_db) == 0)):
            up_inv_dash = st.file_uploader("Tải file Bảng Kê Hóa Đơn (Excel/CSV)", type=['xlsx', 'xls', 'csv'], key='up_inv_dash')
            if up_inv_dash:
                df_raw_dash, col_labels_dash = doc_file_hoa_don_chuan(up_inv_dash)
                def find_idx_dash(kw_list, default=0):
                    for i, label in enumerate(col_labels_dash):
                        if any(k in label.lower() for k in kw_list):
                            return i
                    return default

                d_lsc_idx = st.selectbox("Cột Số LSC / RO:", list(range(len(col_labels_dash))), format_func=lambda i: col_labels_dash[i], index=find_idx_dash(['ro hãng', 'ro_hãng'], 0), key='d_lsc')
                d_shd_idx = st.selectbox("Cột Số Hóa Đơn:", list(range(len(col_labels_dash))), format_func=lambda i: col_labels_dash[i], index=find_idx_dash(['hóa đơn đt', 'số hóa đơn'], 1 if len(col_labels_dash)>1 else 0), key='d_shd')
                d_nhd_idx = st.selectbox("Cột Ngày HĐ:", list(range(len(col_labels_dash))), format_func=lambda i: col_labels_dash[i], index=find_idx_dash(['ngày', 'chứng từ'], 0), key='d_nhd')
                d_gt_idx  = st.selectbox("Cột Tiền Thanh Toán:", list(range(len(col_labels_dash))), format_func=lambda i: col_labels_dash[i], index=find_idx_dash(['tổng thanh toán', 'thành tiền', 'tiền'], 0), key='d_gt')

                if st.button("🚀 NẠP HÓA ĐƠN VÀO DASHBOARD", type="primary", use_container_width=True):
                    clean_rows = []
                    b_count = 0
                    for _, r in df_raw_dash.iterrows():
                        lsc = clean_lsc_giu_gach(r[d_lsc_idx]) if pd.notna(r[d_lsc_idx]) else ""
                        
                        # BỎ ĐẦU 01.B (PHỤ KIỆN BÁN LẺ), CHỈ TÍNH DOANH THU DỊCH VỤ SỬA CHỮA (01.S)
                        if str(lsc).strip().upper().startswith('01.B'):
                            b_count += 1
                            continue

                        shd = str(r[d_shd_idx]).strip() if pd.notna(r[d_shd_idx]) else ""
                        if shd.endswith('.0'): shd = shd[:-2]
                        nhd = clean_ngay_chuan(r[d_nhd_idx]) if pd.notna(r[d_nhd_idx]) else ""
                        raw_gt = str(r[d_gt_idx]).replace(',', '').replace(' ', '') if pd.notna(r[d_gt_idx]) else "0"
                        try:
                            gt = float(raw_gt)
                        except ValueError:
                            gt = 0.0

                        if shd or gt > 0 or lsc:
                            clean_rows.append({
                                'Số lệnh sửa chữa': lsc,
                                'Số hóa đơn': shd,
                                'Ngày xuất hóa đơn': nhd,
                                'Giá trị xuất hóa đơn': gt
                            })
                    df_new_inv = pd.DataFrame(clean_rows)
                    save_invoice_db(df_new_inv)
                    st.success(f"✅ Đã nạp thành công **{len(df_new_inv)}** hóa đơn dịch vụ (Đã tự động loại bỏ **{b_count}** hóa đơn phụ kiện 01.B)!")
                    st.rerun()

    with col_up_511:
        with st.expander("📂 2. Nạp Sổ Cái Chi Tiết TK 511 (5112 & 5114)", expanded=(len(df_511_db) == 0)):
            st.caption("Tải file Sổ chi tiết tài khoản (Excel) - Hệ thống tự động nhận diện cột Ngày, Diễn giải, Phát sinh Có và Số R/O.")
            up_511_file = st.file_uploader("Tải file Sổ Chi Tiết TK 511 (Excel/CSV)", type=['xlsx', 'xls', 'csv'], key='up_511_key')
            if up_511_file:
                df_raw_511, col_names_511 = doc_file_so_511(up_511_file)
                st.write("🔎 Xem trước bảng TK 511 vừa nạp:")
                st.dataframe(df_raw_511.head(3), use_container_width=True)

                def tim_index_511(kws, def_idx=0):
                    for i, c in enumerate(col_names_511):
                        if any(k in c.lower() for k in kws):
                            return i
                    return def_idx

                idx_ngay = tim_index_511(['ngày c.từ', 'ngày ct', 'ngày c.t'], 4)
                idx_shd  = tim_index_511(['số hóa đơn', 'số hđ'], 5)
                idx_dg   = tim_index_511(['diễn giải'], 11)
                idx_co   = tim_index_511(['có', 'phát sinh có', 'ps có'], 15)
                idx_ro   = tim_index_511(['số r/o', 'ro', 'lệnh'], 17)

                st.info(f"💡 Tự động nhận diện: Cột Ngày (`{col_names_511[idx_ngay]}`), Tiền Có (`{col_names_511[idx_co]}`), Diễn giải (`{col_names_511[idx_dg]}`), Số R/O (`{col_names_511[idx_ro]}`)")

                if st.button("🚀 XÁC NHẬN NẠP DỮ LIỆU TK 511", type="primary", use_container_width=True):
                    rows_511 = []
                    b_count_511 = 0
                    for _, r in df_raw_511.iterrows():
                        ro_val = str(r.iloc[idx_ro]).strip() if idx_ro < len(r) and pd.notna(r.iloc[idx_ro]) else ""
                        
                        # BỎ CÁC DÒNG RO ĐẦU 01.B (PHỤ KIỆN BÁN LẺ)
                        if ro_val.upper().startswith('01.B'):
                            b_count_511 += 1
                            continue

                        ngay_str = clean_ngay_chuan(r.iloc[idx_ngay]) if idx_ngay < len(r) and pd.notna(r.iloc[idx_ngay]) else ""
                        dg_val = str(r.iloc[idx_dg]).strip() if idx_dg < len(r) and pd.notna(r.iloc[idx_dg]) else ""
                        
                        raw_tien = str(r.iloc[idx_co]).replace(',', '').replace(' ', '') if idx_co < len(r) and pd.notna(r.iloc[idx_co]) else "0"
                        try:
                            tien = float(raw_tien)
                        except ValueError:
                            tien = 0.0

                        # Nhận diện: Nếu diễn giải chứa "bảo hành" -> TK 5114; còn lại -> TK 5112
                        loai_tk = "5114" if any(w in dg_val.lower() for w in ['bảo hành', 'bh', 'đình vũ']) else "5112"

                        if ngay_str and tien != 0 and 'số dư' not in dg_val.lower():
                            rows_511.append({
                                'Ngày': ngay_str,
                                'Số R/O': ro_val,
                                'Loại TK': loai_tk,
                                'Số tiền': tien,
                                'Diễn giải': dg_val
                            })

                    df_new_511 = pd.DataFrame(rows_511)
                    save_511_db(df_new_511)
                    st.success(f"✅ Đã nạp thành công **{len(df_new_511)}** dòng doanh thu vào TK 511 (Đã loại bỏ **{b_count_511}** dòng phụ kiện 01.B)!")
                    st.rerun()

    # XỬ LÝ DASHBOARD HIỂN THỊ THÁNG HIỆN TẠI VÀ DOUBLE CHECK
    if len(df_invoice_db) > 0:
        df_dash_data = df_invoice_db.copy()
        df_dash_data['dt_parsed'] = pd.to_datetime(df_dash_data['Ngày xuất hóa đơn'], format='%d/%m/%Y', errors='coerce')
        df_dash_data = df_dash_data[df_dash_data['dt_parsed'].notna()]

        if len(df_dash_data) > 0:
            df_dash_data['nam'] = df_dash_data['dt_parsed'].dt.year
            df_dash_data['thang'] = df_dash_data['dt_parsed'].dt.month

            # Chuẩn bị dữ liệu TK 511 nếu có
            df_511_clean = pd.DataFrame()
            if len(df_511_db) > 0:
                df_511_clean = df_511_db.copy()
                df_511_clean['dt_parsed'] = pd.to_datetime(df_511_clean['Ngày'], format='%d/%m/%Y', errors='coerce')
                df_511_clean = df_511_clean[df_511_clean['dt_parsed'].notna()]
                df_511_clean['nam'] = df_511_clean['dt_parsed'].dt.year
                df_511_clean['thang'] = df_511_clean['dt_parsed'].dt.month

            # Bộ lọc Năm / Tháng - MẶC ĐỊNH LÀ THÁNG HIỆN TẠI (Tháng 09/2026)
            col_sel1, col_sel2 = st.columns([2, 4])
            with col_sel1:
                years_avail = sorted(df_dash_data['nam'].unique(), reverse=True)
                sel_year = st.selectbox("📅 Chọn Năm:", years_avail, index=0, key='dash_yr')

                months_avail = sorted(df_dash_data[df_dash_data['nam'] == sel_year]['thang'].unique())
                curr_m = datetime.now().month
                default_m_idx = months_avail.index(curr_m) if curr_m in months_avail else (months_avail.index(9) if 9 in months_avail else len(months_avail)-1)
                sel_month = st.selectbox("📆 Chọn Tháng Báo Cáo:", months_avail, index=default_m_idx, format_func=lambda m: f"Tháng {m:02d}", key='dash_mth')

            # Lọc dữ liệu Hóa Đơn trong tháng
            df_month = df_dash_data[(df_dash_data['nam'] == sel_year) & (df_dash_data['thang'] == sel_month)].copy()
            tong_tien_hd = df_month['Giá trị xuất hóa đơn'].sum()
            tong_so_hd = len(df_month)

            # Lọc dữ liệu TK 511 trong tháng
            tien_5112 = 0.0
            tien_5114 = 0.0
            if len(df_511_clean) > 0:
                df_511_m = df_511_clean[(df_511_clean['nam'] == sel_year) & (df_511_clean['thang'] == sel_month)]
                tien_5112 = df_511_m[df_511_m['Loại TK'] == '5112']['Số tiền'].sum()
                tien_5114 = df_511_m[df_511_m['Loại TK'] == '5114']['Số tiền'].sum()
            
            tong_511 = tien_5112 + tien_5114
            chenh_lech = tong_tien_hd - tong_511

            # KHUNG KPI DOUBLE CHECK
            st.markdown("---")
            m1, m2, m3, m4 = st.columns(4)
            m1.metric(f"🧾 Tổng Xuất HĐ Dịch Vụ (01.S)", f"{tong_tien_hd:,.0f} đ", f"{tong_so_hd:,} hóa đơn")
            m2.metric(f"🛡️ Doanh Thu Bảo Hành (TK 5114)", f"{tien_5114:,.0f} đ", "HĐ Dịch Vụ / Hãng")
            m3.metric(f"💰 Tổng TK 511 (5112 + 5114)", f"{tong_511:,.0f} đ", f"5112: {tien_5112:,.0f} đ")
            
            if tong_511 > 0:
                if abs(chenh_lech) < 1000:
                    m4.metric("🔍 Double Check (HĐ vs 511)", "✅ KHỚP 100%", "0 đ chênh lệch", delta_color="normal")
                else:
                    m4.metric("🔍 Double Check (HĐ vs 511)", f"{abs(chenh_lech):,.0f} đ", f"{'Thừa HĐ' if chenh_lech > 0 else 'Thiếu HĐ'}", delta_color="inverse")
            else:
                m4.metric("🔍 Double Check (HĐ vs 511)", "Chưa nạp Sổ 511", "Chờ file đối soát")

            # BẢNG TỔNG HỢP THEO NGÀY CÓ CỘNG LŨY KẾ & DOUBLE CHECK
            df_month['ngay_num'] = df_month['dt_parsed'].dt.day
            df_month['ngay_str'] = df_month['dt_parsed'].dt.strftime('%d/%m')

            df_daily = df_month.groupby(['ngay_num', 'ngay_str']).agg(
                Tong_Tien=('Giá trị xuất hóa đơn', 'sum'),
                So_Luong_HD=('Số hóa đơn', 'count')
            ).reset_index().sort_values('ngay_num')

            # TÍNH CỘNG LŨY KẾ TỪNG NGÀY
            df_daily['Luy_Ke_HD'] = df_daily['Tong_Tien'].cumsum()

            # GHÉP DỮ LIỆU TK 511 THEO TỪNG NGÀY ĐỂ DOUBLE CHECK NẾU CÓ FILE
            if len(df_511_clean) > 0:
                df_511_m['ngay_num'] = df_511_m['dt_parsed'].dt.day
                df_511_daily = df_511_m.groupby('ngay_num')['Số tiền'].sum().reset_index().rename(columns={'Số tiền': 'Tien_511'})
                df_daily = pd.merge(df_daily, df_511_daily, on='ngay_num', how='left').fillna(0)
                df_daily['Luy_Ke_511'] = df_daily['Tien_511'].cumsum()
                df_daily['Chenh_Lech_Ngay'] = df_daily['Tong_Tien'] - df_daily['Tien_511']

            st.markdown(f"#### 📊 Biểu Đồ Doanh Thu Dịch Vụ Theo Từng Ngày (Tháng {sel_month:02d}/{sel_year})")
            chart_view = df_daily.set_index('ngay_str')[['Tong_Tien']]
            chart_view.columns = ['Giá trị HĐ (VNĐ)']
            st.bar_chart(chart_view, height=360, use_container_width=True)

            st.markdown(f"#### 📋 Chi Tiết Bảng Kê Doanh Thu Từng Ngày & Cộng Lũy Kế (Tháng {sel_month:02d}/{sel_year})")
            
            tbl_view = pd.DataFrame()
            tbl_view['Ngày'] = df_daily['ngay_str']
            tbl_view['Số Lượng HĐ'] = df_daily['So_Luong_HD']
            tbl_view['Tổng Tiền Xuất HĐ (VNĐ)'] = df_daily['Tong_Tien'].apply(lambda x: f"{x:,.0f} đ")
            tbl_view['Cộng Lũy Kế Xuất HĐ (VNĐ)'] = df_daily['Luy_Ke_HD'].apply(lambda x: f"{x:,.0f} đ")
            
            if 'Tien_511' in df_daily.columns:
                tbl_view['Sổ 511 (5112+5114)'] = df_daily['Tien_511'].apply(lambda x: f"{x:,.0f} đ")
                tbl_view['Lũy Kế TK 511'] = df_daily['Luy_Ke_511'].apply(lambda x: f"{x:,.0f} đ")
                tbl_view['Trạng Thái Double Check'] = df_daily['Chenh_Lech_Ngay'].apply(lambda x: "✅ Khớp" if abs(x) < 1000 else f"Lệch: {x:,.0f} đ")

            st.dataframe(tbl_view, use_container_width=True, hide_index=True)

            with st.expander("🔎 Xem danh sách chi tiết tất cả hóa đơn dịch vụ (01.S) trong tháng này"):
                st.dataframe(
                    df_month.sort_values('dt_parsed', ascending=False)[['Số lệnh sửa chữa', 'Số hóa đơn', 'Ngày xuất hóa đơn', 'Giá trị xuất hóa đơn']],
                    use_container_width=True,
                    hide_index=True
                )
        else:
            st.warning("⚠️ Không tìm thấy định dạng ngày hợp lệ trong Bảng Kê Hóa Đơn.")
    else:
        st.info("ℹ️ Hệ thống chưa có dữ liệu Hóa Đơn. Vui lòng mở khung **1. Nạp Bảng Kê Hóa Đơn** ở trên để bắt đầu.")

# TAB 1: BẢNG TÍNH WEB VỚI BỘ LỌC HIDE / UNHIDE CỘT
with tab_work:
    df_hoanthanh = df_master[df_master['Trạng thái'].isin(TRANG_THAI_HOAN_THANH)]
    df_chuahoanthanh = df_master[~df_master['Trạng thái'].isin(TRANG_THAI_HOAN_THANH + ['Đã hủy'])]
    chua_hoan_thanh_cnt = len(df_chuahoanthanh)

    df_kh_total = df_hoanthanh[(df_hoanthanh['KH thanh toán'] > 0) & (df_hoanthanh['Phân loại KH'] != 'GSM Công nợ')]
    kh_da_hd_cnt = df_kh_total[df_kh_total['Số hóa đơn'].notna() & (~df_kh_total['Số hóa đơn'].astype(str).str.strip().isin(['', 'nan', 'None', '0']))].shape[0]
    kh_total_cnt = len(df_kh_total)
    kh_chua_hd_cnt = kh_total_cnt - kh_da_hd_cnt
    kh_chua_hd_amt = df_kh_total[df_kh_total['Số hóa đơn'].isna() | df_kh_total['Số hóa đơn'].astype(str).str.strip().isin(['', 'nan', 'None', '0'])]['KH thanh toán'].sum()

    df_gsm_debt = df_hoanthanh[(df_hoanthanh['KH thanh toán'] > 0) & (df_hoanthanh['Phân loại KH'] == 'GSM Công nợ')]
    gsm_debt_cnt = len(df_gsm_debt)
    gsm_debt_amt = df_gsm_debt['KH thanh toán'].sum()

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
    k2.metric("🚕 GSM Công Nợ (Đã xong)", f"{gsm_debt_cnt:,} lệnh nợ", f"{gsm_debt_amt:,.0f} đ")
    k3.metric("🛡️ Bảo Hành Hãng (W)", f"{bh_total_cnt:,} lệnh", f"Chưa duyệt: {bh_chua_duyet_cnt}/{bh_total_cnt}")
    k4.metric("🏢 Bảo Hiểm (Insurance)", f"{bh_total_all_cnt:,} lệnh", f"Chưa xuất HĐ: {bh_chua_hd_cnt}/{bh_total_all_cnt}")
    k5.metric("⏳ Đang Làm / Báo Giá", f"{chua_hoan_thanh_cnt:,} lệnh", "Chờ hoàn thành")

    if tong_chua_hd_cnt > 0:
        st.error(
            f"🚨 **CẢNH BÁO CHƯA XUẤT HÓA ĐƠN (CHỈ LỆNH ĐÃ HOÀN THÀNH - KH & BẢO HIỂM): {tong_chua_hd_cnt} lệnh | Tổng tiền: {tong_chua_hd_amt:,.0f} đ** "
            f"\n*(Chi tiết: Khách hàng thường: **{kh_chua_hd_cnt}** lệnh ({kh_chua_hd_amt:,.0f} đ) | "
            f"Bảo hiểm: **{bh_chua_hd_cnt}** lệnh ({bh_chua_hd_amt:,.0f} đ). Không tính Báo giá, Đang sửa chữa, Bảo hành, Nội bộ và Nợ GSM)*"
        )

    f_col1, f_col2, f_col3 = st.columns([2, 2, 2])
    with f_col1:
        luong_data = st.selectbox("📂 Chọn luồng dữ liệu xem & quản trị:", [
            "1. KH Thanh Toán (Đã hoàn thành lệnh)",
            "2. GSM Công Nợ (Chỉ các lệnh thuộc file công nợ)",
            "3. Bảo Hành Hãng (W) - Phê duyệt",
            "4. Bảo Hiểm (Insurance)",
            "5. Nội bộ thanh toán",
            "🚨 Lệnh Đã Xong Chưa Xuất HĐ (Chỉ KH & Bảo Hiểm)",
            "6. Xem Lệnh Chưa Hoàn Thành (Báo giá & Đang sửa chữa)",
            "7. Xem Lệnh Đã Hủy",
            "Toàn bộ dữ liệu (Sheet Tổng Hợp)"
        ])
    with f_col2:
        loc_canh_bao_hd = st.selectbox("⚡ Lọc trạng thái HĐ:", ["Tất cả", "Chỉ hiển thị xe CHƯA có hóa đơn", "Đã có hóa đơn"])
    with f_col3:
        tim_kiem_nhanh = st.text_input("🔍 Tìm kiếm (Biển số / LSC / Tên):", "")

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
        df_show = df_kh_total.copy()
        default_cols = cols_base + ['Số tiền thanh toán cuối', 'KH thanh toán'] + cols_hd
        sheet_file_name = "1_KH_Thanh_Toan"
    elif luong_data == "2. GSM Công Nợ (Chỉ các lệnh thuộc file công nợ)":
        df_show = df_gsm_debt.copy()
        default_cols = cols_base + ['Phân loại KH', 'Số tiền thanh toán cuối', 'KH thanh toán'] + cols_hd
        sheet_file_name = "2_GSM_Cong_No"
    elif luong_data == "3. Bảo Hành Hãng (W) - Phê duyệt":
        df_show = df_bh_hang.copy()
        default_cols = cols_base + ['BH hãng thanh toán', 'Phê duyệt bảo hành'] + cols_hd
        sheet_file_name = "3_Bao_Hanh_Hang"
    elif luong_data == "4. Bảo Hiểm (Insurance)":
        df_show = df_bh.copy()
        default_cols = cols_base + ['Số tiền thanh toán cuối', 'BH thanh toán'] + cols_hd
        sheet_file_name = "4_Bao_Hiem"
    elif luong_data == "5. Nội bộ thanh toán":
        df_show = df_hoanthanh[df_hoanthanh['Nội bộ thanh toán'] > 0].copy()
        default_cols = cols_base + ['Số tiền thanh toán cuối', 'Nội bộ thanh toán'] + cols_hd
        sheet_file_name = "5_Noi_Bo"
    elif luong_data == "🚨 Lệnh Đã Xong Chưa Xuất HĐ (Chỉ KH & Bảo Hiểm)":
        mask_target = (df_master['Trạng thái'].isin(TRANG_THAI_HOAN_THANH)) & (
            ((df_master['KH thanh toán'] > 0) & (df_master['Phân loại KH'] != 'GSM Công nợ')) |
            (df_master['BH thanh toán'] > 0)
        ) & (df_master['Số hóa đơn'].isna() | df_master['Số hóa đơn'].astype(str).str.strip().isin(['', 'nan', 'None', '0']))
        df_show = df_master[mask_target].copy()
        default_cols = cols_base + ['Số tiền thanh toán cuối', 'KH thanh toán', 'BH thanh toán'] + cols_hd
        sheet_file_name = "Canh_Bao_Chua_Xuat_HD"
    elif luong_data == "6. Xem Lệnh Chưa Hoàn Thành (Báo giá & Đang sửa chữa)":
        df_show = df_chuahoanthanh.copy()
        default_cols = cols_base + ['Số tiền thanh toán cuối'] + cols_4_thanh_toan
        sheet_file_name = "6_Lenh_Chua_Xong"
    elif luong_data == "7. Xem Lệnh Đã Hủy":
        df_show = df_master[df_master['Trạng thái'] == 'Đã hủy'].copy()
        default_cols = cols_base + ['Số tiền thanh toán cuối']
        sheet_file_name = "7_Lenh_Da_Huy"
    else:
        df_show = df_master.copy()
        default_cols = cols_base + ['Phân loại KH', 'Phê duyệt bảo hành', 'Số tiền thanh toán cuối'] + cols_4_thanh_toan + cols_hd
        sheet_file_name = "Tong_Hop_Toan_Bo"

    with st.expander("👁️ Tùy biến Cột hiển thị (Bấm để Tick chọn Unhide hoặc bỏ chọn để Hide cột)", expanded=False):
        st.caption("Mặc định các cột chi tiết đã được ẨN (Hide) để bảng gọn gàng. Bạn có thể tick thêm bất cứ cột nào bên dưới để MỞ (Unhide) xem lệnh chi tiết:")
        selected_cols = st.multiselect(
            "Chọn các cột bạn muốn xem trên bảng:",
            options=tat_ca_cot_bang,
            default=[c for c in default_cols if c in tat_ca_cot_bang],
            key=f"col_filter_{luong_data}"
        )

    if loc_canh_bao_hd == "Chỉ hiển thị xe CHƯA có hóa đơn":
        df_show = df_show[(df_show['Số hóa đơn'].isna()) | (df_show['Số hóa đơn'].astype(str).str.strip().isin(['', 'nan', 'None', '0']))]
    elif loc_canh_bao_hd == "Đã có hóa đơn":
        df_show = df_show[df_show['Số hóa đơn'].notna() & (~df_show['Số hóa đơn'].astype(str).str.strip().isin(['', 'nan', 'None', '0']))]

    if tim_kiem_nhanh:
        mask = (
            df_show['Biển số'].astype(str).str.contains(tim_kiem_nhanh, case=False, na=False) |
            df_show['Số lệnh sửa chữa'].astype(str).str.contains(tim_kiem_nhanh, case=False, na=False) |
            df_show['Tên khách hàng'].astype(str).str.contains(tim_kiem_nhanh, case=False, na=False)
        )
        df_show = df_show[mask]

    df_show = chuan_hoa_kieu_du_lieu(df_show)
    df_show = df_show.reset_index(drop=True)
    df_show.insert(0, 'STT', range(1, len(df_show) + 1))

    actual_cols = ['STT'] + [c for c in tat_ca_cot_bang if c in selected_cols and c in df_show.columns]
    df_render = df_show[actual_cols].copy()
    cols_to_hide_in_excel = [c for c in tat_ca_cot_bang if c not in selected_cols]

    is_admin = st.session_state.logged_in

    col_cfg = {
        "STT": st.column_config.NumberColumn("STT", disabled=True, pinned=True, width="small"),
        "Số lệnh sửa chữa": st.column_config.TextColumn("Số LSC", disabled=True, pinned=True),
        "Trạng thái": st.column_config.TextColumn("Trạng thái", disabled=True, pinned=True),
        "Cố vấn dịch vụ": st.column_config.TextColumn("Cố vấn dịch vụ", disabled=not is_admin, pinned=True),
        "Tên khách hàng": st.column_config.TextColumn("Tên khách hàng", disabled=not is_admin, pinned=True),
        "Thời gian đóng LSC": st.column_config.TextColumn("Thời gian đóng LSC", disabled=True, pinned=True),
        "Biển số": st.column_config.TextColumn("Biển số", disabled=not is_admin, pinned=True),
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

    if not is_admin:
        st.info("ℹ️ Bạn đang ở chế độ **Chỉ Xem (Read-only)**. Để chỉnh sửa dữ liệu hoặc nạp file, vui lòng đăng nhập quyền Quản trị ở thanh bên trái.")

    edited_df = st.data_editor(
        df_render,
        use_container_width=True,
        height=530,
        column_config=col_cfg,
        disabled=(not is_admin),
        num_rows="fixed",
        hide_index=True,
        key=f"data_editor_table_{luong_data}"
    )

    st.markdown("---")
    c_btn1, c_btn2, c_btn3 = st.columns([3, 3, 4])

    with c_btn1:
        if is_admin:
            if st.button("💾 Lưu Mọi Chỉnh Sửa Trực Tiếp", type="primary", use_container_width=True):
                prog = st.progress(0)
                edit_dict = edited_df.set_index('Số lệnh sửa chữa').to_dict('index')
                total_rows = len(df_master)
                for idx, r_lsc in enumerate(df_master['Số lệnh sửa chữa']):
                    if r_lsc in edit_dict:
                        for c in ['Phê duyệt bảo hành', 'Phân loại KH', 'Số hóa đơn', 'Ngày xuất hóa đơn', 'Giá trị xuất hóa đơn', 'Cố vấn dịch vụ', 'Tên khách hàng', 'Biển số']:
                            if c in edit_dict[r_lsc]:
                                df_master.at[idx, c] = edit_dict[r_lsc][c]
                    if idx % 100 == 0:
                        prog.progress(int((idx / total_rows) * 80))
                        
                save_master(df_master)
                prog.progress(100)
                st.success("✅ Dữ liệu đã lưu thành công vào Hệ Thống.")

    with c_btn2:
        df_export_single = df_show.drop(columns=['STT']) if 'STT' in df_show.columns else df_show
        excel_single_bytes = xuat_excel_don_luong(df_export_single, luong_data, cols_to_hide=cols_to_hide_in_excel)
        st.download_button(
            label=f"⬇️ Tải Riêng Luồng Này (.xlsx)",
            data=excel_single_bytes,
            file_name=f"Bao_Cao_{sheet_file_name}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

    with c_btn3:
        with st.expander("📦 Xuất File Excel Tổng Hợp (Đầy Đủ 8 Sheet)", expanded=False):
            if st.button("🚀 Khởi tạo toàn bộ 8 Sheet Excel", use_container_width=True):
                p_bar_excel = st.progress(0)
                st_txt_excel = st.empty()
                excel_bytes = xuat_excel_da_sheet_with_progress(df_master, p_bar_excel, st_txt_excel)
                st.download_button(
                    label="⬇️ TẢI BÁO CÁO 8 SHEET VỀ MÁY",
                    data=excel_bytes,
                    file_name="Bao_Cao_VinFast_Toan_Bo_8_Sheet.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary",
                    use_container_width=True
                )

# CÁC TAB CHỨC NĂNG KHI ĐÃ ĐĂNG NHẬP
if is_admin:
    # TAB 2: NẠP DỮ LIỆU DMS
    with tab_import:
        st.subheader("Nạp file dữ liệu phân phối VinFast định kỳ")
        up_db = st.file_uploader("Kéo thả file Database mới vào đây", type=['csv', 'xlsx'], key='up_dms')

        if up_db:
            df_raw = doc_file_db(up_db)
            st.write("🔎 Xem trước file vừa tải lên:")
            st.dataframe(df_raw.head(3), use_container_width=True)

            if st.button("🚀 BẮT ĐẦU NẠP VÀ ĐỒNG BỘ DỮ LIỆU", type="primary", use_container_width=True):
                df_inc = pd.DataFrame()
                for c in TAT_CA_COT:
                    df_inc[c] = df_raw[c] if c in df_raw.columns else None

                for c in COT_TIEN:
                    df_inc[c] = pd.to_numeric(clean_tien_series(df_inc[c]), errors='coerce').fillna(0)

                df_inc['Số lệnh sửa chữa'] = df_inc['Số lệnh sửa chữa'].astype(str).apply(clean_lsc_giu_gach)
                df_inc['Phân loại KH'] = "KH Thông Thường"
                df_inc.loc[(df_inc['BH hãng thanh toán'] > 0) & (df_inc['Phê duyệt bảo hành'].isna()), 'Phê duyệt bảo hành'] = "Chờ duyệt"
                df_inc = dong_bo_hoa_don(df_inc)
                df_inc = chuan_hoa_kieu_du_lieu(df_inc)

                p_bar_dms = st.progress(0)
                txt_dms = st.empty()
                txt_dms.write("⏳ Đang đối soát và cập nhật dữ liệu... (10%)")

                master_idx_map = {lsc: idx for idx, lsc in enumerate(df_master['Số lệnh sửa chữa'])}
                new_records = []
                status_updated_cnt = 0
                total_inc = len(df_inc)

                for i, (_, row) in enumerate(df_inc.iterrows()):
                    lsc = row['Số lệnh sửa chữa']
                    if not lsc or lsc in ['nan', 'None', '']:
                        continue
                    if lsc not in master_idx_map:
                        new_records.append(row)
                    else:
                        m_idx = master_idx_map[lsc]
                        old_status = str(df_master.at[m_idx, 'Trạng thái'])
                        new_status = str(row['Trạng thái'])
                        if old_status != new_status:
                            df_master.at[m_idx, 'Trạng thái'] = new_status
                            status_updated_cnt += 1
                        for col in COT_TIEN + ['Thời gian đóng LSC', 'Xe GSM']:
                            df_master.at[m_idx, col] = row[col]
                    
                    if i % 30 == 0:
                        pct = int(10 + (i / total_inc) * 80)
                        p_bar_dms.progress(pct)

                if new_records:
                    df_master = pd.concat([df_master, pd.DataFrame(new_records)], ignore_index=True)

                save_master(df_master)
                p_bar_dms.progress(100)
                txt_dms.empty()
                st.success(f"✅ ĐÃ CHẠY XONG CHU TRÌNH! Thêm **{len(new_records)}** lệnh mới và cập nhật trạng thái cho **{status_updated_cnt}** lệnh.")

    # TAB 3: KHỚP HÓA ĐƠN
    with tab_inv:
        st.subheader("Khớp file Hóa Đơn kế toán với Hệ Thống")
        up_inv = st.file_uploader("Tải file Bảng Kê Hóa Đơn", type=['csv', 'xlsx'], key='up_inv_tab')

        if up_inv:
            df_inv, col_labels = doc_file_hoa_don_chuan(up_inv)
            
            def find_idx(kw_list, default=0):
                for i, label in enumerate(col_labels):
                    if any(k in label.lower() for k in kw_list):
                        return i
                return default

            idx_lsc = find_idx(['ro hãng', 'ro_hãng'], 0)
            idx_shd = find_idx(['hóa đơn đt', 'số hóa đơn đt', 'số hóa đơn'], 1 if len(col_labels) > 1 else 0)
            idx_nhd = find_idx(['ngày', 'chứng từ - ngày'], 0)
            idx_gt  = find_idx(['tổng thanh toán', 'thanh toán', 'thành tiền'], 0)

            c1, c2, c3, c4 = st.columns(4)
            opt_indices = list(range(len(col_labels)))
            sel_lsc_idx = c1.selectbox("📌 Cột Số Lệnh SC:", opt_indices, format_func=lambda i: col_labels[i], index=idx_lsc)
            sel_shd_idx = c2.selectbox("🧾 Cột Số Hóa Đơn:", opt_indices, format_func=lambda i: col_labels[i], index=idx_shd)
            
            opt_with_none = [-1] + opt_indices
            sel_nhd_idx = c3.selectbox("📅 Cột Ngày Hóa Đơn:", opt_with_none, format_func=lambda i: "Bỏ qua" if i == -1 else col_labels[i], index=idx_nhd + 1 if idx_nhd != 0 else 0)
            sel_gt_idx  = c4.selectbox("💰 Cột Giá Trị Hóa Đơn:", opt_with_none, format_func=lambda i: "Bỏ qua" if i == -1 else col_labels[i], index=idx_gt + 1 if idx_gt != 0 else 0)

            if st.button("🚀 BẮT ĐẦU KHỚP HÓA ĐƠN", type="primary", use_container_width=True):
                p_bar_inv = st.progress(0)
                txt_inv = st.empty()
                txt_inv.write("⏳ Đang tổng hợp và gom các hóa đơn theo LSC... (15%)")
                p_bar_inv.progress(15)

                exact_map = {clean_lsc_giu_gach(lsc): idx for idx, lsc in enumerate(df_master['Số lệnh sửa chữa'])}
                norm_map = {norm_lsc_key(lsc): idx for idx, lsc in enumerate(df_master['Số lệnh sửa chữa'])}
                
                inv_aggregated = {}
                raw_invoices_to_save = []
                total_inv_rows = len(df_inv)

                for r_idx in range(total_inv_rows):
                    raw_lsc = str(df_inv.iloc[r_idx, sel_lsc_idx])
                    key_exact = clean_lsc_giu_gach(raw_lsc)

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

                    if shd_val or gt_val > 0 or key_exact:
                        raw_invoices_to_save.append({
                            'Số lệnh sửa chữa': key_exact,
                            'Số hóa đơn': shd_val,
                            'Ngày xuất hóa đơn': nhd_val,
                            'Giá trị xuất hóa đơn': gt_val
                        })

                    if not key_exact or len(key_exact) < 4:
                        continue

                    if key_exact not in inv_aggregated:
                        inv_aggregated[key_exact] = {
                            'so_hd': [shd_val] if shd_val and shd_val not in ['', 'nan', 'None'] else [],
                            'ngay_hd': [nhd_val] if nhd_val else [],
                            'tong_tien': gt_val
                        }
                    else:
                        if shd_val and shd_val not in ['', 'nan', 'None'] and shd_val not in inv_aggregated[key_exact]['so_hd']:
                            inv_aggregated[key_exact]['so_hd'].append(shd_val)
                        if nhd_val and nhd_val not in inv_aggregated[key_exact]['ngay_hd']:
                            inv_aggregated[key_exact]['ngay_hd'].append(nhd_val)
                        inv_aggregated[key_exact]['tong_tien'] += gt_val

                if raw_invoices_to_save:
                    save_invoice_db(pd.DataFrame(raw_invoices_to_save))

                matched_records = []
                total_keys = len(inv_aggregated)

                for k_i, (k_lsc, val_dict) in enumerate(inv_aggregated.items()):
                    m_idx = exact_map.get(k_lsc)
                    if m_idx is None:
                        m_idx = norm_map.get(norm_lsc_key(k_lsc))

                    if m_idx is not None:
                        if str(df_master.at[m_idx, 'Trạng thái']) == 'Đã hủy':
                            continue

                        shd_str = ", ".join(val_dict['so_hd'])
                        nhd_str = ", ".join(val_dict['ngay_hd'])
                        gt_tong = val_dict['tong_tien']

                        df_master.at[m_idx, 'Số hóa đơn'] = shd_str
                        df_master.at[m_idx, 'Ngày xuất hóa đơn'] = nhd_str
                        df_master.at[m_idx, 'Giá trị xuất hóa đơn'] = gt_tong

                        matched_records.append({
                            'Số LSC': k_lsc,
                            'Số Hóa Đơn': shd_str,
                            'Ngày HĐ': nhd_str,
                            'Giá Trị HĐ': f"{gt_tong:,.0f}" if gt_tong else "0"
                        })

                    if k_i % 30 == 0:
                        pct = int(15 + (k_i / max(total_keys, 1)) * 75)
                        p_bar_inv.progress(pct)

                save_master(df_master)
                p_bar_inv.progress(100)
                txt_inv.empty()
                st.success(f"✅ ĐÃ CHẠY XONG CHU TRÌNH! Khớp và gộp thành công **{len(matched_records)}** lệnh sửa chữa và tự động đồng bộ sang Dashboard!")

    # TAB 4: IMPORT BẢO HÀNH
    with tab_bh_import:
        st.subheader("🛡️ Import Danh Sách Phê Duyệt Bảo Hành Hãng Tự Động")
        up_bh_file = st.file_uploader("Tải file duyệt bảo hành (XLSX / CSV)", type=['csv', 'xlsx'], key='up_bh_file')
        
        if up_bh_file:
            df_bh_input = doc_file_db(up_bh_file)
            st.write("🔎 Xem trước file bảo hành:")
            st.dataframe(df_bh_input.head(3), use_container_width=True)
            
            bh_cols = list(df_bh_input.columns)
            b1, b2, b3, b4 = st.columns(4)
            
            idx_bh_lsc = next((i for i, c in enumerate(bh_cols) if any(k in c.lower() for k in ['lệnh', 'lsc', 'ro'])), 0)
            idx_bh_tt = next((i for i, c in enumerate(bh_cols) if any(k in c.lower() for k in ['duyệt', 'trạng thái', 'status'])), 0)
            idx_bh_shd = next((i for i, c in enumerate(bh_cols) if any(k in c.lower() for k in ['hóa đơn', 'số hđ'])), 0)
            idx_bh_nhd = next((i for i, c in enumerate(bh_cols) if any(k in c.lower() for k in ['ngày'])), 0)

            s_bh_lsc = b1.selectbox("Cột Số LSC:", bh_cols, index=idx_bh_lsc)
            s_bh_tt  = b2.selectbox("Cột Trạng Thái Duyệt (Đã duyệt/Từ chối...):", bh_cols, index=idx_bh_tt)
            s_bh_shd = b3.selectbox("Cột Số Hóa Đơn (nếu có):", [None] + bh_cols, index=idx_bh_shd + 1 if idx_bh_shd != 0 else 0)
            s_bh_nhd = b4.selectbox("Cột Ngày Hóa Đơn (nếu có):", [None] + bh_cols, index=idx_bh_nhd + 1 if idx_bh_nhd != 0 else 0)

            if st.button("🚀 BẮT ĐẦU CẬP NHẬT PHÊ DUYỆT BẢO HÀNH", type="primary", use_container_width=True):
                p_bar_bh = st.progress(0)
                txt_bh = st.empty()
                txt_bh.write("⏳ Đang đối soát danh sách bảo hành...")

                exact_map = {clean_lsc_giu_gach(lsc): idx for idx, lsc in enumerate(df_master['Số lệnh sửa chữa'])}
                norm_map = {norm_lsc_key(lsc): idx for idx, lsc in enumerate(df_master['Số lệnh sửa chữa'])}
                
                bh_updated = 0
                total_bh_rows = len(df_bh_input)

                for r_i, (_, r) in enumerate(df_bh_input.iterrows()):
                    key_exact = clean_lsc_giu_gach(r[s_bh_lsc])
                    if not key_exact:
                        continue
                    m_idx = exact_map.get(key_exact)
                    if m_idx is None:
                        m_idx = norm_map.get(norm_lsc_key(key_exact))

                    if m_idx is not None:
                        val_tt = str(r[s_bh_tt]).strip()
                        if any(w in val_tt.lower() for w in ['đã duyệt', 'duyệt', 'approved', 'pass', 'ok']):
                            norm_tt = "Đã duyệt"
                        elif any(w in val_tt.lower() for w in ['từ chối', 'reject', 'không duyệt']):
                            norm_tt = "Từ chối"
                        else:
                            norm_tt = "Chờ duyệt"
                        
                        df_master.at[m_idx, 'Phê duyệt bảo hành'] = norm_tt
                        
                        if s_bh_shd and pd.notna(r[s_bh_shd]):
                            df_master.at[m_idx, 'Số hóa đơn'] = str(r[s_bh_shd]).strip()
                        if s_bh_nhd and pd.notna(r[s_bh_nhd]):
                            df_master.at[m_idx, 'Ngày xuất hóa đơn'] = clean_ngay_chuan(r[s_bh_nhd])
                        bh_updated += 1
                        
                    if r_i % 30 == 0:
                        pct = int((r_i / total_bh_rows) * 90)
                        p_bar_bh.progress(pct)

                save_master(df_master)
                p_bar_bh.progress(100)
                txt_bh.empty()
                st.success(f"✅ ĐÃ CHẠY XONG CHU TRÌNH! Đã cập nhật phê duyệt cho **{bh_updated}** lệnh bảo hành.")

    # TAB 5: QUẢN LÝ CÔNG NỢ GSM
    with tab_gsm_import:
        st.subheader("🚕 Quản Lý Danh Sách Công Nợ GSM (Linh Hoạt Cả 2 Cách)")
        
        col_t1, col_t2 = st.columns([3, 2])
        with col_t1:
            st.info("💡 **Ghi chú:** Các lệnh trong danh sách này sẽ được chuyển sang 'GSM Công nợ' và KHÔNG bị tính vào cảnh báo nợ HĐ của khách hàng lẻ.")
        with col_t2:
            file_mau_bytes = tao_file_mau_gsm()
            st.download_button(
                label="📄 Tải Form File Mẫu Excel GSM Chuẩn (.xlsx)",
                data=file_mau_bytes,
                file_name="Mau_Import_Cong_No_GSM.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )

        st.markdown("---")
        c_tab_file, c_tab_tay = st.tabs(["📂 Cách 1: Nạp File Excel GSM", "✍️ Cách 2: Nhập Tay Lệnh Nợ GSM Trực Tiếp"])

        with c_tab_file:
            up_gsm_file = st.file_uploader("Tải file danh sách lệnh nợ GSM", type=['csv', 'xlsx'], key='up_gsm_file')
            if up_gsm_file:
                df_gsm_in = doc_file_db(up_gsm_file)
                st.write("🔎 Xem trước danh sách trong file:")
                st.dataframe(df_gsm_in.head(3), use_container_width=True)
                
                gsm_cols = list(df_gsm_in.columns)
                idx_gsm_lsc = next((i for i, c in enumerate(gsm_cols) if any(k in c.lower() for k in ['lệnh', 'lsc', 'ro', 'biển'])), 0)
                s_gsm_lsc = st.selectbox("Chọn cột chứa Số LSC (hoặc Biển số) nợ của GSM:", gsm_cols, index=idx_gsm_lsc, key='sel_gsm_col')

                if st.button("🚀 XÁC NHẬN NẠP VÀO CÔNG NỢ GSM", type="primary", use_container_width=True):
                    p_bar_gsm = st.progress(0)
                    exact_map = {clean_lsc_giu_gach(lsc): idx for idx, lsc in enumerate(df_master['Số lệnh sửa chữa'])}
                    bs_map = {norm_lsc_key(bs): idx for idx, bs in enumerate(df_master['Biển số'])}
                    
                    gsm_cnt = 0
                    total_gsm_in = len(df_gsm_in)
                    for r_i, (_, r) in enumerate(df_gsm_in.iterrows()):
                        raw_v = str(r[s_gsm_lsc])
                        k_exact = clean_lsc_giu_gach(raw_v)
                        if not k_exact:
                            continue
                        m_idx = exact_map.get(k_exact)
                        if m_idx is None:
                            m_idx = bs_map.get(norm_lsc_key(k_exact))
                        
                        if m_idx is not None:
                            df_master.at[m_idx, 'Phân loại KH'] = "GSM Công nợ"
                            gsm_cnt += 1
                        if r_i % 20 == 0:
                            p_bar_gsm.progress(int((r_i / total_gsm_in) * 90))

                    save_master(df_master)
                    p_bar_gsm.progress(100)
                    st.success(f"✅ ĐÃ CHẠY XONG CHU TRÌNH! Đã chuyển **{gsm_cnt}** lệnh vào danh sách GSM Công Nợ thành công.")

        with c_tab_tay:
            st.write("##### Nhập nhanh danh sách Số LSC (hoặc Biển số) cần chuyển sang công nợ GSM:")
            nhap_tay_txt = st.text_area("Danh sách mã lệnh / biển số:", height=130, placeholder="C23401-WO-26-08-24-027\nC23401-WO-26-08-30-001\n81G00056")
            
            t_col1, t_col2 = st.columns(2)
            with t_col1:
                if st.button("➕ Thêm Các Lệnh Này Vào Công NỢ GSM", type="primary", use_container_width=True):
                    danh_sach_nhap = [clean_lsc_giu_gach(x) for x in nhap_tay_txt.split('\n') if clean_lsc_giu_gach(x)]
                    if danh_sach_nhap:
                        exact_map = {clean_lsc_giu_gach(lsc): idx for idx, lsc in enumerate(df_master['Số lệnh sửa chữa'])}
                        bs_map = {norm_lsc_key(bs): idx for idx, bs in enumerate(df_master['Biển số'])}
                        
                        added_cnt = 0
                        for k in danh_sach_nhap:
                            m_idx = exact_map.get(k)
                            if m_idx is None:
                                m_idx = bs_map.get(norm_lsc_key(k))
                            if m_idx is not None:
                                df_master.at[m_idx, 'Phân loại KH'] = "GSM Công nợ"
                                added_cnt += 1
                                
                        save_master(df_master)
                        st.success(f"✅ ĐÃ CHẠY XONG CHU TRÌNH! Đã thêm **{added_cnt}** lệnh vào diện công nợ GSM.")
                    else:
                        st.warning("Vui lòng nhập ít nhất 1 mã lệnh hoặc biển số!")

            with t_col2:
                if st.button("🔄 Đưa Các Lệnh Này Quay Lại KH Thông Thường", use_container_width=True):
                    danh_sach_nhap = [clean_lsc_giu_gach(x) for x in nhap_tay_txt.split('\n') if clean_lsc_giu_gach(x)]
                    if danh_sach_nhap:
                        exact_map = {clean_lsc_giu_gach(lsc): idx for idx, lsc in enumerate(df_master['Số lệnh sửa chữa'])}
                        bs_map = {norm_lsc_key(bs): idx for idx, bs in enumerate(df_master['Biển số'])}
                        
                        removed_cnt = 0
                        for k in danh_sach_nhap:
                            m_idx = exact_map.get(k)
                            if m_idx is None:
                                m_idx = bs_map.get(norm_lsc_key(k))
                            if m_idx is not None:
                                df_master.at[m_idx, 'Phân loại KH'] = "KH Thông Thường"
                                removed_cnt += 1
                                
                        save_master(df_master)
                        st.info(f"✅ ĐÃ CHẠY XONG CHU TRÌNH! Đã chuyển **{removed_cnt}** lệnh quay lại KH Thanh Toán thông thường.")

        st.markdown("---")
        st.write("##### 📋 Danh sách các lệnh đang nằm trong Công Nợ GSM hiện tại:")
        df_gsm_hien_tai = df_master[df_master['Phân loại KH'] == 'GSM Công nợ'][['Số lệnh sửa chữa', 'Trạng thái', 'Biển số', 'Tên khách hàng', 'KH thanh toán', 'Số hóa đơn', 'Ngày xuất hóa đơn', 'Phân loại KH']]
        if len(df_gsm_hien_tai) > 0:
            st.dataframe(df_gsm_hien_tai, use_container_width=True)
        else:
            st.info("Chưa có lệnh nào được đánh dấu là nợ GSM.")

    # TAB 6: ĐỐI SOÁT CYBER
    with tab_cyber:
        st.subheader("🔍 Đối Soát Lệnh Đã Hoàn Thành Chưa Up Lên Phần Mềm Cyber")
        st.caption("Chỉ xét các lệnh 'Đã đóng' hoặc 'Sẵn sàng bàn giao'. Form mã lệnh giữ nguyên dấu gạch ngang của DMS cũ.")
        
        up_cyber_file = st.file_uploader("Tải lên file BẢNG TỔNG HỢP LỆNH SỬA CHỮA từ Cyber (Excel)", type=['xlsx', 'xls', 'csv'], key='up_cyber_file')
        
        if up_cyber_file:
            p_bar_cy = st.progress(0)
            txt_cy = st.empty()
            txt_cy.write("⏳ Đang quét dữ liệu Cyber... (20%)")
            p_bar_cy.progress(20)

            cyber_keys = doc_file_cyber(up_cyber_file)
            txt_cy.write("⏳ Đang đối chiếu với các lệnh đã hoàn thành trên DMS... (60%)")
            p_bar_cy.progress(60)

            df_comp = df_master[df_master['Trạng thái'].isin(TRANG_THAI_HOAN_THANH)].copy()
            df_comp['da_up_cyber'] = df_comp['Số lệnh sửa chữa'].apply(
                lambda x: clean_lsc_giu_gach(x) in cyber_keys or norm_lsc_key(x) in cyber_keys
            )
            
            df_chua_up = df_comp[~df_comp['da_up_cyber']].copy()
            df_da_up = df_comp[df_comp['da_up_cyber']].copy()
            p_bar_cy.progress(100)
            txt_cy.empty()
            st.success("✅ ĐÃ CHẠY XONG CHU TRÌNH ĐỐI SOÁT!")
            
            c1, c2, c3 = st.columns(3)
            c1.metric("📌 Tổng Lệnh Đã Hoàn Thành (DMS)", f"{len(df_comp):,} lệnh")
            c2.metric(" Đã Up Lên Cyber", f"{len(df_da_up):,} lệnh")
            c3.metric("🚨 CHƯA UP LÊN CYBER", f"{len(df_chua_up):,} lệnh", delta=f"-{len(df_chua_up)} lệnh", delta_color="inverse")
            
            if len(df_chua_up) > 0:
                st.error(f"⚠️ Phát hiện **{len(df_chua_up)}** lệnh sửa chữa đã xong nhưng CHƯA ĐƯỢC UP LÊN CYBER để xuất hóa đơn!")
                cols_display = ['Số lệnh sửa chữa', 'Trạng thái', 'Biển số', 'Cố vấn dịch vụ', 'Tên khách hàng', 'Số tiền thanh toán cuối', 'KH thanh toán', 'BH thanh toán', 'BH hãng thanh toán']
                cols_valid = [c for c in cols_display if c in df_chua_up.columns]
                
                st.dataframe(df_chua_up[cols_valid], use_container_width=True)
                
                buffer_chua_up = io.BytesIO()
                with pd.ExcelWriter(buffer_chua_up, engine='openpyxl') as wr:
                    df_chua_up[cols_valid].to_excel(wr, index=False, sheet_name='Chua_Up_Cyber')
                buffer_chua_up.seek(0)
                
                st.download_button(
                    label="📥 Tải Danh Sách Các Lệnh Chưa Up Lên Cyber (.xlsx)",
                    data=buffer_chua_up,
                    file_name="Danh_Sach_LSC_Chua_Up_Cyber.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary"
                )
            else:
                st.success("✅ ĐÃ CHẠY XONG CHU TRÌNH! Toàn bộ các lệnh đã hoàn thành đều đã được up lên Cyber đầy đủ.")
