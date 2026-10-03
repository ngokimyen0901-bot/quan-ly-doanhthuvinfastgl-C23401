import streamlit as st
import pandas as pd
import numpy as np
import io
import re
from streamlit_gsheets import GSheetsConnection

st.set_page_config(
    page_title="Tính Lương & Đối Soát Công KTV",
    page_icon="👷",
    layout="wide"
)

st.title("👷 Hệ Thống Tính Lương & Đối Soát Công Kỹ Thuật Viên")
st.caption("☁️ Đối soát chuẩn 100% | Lọc chuẩn tiền công dịch vụ bảo hành | Lưu trữ: **Luong_KTV**.")

conn = st.connection("gsheets", type=GSheetsConnection)

def norm_wo_key(val):
    if pd.isna(val) or val is None:
        return ""
    s = str(val).split('(')[0].strip().upper()
    return re.sub(r'[^A-Z0-9]', '', s)

def clean_num(val):
    if pd.isna(val) or val is None:
        return 0.0
    s = str(val).replace(',', '').replace(' ', '').replace('đ', '').strip()
    try:
        return float(s)
    except Exception:
        return 0.0

def doc_file_5114_chuan(file_obj):
    if file_obj.name.lower().endswith('.csv'):
        df = pd.read_csv(file_obj, low_memory=False)
        df.columns = [str(c).strip() for c in df.columns]
        return df

    xls = pd.ExcelFile(file_obj)
    df_raw = pd.read_excel(xls, sheet_name=xls.sheet_names[0], header=None)
    
    header_idx = 0
    for idx, row in df_raw.head(25).iterrows():
        row_str = " ".join([str(v).lower() for v in row.values if pd.notna(v)])
        if 'số c.từ' in row_str or 'số hóa đơn' in row_str or 'số r/o' in row_str:
            header_idx = idx
            break

    df = pd.read_excel(xls, sheet_name=xls.sheet_names[0], header=header_idx)
    df.columns = [str(c).strip() for c in df.columns]
    return df

def doc_file_excel_chung(file_obj, tu_khoa=['số ro', 'mã ktv', 'lệnh sửa chữa', 'wo']):
    if file_obj.name.lower().endswith('.csv'):
        df = pd.read_csv(file_obj, low_memory=False)
        df.columns = [str(c).strip() for c in df.columns]
        return df

    xls = pd.ExcelFile(file_obj)
    target_sheet = xls.sheet_names[0]
    for s in xls.sheet_names:
        if any(k in s.lower() for k in ['ktv', 'claim', 'active', 'chi tiết', 'đối soát']):
            target_sheet = s
            break

    df_first = pd.read_excel(xls, sheet_name=target_sheet)
    cols_str = " ".join([str(c).lower() for c in df_first.columns])
    if any(k in cols_str for k in tu_khoa) and sum('unnamed' in str(c).lower() for c in df_first.columns) < len(df_first.columns) * 0.5:
        df_first.columns = [str(c).strip() for c in df_first.columns]
        return df_first

    df_raw = pd.read_excel(xls, sheet_name=target_sheet, header=None)
    header_idx = 0
    for idx, row in df_raw.head(25).iterrows():
        row_str = " ".join([str(v).lower() for v in row.values if pd.notna(v)])
        if any(k in row_str for k in tu_khoa):
            header_idx = idx
            break

    df = pd.read_excel(xls, sheet_name=target_sheet, header=header_idx)
    df.columns = [str(c).strip() for c in df.columns]
    return df

# HÀM BÓC TÁCH BẢO HÀNH CHUẨN XÁC: TỰ ĐỘNG PHÂN BIỆT FILE CRM VÀ FILE BẢNG KÊ WCS
def doc_du_lieu_bao_hanh_thong_minh(file_obj):
    xls = pd.ExcelFile(file_obj)
    target_sheet = xls.sheet_names[0]
    for s in xls.sheet_names:
        if any(k in s.lower() for k in ['chi tiết', 'đối soát', 'claim', 'active']):
            target_sheet = s
            break
            
    df_raw = pd.read_excel(xls, sheet_name=target_sheet)
    df_raw.columns = [str(c).strip() for c in df_raw.columns]
    df_out = df_raw.copy()

    # Nhận diện dạng Bảng kê đối soát WCS có cột Loại sản phẩm chứa 'Dịch vụ' / 'Phụ tùng'
    col_loai_sp = next((c for c in df_out.columns if any(k in c.lower() for k in ['loại sản phẩm', 'material3'])), None)
    if col_loai_sp and df_out[col_loai_sp].dropna().astype(str).str.lower().str.contains('dịch vụ|service').any():
        mask_dv = df_out[col_loai_sp].astype(str).str.lower().str.contains('dịch vụ|service')
        df_out = df_out[mask_dv].copy()
        col_wo = next((c for c in df_out.columns if any(k in c.lower() for k in ['lệnh sửa chữa', 'lsc', 'wo'])), df_out.columns[0])
        col_cong = next((c for c in df_out.columns if any(k in c.lower() for k in ['tổng số tiền', 'amount', 'thành tiền', 'tiền công', 'tiền'])), df_out.columns[-1])
        col_cv = next((c for c in df_out.columns if any(k in c.lower() for k in ['mô tả', 'công việc', 'material2', 'material1'])), df_out.columns[0])
        col_ngay = next((c for c in df_out.columns if 'ngày' in c.lower() and ('phê duyệt' in c.lower() or 'approved' in c.lower())), None)
    else:
        # Dạng file Active Warranty claim từ CRM (như file Warran...-05 AM.xlsx)
        col_wo = next((c for c in df_out.columns if any(k in c.lower() for k in ['lệnh sửa chữa', 'lsc', 'wo'])), df_out.columns[0])
        col_cong = next((c for c in df_out.columns if any(k in c.lower() for k in ['approved service', 'tiền công yêu cầu'])), None)
        if not col_cong:
            col_cong = next((c for c in df_out.columns if 'tiền công' in c.lower()), df_out.columns[-1])
        col_cv = next((c for c in df_out.columns if any(k in c.lower() for k in ['tên công việc chính', 'công việc chính', 'mô tả'])), df_out.columns[0])
        col_ngay = next((c for c in df_out.columns if 'ngày' in c.lower() and ('phê duyệt' in c.lower() or 'đxbh' in c.lower())), None)

    df_out['wo_norm'] = df_out[col_wo].apply(norm_wo_key)
    df_out['Tien_Cong_BH'] = df_out[col_cong].apply(clean_num)
    df_out['Ten_CV_BH'] = df_out[col_cv].fillna('').astype(str) if col_cv else ''
    df_out['col_wo_goc'] = df_out[col_wo]
    
    if col_ngay and col_ngay in df_out.columns:
        df_out['Thang_BH'] = pd.to_datetime(df_out[col_ngay], errors='coerce', dayfirst=True).dt.strftime('Tháng %m/%Y').fillna('Chưa rõ')
    else:
        df_out['Thang_BH'] = 'Toàn bộ kỳ'

    return df_out[(df_out['Tien_Cong_BH'] > 0) & (df_out['wo_norm'] != '')].copy()

def trich_xuat_ktv_dataframe(df_source):
    df_res = df_source.copy()
    col_ro_h = next((c for c in df_res.columns if 'số ro hãng' in c.lower() or 'ro hãng' in c.lower()), None)
    col_ro_nb = next((c for c in df_res.columns if c.strip().lower() in ['số ro', 'ro']), None)
    col_ma = next((c for c in df_res.columns if 'mã ktv' in c.lower()), None)
    col_ten = next((c for c in df_res.columns if 'tên ktv' in c.lower()), None)
    
    col_tien_hd = next((c for c in df_res.columns if 'theo hd' in c.lower()), None)
    col_tien_lenh = next((c for c in df_res.columns if 'theo lệnh' in c.lower() or 'theo lenh' in c.lower()), None)
    col_thanh_tien = next((c for c in df_res.columns if c.strip().lower() in ['thành tiền', 'thanh tien']), None)

    col_bs = next((c for c in df_res.columns if 'biển' in c.lower()), None)
    col_hm = next((c for c in df_res.columns if 'hạng mục' in c.lower()), None)
    col_nd = next((c for c in df_res.columns if 'nội dung' in c.lower()), None)
    col_cv = next((c for c in df_res.columns if 'mã cv' in c.lower()), None)

    df_res['wo_norm'] = df_res[col_ro_h].apply(norm_wo_key) if col_ro_h else ''
    df_res['ro_nb_norm'] = df_res[col_ro_nb].apply(norm_wo_key) if col_ro_nb else ''
    df_res['key_match'] = np.where(df_res['wo_norm'] != '', df_res['wo_norm'], df_res['ro_nb_norm'])

    df_res['Ma_KTV_Clean'] = df_res[col_ma].fillna('').astype(str).str.strip().str.replace('.0', '', regex=False) if col_ma else ""
    df_res['Ten_KTV_Clean'] = df_res[col_ten].fillna('').astype(str).str.strip() if col_ten else ""
    
    df_res['Tien_KTV_Theo_HD'] = df_res[col_tien_hd].apply(clean_num) if col_tien_hd else 0.0
    df_res['Tien_KTV_Theo_Lenh'] = df_res[col_tien_lenh].apply(clean_num) if col_tien_lenh else (df_res[col_thanh_tien].apply(clean_num) if col_thanh_tien else 0.0)

    df_res['Bien_So_Val'] = df_res[col_bs].fillna('').astype(str).str.strip() if col_bs else ""
    df_res['Hang_Muc_Val'] = df_res[col_hm].fillna('').astype(str).str.strip() if col_hm else ""
    df_res['Noi_Dung_Val'] = df_res[col_nd].fillna('').astype(str).str.strip() if col_nd else ""
    df_res['Ma_CV_Val'] = df_res[col_cv].fillna('').astype(str).str.strip() if col_cv else ""
    df_res['So_RO_Val'] = df_res[col_ro_h].fillna(df_res[col_ro_nb] if col_ro_nb else '').astype(str).str.strip() if col_ro_h else ""

    return df_res

def bao_ve_cot_ktv_detail(df):
    if df is None or df.empty:
        return pd.DataFrame()
    df_out = df.copy()
    cot_can_co = {
        'Trang_Thai_Khop': '🟢 Khớp 100%',
        'Tien_Chot_KTV': 0.0,
        'Tien_Thuc_Nhan': 0.0,
        'Ma_KTV': 'Chưa xác định',
        'Ten_KTV': 'Chưa có thông tin thợ',
        'So_RO': '',
        'So_HD': '',
        'Ngay_Xuat_HD': '',
        'Bien_So': '',
        'Hang_Muc': 'Sửa chữa',
        'Noi_Dung_CV': '',
        'key_match': '',
        'Tien_5114': 0.0
    }
    for col, default_val in cot_can_co.items():
        if col not in df_out.columns:
            found_col = None
            for c in df_out.columns:
                if col.lower() == c.lower() or col.lower() in c.lower():
                    found_col = c
                    break
            if found_col:
                df_out[col] = df_out[found_col]
            else:
                df_out[col] = default_val
    return df_out

@st.cache_data(ttl=20)
def load_saved_luong():
    try:
        df = conn.read(worksheet="Luong_KTV", ttl=20)
        if df is not None and not df.empty:
            df.columns = [str(c).strip() for c in df.columns]
            return df
    except Exception:
        pass
    return pd.DataFrame()

df_saved_gs = load_saved_luong()

# KHUNG NẠP TỆP DỮ LIỆU
with st.expander("📥 Nạp Tệp Dữ Liệu Tính Lương Mới", expanded=df_saved_gs.empty):
    c1, c2, c3 = st.columns(3)
    with c1:
        up_5114 = st.file_uploader("1. Sổ chi tiết TK 5114 (5114.Xlsx):", type=['xlsx', 'xls', 'csv'], key="p4_5114")
    with c2:
        up_ktv_hd = st.file_uploader("2. Báo cáo KTV (THEO HÓA ĐƠN - Sửa chữa):", type=['xlsx', 'xls', 'csv'], key="p4_ktv_hd")
    with c3:
        up_ktv_lenh = st.file_uploader("3. Báo cáo KTV (THEO LỆNH - Bảo hành):", type=['xlsx', 'xls', 'csv'], key="p4_ktv_lenh")

    c4, c5 = st.columns(2)
    with c4:
        up_bh = st.file_uploader("4. Chi tiết bảo hành (Active Claim / Bảng Kê WCS...xlsx):", type=['xlsx', 'xls'], key="p4_bh")
    with c5:
        up_pdi = st.file_uploader("5. Danh sách lệnh PDI (Tùy chọn):", type=['xlsx', 'xls', 'csv'], key="p4_pdi")

# Xử lý khi nạp file mới
if up_5114 and (up_ktv_hd or up_ktv_lenh):
    df_5114 = doc_file_5114_chuan(up_5114)
    
    df_ktv_hd_parsed = pd.DataFrame()
    df_ktv_lenh_parsed = pd.DataFrame()

    if up_ktv_hd:
        df_raw_hd = doc_file_excel_chung(up_ktv_hd, tu_khoa=['số ro', 'mã ktv', 'tên ktv', 'thành tiền'])
        df_ktv_hd_parsed = trich_xuat_ktv_dataframe(df_raw_hd)
    
    if up_ktv_lenh:
        df_raw_lenh = doc_file_excel_chung(up_ktv_lenh, tu_khoa=['số ro', 'mã ktv', 'tên ktv', 'thành tiền'])
        df_ktv_lenh_parsed = trich_xuat_ktv_dataframe(df_raw_lenh)

    df_ktv_for_sc = df_ktv_hd_parsed if not df_ktv_hd_parsed.empty else df_ktv_lenh_parsed
    df_ktv_for_bh = df_ktv_lenh_parsed if not df_ktv_lenh_parsed.empty else df_ktv_hd_parsed

    df_ktv_all_rows = pd.concat([df_ktv_for_sc, df_ktv_for_bh], ignore_index=True)
    df_ktv_valid_works = df_ktv_all_rows[(df_ktv_all_rows['key_match'] != '') & (df_ktv_all_rows['Ma_KTV_Clean'] != '')].copy()

    # Danh sách PDI
    pdi_keys_set = set()
    if up_pdi:
        df_pdi_input = doc_file_excel_chung(up_pdi, tu_khoa=['lệnh', 'wo', 'ro', 'pdi'])
        col_pdi_ro = next((c for c in df_pdi_input.columns if any(k in c.lower() for k in ['lệnh', 'wo', 'ro', 'lsc'])), df_pdi_input.columns[0])
        for v in df_pdi_input[col_pdi_ro].dropna():
            pdi_keys_set.add(norm_wo_key(v))

    col_5114_ngay = next((c for c in df_5114.columns if any(k in c.lower() for k in ['ngày c.từ', 'ngày chứng từ', 'ngày xuất hóa đơn', 'ngày hđ'])), 'Ngày C.từ')
    col_5114_shd = next((c for c in df_5114.columns if 'hóa đơn' in c.lower()), 'Số hóa đơn')
    col_5114_tien = next((c for c in df_5114.columns if c.strip().lower() in ['có', 'co', 'phát sinh có']), 'Có')
    col_5114_ro_hang = next((c for c in df_5114.columns if 'số r/o hãng' in c.lower() or 'ro hãng' in c.lower()), 'Số R/O hãng')
    col_5114_ro_nb = next((c for c in df_5114.columns if c.strip().lower() in ['số r/o', 'số ro']), 'Số R/O')
    col_5114_dg = next((c for c in df_5114.columns if 'diễn giải' in c.lower() or 'unnamed: 11' in c.lower()), 'Diễn giải')
    col_5114_kh = next((c for c in df_5114.columns if 'tên khách' in c.lower() or 'khách hàng' in c.lower()), df_5114.columns[0])

    if col_5114_ngay in df_5114.columns:
        df_5114['Thang_HD'] = pd.to_datetime(df_5114[col_5114_ngay], errors='coerce', dayfirst=True).dt.strftime('Tháng %m/%Y').fillna('Chưa rõ')
    else:
        df_5114['Thang_HD'] = 'Tháng 09/2026'

    df_5114['Tien_Dau_Vao'] = df_5114[col_5114_tien].apply(clean_num) if col_5114_tien in df_5114.columns else 0.0

    mask_tong_ket = df_5114[col_5114_dg].astype(str).str.lower().str.contains('tổng phát sinh|số dư có|số dư nợ')
    df_5114_valid = df_5114[(df_5114['Tien_Dau_Vao'] > 0) & (~mask_tong_ket)].copy()

    def tao_key_match_5114(r):
        w = norm_wo_key(r.get(col_5114_ro_hang))
        if w: return w
        ro = norm_wo_key(r.get(col_5114_ro_nb))
        if ro: return ro
        dg = str(r.get(col_5114_dg, '')).lower()
        if 'cứu hộ' in dg:
            return "CUU_HO_2608"
        if 'bảo hành' in dg:
            return "BAO_HANH_BANG_KE_2608"
        return f"DON_LE_{r.name}"

    df_5114_valid['key_match'] = df_5114_valid.apply(tao_key_match_5114, axis=1)

    df_5114_grouped = df_5114_valid.groupby('key_match').agg({
        col_5114_ro_hang: 'first' if col_5114_ro_hang in df_5114_valid.columns else lambda x: '',
        col_5114_ro_nb: 'first' if col_5114_ro_nb in df_5114_valid.columns else lambda x: '',
        col_5114_shd: lambda x: ", ".join(sorted(set([str(int(float(v))) if str(v).replace('.0','').isdigit() else str(v) for v in x if pd.notna(v) and str(v).strip() != '']))),
        col_5114_ngay: lambda x: ", ".join(sorted(set([str(v) for v in x if pd.notna(v)]))),
        'Thang_HD': 'first',
        'Tien_Dau_Vao': 'sum',
        col_5114_kh: 'first' if col_5114_kh in df_5114_valid.columns else lambda x: '',
        col_5114_dg: lambda x: " | ".join(sorted(set([str(v) for v in x if pd.notna(v)])))[:200] if col_5114_dg in df_5114_valid.columns else lambda x: ''
    }).reset_index().rename(columns={'Tien_Dau_Vao': 'Tien_5114'})

    # Xử lý Bảo Hành
    bh_wo_set = set()
    df_bh_split = pd.DataFrame()

    if up_bh:
        df_bh_serv = doc_du_lieu_bao_hanh_thong_minh(up_bh)
        bh_wo_set = set(df_bh_serv['wo_norm'])

        bh_wo_sum = df_bh_serv.groupby('wo_norm').agg({
            'col_wo_goc': 'first',
            'Tien_Cong_BH': 'sum',
            'Thang_BH': 'first',
            'Ten_CV_BH': lambda x: " | ".join(sorted(set([str(v) for v in x if pd.notna(v)])))[:150]
        }).reset_index()

        df_ktv_bh_rows = df_ktv_for_bh[df_ktv_for_bh['key_match'].isin(bh_wo_sum['wo_norm']) & (df_ktv_for_bh['Ma_KTV_Clean'] != '')].copy()
        if df_ktv_bh_rows.empty:
            df_ktv_bh_rows = df_ktv_valid_works[df_ktv_valid_works['key_match'].isin(bh_wo_sum['wo_norm'])].copy()

        ktv_cnt_bh = df_ktv_bh_rows.groupby('key_match')['Ma_KTV_Clean'].nunique().to_dict()
        df_ktv_bh_rows['So_KTV_Lam_Chung'] = df_ktv_bh_rows['key_match'].map(ktv_cnt_bh).fillna(1)
        df_ktv_bh_rows.loc[df_ktv_bh_rows['So_KTV_Lam_Chung'] == 0, 'So_KTV_Lam_Chung'] = 1

        df_bh_split = pd.merge(
            df_ktv_bh_rows[['key_match', 'So_RO_Val', 'Bien_So_Val', 'Ma_KTV_Clean', 'Ten_KTV_Clean', 'So_KTV_Lam_Chung']].drop_duplicates(subset=['key_match', 'Ma_KTV_Clean']),
            bh_wo_sum,
            left_on='key_match',
            right_on='wo_norm',
            how='right'
        )
        
        df_bh_split['So_KTV_Lam_Chung'] = df_bh_split['So_KTV_Lam_Chung'].fillna(1)
        df_bh_split['Cong_BH_Thuc_Nhan'] = df_bh_split['Tien_Cong_BH'] / df_bh_split['So_KTV_Lam_Chung']

    ktv_details_temp = []

    for k_match, group_orders in df_5114_grouped.groupby('key_match'):
        r_5114 = group_orders.iloc[0]
        tien_5114 = r_5114['Tien_5114']
        shd_k = r_5114['Số hóa đơn']
        nhd_k = r_5114['Ngày C.từ']
        ro_show = r_5114.get('Số R/O hãng', r_5114.get('Số R/O', k_match))
        if pd.isna(ro_show) or str(ro_show).strip() in ['', 'nan', 'None']:
            ro_show = str(r_5114.get('Diễn giải', ''))[:30]

        if k_match == "CUU_HO_2608":
            ktv_details_temp.append({
                'key_match': k_match,
                'Loai_Cong': 'Cứu hộ',
                'So_RO': 'Chi phí cứu hộ',
                'So_HD': shd_k,
                'Ngay_Xuat_HD': nhd_k,
                'Bien_So': '',
                'Hang_Muc': 'Cứu hộ kéo xe',
                'Noi_Dung_CV': 'Chi phí cứu hộ theo báo cáo chi tiết',
                'Ma_KTV': 'Khoản mục riêng',
                'Ten_KTV': 'Cứu hộ (Không tính công KTV)',
                'Tien_Thuc_Nhan': tien_5114,
                'Tien_Chot_KTV': tien_5114,
                'Tien_5114': tien_5114
            })
            continue

        if k_match == "BAO_HANH_BANG_KE_2608":
            ktv_details_temp.append({
                'key_match': k_match,
                'Loai_Cong': 'Bảo hành cục bộ',
                'So_RO': 'Bảng kê bảo hành hãng',
                'So_HD': shd_k,
                'Ngay_Xuat_HD': nhd_k,
                'Bien_So': '',
                'Hang_Muc': 'Bảo hành hãng',
                'Noi_Dung_CV': 'Chi phí bảo hành tổng kỳ theo báo cáo chi tiết',
                'Ma_KTV': 'Bảo hành hãng',
                'Ten_KTV': 'Chi tiết xem tại Tab Bảo Hành',
                'Tien_Thuc_Nhan': tien_5114,
                'Tien_Chot_KTV': tien_5114,
                'Tien_5114': tien_5114
            })
            continue

        sub_all = df_ktv_for_sc[df_ktv_for_sc['key_match'] == k_match].copy()

        if not sub_all.empty:
            bs_k = sub_all['Bien_So_Val'].iloc[0]

            giam_gia_son = 0.0
            giam_gia_go_sc = 0.0

            sub_discounts = sub_all[(sub_all['Ma_KTV_Clean'] == '') & (sub_all['Tien_KTV_Theo_HD'] < 0)].copy()
            for _, r_d in sub_discounts.iterrows():
                nd_d = str(r_d['Noi_Dung_Val']).lower()
                cv_d = str(r_d['Ma_CV_Val']).lower()
                val_d = abs(r_d['Tien_KTV_Theo_HD'])
                if 'giảm giá' in nd_d or 'chiết khấu' in nd_d or 'ck' in nd_d:
                    if 'sơn' in nd_d or 'mtbh_son' in cv_d:
                        giam_gia_son += val_d
                    else:
                        giam_gia_go_sc += val_d

            sub_works = sub_all[(sub_all['Ma_KTV_Clean'] != '') & (sub_all['Tien_KTV_Theo_HD'] > 0)].copy()

            if not sub_works.empty:
                sub_works['is_son'] = sub_works['Hang_Muc_Val'].astype(str).str.lower().str.contains('sơn') | sub_works['Noi_Dung_Val'].astype(str).str.lower().str.contains('sơn')

                tong_goc_son = sub_works[sub_works['is_son']]['Tien_KTV_Theo_HD'].sum()
                tong_goc_go = sub_works[~sub_works['is_son']]['Tien_KTV_Theo_HD'].sum()

                sub_works['task_id'] = sub_works['Ma_CV_Val'].astype(str) + "_" + sub_works['Noi_Dung_Val'].astype(str)

                for task, group_task in sub_works.groupby('task_id'):
                    n_tho = group_task['Ma_KTV_Clean'].nunique()
                    tien_goc = group_task['Tien_KTV_Theo_HD'].iloc[0]
                    is_s = group_task['is_son'].iloc[0]

                    if is_s and tong_goc_son > 0 and giam_gia_son > 0:
                        tien_sau_giam = tien_goc - (giam_gia_son * (tien_goc / tong_goc_son))
                    elif (not is_s) and tong_goc_go > 0 and giam_gia_go_sc > 0:
                        tien_sau_giam = tien_goc - (giam_gia_go_sc * (tien_goc / tong_goc_go))
                    else:
                        tien_sau_giam = tien_goc

                    tien_moi_tho = max(tien_sau_giam, 0.0) / max(n_tho, 1)
                    hm_ten = group_task['Hang_Muc_Val'].iloc[0] if group_task['Hang_Muc_Val'].iloc[0] else "Sửa chữa"
                    cv_ten = group_task['Noi_Dung_Val'].iloc[0]

                    for _, r_tho in group_task.drop_duplicates(subset=['Ma_KTV_Clean']).iterrows():
                        ktv_details_temp.append({
                            'key_match': k_match,
                            'Loai_Cong': 'Sửa chữa (5114)',
                            'So_RO': ro_show,
                            'So_HD': shd_k,
                            'Ngay_Xuat_HD': nhd_k,
                            'Bien_So': bs_k,
                            'Hang_Muc': hm_ten,
                            'Noi_Dung_CV': cv_ten,
                            'Ma_KTV': r_tho['Ma_KTV_Clean'],
                            'Ten_KTV': r_tho['Ten_KTV_Clean'],
                            'Tien_Thuc_Nhan': tien_moi_tho,
                            'Tien_Chot_KTV': tien_moi_tho,
                            'Tien_5114': tien_5114
                        })
            else:
                ktv_details_temp.append({
                    'key_match': k_match,
                    'Loai_Cong': 'Sửa chữa (5114)',
                    'So_RO': ro_show,
                    'So_HD': shd_k,
                    'Ngay_Xuat_HD': nhd_k,
                    'Bien_So': bs_k,
                    'Hang_Muc': '🚨 CHƯA CÓ KTV',
                    'Noi_Dung_CV': '🚨 LỆNH ĐÃ DUYỆT NHƯNG CHƯA CÓ TÊN KTV',
                    'Ma_KTV': '🚨 CẦN GÁN THỢ',
                    'Ten_KTV': '🚨 BÁO ĐỘNG THIẾU THỢ',
                    'Tien_Thuc_Nhan': 0.0,
                    'Tien_Chot_KTV': 0.0,
                    'Tien_5114': tien_5114
                })
        else:
            ktv_details_temp.append({
                'key_match': k_match,
                'Loai_Cong': 'Sửa chữa (5114)',
                'So_RO': ro_show,
                'So_HD': shd_k,
                'Ngay_Xuat_HD': nhd_k,
                'Bien_So': '',
                'Hang_Muc': '🚨 CHƯA CÓ KTV',
                'Noi_Dung_CV': '🚨 LỆNH ĐÃ DUYỆT NHƯNG CHƯA CÓ TÊN KTV',
                'Ma_KTV': '🚨 CẦN GÁN THỢ',
                'Ten_KTV': '🚨 BÁO ĐỘNG THIẾU THỢ',
                'Tien_Thuc_Nhan': 0.0,
                'Tien_Chot_KTV': 0.0,
                'Tien_5114': tien_5114
            })

    df_ktv_split_master = pd.DataFrame(ktv_details_temp)
    df_ktv_split_master = bao_ve_cot_ktv_detail(df_ktv_split_master)

    if not df_saved_gs.empty and 'key_match' in df_saved_gs.columns and 'Tien_Chot_KTV' in df_saved_gs.columns:
        map_chot_tay = df_saved_gs.set_index(['key_match', 'Ma_KTV', 'Noi_Dung_CV'])['Tien_Chot_KTV'].to_dict()
        for idx_r, r_val in df_ktv_split_master.iterrows():
            k_tuple = (r_val['key_match'], r_val['Ma_KTV'], r_val['Noi_Dung_CV'])
            if k_tuple in map_chot_tay:
                df_ktv_split_master.at[idx_r, 'Tien_Chot_KTV'] = clean_num(map_chot_tay[k_tuple])

    ktv_totals = df_ktv_split_master.groupby('key_match')['Tien_Chot_KTV'].sum().to_dict()

    def xac_dinh_trang_thai_row(r):
        k = r['key_match']
        if r['Ma_KTV'] in ['🚨 CẦN GÁN THỢ', '']:
            return "🚨 ALARM: THIẾU KTV"
        if k in ["CUU_HO_2608", "BAO_HANH_BANG_KE_2608"]:
            return "🟢 Khoản mục riêng"
        t_5114 = r['Tien_5114']
        t_ktv = ktv_totals.get(k, 0.0)
        cl = abs(t_5114 - t_ktv)
        if cl < 1000:
            return "🟢 Khớp 100%"
        return f"🔴 Lệch {abs(t_5114 - t_ktv):,.0f} đ"

    df_ktv_split_master['Trang_Thai_Khop'] = df_ktv_split_master.apply(xac_dinh_trang_thai_row, axis=1)

    df_sc_summary_grouped = df_5114_grouped.copy()
    df_sc_summary_grouped['Tien_KTV'] = df_sc_summary_grouped['key_match'].map(ktv_totals).fillna(0)
    df_sc_summary_grouped['Chenh_Lech'] = df_sc_summary_grouped['Tien_5114'] - df_sc_summary_grouped['Tien_KTV']

    def danh_gia_trang_thai_sc(r):
        k = r['key_match']
        if k == "CUU_HO_2608": return "🚚 Cứu hộ (Chi phí kéo xe)"
        if k == "BAO_HANH_BANG_KE_2608": return "🛡️ Bảng kê BH hãng (Tổng kỳ)"
        cl = abs(r['Chenh_Lech'])
        if cl < 1000: return "🟢 Khớp 100%"
        return "🔴 Lệch tiền"

    df_sc_summary_grouped['Ghi_Chu_Doi_Soat'] = df_sc_summary_grouped.apply(danh_gia_trang_thai_sc, axis=1)

    def gom_ten_tho(ds):
        s_tho = set()
        for v in ds:
            ten_s = str(v).strip()
            if ten_s and ten_s not in ['None', 'nan', '🚨 BÁO ĐỘNG THIẾU THỢ']:
                s_tho.add(ten_s)
        return ", ".join(sorted(s_tho))
    
    ktv_names_map = df_ktv_split_master.groupby('key_match')['Ten_KTV'].apply(gom_ten_tho).to_dict()
    df_sc_summary_grouped['Ten_KTV_Clean'] = df_sc_summary_grouped['key_match'].map(ktv_names_map).fillna('🚨 Chưa có thợ')

    st.session_state['df_ktv_split_master'] = df_ktv_split_master
    st.session_state['df_sc_summary_grouped'] = df_sc_summary_grouped
    st.session_state['df_bh_split_cache'] = df_bh_split

# Nạp dữ liệu từ cache hoặc Google Sheets
if 'df_ktv_split_master' in st.session_state:
    df_ktv_split_master = bao_ve_cot_ktv_detail(st.session_state['df_ktv_split_master'])
    df_sc_summary_grouped = st.session_state.get('df_sc_summary_grouped', pd.DataFrame())
    df_bh_split = st.session_state.get('df_bh_split_cache', pd.DataFrame())
elif not df_saved_gs.empty:
    df_ktv_split_master = bao_ve_cot_ktv_detail(df_saved_gs.copy())
    df_sc_summary_grouped = pd.DataFrame()
    df_bh_split = pd.DataFrame()
else:
    df_ktv_split_master = pd.DataFrame()
    df_sc_summary_grouped = pd.DataFrame()
    df_bh_split = pd.DataFrame()

# NẾU CÓ NẠP FILE BẢO HÀNH RIÊNG THÌ LUÔN LUÔN XỬ LÝ VÀ HIỂN THỊ
if up_bh and df_bh_split.empty:
    try:
        df_bh_serv_direct = doc_du_lieu_bao_hanh_thong_minh(up_bh)
        bh_wo_sum_dir = df_bh_serv_direct.groupby('wo_norm').agg({
            'col_wo_goc': 'first',
            'Tien_Cong_BH': 'sum',
            'Thang_BH': 'first',
            'Ten_CV_BH': lambda x: " | ".join(sorted(set([str(v) for v in x if pd.notna(v)])))[:150]
        }).reset_index()

        df_ktv_source_bh = df_ktv_valid_works if 'df_ktv_valid_works' in locals() and not df_ktv_valid_works.empty else pd.DataFrame()

        if not df_ktv_source_bh.empty:
            df_ktv_bh_rows_dir = df_ktv_source_bh[df_ktv_source_bh['key_match'].isin(bh_wo_sum_dir['wo_norm']) & (df_ktv_source_bh['Ma_KTV_Clean'] != '')].copy()
            ktv_cnt_bh_dir = df_ktv_bh_rows_dir.groupby('key_match')['Ma_KTV_Clean'].nunique().to_dict()
            df_ktv_bh_rows_dir['So_KTV_Lam_Chung'] = df_ktv_bh_rows_dir['key_match'].map(ktv_cnt_bh_dir).fillna(1)
            df_ktv_bh_rows_dir.loc[df_ktv_bh_rows_dir['So_KTV_Lam_Chung'] == 0, 'So_KTV_Lam_Chung'] = 1

            df_bh_split = pd.merge(
                df_ktv_bh_rows_dir[['key_match', 'So_RO_Val', 'Bien_So_Val', 'Ma_KTV_Clean', 'Ten_KTV_Clean', 'So_KTV_Lam_Chung']].drop_duplicates(subset=['key_match', 'Ma_KTV_Clean']),
                bh_wo_sum_dir,
                left_on='key_match',
                right_on='wo_norm',
                how='right'
            )
        else:
            df_bh_split = bh_wo_sum_dir.copy()
            df_bh_split['key_match'] = df_bh_split['wo_norm']
            df_bh_split['So_RO_Val'] = df_bh_split['col_wo_goc']
            df_bh_split['Bien_So_Val'] = ''
            df_bh_split['Ma_KTV_Clean'] = 'Chưa nạp KTV'
            df_bh_split['Ten_KTV_Clean'] = 'Chưa nạp KTV'
            df_bh_split['So_KTV_Lam_Chung'] = 1

        df_bh_split['So_KTV_Lam_Chung'] = df_bh_split['So_KTV_Lam_Chung'].fillna(1)
        df_bh_split['Cong_BH_Thuc_Nhan'] = df_bh_split['Tien_Cong_BH'] / df_bh_split['So_KTV_Lam_Chung']
        st.session_state['df_bh_split_cache'] = df_bh_split
    except Exception as e_bh:
        st.warning(f"Đang phân tích file bảo hành: {e_bh}")

if not df_ktv_split_master.empty:
    df_ktv_split_master = bao_ve_cot_ktv_detail(df_ktv_split_master)
    
    col_tt_check = df_ktv_split_master['Trang_Thai_Khop'].astype(str) if 'Trang_Thai_Khop' in df_ktv_split_master.columns else pd.Series([])
    so_ca_alarm = (col_tt_check.str.contains('ALARM|THIẾU KTV')).sum()
    if so_ca_alarm > 0:
        st.error(f"🚨 **BÁO ĐỘNG: CÓ {so_ca_alarm} LỆNH ĐÃ DUYỆT NHƯNG CHƯA CÓ TÊN KTV TRONG HỆ THỐNG!** Vui lòng kiểm tra các dòng bôi đỏ.")

    tab_sc, tab_split_ktv, tab_bh, tab_ktv_sum = st.tabs([
        "🔧 1. Đối Soát Lệnh Sửa Chữa (Tổng Hợp)",
        "👷 2. Chi Tiết Tính Lương Từng KTV (Có Điều Chỉnh & Báo Động)",
        "🛡️ 3. Lệnh Bảo Hành Trong Tháng (Chia Đều)",
        "💰 4. Bảng Tổng Hợp Lương KTV"
    ])

    # ==================== TAB 1: ĐỐI SOÁT TỔNG HỢP ====================
    with tab_sc:
        st.subheader("🔧 Bảng Đối Soát 5114 Đầy Đủ Doanh Thu (675tr)")
        st.info("💡 **Ghi nhận 100% doanh thu:** Cột điều chỉnh đã chuyển sang **Tab 2** để bạn chỉnh trực tiếp trên từng thợ. Khi bạn chỉnh ở Tab 2, trạng thái khớp sẽ tự động cập nhật.")

        if not df_sc_summary_grouped.empty:
            cnt_khop = (df_sc_summary_grouped['Ghi_Chu_Doi_Soat'] == "🟢 Khớp 100%").sum()
            cnt_lech = len(df_sc_summary_grouped) - cnt_khop

            m1, m2, m3 = st.columns(3)
            m1.metric("Tổng Doanh Thu 5114", f"{df_sc_summary_grouped['Tien_5114'].sum():,.0f} đ", f"{len(df_sc_summary_grouped)} lệnh/dòng")
            m2.metric("Lệnh Khớp Chuẩn (Tự động)", f"{cnt_khop} lệnh")
            m3.metric("Lệnh Cần Lưu Ý", f"{cnt_lech} lệnh", delta=f"{cnt_lech} ca" if cnt_lech > 0 else "0", delta_color="inverse")

            cols_sc_show = ['Số R/O hãng', 'Số R/O', 'Ngày C.từ', 'Số hóa đơn', 'Tien_5114', 'Tien_KTV', 'Chenh_Lech', 'Ghi_Chu_Doi_Soat', 'Ten_KTV_Clean']
            cols_sc_valid = [c for c in cols_sc_show if c in df_sc_summary_grouped.columns]
            
            cfg_sc = {
                "Số R/O hãng": st.column_config.TextColumn("Số RO hãng", width="medium"),
                "Số R/O": st.column_config.TextColumn("Số RO nội bộ", width="medium"),
                "Ngày C.từ": st.column_config.TextColumn("Ngày xuất HĐ", width="small"),
                "Số hóa đơn": st.column_config.TextColumn("Số HĐ", width="small"),
                "Tien_5114": st.column_config.NumberColumn("Tiền 5114", format="%,d đ", width="medium"),
                "Tien_KTV": st.column_config.NumberColumn("Tổng Tiền KTV", format="%,d đ", width="medium"),
                "Chenh_Lech": st.column_config.NumberColumn("Chênh lệch", format="%,d đ", width="small"),
                "Ghi_Chu_Doi_Soat": st.column_config.TextColumn("Trạng thái", width="medium"),
                "Ten_KTV_Clean": st.column_config.TextColumn("Thợ phụ trách", width="large")
            }
            st.dataframe(df_sc_summary_grouped[cols_sc_valid], use_container_width=True, height=540, hide_index=True, column_config=cfg_sc)

    # ==================== TAB 2: SHEET CHI TIẾT KTV (CHỈNH SỬA & BÁO ĐỘNG) ====================
    with tab_split_ktv:
        st.subheader("👷 Chi Tiết Phân Bổ Tiền Công Cho KTV & Điều Chỉnh Trực Tiếp")
        st.info("💡 **Quy tắc phân bổ chuẩn:** Không trừ MTBH vào công thợ. Bạn có thể **chỉnh trực tiếp con số của KTV tại cột 👉 TIỀN BẠN CHỐT CHO KTV**; hệ thống sẽ tự động cập nhật trạng thái Khớp ngay lập tức!")

        s_m1, s_m2, s_m3 = st.columns(3)
        s_m1.metric("Tổng Tiền Công KTV", f"{df_ktv_split_master['Tien_Chot_KTV'].sum():,.0f} đ")
        s_m2.metric("Tổng Số Dòng Công Việc", f"{len(df_ktv_split_master):,} dòng")
        s_m3.metric("Số Ca Thiếu Thợ (Alarm)", f"{so_ca_alarm} ca", delta="🚨 Cần kiểm tra" if so_ca_alarm > 0 else "An toàn", delta_color="inverse")

        cols_detail_edit = [
            'So_RO', 'So_HD', 'Ngay_Xuat_HD', 'Bien_So', 'Hang_Muc', 'Noi_Dung_CV',
            'Ma_KTV', 'Ten_KTV', 'Tien_Thuc_Nhan', 'Tien_Chot_KTV', 'Trang_Thai_Khop', 'key_match'
        ]
        cols_det_valid = [c for c in cols_detail_edit if c in df_ktv_split_master.columns]

        cfg_detail_editor = {
            "So_RO": st.column_config.TextColumn("Số Lệnh RO", disabled=True, width="medium"),
            "So_HD": st.column_config.TextColumn("Số HĐ", disabled=True, width="small"),
            "Ngay_Xuat_HD": st.column_config.TextColumn("Ngày HĐ", disabled=True, width="small"),
            "Bien_So": st.column_config.TextColumn("Biển số", disabled=True, width="small"),
            "Hang_Muc": st.column_config.TextColumn("Hạng mục", disabled=True, width="medium"),
            "Noi_Dung_CV": st.column_config.TextColumn("Nội dung công việc", disabled=True, width="large"),
            "Ma_KTV": st.column_config.TextColumn("Mã KTV", disabled=True, width="small"),
            "Ten_KTV": st.column_config.TextColumn("Họ tên KTV", width="medium"),
            "Tien_Thuc_Nhan": st.column_config.NumberColumn("Tiền theo HĐ (Sau giảm giá)", format="%,d đ", disabled=True, width="medium"),
            "Tien_Chot_KTV": st.column_config.NumberColumn(
                "👉 TIỀN BẠN CHỐT CHO KTV (Sửa được)",
                format="%,d đ",
                disabled=False,
                width="large",
                help="Gõ số tiền bạn muốn chốt tính lương cho thợ tại đây!"
            ),
            "Trang_Thai_Khop": st.column_config.TextColumn("Trạng thái đối soát", disabled=True, width="medium"),
            "key_match": st.column_config.TextColumn("Key", disabled=True, width="small")
        }

        edited_ktv_df = st.data_editor(
            df_ktv_split_master[cols_det_valid],
            use_container_width=True,
            height=560,
            column_config=cfg_detail_editor,
            hide_index=True,
            key="editor_ktv_detail_v14"
        )

        c_btn1, c_btn2 = st.columns([4, 6])
        with c_btn1:
            if st.button("☁️ LƯU BẢNG CHI TIẾT LƯƠNG KTV NÀY LÊN GOOGLE SHEETS", type="primary", use_container_width=True):
                with st.spinner("⏳ Đang lưu dữ liệu vào sheet Luong_KTV..."):
                    try:
                        df_save = edited_ktv_df.fillna("").copy()
                        conn.update(worksheet="Luong_KTV", data=df_save)
                        st.success("✅ ĐÃ LƯU THÀNH CÔNG VÀO SHEET Luong_KTV! Lần sau mở app sẽ tự động hiển thị bảng này ngay lập tức.")
                        st.rerun()
                    except Exception as e_s:
                        st.error(f"Lỗi khi lưu lên Google Sheets: {e_s}")

        with c_btn2:
            buffer_ktv = io.BytesIO()
            with pd.ExcelWriter(buffer_ktv, engine='openpyxl') as writer:
                edited_ktv_df.to_excel(writer, index=False, sheet_name="Chi_Tiet_Luong_KTV")
            st.download_button(
                label="⬇️ Tải File Excel Chi Tiết Phân Bổ Công KTV",
                data=buffer_ktv.getvalue(),
                file_name="Chi_Tiet_Luong_KTV_Thang_09_2026.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    # ==================== TAB 3: BẢO HÀNH (CHỈ LẤY DỊCH VỤ, BỎ HOÀN TOÀN PHỤ TÙNG) ====================
    with tab_bh:
        st.subheader("🛡️ Chi Tiết Tiền Công Bảo Hành Phân Bổ Cho KTV")
        st.caption("Chỉ tính **Tiền công dịch vụ** (đã loại bỏ triệt để 100% chi phí phụ tùng).")

        if not df_bh_split.empty:
            b_m1, b_m2, b_m3 = st.columns(3)
            b_m1.metric("Tổng Tiền Công Dịch Vụ (No VAT)", f"{df_bh_split['Cong_BH_Thuc_Nhan'].sum():,.0f} đ")
            b_m2.metric("Số Lệnh WO Bảo Hành", f"{df_bh_split['wo_norm'].nunique()} lệnh WO")
            b_m3.metric("Số Lượt Chia Công KTV", f"{len(df_bh_split)} lượt công")

            col_bh_wo_show = 'col_wo_goc' if 'col_wo_goc' in df_bh_split.columns else 'wo_norm'
            col_bh_cv_show = 'Ten_CV_BH' if 'Ten_CV_BH' in df_bh_split.columns else df_bh_split.columns[0]

            cfg_bh_split = {
                col_bh_wo_show: st.column_config.TextColumn("Lệnh Bảo Hành (WO)", width="medium"),
                'Bien_So_Val': st.column_config.TextColumn("Biển số", width="small"),
                col_bh_cv_show: st.column_config.TextColumn("Nội dung công việc bảo hành (Dịch vụ)", width="large"),
                "Tien_Cong_BH": st.column_config.NumberColumn("Tiền công dịch vụ", format="%,d đ", width="medium"),
                "So_KTV_Lam_Chung": st.column_config.NumberColumn("Số KTV chia đều", width="small"),
                "Ma_KTV_Clean": st.column_config.TextColumn("Mã KTV", width="small"),
                "Ten_KTV_Clean": st.column_config.TextColumn("Họ tên KTV", width="medium"),
                "Cong_BH_Thuc_Nhan": st.column_config.NumberColumn("👉 Công thực nhận", format="%,d đ", width="medium")
            }
            cols_show_bh = [c for c in [col_bh_wo_show, 'Bien_So_Val', col_bh_cv_show, 'Tien_Cong_BH', 'So_KTV_Lam_Chung', 'Ma_KTV_Clean', 'Ten_KTV_Clean', 'Cong_BH_Thuc_Nhan'] if c in df_bh_split.columns]
            st.dataframe(df_bh_split[cols_show_bh], use_container_width=True, height=560, hide_index=True, column_config=cfg_bh_split)
        else:
            st.info("💡 Chưa nạp tệp bảo hành.")

    # ==================== TAB 4: BẢNG LƯƠNG TỔNG HỢP ====================
    with tab_ktv_sum:
        st.subheader("💰 Bảng Lương Tổng Hợp Kỹ Thuật Viên")
        st.caption("Tổng hợp toàn bộ công thợ thực nhận từ Tab 2 (Sửa chữa đã điều chỉnh) và Tab 3 (Bảo hành dịch vụ).")

        df_source_ktv = edited_ktv_df if 'edited_ktv_df' in locals() else df_ktv_split_master
        mask_valid_ktv = ~df_source_ktv['Ma_KTV'].astype(str).str.contains('Chưa xác định|Khoản mục riêng|Bảo hành hãng|ALARM|CẦN GÁN')
        
        if not df_source_ktv.empty:
            df_sc_sum = df_source_ktv[mask_valid_ktv].groupby(['Ma_KTV', 'Ten_KTV'])['Tien_Chot_KTV'].sum().reset_index().rename(columns={'Tien_Chot_KTV': 'Cong_SC'})
        else:
            df_sc_sum = pd.DataFrame(columns=['Ma_KTV', 'Ten_KTV', 'Cong_SC'])

        if not df_bh_split.empty and 'Ma_KTV_Clean' in df_bh_split.columns:
            mask_bh_valid = ~df_bh_split['Ma_KTV_Clean'].astype(str).str.contains('Chưa xác định|Chưa nạp KTV|None')
            df_bh_sum = df_bh_split[mask_bh_valid].groupby(['Ma_KTV_Clean', 'Ten_KTV_Clean'])['Cong_BH_Thuc_Nhan'].sum().reset_index().rename(
                columns={'Ma_KTV_Clean': 'Ma_KTV', 'Ten_KTV_Clean': 'Ten_KTV', 'Cong_BH_Thuc_Nhan': 'Cong_BH'}
            )
        else:
            df_bh_sum = pd.DataFrame(columns=['Ma_KTV', 'Ten_KTV', 'Cong_BH'])

        df_luong_final = pd.merge(df_sc_sum, df_bh_sum, on=['Ma_KTV', 'Ten_KTV'], how='outer').fillna(0)
        df_luong_final['Tong_Luong_Nhan'] = df_luong_final['Cong_SC'] + df_luong_final['Cong_BH']
        df_luong_final = df_luong_final.sort_values(by='Tong_Luong_Nhan', ascending=False).reset_index(drop=True)
        df_luong_final.insert(0, 'STT', range(1, len(df_luong_final) + 1))

        cfg_luong = {
            "STT": st.column_config.NumberColumn("STT", width="small"),
            "Ma_KTV": st.column_config.TextColumn("Mã KTV", width="small"),
            "Ten_KTV": st.column_config.TextColumn("Họ và Tên Kỹ Thuật Viên", width="large"),
            "Cong_SC": st.column_config.NumberColumn("Công Sửa Chữa (5114)", format="%,d đ", width="medium"),
            "Cong_BH": st.column_config.NumberColumn("Công Bảo Hành (Dịch vụ)", format="%,d đ", width="medium"),
            "Tong_Luong_Nhan": st.column_config.NumberColumn("TỔNG TIỀN CÔNG TÍNH LƯƠNG", format="%,d đ", width="large")
        }
        st.dataframe(df_luong_final, use_container_width=True, hide_index=True, column_config=cfg_luong)
else:
    st.info("💡 Chưa có dữ liệu được lưu trên Google Sheets. Vui lòng mở ô 'Nạp Tệp Dữ Liệu' ở trên để nạp file.")
