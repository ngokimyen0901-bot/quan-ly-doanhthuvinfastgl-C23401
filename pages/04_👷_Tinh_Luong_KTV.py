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
st.caption("☁️ Quản lý đa kỳ lương | Tự động lưu trữ lịch sử hàng tháng | Lưu trữ: **Luong_KTV**.")

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

    col_loai_sp = next((c for c in df_out.columns if any(k in c.lower() for k in ['loại sản phẩm', 'material3'])), None)
    if col_loai_sp and df_out[col_loai_sp].dropna().astype(str).str.lower().str.contains('dịch vụ|service').any():
        mask_dv = df_out[col_loai_sp].astype(str).str.lower().str.contains('dịch vụ|service')
        df_out = df_out[mask_dv].copy()
        col_wo = next((c for c in df_out.columns if any(k in c.lower() for k in ['lệnh sửa chữa', 'lsc', 'wo'])), df_out.columns[0])
        col_cong = next((c for c in df_out.columns if any(k in c.lower() for k in ['tổng số tiền', 'amount', 'thành tiền', 'tiền công', 'tiền'])), df_out.columns[-1])
        col_cv = next((c for c in df_out.columns if any(k in c.lower() for k in ['mô tả', 'công việc', 'material2', 'material1'])), df_out.columns[0])
        col_ngay = next((c for c in df_out.columns if 'ngày' in c.lower() and ('phê duyệt' in c.lower() or 'approved' in c.lower())), None)
    else:
        col_wo = next((c for c in df_out.columns if any(k in c.lower() for k in ['lệnh sửa chữa', 'lsc', 'wo'])), df_out.columns[0])
        col_cong = next((c for c in df_out.columns if any(k in c.lower() for k in ['approved service', 'tiền công yêu cầu'])), None)
        if not col_cong:
            col_cong = next((c for c in df_out.columns if 'tiền công' in c.lower()), df_out.columns[-1])
        
        col_cv = None
        for c in df_out.columns:
            if 'tên công việc chính' in c.lower() or 'tên công việc' in c.lower():
                col_cv = c
                break
        if not col_cv:
            for c in df_out.columns:
                if 'mô tả' in c.lower():
                    col_cv = c
                    break
        if not col_cv:
            col_cv = df_out.columns[0]
            
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
        'Ky_Luong': 'Tháng 09/2026',
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

# Xử lý tính toán khi nạp file mới
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
        }).reset_index().rename(columns={'Tien_Cong_BH': 'Tien_Nha_May_Duyet'})

        df_ktv_bh_rows = df_ktv_for_bh[df_ktv_for_bh['key_match'].isin(bh_wo_sum['wo_norm']) & (df_ktv_for_bh['Ma_KTV_Clean'] != '')].copy()
        if df_ktv_bh_rows.empty:
            df_ktv_bh_rows = df_ktv_valid_works[df_ktv_valid_works['key_match'].isin(bh_wo_sum['wo_norm'])].copy()

        ktv_bh_tot_map = df_ktv_bh_rows.groupby('key_match')['Tien_KTV_Theo_Lenh'].sum().to_dict() if not df_ktv_bh_rows.empty else {}
        ktv_cnt_bh = df_ktv_bh_rows.groupby('key_match')['Ma_KTV_Clean'].nunique().to_dict() if not df_ktv_bh_rows.empty else {}

        df_ktv_bh_rows['So_KTV_Lam_Chung'] = df_ktv_bh_rows['key_match'].map(ktv_cnt_bh).fillna(1)
        df_ktv_bh_rows.loc[df_ktv_bh_rows['So_KTV_Lam_Chung'] == 0, 'So_KTV_Lam_Chung'] = 1

        df_bh_split = pd.merge(
            df_ktv_bh_rows[['key_match', 'So_RO_Val', 'Bien_So_Val', 'Ma_KTV_Clean', 'Ten_KTV_Clean', 'So_KTV_Lam_Chung']].drop_duplicates(subset=['key_match', 'Ma_KTV_Clean']),
            bh_wo_sum,
            left_on='key_match',
            right_on='wo_norm',
            how='right'
        )

        df_bh_split['Tien_KTV_He_Thong'] = df_bh_split['wo_norm'].map(ktv_bh_tot_map).fillna(0.0)
        df_bh_split['Chenh_Lech_Duyet'] = df_bh_split['Tien_Nha_May_Duyet'] - df_bh_split['Tien_KTV_He_Thong']

        def xac_dinh_canh_bao_bh(r):
            if pd.isna(r['Ma_KTV_Clean']) or str(r['Ma_KTV_Clean']).strip() in ['', 'None', 'nan']:
                return "🟡 CHƯA CÓ KTV TRÊN LỆNH"
            cl = abs(r['Chenh_Lech_Duyet'])
            if cl < 1000:
                return "🟢 Khớp duyệt nhà máy"
            return f"🔴 LỆCH DUYỆT (Coi lại: lệch {abs(r['Chenh_Lech_Duyet']):,.0f} đ)"

        df_bh_split['Canh_Bao_Duyet'] = df_bh_split.apply(xac_dinh_canh_bao_bh, axis=1)
        df_bh_split['Ma_KTV_Clean'] = df_bh_split['Ma_KTV_Clean'].fillna('Chưa có thợ')
        df_bh_split['Ten_KTV_Clean'] = df_bh_split['Ten_KTV_Clean'].fillna('Chưa có KTV')
        df_bh_split['So_KTV_Lam_Chung'] = df_bh_split['So_KTV_Lam_Chung'].fillna(1)
        df_bh_split['Cong_BH_Thuc_Nhan'] = df_bh_split['Tien_Nha_May_Duyet'] / df_bh_split['So_KTV_Lam_Chung']

    ktv_details_temp = []

    for k_match, group_orders in df_5114_grouped.groupby('key_match'):
        r_5114 = group_orders.iloc[0]
        tien_5114 = r_5114['Tien_5114']
        shd_k = r_5114['Số hóa đơn']
        nhd_k = r_5114['Ngày C.từ']
        ro_show = r_5114.get('Số R/O hãng', r_5114.get('Số R/O', k_match))
        if pd.isna(ro_show) or str(ro_show).strip() in ['', 'nan', 'None']:
            ro_show = str(r_5114.get('Diễn giải', ''))[:30]

        thang_hd_row = r_5114.get('Thang_HD', 'Tháng 09/2026')

        if k_match == "CUU_HO_2608":
            ktv_details_temp.append({
                'Ky_Luong': thang_hd_row,
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
                'Ky_Luong': thang_hd_row,
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
                            'Ky_Luong': thang_hd_row,
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
                    'Ky_Luong': thang_hd_row,
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
                'Ky_Luong': thang_hd_row,
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

    ktv_totals = df_ktv_split_master.groupby('key_match')['Tien_Chot_KTV'].sum().to_dict()

    def xac_dinh_trang_thai_row(r):
        k = r['key_match']
        if r['Ma_KTV'] in ['🚨 CẦN GÁN THỢ', 'Chưa xác định', '']:
            return "🟡 CẢNH BÁO: CHƯA CÓ KTV"
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

    # TỰ ĐỘNG GHI THẲNG VÀO SHEET Luong_KTV MỖI KHI NẠP TÍNH TOÁN
    try:
        df_auto_save = df_ktv_split_master.fillna("").copy()
        conn.update(worksheet="Luong_KTV", data=df_auto_save)
        st.toast("✅ Đã tự động sao lưu toàn bộ dữ liệu lên Google Sheets (sheet Luong_KTV)!", icon="☁️")
    except Exception as e_as:
        pass

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

# Nạp file bảo hành độc lập nếu có
if up_bh and df_bh_split.empty:
