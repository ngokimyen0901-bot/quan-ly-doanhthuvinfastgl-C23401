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
st.caption("☁️ Đối soát chuẩn theo **Tiền dịch vụ (Theo HD)** | Tách bạch Sửa chữa & Bảo hành | Lưu trữ: **Luong_KTV**.")

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
        if 'chi tiết' in s.lower() and 'đối soát' in s.lower():
            target_sheet = s
            break
        elif 'claim' in s.lower() or 'active' in s.lower():
            target_sheet = s
            break
            
    df_raw = pd.read_excel(xls, sheet_name=target_sheet)
    df_raw.columns = [str(c).strip() for c in df_raw.columns]
    
    if 'Material3' in df_raw.columns and 'Amount' in df_raw.columns and 'WO' in df_raw.columns:
        df_serv = df_raw[df_raw['Material3'].astype(str).str.lower().str.contains('dịch vụ|service')].copy()
        col_wo = 'WO'
        col_cong = 'Amount'
        col_cv = 'Material2' if 'Material2' in df_serv.columns else 'Material1'
        col_ngay = 'WC- Approved' if 'WC- Approved' in df_serv.columns else None
    else:
        col_wo = next((c for c in df_raw.columns if any(k in c.lower() for k in ['lệnh sửa chữa', 'lsc', 'wo'])), df_raw.columns[0])
        col_cong = next((c for c in df_raw.columns if any(k in c.lower() for k in ['approved service', 'tiền công yêu cầu', 'tiền công'])), None)
        if not col_cong:
            col_cong = next((c for c in df_raw.columns if 'tiền' in c.lower()), df_raw.columns[-1])
        col_cv = next((c for c in df_raw.columns if any(k in c.lower() for k in ['tên công việc chính', 'công việc chính', 'mô tả'])), df_raw.columns[0])
        col_ngay = next((c for c in df_raw.columns if 'ngày' in c.lower() and ('phê duyệt' in c.lower() or 'đxbh' in c.lower())), None)
        df_serv = df_raw.copy()

    df_serv['wo_norm'] = df_serv[col_wo].apply(norm_wo_key)
    df_serv['Tien_Cong_BH'] = df_serv[col_cong].apply(clean_num)
    df_serv['Ten_CV_BH'] = df_serv[col_cv].fillna('').astype(str) if col_cv else ''
    
    if col_ngay and col_ngay in df_serv.columns:
        df_serv['Thang_BH'] = pd.to_datetime(df_serv[col_ngay], errors='coerce', dayfirst=True).dt.strftime('Tháng %m/%Y').fillna('Chưa rõ')
    else:
        df_serv['Thang_BH'] = 'Toàn bộ kỳ'
        
    df_serv['col_wo_goc'] = df_serv[col_wo]
    return df_serv[(df_serv['Tien_Cong_BH'] > 0) & (df_serv['wo_norm'] != '')].copy()

def trich_xuat_ktv_dataframe(df_source):
    df_res = df_source.copy()
    col_ro_h = next((c for c in df_res.columns if 'số ro hãng' in c.lower() or 'ro hãng' in c.lower()), None)
    col_ro_nb = next((c for c in df_res.columns if c.strip().lower() in ['số ro', 'ro']), None)
    col_ma = next((c for c in df_res.columns if 'mã ktv' in c.lower()), None)
    col_ten = next((c for c in df_res.columns if 'tên ktv' in c.lower()), None)
    
    # 3 CỘT TIỀN TRONG BÁO CÁO KTV
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
    
    # TIỀN THEO HÓA ĐƠN vs TIỀN THEO LỆNH
    df_res['Tien_KTV_Theo_HD'] = df_res[col_tien_hd].apply(clean_num) if col_tien_hd else 0.0
    df_res['Tien_KTV_Theo_Lenh'] = df_res[col_tien_lenh].apply(clean_num) if col_tien_lenh else (df_res[col_thanh_tien].apply(clean_num) if col_thanh_tien else 0.0)

    df_res['Bien_So_Val'] = df_res[col_bs].fillna('').astype(str).str.strip() if col_bs else ""
    df_res['Hang_Muc_Val'] = df_res[col_hm].fillna('').astype(str).str.strip() if col_hm else ""
    df_res['Noi_Dung_Val'] = df_res[col_nd].fillna('').astype(str).str.strip() if col_nd else ""
    df_res['Ma_CV_Val'] = df_res[col_cv].fillna('').astype(str).str.strip() if col_cv else ""
    df_res['So_RO_Val'] = df_res[col_ro_h].fillna(df_res[col_ro_nb] if col_ro_nb else '').astype(str).str.strip() if col_ro_h else ""

    return df_res

@st.cache_data(ttl=30)
def load_saved_luong():
    try:
        df = conn.read(worksheet="Luong_KTV", ttl=30)
        if df is not None and not df.empty:
            df.columns = [str(c).strip() for c in df.columns]
            return df
    except Exception:
        pass
    return pd.DataFrame()

df_saved_gs = load_saved_luong()

# ----------------- KHU VỰC NẠP TỆP -----------------
st.markdown("### 📥 Nạp Tệp Dữ Liệu Tính Lương")
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

    # 1. Danh sách PDI
    pdi_keys_set = set()
    if up_pdi:
        df_pdi_input = doc_file_excel_chung(up_pdi, tu_khoa=['lệnh', 'wo', 'ro', 'pdi'])
        col_pdi_ro = next((c for c in df_pdi_input.columns if any(k in c.lower() for k in ['lệnh', 'wo', 'ro', 'lsc'])), df_pdi_input.columns[0])
        for v in df_pdi_input[col_pdi_ro].dropna():
            pdi_keys_set.add(norm_wo_key(v))

    # 2. Xử lý TOÀN BỘ doanh thu 5114
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

    # TỔNG TIỀN KTV ĐỐI SOÁT: LẤY CHUẨN THEO CỘT `Tien_KTV_Theo_HD`
    ktv_per_ro_sc = df_ktv_for_sc[df_ktv_for_sc['key_match'] != ''].groupby('key_match').agg({
        'So_RO_Val': 'first',
        'Bien_So_Val': 'first',
        'Tien_KTV_Theo_HD': 'sum',
        'Tien_KTV_Theo_Lenh': 'sum',
        'Ma_KTV_Clean': lambda x: ", ".join(sorted(set([str(v) for v in x if str(v).strip() and str(v) != 'nan']))),
        'Ten_KTV_Clean': lambda x: ", ".join(sorted(set([str(v) for v in x if str(v).strip() and str(v) != 'nan']))),
        'Noi_Dung_Val': lambda x: " | ".join(sorted(set([str(v) for v in x if str(v).strip() and str(v) != 'nan'])))[:150]
    }).reset_index()

    ktv_per_ro_sc['Tien_KTV'] = np.where(ktv_per_ro_sc['Tien_KTV_Theo_HD'] > 0, ktv_per_ro_sc['Tien_KTV_Theo_HD'], ktv_per_ro_sc['Tien_KTV_Theo_Lenh'])

    # 3. Xử lý File Bảo Hành (Active Warranty / BangKe) - ĐỌC SIÊU THÔNG MINH
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
            df_ktv_bh_rows = df_ktv_works_all[df_ktv_works_all['key_match'].isin(bh_wo_sum['wo_norm'])].copy()

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

    def phan_loai_lenh(r):
        k = r['key_match']
        if k == "CUU_HO_2608":
            return "CUU_HO"
        if k == "BAO_HANH_BANG_KE_2608":
            return "BAO_HANH_TONG"
        if k in pdi_keys_set:
            return "PDI"
        if k in bh_wo_set:
            return "BAO_HANH"
        return "SUA_CHUA_5114"

    cols_5114_sub = ['key_match', 'Thang_HD', 'Tien_5114']
    if col_5114_ro_hang in df_5114_grouped.columns: cols_5114_sub.append(col_5114_ro_hang)
    if col_5114_ro_nb in df_5114_grouped.columns: cols_5114_sub.append(col_5114_ro_nb)
    if col_5114_shd in df_5114_grouped.columns: cols_5114_sub.append(col_5114_shd)
    if col_5114_ngay in df_5114_grouped.columns: cols_5114_sub.append(col_5114_ngay)
    if col_5114_dg in df_5114_grouped.columns: cols_5114_sub.append(col_5114_dg)

    df_merged = pd.merge(
        df_5114_grouped[cols_5114_sub],
        ktv_per_ro_sc,
        on='key_match',
        how='left'
    )
    df_merged['Tien_KTV'] = df_merged['Tien_KTV'].fillna(0)
    df_merged['Chenh_Lech'] = df_merged['Tien_5114'] - df_merged['Tien_KTV']
    df_merged['Nhom_Phan_Loai'] = df_merged.apply(phan_loai_lenh, axis=1)

    def danh_gia_trang_thai(r):
        k = r['key_match']
        if k == "CUU_HO_2608":
            return "🚚 Cứu hộ (Chi phí kéo xe)"
        if k == "BAO_HANH_BANG_KE_2608":
            return "🛡️ Bảng kê BH hãng (Tổng kỳ)"
        cl = abs(r['Chenh_Lech'])
        dg = str(r.get(col_5114_dg, '')).lower()
        if cl < 1000:
            return "🟢 Khớp 100%"
        elif "khấu trừ" in dg or "bảo hiểm" in dg:
            return "🔴 Lệch (BH khấu trừ)"
        elif "thuê ngoài" in dg:
            return "🔴 Lệch (Thuê ngoài)"
        return "🔴 Lệch (Cần rà soát)"

    df_merged['Ghi_Chu_Doi_Soat'] = df_merged.apply(danh_gia_trang_thai, axis=1)
    df_merged['Tien_Cong_Chot_Luong'] = df_merged['Tien_5114']
    df_merged['Ghi_Chu_Thu_Cong'] = ""

    # Khôi phục dữ liệu đã lưu
    if not df_saved_gs.empty and 'key_match' in df_saved_gs.columns:
        map_tien_chot = df_saved_gs.set_index('key_match')['Tien_Cong_Chot_Luong'].to_dict() if 'Tien_Cong_Chot_Luong' in df_saved_gs.columns else {}
        map_note_tc = df_saved_gs.set_index('key_match')['Ghi_Chu_Thu_Cong'].to_dict() if 'Ghi_Chu_Thu_Cong' in df_saved_gs.columns else {}
        
        for idx_r, r_val in df_merged.iterrows():
            k_w = r_val['key_match']
            if k_w in map_tien_chot:
                df_merged.at[idx_r, 'Tien_Cong_Chot_Luong'] = clean_num(map_tien_chot[k_w])
            if k_w in map_note_tc:
                df_merged.at[idx_r, 'Ghi_Chu_Thu_Cong'] = str(map_note_tc[k_w])

    st.markdown("---")
    all_months = sorted(list(set([str(m) for m in df_merged['Thang_HD'].unique() if str(m) not in ['', 'nan', 'None']])))
    
    fl_col1, fl_col2 = st.columns([4, 6])
    with fl_col1:
        sel_thang_nam = st.selectbox("📅 LỰA CHỌN THỜI GIAN THEO THÁNG / NĂM (NGÀY XUẤT HĐ):", ["Tất cả các tháng"] + all_months, index=0)

    df_merged_filtered = df_merged.copy()
    if sel_thang_nam != "Tất cả các tháng":
        df_merged_filtered = df_merged_filtered[df_merged_filtered['Thang_HD'] == sel_thang_nam]

    df_sc_remaining = df_merged_filtered[df_merged_filtered['Nhom_Phan_Loai'].isin(['SUA_CHUA_5114', 'CUU_HO', 'BAO_HANH_TONG'])].copy()
    df_pdi_only = df_merged_filtered[df_merged_filtered['Nhom_Phan_Loai'] == 'PDI'].copy()

    # =========================================================================
    # LOGIC CHIA TIỀN CÔNG THỢ CHUẨN XÁC:
    # 1. Với Sửa chữa: BẮT BUỘC chỉ tính các dòng có `Tien_KTV_Theo_HD > 0`!
    #    (Dòng nào trên HĐ = 0 đ như CNPM bảo hành thì BỎ RA KHỎI bảng sửa chữa 5114)
    # 2. Công việc nào 1 thợ làm -> Thợ nhận trọn 100% tiền HĐ của công việc đó!
    #    Chỉ khi cùng 1 công việc có nhiều thợ cùng đứng tên mới chia đều!
    # =========================================================================
    ktv_records_split = []

    for _, r_sc in df_sc_remaining.iterrows():
        k_match = r_sc['key_match']
        tien_lenh_chot = clean_num(r_sc['Tien_Cong_Chot_Luong'])
        shd_k = r_sc.get(col_5114_shd, '')
        nhd_k = r_sc.get(col_5114_ngay, '')
        ro_show = r_sc.get(col_5114_ro_hang, r_sc.get(col_5114_ro_nb, k_match))
        if pd.isna(ro_show) or str(ro_show).strip() in ['', 'nan', 'None']:
            ro_show = str(r_sc.get(col_5114_dg, ''))[:30]
        bs_k = r_sc.get('Bien_So_Val', '')

        if k_match == "CUU_HO_2608":
            ktv_records_split.append({
                'Loai_Cong': 'Cứu hộ',
                'So_RO': 'Chi phí cứu hộ',
                'So_HD': shd_k,
                'Ngay_Xuat_HD': nhd_k,
                'Bien_So': '',
                'Hang_Muc': 'Cứu hộ kéo xe',
                'Noi_Dung_CV': 'Chi phí cứu hộ theo báo cáo chi tiết',
                'Ma_KTV': 'Khoản mục riêng',
                'Ten_KTV': 'Cứu hộ (Không tính công KTV)',
                'Tien_Thuc_Nhan': tien_lenh_chot
            })
            continue

        if k_match == "BAO_HANH_BANG_KE_2608":
            ktv_records_split.append({
                'Loai_Cong': 'Bảo hành cục bộ',
                'So_RO': 'Bảng kê bảo hành hãng',
                'So_HD': shd_k,
                'Ngay_Xuat_HD': nhd_k,
                'Bien_So': '',
                'Hang_Muc': 'Bảo hành hãng',
                'Noi_Dung_CV': 'Chi phí bảo hành tổng kỳ theo báo cáo chi tiết',
                'Ma_KTV': 'Bảo hành hãng',
                'Ten_KTV': 'Chi tiết xem tại Tab Bảo Hành',
                'Tien_Thuc_Nhan': tien_lenh_chot
            })
            continue

        sub_works_all = df_ktv_for_sc[df_ktv_for_sc['key_match'] == k_match].copy()

        # LỌC CHUẨN XÁC: Chỉ lấy các dòng có tiền theo Hóa Đơn > 0!
        # (Loại bỏ các dòng CNPM hoặc bảo hành có tiền HĐ = 0 đ)
        sub_works = sub_works_all[sub_works_all['Tien_KTV_Theo_HD'] > 0].copy()

        if sub_works.empty:
            sub_works = sub_works_all[sub_works_all['Tien_KTV_Val'] > 0].copy()

        if not sub_works.empty:
            sub_works['task_id'] = sub_works['Ma_CV_Val'].astype(str) + "_" + sub_works['Noi_Dung_Val'].astype(str)
            sum_ktv_tasks = sub_works['Tien_KTV_Theo_HD'].sum() if sub_works['Tien_KTV_Theo_HD'].sum() > 0 else sub_works['Tien_KTV_Val'].sum()
            
            he_so = (tien_lenh_chot / sum_ktv_tasks) if (sum_ktv_tasks > 0 and abs(tien_lenh_chot - sum_ktv_tasks) > 1000) else 1.0

            for task, group_task in sub_works.groupby('task_id'):
                n_tho_lam_cv_nay = group_task['Ma_KTV_Clean'].nunique()
                tien_goc_cv = group_task['Tien_KTV_Theo_HD'].iloc[0] if group_task['Tien_KTV_Theo_HD'].iloc[0] > 0 else group_task['Tien_KTV_Val'].iloc[0]
                tien_cong_cv = tien_goc_cv * he_so
                tien_moi_tho = tien_cong_cv / max(n_tho_lam_cv_nay, 1)

                hm_ten = group_task['Hang_Muc_Val'].iloc[0] if group_task['Hang_Muc_Val'].iloc[0] else "Sửa chữa"
                cv_ten = group_task['Noi_Dung_Val'].iloc[0] if group_task['Noi_Dung_Val'].iloc[0] else group_task['Ma_CV_Val'].iloc[0]

                for _, r_tho in group_task.drop_duplicates(subset=['Ma_KTV_Clean']).iterrows():
                    ktv_records_split.append({
                        'Loai_Cong': 'Sửa chữa (5114)',
                        'So_RO': ro_show,
                        'So_HD': shd_k,
                        'Ngay_Xuat_HD': nhd_k,
                        'Bien_So': bs_k,
                        'Hang_Muc': hm_ten,
                        'Noi_Dung_CV': cv_ten,
                        'Ma_KTV': r_tho['Ma_KTV_Clean'],
                        'Ten_KTV': r_tho['Ten_KTV_Clean'],
                        'Tien_Thuc_Nhan': tien_moi_tho
                    })
        else:
            ktv_records_split.append({
                'Loai_Cong': 'Sửa chữa (5114)',
                'So_RO': ro_show,
                'So_HD': shd_k,
                'Ngay_Xuat_HD': nhd_k,
                'Bien_So': bs_k,
                'Hang_Muc': 'Chưa xác định',
                'Noi_Dung_CV': str(r_sc.get(col_5114_dg, 'Chưa có thông tin công việc'))[:100],
                'Ma_KTV': 'Chưa xác định',
                'Ten_KTV': 'Chưa có thông tin thợ',
                'Tien_Thuc_Nhan': tien_lenh_chot
            })

    df_ktv_detail_sheet = pd.DataFrame(ktv_records_split)

    tab_sc, tab_split_ktv, tab_bh, tab_pdi, tab_ktv_sum = st.tabs([
        "🔧 1. Đối Soát Lệnh Sửa Chữa (Toàn Bộ 5114)",
        "👷 2. Chi Tiết Tính Lương Từng KTV (Sheet Riêng)",
        "🛡️ 3. Lệnh Bảo Hành Trong Tháng (Chia Đều)",
        "🚗 4. Lệnh PDI (Tách Riêng)",
        "💰 5. Bảng Tổng Hợp Lương KTV"
    ])

    # TAB 1: ĐỐI SOÁT 5114
    with tab_sc:
        st.subheader("🔧 Bảng Đối Soát 5114 Đầy Đủ Doanh Thu (Đã Gom Dòng Theo RO)")
        st.info("💡 **Ghi nhận 100% doanh thu:** Bao gồm các lệnh có WO hãng, lệnh chỉ có số RO 01.S, khoản Cứu hộ và Bảng kê bảo hành hãng.")

        cnt_khop = (df_sc_remaining['Ghi_Chu_Doi_Soat'] == "🟢 Khớp 100%").sum()
        cnt_lech = len(df_sc_remaining) - cnt_khop
        
        m1, m2, m3 = st.columns(3)
        m1.metric("Tổng Doanh Thu 5114", f"{df_sc_remaining['Tien_5114'].sum():,.0f} đ", f"{len(df_sc_remaining)} dòng/lệnh")
        m2.metric("Lệnh Khớp Chuẩn (Tự động)", f"{cnt_khop} lệnh")
        m3.metric("Lệnh Cần Lưu Ý / Bạn Quyết Định", f"{cnt_lech} lệnh", delta=f"{cnt_lech} ca" if cnt_lech > 0 else "0", delta_color="inverse")

        f_c1, f_c2 = st.columns([4, 6])
        with f_c1:
            loc_xem = st.selectbox("🔍 Lọc hiển thị:", ["Tất cả lệnh sửa chữa", "🔴 Chỉ xem các lệnh LỆCH / CỨU HỘ", "🟢 Chỉ xem các lệnh KHỚP 100%"])
        
        df_display_sc = df_sc_remaining.copy()
        if "Chỉ xem các lệnh LỆCH" in loc_xem:
            df_display_sc = df_display_sc[df_display_sc['Ghi_Chu_Doi_Soat'] != "🟢 Khớp 100%"]
        elif "Chỉ xem các lệnh KHỚP" in loc_xem:
            df_display_sc = df_display_sc[df_display_sc['Ghi_Chu_Doi_Soat'] == "🟢 Khớp 100%"]

        col_show_ro = col_5114_ro_hang if col_5114_ro_hang in df_display_sc.columns else col_5114_ro_nb
        
        cols_edit = [
            'Thang_HD', col_show_ro, col_5114_ro_nb, col_5114_ngay, col_5114_shd, 
            'Tien_5114', 'Tien_KTV', 'Chenh_Lech', 'Tien_Cong_Chot_Luong',
            'Ghi_Chu_Doi_Soat', 'Ma_KTV_Clean', 'Ten_KTV_Clean', 'Bien_So_Val',
            'Ghi_Chu_Thu_Cong', 'key_match'
        ]
        cols_valid = [c for c in cols_edit if c in df_display_sc.columns]

        cfg_edit = {
            "Thang_HD": st.column_config.TextColumn("Tháng HĐ", disabled=True, width="small"),
            col_show_ro: st.column_config.TextColumn("Số Lệnh (RO hãng)", disabled=True, width="medium"),
            col_5114_ro_nb: st.column_config.TextColumn("Số RO nội bộ (01.S)", disabled=True, width="medium"),
            col_5114_ngay: st.column_config.TextColumn("Ngày xuất HĐ", disabled=True, width="small"),
            col_5114_shd: st.column_config.TextColumn("Số HĐ", disabled=True, width="small"),
            "Tien_5114": st.column_config.NumberColumn("Tiền 5114", format="%,d đ", disabled=True, width="medium"),
            "Tien_KTV": st.column_config.NumberColumn("Tiền KTV (HĐ)", format="%,d đ", disabled=True, width="medium"),
            "Chenh_Lech": st.column_config.NumberColumn("Chênh lệch", format="%,d đ", disabled=True, width="small"),
            "Tien_Cong_Chot_Luong": st.column_config.NumberColumn(
                "👉 SỐ TIỀN BẠN ĐIỀU CHỈNH (Tính công thợ)",
                format="%,d đ",
                disabled=False,
                width="large",
                help="Sửa trực tiếp con số này nếu bạn muốn chốt số tiền khác cho thợ!"
            ),
            "Ghi_Chu_Doi_Soat": st.column_config.TextColumn("Trạng thái", disabled=True, width="medium"),
            "Ma_KTV_Clean": st.column_config.TextColumn("Mã KTV", disabled=True, width="small"),
            "Ten_KTV_Clean": st.column_config.TextColumn("KTV làm lệnh", disabled=True, width="large"),
            "Bien_So_Val": st.column_config.TextColumn("Biển số", disabled=True, width="small"),
            "Ghi_Chu_Thu_Cong": st.column_config.TextColumn("Ghi chú thủ công", disabled=False, width="medium"),
            "key_match": st.column_config.TextColumn("Mã chuẩn", disabled=True, width="small")
        }

        edited_sc_df = st.data_editor(
            df_display_sc[cols_valid],
            use_container_width=True,
            height=560,
            column_config=cfg_edit,
            hide_index=True,
            key="editor_doi_soat_v7"
        )

        c_save1, c_save2 = st.columns([3.5, 6.5])
        with c_save1:
            if st.button("☁️ Lưu Bảng Đối Soát 5114 Lên Google Sheets", type="primary", use_container_width=True):
                with st.spinner("⏳ Đang lưu vào sheet Luong_KTV..."):
                    dict_updated = edited_sc_df.set_index('key_match').to_dict('index')
                    for idx_r, r_val in df_merged.iterrows():
                        k_w = r_val['key_match']
                        if k_w in dict_updated:
                            df_merged.at[idx_r, 'Tien_Cong_Chot_Luong'] = dict_updated[k_w].get('Tien_Cong_Chot_Luong', df_merged.at[idx_r, 'Tien_Cong_Chot_Luong'])
                            df_merged.at[idx_r, 'Ghi_Chu_Thu_Cong'] = dict_updated[k_w].get('Ghi_Chu_Thu_Cong', df_merged.at[idx_r, 'Ghi_Chu_Thu_Cong'])
                    try:
                        df_to_save_clean = df_merged.fillna("").copy()
                        conn.update(worksheet="Luong_KTV", data=df_to_save_clean)
                        st.success("✅ Đã lưu bảng đối soát lên Google Sheets thành công!")
                        st.rerun()
                    except Exception as e_s:
                        st.error(f"Lỗi khi lưu: {e_s}")

    # TAB 2: SHEET RIÊNG CHI TIẾT TÍNH LƯƠNG TỪNG KTV (CHỈ TÍNH TIỀN CÔNG XUẤT HÓA ĐƠN)
    with tab_split_ktv:
        st.subheader("👷 Chi Tiết Phân Bổ Tiền Công Theo Hóa Đơn Cho KTV")
        st.info("💡 **Tách bạch chuẩn xác:** Chỉ tính tiền công cho các hạng mục có xuất hóa đơn (Ví dụ: Bảo dưỡng 300k cho thợ 1014). Các hạng mục bảo hành 0đ trên HĐ (như CNPM) tự động tách sang Tab Bảo Hành!")

        if not df_ktv_detail_sheet.empty:
            s_m1, s_m2, s_m3 = st.columns(3)
            s_m1.metric("Tổng Tiền Công Được Chia", f"{df_ktv_detail_sheet['Tien_Thuc_Nhan'].sum():,.0f} đ")
            s_m2.metric("Tổng Số Dòng Công Việc", f"{len(df_ktv_detail_sheet):,} dòng")
            s_m3.metric("Số KTV Tham Gia", f"{df_ktv_detail_sheet['Ma_KTV'].nunique()} KTV")

            cfg_detail_sheet = {
                "Loai_Cong": st.column_config.TextColumn("Phân loại", width="small"),
                "So_RO": st.column_config.TextColumn("Mã Lệnh SC", width="medium"),
                "So_HD": st.column_config.TextColumn("Số HĐ", width="small"),
                "Ngay_Xuat_HD": st.column_config.TextColumn("Ngày HĐ", width="small"),
                "Bien_So": st.column_config.TextColumn("Biển số", width="small"),
                "Hang_Muc": st.column_config.TextColumn("Hạng mục", width="medium"),
                "Noi_Dung_CV": st.column_config.TextColumn("👉 NỘI DUNG CÔNG VIỆC CHI TIẾT", width="large"),
                "Ma_KTV": st.column_config.TextColumn("Mã KTV", width="small"),
                "Ten_KTV": st.column_config.TextColumn("Họ tên KTV", width="medium"),
                "Tien_Thuc_Nhan": st.column_config.NumberColumn("👉 TIỀN KTV THỰC NHẬN", format="%,d đ", width="medium")
            }

            st.dataframe(
                df_ktv_detail_sheet,
                use_container_width=True,
                height=560,
                hide_index=True,
                column_config=cfg_detail_sheet
            )

            c_btn_l1, c_btn_l2 = st.columns([4, 6])
            with c_btn_l1:
                if st.button("☁️ LƯU BẢNG CHI TIẾT LƯƠNG KTV NÀY LÊN GOOGLE SHEETS", type="primary", use_container_width=True):
                    with st.spinner("⏳ Đang lưu chi tiết vào sheet Luong_KTV..."):
                        try:
                            df_save_detail = df_ktv_detail_sheet.fillna("").copy()
                            conn.update(worksheet="Luong_KTV", data=df_save_detail)
                            st.success("✅ ĐÃ LƯU THÀNH CÔNG VÀO SHEET Luong_KTV! Mở Google Sheets bạn sẽ thấy đầy đủ dữ liệu ngay!")
                            st.rerun()
                        except Exception as e_gs:
                            st.error(f"Lỗi khi lưu lên Google Sheets: {e_gs}")

            with c_btn_l2:
                buffer_ktv = io.BytesIO()
                with pd.ExcelWriter(buffer_ktv, engine='openpyxl') as writer:
                    df_ktv_detail_sheet.to_excel(writer, index=False, sheet_name="Chi_Tiet_Luong_KTV")
                st.download_button(
                    label="⬇️ Tải File Excel Chi Tiết Phân Bổ Công KTV",
                    data=buffer_ktv.getvalue(),
                    file_name=f"Chi_Tiet_Luong_KTV_{sel_thang_nam.replace('/', '_')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )

    # TAB 3: BẢO HÀNH (KHÔNG BỊ LỌC MẤT DỮ LIỆU)
    with tab_bh:
        st.subheader("🛡️ Chi Tiết Tiền Công Bảo Hành Phân Bổ Cho KTV")
        st.caption("Chỉ tính **Tiền công dịch vụ** (loại bỏ chi phí phụ tùng). Tra cứu thợ từ file KTV theo Lệnh và chia đều nếu làm chung.")

        if up_bh and not df_bh_split.empty:
            available_bh_months = sorted(list(set([str(m) for m in df_bh_split['Thang_BH'].unique() if str(m) not in ['', 'nan', 'None']])))
            
            c_f_bh1, c_f_bh2 = st.columns([4, 6])
            with c_f_bh1:
                sel_bh_filter = st.selectbox(
                    "📅 Lọc theo tháng phê duyệt bảo hành:", 
                    ["Hiển thị toàn bộ tệp bảo hành đã nạp (Khuyên dùng)"] + available_bh_months,
                    index=0
                )

            df_bh_view = df_bh_split.copy()
            if sel_bh_filter != "Hiển thị toàn bộ tệp bảo hành đã nạp (Khuyên dùng)":
                df_bh_view = df_bh_view[df_bh_view['Thang_BH'] == sel_bh_filter]

            b_m1, b_m2, b_m3 = st.columns(3)
            b_m1.metric("Tổng Tiền Công Bảo Hành (No VAT)", f"{df_bh_view['Cong_BH_Thuc_Nhan'].sum():,.0f} đ")
            b_m2.metric("Số Lệnh WO Bảo Hành", f"{df_bh_view['wo_norm'].nunique()} lệnh WO")
            b_m3.metric("Số Lượt Chia Công KTV", f"{len(df_bh_view)} lượt công")

            col_bh_wo_show = 'col_wo_goc' if 'col_wo_goc' in df_bh_view.columns else 'wo_norm'
            col_bh_cv_show = 'Ten_CV_BH' if 'Ten_CV_BH' in df_bh_view.columns else df_bh_view.columns[0]

            cfg_bh_split = {
                col_bh_wo_show: st.column_config.TextColumn("Lệnh Bảo Hành (WO)", width="medium"),
                'Bien_So_Val': st.column_config.TextColumn("Biển số", width="small"),
                col_bh_cv_show: st.column_config.TextColumn("Nội dung công việc bảo hành", width="large"),
                "Tien_Cong_BH": st.column_config.NumberColumn("Tổng tiền công", format="%,d đ", width="medium"),
                "So_KTV_Lam_Chung": st.column_config.NumberColumn("Số KTV chia đều", width="small"),
                "Ma_KTV_Clean": st.column_config.TextColumn("Mã KTV", width="small"),
                "Ten_KTV_Clean": st.column_config.TextColumn("Họ tên KTV", width="medium"),
                "Cong_BH_Thuc_Nhan": st.column_config.NumberColumn("👉 Công thực nhận", format="%,d đ", width="medium")
            }
            cols_show_bh = [c for c in [col_bh_wo_show, 'Bien_So_Val', col_bh_cv_show, 'Tien_Cong_BH', 'So_KTV_Lam_Chung', 'Ma_KTV_Clean', 'Ten_KTV_Clean', 'Cong_BH_Thuc_Nhan'] if c in df_bh_view.columns]
            st.dataframe(df_bh_view[cols_show_bh], use_container_width=True, height=560, hide_index=True, column_config=cfg_bh_split)
        else:
            st.warning("⚠️️ Vui lòng nạp file '4. Chi tiết bảo hành (Active Claim / Bảng Kê WCS)' ở trên để tính toán.")

    # TAB 4: PDI
    with tab_pdi:
        st.subheader(f"🚗 Danh Sách Lệnh PDI Được Tách Riêng ({sel_thang_nam})")
        st.caption("Các lệnh thuộc danh sách PDI được tách riêng hoàn toàn, không tính lẫn vào công sửa chữa thông thường.")

        if not df_pdi_only.empty:
            p_m1, p_m2 = st.columns(2)
            p_m1.metric("Tổng Tiền Công PDI", f"{df_pdi_only['Tien_KTV'].sum():,.0f} đ", f"{len(df_pdi_only)} lệnh")
            p_m2.metric("Số Lượt KTV Làm PDI", f"{len(df_pdi_only)} lượt")

            cols_pdi_show = ['Thang_HD', col_show_ro, 'Tien_KTV', 'Ma_KTV_Clean', 'Ten_KTV_Clean', 'Bien_So_Val', 'Noi_Dung_Val']
            cols_pdi_valid = [c for c in cols_pdi_show if c in df_pdi_only.columns]
            cfg_pdi = {
                "Thang_HD": st.column_config.TextColumn("Tháng xuất HĐ", width="small"),
                col_show_ro: st.column_config.TextColumn("Số Lệnh PDI", width="medium"),
                "Tien_KTV": st.column_config.NumberColumn("Tiền công PDI", format="%,d đ", width="medium"),
                "Ma_KTV_Clean": st.column_config.TextColumn("Mã KTV", width="small"),
                "Ten_KTV_Clean": st.column_config.TextColumn("Tên KTV", width="medium"),
                "Bien_So_Val": st.column_config.TextColumn("Biển số / Mã PDI", width="small"),
                "Noi_Dung_Val": st.column_config.TextColumn("Nội dung công việc", width="large")
            }
            st.dataframe(df_pdi_only[cols_pdi_valid], use_container_width=True, hide_index=True, column_config=cfg_pdi)
        else:
            st.info("💡 Chưa có lệnh PDI nào trong tháng này (hoặc chưa nạp danh sách PDI ở ô số 5).")

    # TAB 5: BẢNG LƯƠNG TỔNG HỢP
    with tab_ktv_sum:
        st.subheader(f"💰 Bảng Lương Tổng Hợp Kỹ Thuật Viên ({sel_thang_nam})")
        st.caption("Tổng hợp toàn bộ công thợ thực nhận từ Tab 2 (Sửa chữa đã xét từng công việc), Tab 3 (Bảo hành), và Tab 4 (PDI).")

        if not df_ktv_detail_sheet.empty:
            mask_valid_ktv = ~df_ktv_detail_sheet['Ma_KTV'].isin(['Chưa xác định', 'Khoản mục riêng', 'Bảo hành hãng'])
            df_sc_sum = df_ktv_detail_sheet[mask_valid_ktv].groupby(['Ma_KTV', 'Ten_KTV'])['Tien_Thuc_Nhan'].sum().reset_index().rename(columns={'Tien_Thuc_Nhan': 'Cong_SC'})
        else:
            df_sc_sum = pd.DataFrame(columns=['Ma_KTV', 'Ten_KTV', 'Cong_SC'])

        if up_bh and not df_bh_split.empty:
            df_bh_target = df_bh_split.copy()
            if sel_thang_nam != "Tất cả các tháng" and 'Thang_BH' in df_bh_target.columns:
                df_bh_target = df_bh_target[df_bh_target['Thang_BH'] == sel_thang_nam]
            df_bh_sum = df_bh_target.groupby(['Ma_KTV_Clean', 'Ten_KTV_Clean'])['Cong_BH_Thuc_Nhan'].sum().reset_index().rename(
                columns={'Ma_KTV_Clean': 'Ma_KTV', 'Ten_KTV_Clean': 'Ten_KTV', 'Cong_BH_Thuc_Nhan': 'Cong_BH'}
            )
        else:
            df_bh_sum = pd.DataFrame(columns=['Ma_KTV', 'Ten_KTV', 'Cong_BH'])

        ktv_pdi_list = []
        for _, r_p in df_pdi_only.iterrows():
            ma_ktvs = [x.strip() for x in str(r_p['Ma_KTV_Clean']).split(',') if x.strip() and x.strip() != 'nan']
            ten_ktvs = [x.strip() for x in str(r_p['Ten_KTV_Clean']).split(',') if x.strip() and x.strip() != 'nan']
            so_nguoi = max(len(ma_ktvs), 1)
            tien_chia = clean_num(r_p['Tien_KTV']) / so_nguoi
            for idx_k, m_k in enumerate(ma_ktvs):
                t_k = ten_ktvs[idx_k] if idx_k < len(ten_ktvs) else m_k
                ktv_pdi_list.append({'Ma_KTV': m_k, 'Ten_KTV': t_k, 'Cong_PDI': tien_chia})

        df_pdi_sum = pd.DataFrame(ktv_pdi_list).groupby(['Ma_KTV', 'Ten_KTV'])['Cong_PDI'].sum().reset_index() if ktv_pdi_list else pd.DataFrame(columns=['Ma_KTV', 'Ten_KTV', 'Cong_PDI'])

        df_luong_final = pd.merge(df_sc_sum, df_bh_sum, on=['Ma_KTV', 'Ten_KTV'], how='outer').fillna(0)
        df_luong_final = pd.merge(df_luong_final, df_pdi_sum, on=['Ma_KTV', 'Ten_KTV'], how='outer').fillna(0)
        df_luong_final['Tong_Luong_Nhan'] = df_luong_final['Cong_SC'] + df_luong_final['Cong_BH'] + df_luong_final.get('Cong_PDI', 0)
        df_luong_final = df_luong_final.sort_values(by='Tong_Luong_Nhan', ascending=False).reset_index(drop=True)
        df_luong_final.insert(0, 'STT', range(1, len(df_luong_final) + 1))

        cfg_luong = {
            "STT": st.column_config.NumberColumn("STT", width="small"),
            "Ma_KTV": st.column_config.TextColumn("Mã KTV", width="small"),
            "Ten_KTV": st.column_config.TextColumn("Họ và Tên Kỹ Thuật Viên", width="large"),
            "Cong_SC": st.column_config.NumberColumn("Công Sửa Chữa (5114)", format="%,d đ", width="medium"),
            "Cong_BH": st.column_config.NumberColumn("Công Bảo Hành (Đã Chia)", format="%,d đ", width="medium"),
            "Cong_PDI": st.column_config.NumberColumn("Công PDI (Tách Riêng)", format="%,d đ", width="medium"),
            "Tong_Luong_Nhan": st.column_config.NumberColumn("TỔNG TIỀN CÔNG TÍNH LƯƠNG", format="%,d đ", width="large")
        }
        st.dataframe(df_luong_final, use_container_width=True, hide_index=True, column_config=cfg_luong)
else:
    st.info("💡 Vui lòng nạp tệp **'1. Sổ chi tiết TK 5114 (5114.Xlsx)'** và ít nhất 1 tệp **Báo cáo doanh thu KTV** phía trên để bắt đầu.")
