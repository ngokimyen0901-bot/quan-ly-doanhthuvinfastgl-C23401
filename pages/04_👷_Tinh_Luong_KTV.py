import streamlit as st
import pandas as pd
import numpy as np
import io
import re

# Bọc an toàn set_page_config cho ứng dụng đa trang
try:
    st.set_page_config(
        page_title="Tính Lương & Đối Soát Công KTV",
        page_icon="👷",
        layout="wide"
    )
except Exception:
    pass

st.title("👷 Hệ Thống Tính Lương & Đối Soát Công Kỹ Thuật Viên")
st.caption("☁️ Quản lý đa kỳ lương | Tự động lưu trữ lịch sử hàng tháng | Lưu trữ: **Luong_KTV**.")

# Khởi tạo kết nối Google Sheets an toàn, tránh treo trang
@st.cache_resource(ttl=3600)
def get_gsheet_connection():
    try:
        from streamlit_gsheets import GSheetsConnection
        return st.connection("gsheets", type=GSheetsConnection)
    except Exception:
        return None

conn = get_gsheet_connection()

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
        if 'số c.từ' in row_str or 'số hóa đơn' in row_str or 'số r/o' in row_str or 'lệnh sửa chữa' in row_str:
            header_idx = idx
            break

    df = pd.read_excel(xls, sheet_name=xls.sheet_names[0], header=header_idx)
    df.columns = [str(c).strip() for c in df.columns]
    return df

def doc_file_excel_chung(file_obj, tu_khoa=['số ro', 'mã ktv', 'lệnh sửa chữa', 'wo']):
    if file_obj is None:
        return pd.DataFrame()
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

def doc_du_lieu_bao_hanh_thong_minh(file_obj, ky_mac_dinh='Tháng 09/2026'):
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
        df_out = df_out[df_out[col_loai_sp].astype(str).str.lower().str.contains('dịch vụ|service')].copy()
    
    col_wo = next((c for c in df_out.columns if any(k in c.lower() for k in ['lệnh sửa chữa', 'lsc', 'wo', 'số ro'])), df_out.columns[0])
    col_cong = next((c for c in df_out.columns if any(k in c.lower() for k in ['approved service', 'tiền công yêu cầu', 'tổng số tiền', 'amount', 'thành tiền', 'tiền'])), df_out.columns[-1])
    col_cv = next((c for c in df_out.columns if any(k in c.lower() for k in ['tên công việc', 'mô tả', 'công việc'])), df_out.columns[0])

    df_out['wo_norm'] = df_out[col_wo].apply(norm_wo_key)
    df_out['Tien_Cong_BH'] = df_out[col_cong].apply(clean_num)
    df_out['Ten_CV_BH'] = df_out[col_cv].fillna('').astype(str) if col_cv else ''
    df_out['Thang_BH'] = ky_mac_dinh

    return df_out[(df_out['Tien_Cong_BH'] > 0) & (df_out['wo_norm'] != '')].copy()

def trich_xuat_ktv_dataframe(df_source):
    if df_source is None or df_source.empty:
        return pd.DataFrame()
    df_res = df_source.copy()
    col_ro_h = next((c for c in df_res.columns if 'số ro hãng' in c.lower() or 'ro hãng' in c.lower() or 'số lệnh sửa chữa' in c.lower()), None)
    col_ro_nb = next((c for c in df_res.columns if c.strip().lower() in ['số ro', 'ro']), None)
    col_ma = next((c for c in df_res.columns if 'mã ktv' in c.lower()), None)
    col_ten = next((c for c in df_res.columns if 'tên ktv' in c.lower() or 'kỹ thuật viên' in c.lower()), None)
    
    col_tien_hd = next((c for c in df_res.columns if 'theo hd' in c.lower() or 'theo hóa đơn' in c.lower()), None)
    col_tien_lenh = next((c for c in df_res.columns if 'theo lệnh' in c.lower() or 'theo lenh' in c.lower()), None)
    col_thanh_tien = next((c for c in df_res.columns if c.strip().lower() in ['thành tiền', 'thanh tien', 'tổng tiền công']), None)

    col_bs = next((c for c in df_res.columns if 'biển' in c.lower()), None)
    col_hm = next((c for c in df_res.columns if 'hạng mục' in c.lower()), None)
    col_nd = next((c for c in df_res.columns if 'nội dung' in c.lower() or 'công việc' in c.lower()), None)
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

def load_saved_luong():
    if conn is None:
        return pd.DataFrame()
    try:
        df = conn.read(worksheet="Luong_KTV", ttl=20)
        if df is not None and not df.empty:
            df.columns = [str(c).strip() for c in df.columns]
            return df
    except Exception:
        pass
    return pd.DataFrame()

df_saved_gs = load_saved_luong()

# KHUNG NẠP DỮ LIỆU
with st.expander("📥 Nạp Tệp Dữ Liệu Tính Lương Mới", expanded=df_saved_gs.empty):
    c1, c2, c3 = st.columns(3)
    with c1:
        up_5114 = st.file_uploader("1. Sổ chi tiết TK 5114 hoặc Master DMS:", type=['xlsx', 'xls', 'csv'], key="p4_5114")
    with c2:
        up_ktv_hd = st.file_uploader("2. Báo cáo KTV (THEO HÓA ĐƠN - Sửa chữa):", type=['xlsx', 'xls', 'csv'], key="p4_ktv_hd")
    with c3:
        up_ktv_lenh = st.file_uploader("3. Báo cáo KTV (THEO LỆNH - Bảo hành):", type=['xlsx', 'xls', 'csv'], key="p4_ktv_lenh")

    c4, c5 = st.columns(2)
    with c4:
        up_bh = st.file_uploader("4. Chi tiết bảo hành (Active Claim / Bảng Kê WCS):", type=['xlsx', 'xls'], key="p4_bh")
    with c5:
        up_pdi = st.file_uploader("5. Danh sách lệnh PDI (Tùy chọn):", type=['xlsx', 'xls', 'csv'], key="p4_pdi")

# Xử lý tính toán khi nạp file
if up_5114 and (up_ktv_hd or up_ktv_lenh):
    df_5114 = doc_file_5114_chuan(up_5114)
    df_ktv_hd_parsed = trich_xuat_ktv_dataframe(doc_file_excel_chung(up_ktv_hd)) if up_ktv_hd else pd.DataFrame()
    df_ktv_lenh_parsed = trich_xuat_ktv_dataframe(doc_file_excel_chung(up_ktv_lenh)) if up_ktv_lenh else pd.DataFrame()

    df_ktv_for_sc = df_ktv_hd_parsed if not df_ktv_hd_parsed.empty else df_ktv_lenh_parsed
    df_ktv_for_bh = df_ktv_lenh_parsed if not df_ktv_lenh_parsed.empty else df_ktv_hd_parsed

    col_5114_ngay = next((c for c in df_5114.columns if any(k in c.lower() for k in ['ngày c.từ', 'ngày chứng từ', 'ngày xuất hóa đơn', 'ngày hđ', 'thời gian đóng lsc'])), 'Ngày C.từ')
    col_5114_shd = next((c for c in df_5114.columns if 'hóa đơn' in c.lower()), 'Số hóa đơn')
    col_5114_tien = next((c for c in df_5114.columns if c.strip().lower() in ['có', 'co', 'phát sinh có', 'tổng tiền công', 'tổng có vat', 'tiền công']), 'Có')
    col_5114_ro_hang = next((c for c in df_5114.columns if 'số r/o hãng' in c.lower() or 'ro hãng' in c.lower() or 'số lệnh sửa chữa' in c.lower()), 'Số R/O hãng')
    col_5114_ro_nb = next((c for c in df_5114.columns if c.strip().lower() in ['số r/o', 'số ro']), 'Số R/O')
    col_5114_dg = next((c for c in df_5114.columns if 'diễn giải' in c.lower() or 'unnamed: 11' in c.lower() or 'yêu cầu khách hàng' in c.lower()), 'Diễn giải')
    col_5114_kh = next((c for c in df_5114.columns if 'tên khách' in c.lower() or 'khách hàng' in c.lower()), df_5114.columns[0])

    if col_5114_ngay in df_5114.columns:
        df_5114['Thang_HD'] = pd.to_datetime(df_5114[col_5114_ngay], errors='coerce', dayfirst=True).dt.strftime('Tháng %m/%Y').fillna('Chưa rõ')
    else:
        df_5114['Thang_HD'] = 'Tháng 09/2026'

    df_5114['Tien_Dau_Vao'] = df_5114[col_5114_tien].apply(clean_num) if col_5114_tien in df_5114.columns else 0.0
    mask_tong_ket = df_5114[col_5114_dg].astype(str).str.lower().str.contains('tổng phát sinh|số dư có|số dư nợ') if col_5114_dg in df_5114.columns else pd.Series([False]*len(df_5114))
    df_5114_valid = df_5114[(df_5114['Tien_Dau_Vao'] > 0) & (~mask_tong_ket)].copy()

    def tao_key_match_5114(r):
        w = norm_wo_key(r.get(col_5114_ro_hang))
        if w: return w
        ro = norm_wo_key(r.get(col_5114_ro_nb))
        if ro: return ro
        dg = str(r.get(col_5114_dg, '')).lower()
        if 'cứu hộ' in dg: return "CUU_HO_2608"
        if 'bảo hành' in dg: return "BAO_HANH_BANG_KE_2608"
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

    ktv_details_temp = []
    for k_match, group_orders in df_5114_grouped.groupby('key_match'):
        r_5114 = group_orders.iloc[0]
        tien_5114 = r_5114['Tien_5114']
        shd_k = r_5114.get('Số hóa đơn', '')
        nhd_k = r_5114.get('Ngày C.từ', '')
        ro_show = r_5114.get('Số R/O hãng', r_5114.get('Số R/O', k_match))
        thang_hd_row = r_5114.get('Thang_HD', 'Tháng 09/2026')

        sub_all = df_ktv_for_sc[df_ktv_for_sc['key_match'] == k_match].copy()
        if not sub_all.empty:
            bs_k = sub_all['Bien_So_Val'].iloc[0]
            sub_works = sub_all[(sub_all['Ma_KTV_Clean'] != '') & (sub_all['Tien_KTV_Theo_HD'] > 0)].copy()
            if not sub_works.empty:
                for _, r_tho in sub_works.iterrows():
                    ktv_details_temp.append({
                        'Ky_Luong': thang_hd_row,
                        'key_match': k_match,
                        'Loai_Cong': 'Sửa chữa (5114)',
                        'So_RO': ro_show,
                        'So_HD': shd_k,
                        'Ngay_Xuat_HD': nhd_k,
                        'Bien_So': bs_k,
                        'Hang_Muc': r_tho['Hang_Muc_Val'],
                        'Noi_Dung_CV': r_tho['Noi_Dung_Val'],
                        'Ma_KTV': r_tho['Ma_KTV_Clean'],
                        'Ten_KTV': r_tho['Ten_KTV_Clean'],
                        'Tien_Thuc_Nhan': r_tho['Tien_KTV_Theo_HD'],
                        'Tien_Chot_KTV': r_tho['Tien_KTV_Theo_HD'],
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
    st.session_state['df_ktv_split_master'] = df_ktv_split_master

    # Xử lý bảo hành đối soát chuẩn xác
    if up_bh:
        df_bh_serv = doc_du_lieu_bao_hanh_thong_minh(up_bh, ky_mac_dinh='Tháng 09/2026')
        
        # Gom tổng tiền Nhà Máy duyệt theo từng Lệnh duy nhất
        bh_wo_sum = df_bh_serv.groupby('wo_norm').agg({
            'Tien_Cong_BH': 'sum',
            'Thang_BH': 'first',
            'Ten_CV_BH': lambda x: " | ".join(sorted(set([str(v) for v in x if pd.notna(v) and str(v).strip() != ''])))[:150]
        }).reset_index().rename(columns={'Tien_Cong_BH': 'Tien_Nha_May_Duyet'})

        # Lọc KTV làm bảo hành (chỉ lấy KTV có tiền theo lệnh > 0, loại bỏ thợ làm bảo dưỡng tính theo HĐ)
        df_ktv_bh_chuan = df_ktv_for_bh[
            (df_ktv_for_bh['key_match'].isin(bh_wo_sum['wo_norm'])) & 
            (df_ktv_for_bh['Ma_KTV_Clean'] != '') & 
            (df_ktv_for_bh['Tien_KTV_Theo_Lenh'] > 0)
        ].copy()

        mask_bh_task = df_ktv_bh_chuan['Hang_Muc_Val'].str.lower().str.contains('bảo hành|bh|cập nhật') | \
                       df_ktv_bh_chuan['Noi_Dung_Val'].str.lower().str.contains('cập nhật|bảo hành|thay thế|frs')
        if mask_bh_task.any():
            df_ktv_bh_chuan = df_ktv_bh_chuan[mask_bh_task].copy()

        ktv_bh_tot_map = df_ktv_bh_chuan.groupby('key_match')['Tien_KTV_Theo_Lenh'].sum().to_dict()
        
        def gom_ktv(sub_df):
            names = []
            for _, r in sub_df.drop_duplicates(subset=['Ma_KTV_Clean']).iterrows():
                names.append(str(r['Ten_KTV_Clean'] if r['Ten_KTV_Clean'] else r['Ma_KTV_Clean']).strip())
            return ", ".join(names) if names else "Chưa gán thợ BH"

        ktv_names_bh_map = df_ktv_bh_chuan.groupby('key_match').apply(gom_ktv).to_dict()
        bs_bh_map = df_ktv_bh_chuan.groupby('key_match')['Bien_So_Val'].first().to_dict()
        ro_bh_map = df_ktv_bh_chuan.groupby('key_match')['So_RO_Val'].first().to_dict()

        # Tạo bảng đối soát: 1 Lệnh = 1 Dòng
        df_bh_split = bh_wo_sum.copy()
        df_bh_split['So_RO'] = df_bh_split['wo_norm'].map(ro_bh_map).fillna(df_bh_split['wo_norm'])
        df_bh_split['Bien_So'] = df_bh_split['wo_norm'].map(bs_bh_map).fillna('')
        df_bh_split['KTV_Bao_Hanh'] = df_bh_split['wo_norm'].map(ktv_names_bh_map).fillna('Chưa gán thợ BH')
        df_bh_split['Tien_KTV_He_Thong'] = df_bh_split['wo_norm'].map(ktv_bh_tot_map).fillna(0.0)
        df_bh_split['Chenh_Lech_Duyet'] = df_bh_split['Tien_Nha_May_Duyet'] - df_bh_split['Tien_KTV_He_Thong']
        
        def check_status(r):
            if r['KTV_Bao_Hanh'] == 'Chưa gán thợ BH':
                return "🟡 CHƯA CÓ KTV BẢO HÀNH"
            if abs(r['Chenh_Lech_Duyet']) < 1000:
                return "🟢 Khớp duyệt nhà máy"
            return f"🔴 LỆCH DUYỆT (Lệch: {r['Chenh_Lech_Duyet']:,.0f} đ)"

        df_bh_split['Trang_Thai_Duyet'] = df_bh_split.apply(check_status, axis=1)

        cols_bh = ['Thang_BH', 'So_RO', 'Bien_So', 'Ten_CV_BH', 'KTV_Bao_Hanh', 'Tien_Nha_May_Duyet', 'Tien_KTV_He_Thong', 'Chenh_Lech_Duyet', 'Trang_Thai_Duyet']
        st.session_state['df_bh_split_cache'] = df_bh_split[[c for c in cols_bh if c in df_bh_split.columns]]

# Nạp dữ liệu hiển thị từ Cache hoặc Google Sheets
df_ktv_split_master = st.session_state.get('df_ktv_split_master', df_saved_gs)
df_bh_split = st.session_state.get('df_bh_split_cache', pd.DataFrame())

if not df_ktv_split_master.empty:
    df_ktv_split_master = bao_ve_cot_ktv_detail(df_ktv_split_master)
    
    st.markdown("---")
    danh_sach_ky = sorted(list(set([str(m) for m in df_ktv_split_master['Ky_Luong'].unique() if str(m) not in ['', 'nan', 'None']])))
    if not danh_sach_ky:
        danh_sach_ky = ['Tháng 09/2026']

    sel_ky = st.selectbox("📅 CHỌN KỲ LƯƠNG ĐỐI SOÁT:", ["Tất cả các kỳ lương"] + danh_sach_ky, index=1 if len(danh_sach_ky) > 0 else 0)

    df_active = df_ktv_split_master.copy()
    if sel_ky != "Tất cả các kỳ lương":
        df_active = df_active[df_active['Ky_Luong'] == sel_ky]

    tab1, tab2, tab3 = st.tabs(["🔧 1. Chi Tiết Lương KTV", "🛡️ 2. Đối Soát Bảo Hành Hãng", "💰 3. Bảng Tổng Hợp Lương KTV"])
    
    with tab1:
        st.subheader(f"🔧 Chi Tiết Công Việc KTV ({sel_ky})")
        cols_t1 = ['So_RO', 'Bien_So', 'So_HD', 'Ngay_Xuat_HD', 'Hang_Muc', 'Noi_Dung_CV', 'Ma_KTV', 'Ten_KTV', 'Tien_Chot_KTV', 'Trang_Thai_Khop']
        cols_t1_v = [c for c in cols_t1 if c in df_active.columns]
        st.dataframe(df_active[cols_t1_v], use_container_width=True, height=520, hide_index=True)

    with tab2:
        st.subheader(f"🛡️ Đối Soát Bảo Hành Hãng (1 Lệnh / 1 Dòng Duy Nhất)")
        if not df_bh_split.empty:
            cfg_bh = {
                "Thang_BH": st.column_config.TextColumn("Kỳ Bảo Hành", width="small"),
                "So_RO": st.column_config.TextColumn("Số LSC / RO", width="medium"),
                "Bien_So": st.column_config.TextColumn("Biển Số", width="small"),
                "Ten_CV_BH": st.column_config.TextColumn("Nội Dung Công Việc", width="large"),
                "KTV_Bao_Hanh": st.column_config.TextColumn("KTV Bảo Hành", width="medium"),
                "Tien_Nha_May_Duyet": st.column_config.NumberColumn("Nhà Máy Duyệt", format="%,d đ"),
                "Tien_KTV_He_Thong": st.column_config.NumberColumn("KTV Nhận", format="%,d đ"),
                "Chenh_Lech_Duyet": st.column_config.NumberColumn("Chênh Lệch", format="%,d đ"),
                "Trang_Thai_Duyet": st.column_config.TextColumn("Trạng Thái Duyệt", width="medium")
            }
            st.dataframe(df_bh_split, column_config=cfg_bh, use_container_width=True, height=520, hide_index=True)
        else:
            st.info("Chưa nạp tệp chi tiết bảo hành (Active Claim / Bảng Kê WCS).")

    with tab3:
        st.subheader(f"💰 Bảng Tổng Hợp Lương Theo Kỹ Thuật Viên ({sel_ky})")
        
        # XỬ LÝ AN TOÀN TRIỆT ĐỂ: Tránh lỗi TypeError do ô trống NaN
        s_ma = df_active['Ma_KTV'].fillna('').astype(str)
        mask_hop_le = (~s_ma.str.contains('🚨|Chưa xác định|Khoản mục riêng|Bảo hành hãng', regex=True)) & (s_ma.str.strip() != '')
        df_real = df_active[mask_hop_le].copy()

        if not df_real.empty:
            df_real['Tien_Chot_KTV'] = pd.to_numeric(df_real['Tien_Chot_KTV'], errors='coerce').fillna(0.0)
            
            sum_ktv = df_real.groupby(['Ma_KTV', 'Ten_KTV'])['Tien_Chot_KTV'].agg(
                Tong_Cong='sum',
                So_Viec='count'
            ).reset_index()
            
            sum_ktv.columns = ['Mã KTV', 'Tên Kỹ Thuật Viên', 'Tổng Tiền Công (VNĐ)', 'Số Lượng Việc']
            
            cfg_sum = {
                "Tổng Tiền Công (VNĐ)": st.column_config.NumberColumn("Tổng Tiền Công", format="%,d đ"),
                "Số Lượng Việc": st.column_config.NumberColumn("Số Công Việc", format="%,d"),
            }
            st.dataframe(sum_ktv, column_config=cfg_sum, use_container_width=True, hide_index=True)
        else:
            st.info("Chưa có dữ liệu KTV hợp lệ để tổng hợp.")
else:
    st.info("💡 Vui lòng tải các tệp dữ liệu lên khung phía trên để bắt đầu tính toán và đối soát.")
