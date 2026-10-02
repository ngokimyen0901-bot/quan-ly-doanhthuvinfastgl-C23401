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

st.title("👷 Tính Lương & Phân Bổ Công Kỹ Thuật Viên")
st.caption("☁️ Phân luồng: **Bảo Hành Tháng** | **Sửa Chữa 5114** | **PDI Tách Riêng** | Lưu trữ: **Luong_KTV**.")

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

def doc_file_excel_da_nang(file_obj, tu_khoa_nhan_dien=['số r/o', 'số ro', 'lệnh', 'lsc', 'wo', 'hóa đơn', 'mã c.từ', 'số c.từ', 'mã ktv']):
    if file_obj.name.lower().endswith('.csv'):
        df = pd.read_csv(file_obj, low_memory=False)
        df.columns = [str(c).strip() for c in df.columns]
        return df
        
    xls = pd.ExcelFile(file_obj)
    target_sheet = xls.sheet_names[0]
    for s in xls.sheet_names:
        if any(k in s.lower() for k in ['chi tiết', 'đối soát', '5114', 'ktv', 'claim', 'active', 'pdi']):
            target_sheet = s
            break
            
    df_first = pd.read_excel(xls, sheet_name=target_sheet)
    cols_first_str = " ".join([str(c).lower() for c in df_first.columns])
    if any(k in cols_first_str for k in tu_khoa_nhan_dien) and sum('unnamed' in str(c).lower() for c in df_first.columns) < len(df_first.columns) * 0.5:
        df_first.columns = [str(c).strip() for c in df_first.columns]
        return df_first

    df_raw = pd.read_excel(xls, sheet_name=target_sheet, header=None)
    header_idx = 0
    for idx, row in df_raw.head(25).iterrows():
        row_str = " ".join([str(v).lower() for v in row.values if pd.notna(v)])
        if any(k in row_str for k in tu_khoa_nhan_dien):
            header_idx = idx
            break
            
    df = pd.read_excel(xls, sheet_name=target_sheet, header=header_idx)
    df.columns = [str(c).strip() for c in df.columns]
    return df

def tim_cot_lsc_linh_hoat(df):
    cols = [str(c).strip() for c in df.columns]
    tu_khoa = [
        'số r/o hãng', 'r/o hãng', 'ro hãng', 'lệnh hãng', 
        'lệnh sửa chữa', 'lsc', 'số lệnh sửa chữa', 'số lệnh', 
        'số r/o', 'số ro', 'wo', 'ro'
    ]
    for kw in tu_khoa:
        for c in cols:
            if c.lower() == kw or (kw in c.lower() and 'ngày' not in c.lower() and 'trạng thái' not in c.lower()):
                return c
    for c in cols:
        sample_vals = df[c].dropna().astype(str).head(20).tolist()
        if any('wo-' in v.lower() or 'c23401' in v.lower() for v in sample_vals):
            return c
    return cols[0] if cols else 'Số R/O'

@st.cache_data(ttl=60)
def load_master_ref():
    try:
        df = conn.read(worksheet="MasterData", ttl=60)
        if df is not None and not df.empty:
            df.columns = [str(c).strip() for c in df.columns]
            col_lsc = next((c for c in df.columns if any(k in c.lower() for k in ['lệnh sửa chữa', 'lsc', 'wo'])), 'Số lệnh sửa chữa')
            df['wo_norm'] = df[col_lsc].apply(norm_wo_key)
            return df
    except Exception:
        pass
    return pd.DataFrame()

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

df_master_ref = load_master_ref()
df_saved_gs = load_saved_luong()

# ----------------- KHU VỰC NẠP FILE -----------------
st.markdown("### 📥 Nạp Tệp Dữ Liệu")
c1, c2, c3, c4 = st.columns(4)
with c1:
    up_5114 = st.file_uploader("1. Sổ chi tiết TK 5114:", type=['xlsx', 'xls', 'csv'], key="p4_5114")
with c2:
    up_ktv = st.file_uploader("2. Báo cáo doanh thu KTV:", type=['xlsx', 'xls', 'csv'], key="p4_ktv")
with c3:
    up_bh = st.file_uploader("3. Chi tiết bảo hành (Active Claim):", type=['xlsx', 'xls'], key="p4_bh")
with c4:
    up_pdi = st.file_uploader("4. Danh sách lệnh PDI (Tùy chọn):", type=['xlsx', 'xls', 'csv'], key="p4_pdi")

if up_5114 and up_ktv:
    df_5114 = doc_file_excel_da_nang(up_5114, tu_khoa_nhan_dien=['số r/o', 'số ro', 'lệnh', 'lsc', 'wo', 'hóa đơn', 'mã c.từ', 'số c.từ', 'có'])
    df_ktv = doc_file_excel_da_nang(up_ktv, tu_khoa_nhan_dien=['số ro', 'mã ktv', 'tên ktv', 'thành tiền', 'hạng mục'])

    # 1. Đọc danh sách PDI
    pdi_keys_set = set()
    if up_pdi:
        df_pdi_input = doc_file_excel_da_nang(up_pdi, tu_khoa_nhan_dien=['lệnh', 'wo', 'ro', 'pdi', 'số'])
        col_pdi_ro = tim_cot_lsc_linh_hoat(df_pdi_input)
        for v in df_pdi_input[col_pdi_ro].dropna():
            pdi_keys_set.add(norm_wo_key(v))

    # 2. Xử lý File 5114
    col_5114_ro = tim_cot_lsc_linh_hoat(df_5114)
    col_5114_shd = next((c for c in df_5114.columns if 'hóa đơn' in c.lower()), next((c for c in df_5114.columns if 'chứng từ' in c.lower()), df_5114.columns[0]))
    col_5114_ngay = next((c for c in df_5114.columns if 'ngày' in c.lower()), None)
    col_5114_tien = next((c for c in df_5114.columns if c.strip().lower() in ['có', 'co', 'phát sinh có', 'thành tiền', 'tiền']), None)
    if not col_5114_tien:
        col_5114_tien = next((c for c in df_5114.columns if 'tiền' in c.lower() or 'doanh thu' in c.lower()), df_5114.columns[-1])
    col_5114_kh = next((c for c in df_5114.columns if 'tên khách' in c.lower() or 'khách hàng' in c.lower()), df_5114.columns[0])
    col_5114_dg = next((c for c in df_5114.columns if 'diễn giải' in c.lower() or 'nội dung' in c.lower()), df_5114.columns[0])

    df_5114['wo_norm'] = df_5114[col_5114_ro].apply(norm_wo_key)
    df_5114['Tien_5114'] = df_5114[col_5114_tien].apply(clean_num)
    if col_5114_ngay:
        df_5114['Thang'] = pd.to_datetime(df_5114[col_5114_ngay], errors='coerce', dayfirst=True).dt.strftime('Tháng %m/%Y').fillna('Chưa rõ')
    else:
        df_5114['Thang'] = 'Toàn bộ kỳ'

    df_5114_clean = df_5114[(df_5114['Tien_5114'] > 0) & (df_5114['wo_norm'] != '')].copy()

    # 3. Xử lý File KTV
    col_ktv_ro = tim_cot_lsc_linh_hoat(df_ktv)
    col_ktv_ma = next((c for c in df_ktv.columns if 'mã ktv' in c.lower()), 'Mã KTV')
    col_ktv_ten = next((c for c in df_ktv.columns if 'tên ktv' in c.lower()), 'Tên KTV')
    col_ktv_cv = next((c for c in df_ktv.columns if 'mã cv' in c.lower() or 'cố vấn' in c.lower()), 'Mã CV')
    col_ktv_tien = next((c for c in df_ktv.columns if 'thành tiền' in c.lower() or 'tiền dịch vụ' in c.lower()), 'Thành tiền')
    col_ktv_bs = next((c for c in df_ktv.columns if 'biển' in c.lower()), 'Biển kiểm soát')
    col_ktv_nd = next((c for c in df_ktv.columns if 'nội dung' in c.lower() or 'hạng mục' in c.lower()), 'Nội dung')
    col_ktv_ngay = next((c for c in df_ktv.columns if 'ngày' in c.lower() and ('nghiệm thu' in c.lower() or 'ra xưởng' in c.lower())), None)

    df_ktv['wo_norm'] = df_ktv[col_ktv_ro].apply(norm_wo_key)
    df_ktv['Tien_KTV'] = df_ktv[col_ktv_tien].apply(clean_num) if col_ktv_tien in df_ktv.columns else 0.0
    df_ktv['Ma_KTV_Clean'] = df_ktv[col_ktv_ma].fillna('').astype(str).str.strip().str.replace('.0', '', regex=False) if col_ktv_ma in df_ktv.columns else ""
    df_ktv['Ten_KTV_Clean'] = df_ktv[col_ktv_ten].fillna('').astype(str).str.strip() if col_ktv_ten in df_ktv.columns else ""

    if col_ktv_ngay:
        df_ktv['Thang_KTV'] = pd.to_datetime(df_ktv[col_ktv_ngay], errors='coerce', dayfirst=True).dt.strftime('Tháng %m/%Y').fillna('Chưa rõ')
    else:
        df_ktv['Thang_KTV'] = 'Toàn bộ kỳ'

    agg_dict = {'Tien_KTV': 'sum', 'Thang_KTV': 'first'}
    if col_ktv_ro in df_ktv.columns: agg_dict[col_ktv_ro] = 'first'
    if col_ktv_bs in df_ktv.columns: agg_dict[col_ktv_bs] = 'first'
    if col_ktv_cv in df_ktv.columns: agg_dict[col_ktv_cv] = 'first'
    agg_dict['Ma_KTV_Clean'] = lambda x: ", ".join(sorted(set([str(v) for v in x if str(v).strip() and str(v) != 'nan'])))
    agg_dict['Ten_KTV_Clean'] = lambda x: ", ".join(sorted(set([str(v) for v in x if str(v).strip() and str(v) != 'nan'])))
    if col_ktv_nd in df_ktv.columns:
        agg_dict[col_ktv_nd] = lambda x: " | ".join(sorted(set([str(v) for v in x if str(v).strip() and str(v) != 'nan'])))[:150]

    ktv_per_ro = df_ktv[df_ktv['wo_norm'] != ''].groupby('wo_norm').agg(agg_dict).reset_index()

    if not df_master_ref.empty and 'Cố vấn dịch vụ' in df_master_ref.columns:
        map_cv = df_master_ref.set_index('wo_norm')['Cố vấn dịch vụ'].to_dict()
        ktv_per_ro['Ten_CVDV'] = ktv_per_ro['wo_norm'].map(map_cv).fillna(ktv_per_ro.get(col_ktv_cv, ''))
    else:
        ktv_per_ro['Ten_CVDV'] = ktv_per_ro.get(col_ktv_cv, '')

    # 4. Xử lý File Bảo Hành (Active Warranty)
    bh_wo_set = set()
    df_bh_split = pd.DataFrame()
    bh_wo_sum = pd.DataFrame()

    if up_bh:
        df_bh_raw = doc_file_excel_da_nang(up_bh, tu_khoa_nhan_dien=['lệnh sửa chữa', 'lsc', 'wo', 'tiền công'])
        col_bh_wo = tim_cot_lsc_linh_hoat(df_bh_raw)
        col_bh_cong = next((c for c in df_bh_raw.columns if any(k in c.lower() for k in ['tiền công yêu cầu', 'approved service', 'tiền công'])), None)
        if not col_bh_cong:
            col_bh_cong = next((c for c in df_bh_raw.columns if 'tiền' in c.lower()), df_bh_raw.columns[-1])
        col_bh_cv = next((c for c in df_bh_raw.columns if any(k in c.lower() for k in ['tên công việc chính', 'công việc chính', 'mô tả'])), df_bh_raw.columns[0])
        col_bh_ngay = next((c for c in df_bh_raw.columns if 'ngày' in c.lower() and ('phê duyệt' in c.lower() or 'đxbh' in c.lower())), None)

        df_bh_raw['wo_norm'] = df_bh_raw[col_bh_wo].apply(norm_wo_key)
        df_bh_raw['Tien_Cong_BH'] = df_bh_raw[col_bh_cong].apply(clean_num)
        if col_bh_ngay:
            df_bh_raw['Thang_BH'] = pd.to_datetime(df_bh_raw[col_bh_ngay], errors='coerce', dayfirst=True).dt.strftime('Tháng %m/%Y').fillna('Chưa rõ')
        else:
            df_bh_raw['Thang_BH'] = 'Toàn bộ kỳ'

        df_bh_cong_only = df_bh_raw[(df_bh_raw['Tien_Cong_BH'] > 0) & (df_bh_raw['wo_norm'] != '')].copy()
        bh_wo_set = set(df_bh_cong_only['wo_norm'])

        bh_wo_sum = df_bh_cong_only.groupby('wo_norm').agg({
            col_bh_wo: 'first',
            'Tien_Cong_BH': 'sum',
            'Thang_BH': 'first',
            col_bh_cv: lambda x: " | ".join(sorted(set([str(v) for v in x if pd.notna(v)])))[:150]
        }).reset_index()

        df_ktv_bh_rows = df_ktv[df_ktv['wo_norm'].isin(bh_wo_sum['wo_norm'])].copy()
        df_ktv_bh_rows = df_ktv_bh_rows[df_ktv_bh_rows['Ma_KTV_Clean'] != ''].copy()

        ktv_cnt_per_wo = df_ktv_bh_rows.groupby('wo_norm')['Ma_KTV_Clean'].nunique().to_dict()
        df_ktv_bh_rows['So_KTV_Lam_Chung'] = df_ktv_bh_rows['wo_norm'].map(ktv_cnt_per_wo).fillna(1)
        df_ktv_bh_rows.loc[df_ktv_bh_rows['So_KTV_Lam_Chung'] == 0, 'So_KTV_Lam_Chung'] = 1

        df_bh_split = pd.merge(
            df_ktv_bh_rows[['wo_norm', col_ktv_ro, col_ktv_bs, 'Ma_KTV_Clean', 'Ten_KTV_Clean', 'So_KTV_Lam_Chung']].drop_duplicates(subset=['wo_norm', 'Ma_KTV_Clean']),
            bh_wo_sum,
            on='wo_norm',
            how='left'
        )
        df_bh_split['Cong_BH_Thuc_Nhan'] = df_bh_split['Tien_Cong_BH'] / df_bh_split['So_KTV_Lam_Chung']

    # ----------------- PHÂN LUỒNG RÕ RÀNG -----------------
    def phan_loai_lenh(r):
        w = r['wo_norm']
        if w in pdi_keys_set:
            return "PDI"
        if w in bh_wo_set:
            return "BAO_HANH"
        return "SUA_CHUA_5114"

    df_merged = pd.merge(
        df_5114_clean[['wo_norm', col_5114_ro, col_5114_shd, 'Thang', 'Tien_5114', col_5114_kh, col_5114_dg]],
        ktv_per_ro,
        on='wo_norm',
        how='left'
    )
    df_merged['Tien_KTV'] = df_merged['Tien_KTV'].fillna(0)
    df_merged['Chenh_Lech'] = df_merged['Tien_5114'] - df_merged['Tien_KTV']
    df_merged['Nhom_Phan_Loai'] = df_merged.apply(phan_loai_lenh, axis=1)

    def danh_gia_trang_thai(r):
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

    df_merged['Quyet_Dinh_Nguon'] = df_merged['Ghi_Chu_Doi_Soat'].apply(
        lambda x: "Tự động (Khớp chuẩn)" if "Khớp 100%" in x else "Lấy theo KTV (Hóa đơn)"
    )

    df_merged['Tien_Cong_Chot_Luong'] = df_merged.apply(
        lambda r: r['Tien_5114'] if "Khớp 100%" in r['Ghi_Chu_Doi_Soat'] else r['Tien_KTV'], axis=1
    )
    df_merged['Ghi_Chu_Thu_Cong'] = ""

    # Khôi phục dữ liệu đã lưu
    if not df_saved_gs.empty and 'wo_norm' in df_saved_gs.columns:
        map_nguon = df_saved_gs.set_index('wo_norm')['Quyet_Dinh_Nguon'].to_dict() if 'Quyet_Dinh_Nguon' in df_saved_gs.columns else {}
        map_tien_chot = df_saved_gs.set_index('wo_norm')['Tien_Cong_Chot_Luong'].to_dict() if 'Tien_Cong_Chot_Luong' in df_saved_gs.columns else {}
        map_note_tc = df_saved_gs.set_index('wo_norm')['Ghi_Chu_Thu_Cong'].to_dict() if 'Ghi_Chu_Thu_Cong' in df_saved_gs.columns else {}
        
        for idx_r, r_val in df_merged.iterrows():
            k_w = r_val['wo_norm']
            if k_w in map_nguon and map_nguon[k_w]:
                df_merged.at[idx_r, 'Quyet_Dinh_Nguon'] = str(map_nguon[k_w])
            if k_w in map_tien_chot:
                df_merged.at[idx_r, 'Tien_Cong_Chot_Luong'] = clean_num(map_tien_chot[k_w])
            if k_w in map_note_tc:
                df_merged.at[idx_r, 'Ghi_Chu_Thu_Cong'] = str(map_note_tc[k_w])

    # ----------------- BỘ LỌC THỜI GIAN THEO THÁNG / NĂM -----------------
    st.markdown("---")
    all_months = sorted(list(set([str(m) for m in df_merged['Thang'].unique() if str(m) not in ['', 'nan', 'None']])))
    
    fl_col1, fl_col2 = st.columns([4, 6])
    with fl_col1:
        sel_thang_nam = st.selectbox("📅 LỰA CHỌN THỜI GIAN THEO THÁNG / NĂM:", ["Tất cả các tháng"] + all_months, index=0)

    df_merged_filtered = df_merged.copy()
    if sel_thang_nam != "Tất cả các tháng":
        df_merged_filtered = df_merged_filtered[df_merged_filtered['Thang'] == sel_thang_nam]

    df_sc_remaining = df_merged_filtered[df_merged_filtered['Nhom_Phan_Loai'] == 'SUA_CHUA_5114'].copy()
    df_pdi_only = df_merged_filtered[df_merged_filtered['Nhom_Phan_Loai'] == 'PDI'].copy()

    # ----------------- GIAO DIỆN CÁC TAB CHỨC NĂNG -----------------
    tab_sc, tab_bh, tab_pdi, tab_ktv_sum = st.tabs([
        "🔧 1. Lệnh Sửa Chữa Trong 5114 (Đối Soát & Bạn Quyết Định)",
        "🛡️ 2. Lệnh Bảo Hành Trong Tháng (Chỉ Tiền Công & Chia Đều)",
        "🚗 3. Lệnh PDI (Tách Riêng Theo Danh Sách)",
        "👷 4. Bảng Lương Tổng Hợp Kỹ Thuật Viên"
    ])

    # ==================== TAB 1 ====================
    with tab_sc:
        st.subheader(f"🔧 Các Số WO Còn Lại Trong 5114 ({sel_thang_nam})")
        st.caption("Các lệnh chuẩn **🟢 Khớp 100%** sẽ chạy tự động. Các lệnh lệch do thuê ngoài, bảo hiểm khấu trừ... có biểu tượng **🔴 Lệch** để bạn quyết định giá trị chốt ở cột cuối.")

        cnt_khop = (df_sc_remaining['Ghi_Chu_Doi_Soat'] == "🟢 Khớp 100%").sum()
        cnt_lech = len(df_sc_remaining) - cnt_khop
        
        m1, m2, m3 = st.columns(3)
        m1.metric("Tổng Doanh Thu 5114 (Sửa Chữa)", f"{df_sc_remaining['Tien_5114'].sum():,.0f} đ", f"{len(df_sc_remaining)} lệnh")
        m2.metric("Lệnh Khớp Chuẩn (Tự động chạy)", f"{cnt_khop} lệnh")
        m3.metric("Lệnh Lệch (Bạn quyết định)", f"{cnt_lech} lệnh", delta=f"{cnt_lech} lệnh cần duyệt" if cnt_lech > 0 else "0", delta_color="inverse")

        f_c1, f_c2 = st.columns([4, 6])
        with f_c1:
            loc_xem = st.selectbox("🔍 Lọc hiển thị:", ["Tất cả lệnh sửa chữa", "🔴 Chỉ xem các lệnh LỆCH (Cần bạn duyệt)", "🟢 Chỉ xem các lệnh KHỚP 100%"])
        
        df_display_sc = df_sc_remaining.copy()
        if "Chỉ xem các lệnh LỆCH" in loc_xem:
            df_display_sc = df_display_sc[df_display_sc['Ghi_Chu_Doi_Soat'] != "🟢 Khớp 100%"]
        elif "Chỉ xem các lệnh KHỚP" in loc_xem:
            df_display_sc = df_display_sc[df_display_sc['Ghi_Chu_Doi_Soat'] == "🟢 Khớp 100%"]

        cols_edit = [
            'Thang', col_5114_ro, col_5114_shd, 'Tien_5114', 'Tien_KTV', 'Chenh_Lech',
            'Ghi_Chu_Doi_Soat', 'Ten_CVDV', 'Ma_KTV_Clean', 'Ten_KTV_Clean', col_ktv_bs,
            'Quyet_Dinh_Nguon', 'Tien_Cong_Chot_Luong', 'Ghi_Chu_Thu_Cong', 'wo_norm'
        ]
        cols_valid = [c for c in cols_edit if c in df_display_sc.columns]

        cfg_edit = {
            "Thang": st.column_config.TextColumn("Tháng", disabled=True, width="small"),
            col_5114_ro: st.column_config.TextColumn("Số Lệnh SC (RO)", disabled=True, width="medium"),
            col_5114_shd: st.column_config.TextColumn("Số HĐ", disabled=True, width="small"),
            "Tien_5114": st.column_config.NumberColumn("Tiền 5114", format="%,d đ", disabled=True, width="medium"),
            "Tien_KTV": st.column_config.NumberColumn("Tiền KTV (HĐ)", format="%,d đ", disabled=True, width="medium"),
            "Chenh_Lech": st.column_config.NumberColumn("Chênh lệch", format="%,d đ", disabled=True, width="small"),
            "Ghi_Chu_Doi_Soat": st.column_config.TextColumn("Trạng thái", disabled=True, width="medium"),
            "Ten_CVDV": st.column_config.TextColumn("Cố vấn", disabled=True, width="medium"),
            "Ma_KTV_Clean": st.column_config.TextColumn("Mã KTV", disabled=True, width="small"),
            "Ten_KTV_Clean": st.column_config.TextColumn("Tên KTV thực hiện", disabled=True, width="large"),
            col_ktv_bs: st.column_config.TextColumn("Biển số", disabled=True, width="small"),
            "Quyet_Dinh_Nguon": st.column_config.SelectboxColumn(
                "Nguồn lấy giá trị (Bạn chọn)",
                options=["Tự động (Khớp chuẩn)", "Lấy theo KTV (Hóa đơn)", "Lấy theo 5114", "Nhập số tiền khác"],
                disabled=False,
                width="medium"
            ),
            "Tien_Cong_Chot_Luong": st.column_config.NumberColumn(
                "👉 SỐ TIỀN BẠN QUYẾT ĐỊNH (Chốt tính lương)",
                format="%,d đ",
                disabled=False,
                width="large"
            ),
            "Ghi_Chu_Thu_Cong": st.column_config.TextColumn("Ghi chú lý do lệch", disabled=False, width="medium"),
            "wo_norm": st.column_config.TextColumn("Mã chuẩn", disabled=True, width="small")
        }

        edited_sc_df = st.data_editor(
            df_display_sc[cols_valid],
            use_container_width=True,
            height=540,
            column_config=cfg_edit,
            hide_index=True,
            key="editor_doi_soat_luong_sc"
        )

        c_save1, c_save2 = st.columns([3.5, 6.5])
        with c_save1:
            if st.button("☁️ Lưu Quyết Định Sửa Chữa Lên Google Sheets", type="primary", use_container_width=True):
                with st.spinner("⏳ Đang lưu vào sheet Luong_KTV..."):
                    dict_updated = edited_sc_df.set_index('wo_norm').to_dict('index')
                    for idx_r, r_val in df_merged.iterrows():
                        k_w = r_val['wo_norm']
                        if k_w in dict_updated:
                            df_merged.at[idx_r, 'Quyet_Dinh_Nguon'] = dict_updated[k_w].get('Quyet_Dinh_Nguon', df_merged.at[idx_r, 'Quyet_Dinh_Nguon'])
                            df_merged.at[idx_r, 'Tien_Cong_Chot_Luong'] = dict_updated[k_w].get('Tien_Cong_Chot_Luong', df_merged.at[idx_r, 'Tien_Cong_Chot_Luong'])
                            df_merged.at[idx_r, 'Ghi_Chu_Thu_Cong'] = dict_updated[k_w].get('Ghi_Chu_Thu_Cong', df_merged.at[idx_r, 'Ghi_Chu_Thu_Cong'])
                    try:
                        conn.update(worksheet="Luong_KTV", data=df_merged)
                        st.success("✅ Đã lưu quyết định của bạn vĩnh viễn lên Google Sheets!")
                        st.rerun()
                    except Exception as e_s:
                        st.error(f"Lỗi khi lưu: {e_s}")

    # ==================== TAB 2 ====================
    with tab_bh:
        st.subheader(f"🛡️ Các Số WO Có Trong Bảo Hành ({sel_thang_nam})")
        st.caption("Chỉ tính **Tiền công dịch vụ** (loại bỏ chi phí phụ tùng). Tự động tra cứu mã WO và chia đều cho các KTV cùng tham gia.")

        if up_bh and not df_bh_split.empty:
            df_bh_view = df_bh_split.copy()
            if sel_thang_nam != "Tất cả các tháng" and 'Thang_BH' in df_bh_view.columns:
                df_bh_view = df_bh_view[df_bh_view['Thang_BH'] == sel_thang_nam]

            b_m1, b_m2 = st.columns(2)
            b_m1.metric("Tổng Tiền Công Bảo Hành (No VAT)", f"{df_bh_view['Cong_BH_Thuc_Nhan'].sum():,.0f} đ", f"{df_bh_view['wo_norm'].nunique()} lệnh WO")
            b_m2.metric("Số Lượt Chia Công KTV", f"{len(df_bh_view)} lượt công")

            cfg_bh_split = {
                col_bh_wo: st.column_config.TextColumn("Lệnh Bảo Hành (WO)", width="medium"),
                col_ktv_bs: st.column_config.TextColumn("Biển số", width="small"),
                col_bh_cv: st.column_config.TextColumn("Nội dung bảo hành", width="large"),
                "Tien_Cong_BH": st.column_config.NumberColumn("Tổng tiền công", format="%,d đ", width="medium"),
                "So_KTV_Lam_Chung": st.column_config.NumberColumn("Số KTV chia đều", width="small"),
                "Ma_KTV_Clean": st.column_config.TextColumn("Mã KTV", width="small"),
                "Ten_KTV_Clean": st.column_config.TextColumn("Họ tên KTV", width="medium"),
                "Cong_BH_Thuc_Nhan": st.column_config.NumberColumn("Công thực nhận", format="%,d đ", width="medium")
            }
            cols_show_bh = [col_bh_wo, col_ktv_bs, col_bh_cv, 'Tien_Cong_BH', 'So_KTV_Lam_Chung', 'Ma_KTV_Clean', 'Ten_KTV_Clean', 'Cong_BH_Thuc_Nhan']
            st.dataframe(df_bh_view[cols_show_bh], use_container_width=True, hide_index=True, column_config=cfg_bh_split)
        else:
            st.warning("⚠️ Vui lòng nạp file '3. Chi tiết bảo hành hãng' để chia công bảo hành.")

    # ==================== TAB 3 ====================
    with tab_pdi:
        st.subheader(f"🚗 Danh Sách Lệnh PDI Được Tách Riêng ({sel_thang_nam})")
        st.caption("Các lệnh thuộc danh sách PDI được tách riêng hoàn toàn, không tính lẫn vào công sửa chữa thông thường.")

        if not df_pdi_only.empty:
            p_m1, p_m2 = st.columns(2)
            p_m1.metric("Tổng Tiền Công PDI", f"{df_pdi_only['Tien_KTV'].sum():,.0f} đ", f"{len(df_pdi_only)} lệnh")
            p_m2.metric("Số Lượt KTV Làm PDI", f"{len(df_pdi_only)} lượt")

            cols_pdi_show = ['Thang', col_5114_ro, 'Tien_KTV', 'Ten_CVDV', 'Ma_KTV_Clean', 'Ten_KTV_Clean', col_ktv_bs, col_ktv_nd]
            cols_pdi_valid = [c for c in cols_pdi_show if c in df_pdi_only.columns]
            cfg_pdi = {
                col_5114_ro: st.column_config.TextColumn("Số Lệnh PDI", width="medium"),
                "Tien_KTV": st.column_config.NumberColumn("Tiền công PDI", format="%,d đ", width="medium"),
                "Ten_CVDV": st.column_config.TextColumn("Cố vấn", width="medium"),
                "Ma_KTV_Clean": st.column_config.TextColumn("Mã KTV", width="small"),
                "Ten_KTV_Clean": st.column_config.TextColumn("Tên KTV", width="medium"),
                col_ktv_bs: st.column_config.TextColumn("Biển số / Mã PDI", width="small"),
                col_ktv_nd: st.column_config.TextColumn("Nội dung công việc", width="large")
            }
            st.dataframe(df_pdi_only[cols_pdi_valid], use_container_width=True, hide_index=True, column_config=cfg_pdi)
        else:
            st.info("💡 Chưa có lệnh PDI nào trong tháng này (hoặc chưa nạp danh sách PDI ở ô số 4).")

    # ==================== TAB 4 ====================
    with tab_ktv_sum:
        st.subheader(f"👷 Bảng Lương Tổng Hợp Kỹ Thuật Viên ({sel_thang_nam})")
        st.caption("Công sửa chữa tính theo **giá trị bạn quyết định** (Tab 1), công bảo hành tính theo **tỷ lệ chia đều** (Tab 2), và công PDI được tổng hợp rõ ràng.")

        sc_source = edited_sc_df if 'edited_sc_df' in locals() else df_sc_remaining
        
        # 1. Công sửa chữa thương mại
        ktv_sc_list = []
        for _, r_m in sc_source.iterrows():
            ma_ktvs = [x.strip() for x in str(r_m['Ma_KTV_Clean']).split(',') if x.strip() and x.strip() != 'nan']
            ten_ktvs = [x.strip() for x in str(r_m['Ten_KTV_Clean']).split(',') if x.strip() and x.strip() != 'nan']
            so_nguoi = max(len(ma_ktvs), 1)
            tien_chia = clean_num(r_m['Tien_Cong_Chot_Luong']) / so_nguoi
            for idx_k, m_k in enumerate(ma_ktvs):
                t_k = ten_ktvs[idx_k] if idx_k < len(ten_ktvs) else m_k
                ktv_sc_list.append({'Ma_KTV': m_k, 'Ten_KTV': t_k, 'Cong_SC': tien_chia})

        df_sc_ktv_sum = pd.DataFrame(ktv_sc_list).groupby(['Ma_KTV', 'Ten_KTV'])['Cong_SC'].sum().reset_index() if ktv_sc_list else pd.DataFrame(columns=['Ma_KTV', 'Ten_KTV', 'Cong_SC'])

        # 2. Công bảo hành (đã chia)
        if up_bh and not df_bh_split.empty:
            df_bh_target = df_bh_split.copy()
            if sel_thang_nam != "Tất cả các tháng" and 'Thang_BH' in df_bh_target.columns:
                df_bh_target = df_bh_target[df_bh_target['Thang_BH'] == sel_thang_nam]
            df_bh_ktv_sum = df_bh_target.groupby(['Ma_KTV_Clean', 'Ten_KTV_Clean'])['Cong_BH_Thuc_Nhan'].sum().reset_index().rename(
                columns={'Ma_KTV_Clean': 'Ma_KTV', 'Ten_KTV_Clean': 'Ten_KTV', 'Cong_BH_Thuc_Nhan': 'Cong_BH'}
            )
        else:
            df_bh_ktv_sum = pd.DataFrame(columns=['Ma_KTV', 'Ten_KTV', 'Cong_BH'])

        # 3. Công PDI
        ktv_pdi_list = []
        for _, r_p in df_pdi_only.iterrows():
            ma_ktvs = [x.strip() for x in str(r_p['Ma_KTV_Clean']).split(',') if x.strip() and x.strip() != 'nan']
            ten_ktvs = [x.strip() for x in str(r_p['Ten_KTV_Clean']).split(',') if x.strip() and x.strip() != 'nan']
            so_nguoi = max(len(ma_ktvs), 1)
            tien_chia = clean_num(r_p['Tien_KTV']) / so_nguoi
            for idx_k, m_k in enumerate(ma_ktvs):
                t_k = ten_ktvs[idx_k] if idx_k < len(ten_ktvs) else m_k
                ktv_pdi_list.append({'Ma_KTV': m_k, 'Ten_KTV': t_k, 'Cong_PDI': tien_chia})

        df_pdi_ktv_sum = pd.DataFrame(ktv_pdi_list).groupby(['Ma_KTV', 'Ten_KTV'])['Cong_PDI'].sum().reset_index() if ktv_pdi_list else pd.DataFrame(columns=['Ma_KTV', 'Ten_KTV', 'Cong_PDI'])

        # Gộp tất cả nguồn công
        df_luong_final = pd.merge(df_sc_ktv_sum, df_bh_ktv_sum, on=['Ma_KTV', 'Ten_KTV'], how='outer').fillna(0)
        df_luong_final = pd.merge(df_luong_final, df_pdi_ktv_sum, on=['Ma_KTV', 'Ten_KTV'], how='outer').fillna(0)
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
    st.info("💡 Vui lòng nạp tệp **'1. Sổ chi tiết TK 5114'** và **'2. Báo cáo doanh thu KTV'** phía trên để bắt đầu đối soát và tính lương.")
