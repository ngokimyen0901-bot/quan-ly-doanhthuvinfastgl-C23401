import streamlit as st
import pandas as pd
import numpy as np
import io
import re
import hashlib
from streamlit_gsheets import GSheetsConnection

st.set_page_config(
    page_title="Kiểm Soát Quyết Toán Bảo Hành (VinFast)",
    page_icon="🛡️",
    layout="wide"
)

st.title("🛡️ Kiểm Soát Đề Xuất Bảo Hành & Quyết Toán VinFast")
st.caption("☁️ Kết nối trực tiếp Google Sheets: **VinFast_Master_Database**.")

conn = st.connection("gsheets", type=GSheetsConnection)

def norm_lsc_key(val):
    if pd.isna(val) or val is None:
        return ""
    s = str(val).split('(')[0].strip().upper()
    return re.sub(r'[^A-Z0-9]', '', s)

def clean_num_bh(val):
    if pd.isna(val) or val is None:
        return 0.0
    s = str(val).replace(',', '').replace(' ', '').replace('đ', '').strip()
    try:
        return float(s)
    except Exception:
        return 0.0

def kiem_tra_sau_29082026_bh(lsc_str, ngay_str=""):
    m = re.search(r'WO-(\d{6})-', str(lsc_str).upper())
    if m:
        try:
            return int(m.group(1)) >= 260829
        except Exception:
            pass
    if pd.notna(ngay_str) and str(ngay_str).strip():
        dt = pd.to_datetime(ngay_str, errors='coerce', dayfirst=True)
        if pd.notna(dt):
            return dt >= pd.Timestamp("2026-08-29")
    return True

def tinh_ky_quyet_toan_bh(ngay_val):
    """Tính chu kỳ 23-22: Từ 29/08/2026 trở đi tối thiểu là Kỳ Tháng 09/2026"""
    if pd.isna(ngay_val) or not str(ngay_val).strip():
        return "Chưa duyệt"
    dt = pd.to_datetime(ngay_val, errors='coerce')
    if pd.isna(dt):
        return "Chưa duyệt"
    if dt < pd.Timestamp("2026-08-29"):
        return "Chưa duyệt"
    
    if dt.day <= 22:
        thang = dt.month
        nam = dt.year
    else:
        if dt.month == 12:
            thang = 1
            nam = dt.year + 1
        else:
            thang = dt.month + 1
            nam = dt.year
            
    if nam < 2026 or (nam == 2026 and thang < 9):
        thang = 9
        nam = 2026
        
    return f"Kỳ Tháng {thang:02d}/{nam}"

def xoa_sach_emoji(val):
    """Làm sạch emoji khi đẩy dữ liệu lên Google Sheets"""
    if pd.isna(val) or val is None:
        return ""
    s = str(val)
    return re.sub(r'[🔴🔵🟢⚪⏳⌛🚨💡🛠️🧾📋🔄🔎📌🚗🚕❌⚠️✅🔻]', '', s).strip()

def to_mau_trang_thai(val):
    s = str(val).lower()
    if ('trả về' in s or 'chưa upload' in s or 'chưa đề xuất' in s) and not any(k in s for k in ['bỏ qua', 'không duyệt', 'từ chối', 'không claim']):
        return 'background-color: #ffebee; color: #c62828; font-weight: bold;'
    elif any(k in s for k in ['chờ', 'hậu kiểm', 'chờ duyệt']) and not any(k in s for k in ['bỏ qua', 'không duyệt', 'từ chối', 'không claim']):
        return 'background-color: #e3f2fd; color: #1565c0; font-weight: bold;'
    return ''

def style_bac_thang(row):
    """Tô màu bậc thang: Dòng lệnh chính tô vàng nổi bật, dòng con giữ màu nhẹ"""
    cap = row.get('Cấp', '')
    if cap == 'CHA':
        return ['background-color: #fff9c4; font-weight: bold;' for _ in row]
    else:
        return ['color: #555555;' for _ in row]

if "ds_bo_qua" not in st.session_state:
    st.session_state.ds_bo_qua = set()

# ĐỌC MASTERDATA LÀM DỮ LIỆU ĐỒNG BỘ CHÉO
@st.cache_data(ttl=60)
def lay_du_lieu_master_data():
    try:
        df_m = conn.read(worksheet="MasterData", ttl=60)
        if df_m is not None and not df_m.empty:
            df_m.columns = [str(c).strip() for c in df_m.columns]
            col_lsc = next((c for c in df_m.columns if any(k in c.lower() for k in ['lệnh sửa chữa', 'lsc', 'wo'])), 'Số lệnh sửa chữa')
            df_m['lsc_norm'] = df_m[col_lsc].apply(norm_lsc_key)
            return df_m
    except Exception:
        pass
    return pd.DataFrame()

df_master_ref = lay_du_lieu_master_data()
df_bh_active = None
is_bh_from_gsheets = False

with st.expander("📥 Nạp Tệp Dữ Liệu Đối Soát (Chỉ cần nạp 2 tệp số 2 & 3 là chạy mượt mà)", expanded=False):
    st.info("💡 **Gợi ý:** Chỉ cần nạp **2 tệp cốt lõi (Chi tiết lệnh + Cổng ĐXBH)** là hệ thống tự chạy đối soát và lấy thông tin xe từ MasterData.")
    c_up1, c_up2 = st.columns(2)
    with c_up1:
        up_bh_db = st.file_uploader("1. Database Tổng (Tùy chọn - nếu không có sẽ tự lấy từ MasterData):", type=['csv', 'xlsx', 'xls'], key="bh_up_db")
        up_bh_ct = st.file_uploader("2. Chi Tiết Lệnh DMS (Bắt buộc - chitietlenh.csv):", type=['csv', 'xlsx', 'xls'], key="bh_up_ct")
    with c_up2:
        up_bh_dx = st.file_uploader("3. Cổng Đề Xuất BH (Bắt buộc - dexuatbaohanh.csv):", type=['csv', 'xlsx', 'xls'], key="bh_up_dx")
        up_bh_wcs = st.file_uploader("4. Bảng Kê Nhà Máy Duyệt (Tùy chọn khi có WCS):", type=['xlsx', 'xls'], key="bh_up_wcs")

if up_bh_ct and up_bh_dx:
    try:
        df_b_db = pd.read_excel(up_bh_db) if up_bh_db and up_bh_db.name.lower().endswith(('.xlsx', '.xls')) else (pd.read_csv(up_bh_db, low_memory=False) if up_bh_db else df_master_ref.copy())
        df_b_ct = pd.read_excel(up_bh_ct) if up_bh_ct.name.lower().endswith(('.xlsx', '.xls')) else pd.read_csv(up_bh_ct, low_memory=False)
        df_b_dx = pd.read_excel(up_bh_dx) if up_bh_dx.name.lower().endswith(('.xlsx', '.xls')) else pd.read_csv(up_bh_dx, low_memory=False)

        df_b_db.columns = [str(c).strip() for c in df_b_db.columns]
        df_b_ct.columns = [str(c).strip() for c in df_b_ct.columns]
        df_b_dx.columns = [str(c).strip() for c in df_b_dx.columns]

        col_b_db_lsc = next((c for c in df_b_db.columns if any(k in c.lower() for k in ['lệnh sửa chữa', 'lsc', 'wo'])), 'Số lệnh sửa chữa')
        col_b_ct_lsc = next((c for c in df_b_ct.columns if any(k in c.lower() for k in ['lệnh sửa chữa', 'lsc', 'wo'])), 'Lệnh sửa chữa')
        col_b_dx_lsc = next((c for c in df_b_dx.columns if any(k in c.lower() for k in ['lệnh sửa chữa', 'lsc', 'wo'])), 'Lệnh sửa chữa')

        df_b_db['lsc_norm'] = df_b_db[col_b_db_lsc].apply(norm_lsc_key)
        df_b_ct['lsc_norm'] = df_b_ct[col_b_ct_lsc].apply(norm_lsc_key)
        df_b_dx['lsc_norm'] = df_b_dx[col_b_dx_lsc].apply(norm_lsc_key)

        col_b_bs = next((c for c in df_b_db.columns if 'biển số' in c.lower()), 'Biển số')
        col_b_cv = next((c for c in df_b_db.columns if 'cố vấn' in c.lower() and 'id' not in c.lower()), 'Cố vấn dịch vụ')
        col_b_tg = next((c for c in df_b_db.columns if 'thời gian đóng' in c.lower()), 'Thời gian đóng LSC')
        col_b_sk = next((c for c in df_b_db.columns if 'khung' in c.lower() or 'vin' in c.lower()), 'Số khung')

        col_b_mota = next((c for c in df_b_ct.columns if 'mô tả sản phẩm' in c.lower()), 'Mô tả sản phẩm')
        col_b_masp = next((c for c in df_b_ct.columns if 'mã sản phẩm' in c.lower() or 'mã vật tư' in c.lower()), 'Mã sản phẩm')
        df_b_ct['is_PDI'] = (df_b_ct[col_b_mota].astype(str).str.contains('Pre-delivery', case=False, na=False)) | (df_b_ct[col_b_masp].astype(str).str.strip() == '5301003')

        col_b_st = next((c for c in df_b_db.columns if any(k in c.lower() for k in ['trạng thái lệnh', 'trạng thái lsc', 'trạng thái sửa chữa', 'status']) or c.lower().strip() == 'trạng thái'), None)
        if col_b_st and col_b_st in df_b_db.columns:
            map_b_status = df_b_db.set_index('lsc_norm')[col_b_st].to_dict()
            df_b_ct['Trạng thái DMS'] = df_b_ct['lsc_norm'].map(map_b_status).fillna('Chưa xác định')
        else:
            df_b_ct['Trạng thái DMS'] = 'Đã đóng'

        col_b_pb = next((c for c in df_b_ct.columns if 'classification' in c.lower() or 'p/bill' in c.lower()), 'P/bill classification')
        df_b_w = df_b_ct[(df_b_ct[col_b_pb].astype(str).str.strip().str.upper() == 'W') & (~df_b_ct['is_PDI'])].copy()
        df_b_w = df_b_w[df_b_w[col_b_ct_lsc].apply(kiem_tra_sau_29082026_bh)].copy()

        col_b_ct_mavt = next((c for c in df_b_w.columns if 'mã sản phẩm' in c.lower() or 'mã vật tư' in c.lower()), 'Mã sản phẩm')
        col_b_dx_mavt = next((c for c in df_b_dx.columns if 'mã vật tư' in c.lower() or 'mã sản phẩm' in c.lower()), 'Mã vật tư')
        df_b_w['mavt_norm'] = df_b_w[col_b_ct_mavt].apply(norm_lsc_key)
        df_b_dx['mavt_norm'] = df_b_dx[col_b_dx_mavt].apply(norm_lsc_key)

        col_b_ct_loai = next((c for c in df_b_w.columns if 'loại sản phẩm' in c.lower()), 'Loại sản phẩm')
        col_b_ct_sl = next((c for c in df_b_w.columns if 'số lượng' in c.lower()), 'Số lượng/Nhân công')
        col_b_ct_tien_pt = next((c for c in df_b_w.columns if 'tiền phụ tùng' in c.lower()), 'Tổng tiền phụ tùng')
        col_b_ct_tien_nc = next((c for c in df_b_w.columns if 'tiền nhân công' in c.lower()), 'Tổng tiền nhân công')

        df_b_w['SL_DMS'] = df_b_w[col_b_ct_sl].apply(clean_num_bh)
        df_b_w['TienCong_DMS'] = df_b_w[col_b_ct_tien_nc].apply(clean_num_bh)
        df_b_w['Tiền xưởng (DMS)'] = df_b_w.apply(
            lambda r: clean_num_bh(r.get(col_b_ct_tien_pt, 0)) if clean_num_bh(r.get(col_b_ct_tien_pt, 0)) > 0 else clean_num_bh(r.get(col_b_ct_tien_nc, 0)), axis=1
        )

        col_b_dx_sl = next((c for c in df_b_dx.columns if 'số lượng' in c.lower()), 'Số lượng/số giờ thực tế')
        col_b_dx_tien = next((c for c in df_b_dx.columns if 'tổng số tiền' in c.lower()), 'Tổng số tiền')
        col_b_dx_so_dx = next((c for c in df_b_dx.columns if 'số đề xuất' in c.lower()), 'Số đề xuất bảo hành')
        col_b_dx_tt = next((c for c in df_b_dx.columns if 'trạng thái' in c.lower() and 'phê duyệt' not in c.lower()), 'Trạng thái')
        col_b_dx_ngay_pd = next((c for c in df_b_dx.columns if 'ngày phê duyệt' in c.lower()), 'Ngày phê duyệt')

        df_b_dx['SL_DXBH'] = df_b_dx[col_b_dx_sl].apply(clean_num_bh)
        df_b_dx['Tien_DXBH'] = df_b_dx[col_b_dx_tien].apply(clean_num_bh)
        df_b_dx['Kỳ_Tính_Toán_DXBH'] = df_b_dx[col_b_dx_ngay_pd].apply(tinh_ky_quyet_toan_bh)
        df_b_dx['Ngày_Duyệt_Format'] = pd.to_datetime(df_b_dx[col_b_dx_ngay_pd], errors='coerce').dt.strftime('%d/%m/%Y %H:%M').fillna('Chưa duyệt')

        cols_sub_dx = ['lsc_norm', 'mavt_norm', col_b_dx_so_dx, col_b_dx_tt, 'Ngày_Duyệt_Format', 'Kỳ_Tính_Toán_DXBH', 'SL_DXBH', 'Tien_DXBH']
        df_b_dx_sub = df_b_dx[[c for c in cols_sub_dx if c in df_b_dx.columns]].drop_duplicates(subset=['lsc_norm', 'mavt_norm'])
        df_b_merged = pd.merge(df_b_w, df_b_dx_sub, on=['lsc_norm', 'mavt_norm'], how='left')

        ky_wcs_ten_txt = ""
        if up_bh_wcs:
            xls_w = pd.ExcelFile(up_bh_wcs)
            sheet_w_target = 'Chi tiết_Đối soát' if 'Chi tiết_Đối soát' in xls_w.sheet_names else xls_w.sheet_names[0]
            df_b_wcs_raw = pd.read_excel(up_bh_wcs, sheet_name=sheet_w_target)
            df_b_wcs_raw.columns = [str(c).strip() for c in df_b_wcs_raw.columns]

            col_wcs_wo = next((c for c in df_b_wcs_raw.columns if c.upper() == 'WO' or 'lệnh' in c.lower()), 'WO')
            col_wcs_mat = next((c for c in df_b_wcs_raw.columns if 'material1' in c.lower() or 'mã' in c.lower()), 'Material1')
            col_wcs_sl = next((c for c in df_b_wcs_raw.columns if c.upper() == 'SL' or 'số lượng' in c.lower()), 'SL')
            col_wcs_amt = next((c for c in df_b_wcs_raw.columns if 'amount' in c.lower() or 'tiền' in c.lower()), 'Amount')
            col_wcs_bk = next((c for c in df_b_wcs_raw.columns if 'số bảng kê' in c.lower() or 'mã đl' in c.lower()), 'Mã ĐL & Số bảng kê')

            df_b_wcs_raw['lsc_norm'] = df_b_wcs_raw[col_wcs_wo].apply(norm_lsc_key)
            df_b_wcs_raw['mavt_norm'] = df_b_wcs_raw[col_wcs_mat].apply(norm_lsc_key)
            df_b_wcs_raw['SL_WCS'] = df_b_wcs_raw[col_wcs_sl].apply(clean_num_bh)
            df_b_wcs_raw['Tien_WCS'] = df_b_wcs_raw[col_wcs_amt].apply(clean_num_bh)

            if col_wcs_bk in df_b_wcs_raw.columns and not df_b_wcs_raw[col_b_wcs_bk].dropna().empty:
                wcs_code_txt = str(df_b_wcs_raw[col_b_wcs_bk].dropna().iloc[0]).strip()
                m_ky = re.search(r'WCS-(\d{2})-(\d{2})-(\d{4})', wcs_code_txt)
                if m_ky:
                    ky_wcs_ten_txt = f"Bảng kê {wcs_code_txt} (Kỳ {int(m_ky.group(3)):02d} - Tháng {m_ky.group(2)}/20{m_ky.group(1)})"
                else:
                    ky_wcs_ten_txt = f"Bảng kê {wcs_code_txt}"

            df_b_wcs_raw['Kỳ_WCS_Duyệt'] = ky_wcs_ten_txt
            df_b_wcs_sub = df_b_wcs_raw[['lsc_norm', 'mavt_norm', 'SL_WCS', 'Tien_WCS', 'Kỳ_WCS_Duyệt']].drop_duplicates(subset=['lsc_norm', 'mavt_norm'])
            df_b_merged = pd.merge(df_b_merged, df_b_wcs_sub, on=['lsc_norm', 'mavt_norm'], how='left')
        else:
            df_b_merged['SL_WCS'] = np.nan
            df_b_merged['Tien_WCS'] = np.nan
            df_b_merged['Kỳ_WCS_Duyệt'] = ""

        df_b_db_info = df_b_db[['lsc_norm'] + [c for c in [col_b_bs, col_b_cv, col_b_tg, col_b_sk] if c in df_b_db.columns]].drop_duplicates(subset=['lsc_norm']).copy()
        if 'Biển số' not in df_b_merged.columns and col_b_bs in df_b_db_info.columns:
            df_b_merged['Biển số'] = df_b_merged['lsc_norm'].map(df_b_db_info.set_index('lsc_norm')[col_b_bs].to_dict()).fillna('')
        if 'Cố vấn dịch vụ' not in df_b_merged.columns and col_b_cv in df_b_db_info.columns:
            df_b_merged['Cố vấn dịch vụ'] = df_b_merged['lsc_norm'].map(df_b_db_info.set_index('lsc_norm')[col_b_cv].to_dict()).fillna('')
        if 'Thời gian đóng LSC' not in df_b_merged.columns and col_b_tg in df_b_db_info.columns:
            df_b_merged['Thời gian đóng LSC'] = df_b_merged['lsc_norm'].map(df_b_db_info.set_index('lsc_norm')[col_b_tg].to_dict()).fillna('')
        if 'Số khung' not in df_b_merged.columns and col_b_sk in df_b_db_info.columns:
            df_b_merged['Số khung'] = df_b_merged['lsc_norm'].map(df_b_db_info.set_index('lsc_norm')[col_b_sk].to_dict()).fillna('')

        def dinh_danh_ky_cu_the_bh(r):
            if pd.notna(r.get('Tien_WCS')) and clean_num_bh(r.get('Tien_WCS')) > 0:
                return r.get('Kỳ_WCS_Duyệt', ky_wcs_ten_txt)
            ky_dx = r.get('Kỳ_Tính_Toán_DXBH', 'Chưa duyệt')
            if ky_dx != 'Chưa duyệt':
                m_t = re.search(r'Tháng \d{1,2}/\d{4}', ky_dx)
                t_str = m_t.group(0) if m_t else ky_dx
                return f"Cổng VF ({t_str} - Chờ WCS)"
            return "Chưa duyệt"

        df_b_merged['WCS Số Kỳ Cụ Thể'] = df_b_merged.apply(dinh_danh_ky_cu_the_bh, axis=1)

        ro_da_tung_claim_set = set(df_b_merged[df_b_merged[col_b_dx_so_dx].notna() & (df_b_merged[col_b_dx_so_dx].astype(str).str.strip() != '')]['lsc_norm'])

        def gan_nhan_chuyen_sau_bh(r):
            lsc = r.get('lsc_norm')
            if lsc in st.session_state.ds_bo_qua:
                return "⚪ BỎ QUA"
            
            tt_dx = str(r.get(col_b_dx_tt, '')).strip()
            tt_lower = tt_dx.lower()
            
            if any(k in tt_lower for k in ['không duyệt', 'từ chối', 'bỏ qua']):
                return f"⚪ {tt_dx.upper()}"
            
            if pd.isna(r.get(col_b_dx_so_dx)) or not str(r.get(col_b_dx_so_dx)).strip():
                if lsc in ro_da_tung_claim_set:
                    return "⚪ KHÔNG CLAIM (Xưởng tự chịu)"
                return "🔴 CHƯA UPLOAD ĐỀ XUẤT"

            if 'trả về' in tt_lower:
                return f"🔴 {tt_dx.upper()}"
            if pd.notna(r.get('Tien_WCS')) and clean_num_bh(r.get('Tien_WCS')) > 0:
                return "🟢 ĐÃ DUYỆT (Chốt Bảng Kê WCS)"
            if "cấp 3" in tt_lower and "chờ" in tt_lower:
                return "🔵 Chờ phê duyệt cấp 3"
            if "cấp 2" in tt_lower and "chờ" in tt_lower:
                return "🔵 Chờ phê duyệt cấp 2"
            if "cấp 1" in tt_lower and "chờ" in tt_lower:
                return "🔵 Chờ phê duyệt cấp 1"
            if "hậu kiểm" in tt_lower:
                return "🔵 Chờ hậu kiểm"
            if tt_lower == "phê duyệt" or ("phê duyệt" in tt_lower and "chờ" not in tt_lower):
                ky_tinh = r.get('Kỳ_Tính_Toán_DXBH', 'Chưa duyệt')
                m_t = re.search(r'Tháng \d{1,2}/\d{4}', ky_tinh)
                t_str = m_t.group(0) if m_t else ky_tinh
                return f"🟢 ĐÃ DUYỆT CỔNG ({t_str})"
            return f"🔵 {tt_dx}"

        df_b_merged['Nhãn Trạng Thái'] = df_b_merged.apply(gan_nhan_chuyen_sau_bh, axis=1)
        df_bh_active = df_b_merged.copy()

        hash_str = str(len(df_b_ct)) + str(len(df_b_dx))
        current_bh_hash = hashlib.md5(hash_str.encode()).hexdigest()

        if "last_bh_sync_hash" not in st.session_state:
            st.session_state.last_bh_sync_hash = ""

        if st.session_state.last_bh_sync_hash != current_bh_hash:
            try:
                with st.spinner("⏳ Đang lưu kho vào sheet ChiTiet_BaoHanh..."):
                    cols_bh_sync = [
                        col_b_ct_lsc, 'Biển số', 'Số khung', 'Cố vấn dịch vụ', 'Trạng thái DMS', 'Thời gian đóng LSC',
                        col_b_ct_loai, col_b_ct_mavt, col_b_mota, 'SL_DMS', 'SL_WCS',
                        'Tiền xưởng (DMS)', 'Tien_DXBH', 'Tien_WCS',
                        'Ngày_Duyệt_Format', 'WCS Số Kỳ Cụ Thể', 'Kỳ_Tính_Toán_DXBH', 'Nhãn Trạng Thái', col_b_dx_so_dx
                    ]
                    cols_valid_sync = [c for c in cols_bh_sync if c in df_b_merged.columns]
                    df_bh_sync = df_b_merged[cols_valid_sync].copy()

                    for c_num in ['SL_DMS', 'SL_WCS', 'Tiền xưởng (DMS)', 'Tien_DXBH', 'Tien_WCS']:
                        if c_num in df_bh_sync.columns:
                            df_bh_sync[c_num] = pd.to_numeric(df_bh_sync[c_num], errors='coerce').fillna(0)

                    for c_txt in df_bh_sync.select_dtypes(include='object').columns:
                        df_bh_sync[c_txt] = df_bh_sync[c_txt].apply(xoa_sach_emoji)

                    conn.update(worksheet="ChiTiet_BaoHanh", data=df_bh_sync)
                    st.session_state.last_bh_sync_hash = current_bh_hash
                    st.toast("✅ Đã cập nhật xong sheet ChiTiet_BaoHanh!", icon="🚀")
            except Exception as e_gs_bh:
                st.warning(f"⚠️ Chưa ghi được lên Google Sheets: {e_gs_bh}")

    except Exception as e_proc_bh:
        st.error(f"❌ Có lỗi khi phân tích: {str(e_proc_bh)}")

elif df_bh_active is None:
    try:
        with st.spinner("🔄 Đang tải từ sheet ChiTiet_BaoHanh..."):
            df_bh_from_gsheet = conn.read(worksheet="ChiTiet_BaoHanh", ttl=0)
            if df_bh_from_gsheet is not None and not df_bh_from_gsheet.empty:
                df_bh_active = df_bh_from_gsheet.copy()
                is_bh_from_gsheets = True
                
                col_b_ct_lsc = next((c for c in df_bh_active.columns if 'lệnh sửa chữa' in c.lower()), 'Lệnh sửa chữa')
                col_b_ct_mavt = next((c for c in df_bh_active.columns if 'mã sản phẩm' in c.lower() or 'mã vật tư' in c.lower()), 'Mã sản phẩm')
                col_b_mota = next((c for c in df_bh_active.columns if 'mô tả sản phẩm' in c.lower()), 'Mô tả sản phẩm')
                col_b_ct_loai = next((c for c in df_bh_active.columns if 'loại sản phẩm' in c.lower()), 'Loại sản phẩm')
                
                df_bh_active['lsc_norm'] = df_bh_active[col_b_ct_lsc].apply(norm_lsc_key)

                if not df_master_ref.empty and 'lsc_norm' in df_master_ref.columns:
                    map_bs = df_master_ref.set_index('lsc_norm')['Biển số'].to_dict() if 'Biển số' in df_master_ref.columns else {}
                    map_cv = df_master_ref.set_index('lsc_norm')['Cố vấn dịch vụ'].to_dict() if 'Cố vấn dịch vụ' in df_master_ref.columns else {}
                    map_tt = df_master_ref.set_index('lsc_norm')['Trạng thái'].to_dict() if 'Trạng thái' in df_master_ref.columns else {}
                    map_tg = df_master_ref.set_index('lsc_norm')['Thời gian đóng LSC'].to_dict() if 'Thời gian đóng LSC' in df_master_ref.columns else {}

                    if 'Biển số' not in df_bh_active.columns or df_bh_active['Biển số'].dropna().empty:
                        df_bh_active['Biển số'] = df_bh_active['lsc_norm'].map(map_bs).fillna('')
                    if 'Cố vấn dịch vụ' not in df_bh_active.columns or df_bh_active['Cố vấn dịch vụ'].dropna().empty:
                        df_bh_active['Cố vấn dịch vụ'] = df_bh_active['lsc_norm'].map(map_cv).fillna('')
                    if 'Trạng thái DMS' not in df_bh_active.columns:
                        df_bh_active['Trạng thái DMS'] = df_bh_active['lsc_norm'].map(map_tt).fillna('Đã đóng')
                    if 'Thời gian đóng LSC' not in df_bh_active.columns:
                        df_bh_active['Thời gian đóng LSC'] = df_bh_active['lsc_norm'].map(map_tg).fillna('')

                if 'Trạng thái DMS' not in df_bh_active.columns:
                    df_bh_active['Trạng thái DMS'] = 'Đã đóng'
                if 'WCS Số Kỳ Cụ Thể' not in df_bh_active.columns:
                    df_bh_active['WCS Số Kỳ Cụ Thể'] = 'Chưa duyệt'
                if 'Kỳ_Tính_Toán_DXBH' not in df_bh_active.columns:
                    col_ngay_pd_old = next((c for c in df_bh_active.columns if 'phê duyệt' in c.lower() and 'ngày' in c.lower()), None)
                    df_bh_active['Kỳ_Tính_Toán_DXBH'] = df_bh_active[col_ngay_pd_old].apply(tinh_ky_quyet_toan_bh) if col_ngay_pd_old else 'Chưa duyệt'
                if 'Nhãn Trạng Thái' not in df_bh_active.columns:
                    col_tt_old = next((c for c in df_bh_active.columns if 'trạng thái' in c.lower()), None)
                    df_bh_active['Nhãn Trạng Thái'] = df_bh_active[col_tt_old].fillna('Chưa duyệt') if col_tt_old else 'Chưa duyệt'
                if 'Ngày_Duyệt_Format' not in df_bh_active.columns:
                    col_ngay_pd_old = next((c for c in df_bh_active.columns if 'phê duyệt' in c.lower() and 'ngày' in c.lower()), None)
                    df_bh_active['Ngày_Duyệt_Format'] = pd.to_datetime(df_bh_active[col_ngay_pd_old], errors='coerce').dt.strftime('%d/%m/%Y %H:%M').fillna('Chưa duyệt') if col_ngay_pd_old else 'Chưa duyệt'
                if 'Tiền xưởng (DMS)' not in df_bh_active.columns:
                    df_bh_active['Tiền xưởng (DMS)'] = 0.0
                if 'Tien_DXBH' not in df_bh_active.columns:
                    col_t_dx = next((c for c in df_bh_active.columns if 'claim' in c.lower() or 'tiền' in c.lower()), None)
                    df_bh_active['Tien_DXBH'] = df_bh_active[col_t_dx].apply(clean_num_bh) if col_t_dx else 0.0
                if 'Tien_WCS' not in df_bh_active.columns:
                    df_bh_active['Tien_WCS'] = 0.0
                if 'SL_DMS' not in df_bh_active.columns:
                    df_bh_active['SL_DMS'] = 1.0
                if 'SL_WCS' not in df_bh_active.columns:
                    df_bh_active['SL_WCS'] = 0.0
    except Exception:
        pass

# ----------------- KHU VỰC BỘ LỌC THÔNG MINH GỌN GÀNG -----------------
if df_bh_active is not None and not df_bh_active.empty:
    df_bh_m = df_bh_active.copy()

    fl_b1, fl_b2, fl_b3 = st.columns([2.5, 3.5, 4.0])
    with fl_b1:
        list_b_tt = ["Tất cả trạng thái"] + sorted(list(df_bh_m['Trạng thái DMS'].dropna().unique()))
        sel_b_tt = st.selectbox("1. Trạng Thái LSC:", list_b_tt, index=list_b_tt.index("Đã đóng") if "Đã đóng" in list_b_tt else 0)

    with fl_b2:
        ky_thuc_te_set = set()
        for v in df_bh_m['Kỳ_Tính_Toán_DXBH'].dropna().unique():
            m_ky = re.search(r'Tháng (\d{1,2})/(\d{4})', str(v))
            if m_ky:
                thang_num, nam_num = int(m_ky.group(1)), int(m_ky.group(2))
                if (nam_num > 2026) or (nam_num == 2026 and thang_num >= 9):
                    ky_thuc_te_set.add((nam_num, thang_num))

        sorted_kys = sorted(list(ky_thuc_te_set), key=lambda x: (x[0], x[1]), reverse=True)
        list_b_ky = ["Tất cả các kỳ"] + [f"Kỳ Tháng {m:02d}/{y}" for (y, m) in sorted_kys] + ["Chưa duyệt"]
        sel_b_ky = st.selectbox("2. Kỳ Quyết Toán (Chu kỳ 23-22):", list_b_ky, index=0)

    with fl_b3:
        loc_b_nhan = st.multiselect("3. Trạng Thái Duyệt:", df_bh_m['Nhãn Trạng Thái'].dropna().unique(), default=df_bh_m['Nhãn Trạng Thái'].dropna().unique())

    df_b_filtered = df_bh_m.copy()
    if sel_b_tt != "Tất cả trạng thái":
        df_b_filtered = df_b_filtered[df_b_filtered['Trạng thái DMS'] == sel_b_tt]
    
    if sel_b_ky != "Tất cả các kỳ":
        if sel_b_ky == "Chưa duyệt":
            df_b_filtered = df_b_filtered[df_b_filtered['Kỳ_Tính_Toán_DXBH'] == "Chưa duyệt"]
        else:
            m_sub = re.search(r'Tháng \d{1,2}/\d{4}', sel_b_ky)
            if m_sub:
                chuoi_thang = m_sub.group(0)
                df_b_filtered = df_b_filtered[
                    df_b_filtered['WCS Số Kỳ Cụ Thể'].astype(str).str.contains(chuoi_thang, na=False) |
                    df_b_filtered['Kỳ_Tính_Toán_DXBH'].astype(str).str.contains(chuoi_thang, na=False)
                ]
            else:
                df_b_filtered = df_b_filtered[df_b_filtered['WCS Số Kỳ Cụ Thể'].astype(str).str.contains(sel_b_ky, na=False)]

    if loc_b_nhan:
        df_b_filtered = df_b_filtered[df_b_filtered['Nhãn Trạng Thái'].isin(loc_b_nhan)]

    def is_loi_can_sua_gap(val):
        s = str(val).lower()
        if any(k in s for k in ['không duyệt', 'từ chối', 'bỏ qua', 'không claim']):
            return False
        return any(k in s for k in ['trả về', 'reject'])

    def is_chua_up(val):
        s = str(val).lower()
        if any(k in s for k in ['không duyệt', 'từ chối', 'bỏ qua', 'không claim']):
            return False
        return 'chưa upload' in s or 'chưa đề xuất' in s

    def is_bo_qua_hoac_khong_duyet(val):
        s = str(val).lower()
        return any(k in s for k in ['bỏ qua', 'không duyệt', 'từ chối', 'không claim'])

    def is_da_duyet(val):
        s = str(val).lower()
        return 'đã duyệt' in s or 'duyệt cổng' in s

    cnt_b_ro_total = df_b_filtered['lsc_norm'].nunique()
    cnt_b_do = df_b_filtered[df_b_filtered['Nhãn Trạng Thái'].apply(lambda x: is_loi_can_sua_gap(x) or is_chua_up(x))]['lsc_norm'].nunique()
    cnt_b_cho = df_b_filtered[df_b_filtered['Nhãn Trạng Thái'].str.contains('chờ|hậu kiểm', case=False, na=False)]['lsc_norm'].nunique()
    cnt_b_ok = df_b_filtered[df_b_filtered['Nhãn Trạng Thái'].apply(is_da_duyet)]['lsc_norm'].nunique()

    kb1, kb2, kb3, kb4 = st.columns(4)
    kb1.metric("📌 Tổng RO Bảo Hành", f"{cnt_b_ro_total} RO", f"{len(df_b_filtered)} mục")
    kb2.metric("🔴 Trả Về / Chưa Upload", f"{cnt_b_do} RO", delta=f"{cnt_b_do} ca cần sửa" if cnt_b_do > 0 else "0", delta_color="inverse")
    kb3.metric("🔵 Đang Chờ Phê Duyệt", f"{cnt_b_cho} RO", delta="Cấp 1/2/3", delta_color="off")
    kb4.metric("🟢 Đã Duyệt Quyết Toán", f"{cnt_b_ok} RO")

    t_tab_do, t_tab_xanh, t_b_ro, t_b_split, t_b_detail, t_b_ignore, t_b_inv = st.tabs([
        "🔴 Bị Trả Về",
        "🔵 Chưa Claim",
        "📋 Tất Cả LSC",
        "🔄 Tách Đợt",
        "🔎 Chi Tiết Vật Tư",
        "⚪ Bỏ Qua",
        "🧾 Hóa Đơn NM"
    ])

    # HÀM TẠO BẢNG BẬC THANG DUY NHẤT (VĂN BẢN SẠCH 100%, KHÔNG KÝ TỰ LẠ, DỄ COPY)
    def render_tree_table(df_grouped, df_raw_all, empty_msg, tab_key):
        if df_grouped.empty:
            st.success(empty_msg)
            return

        c_t1, c_t2 = st.columns([6, 4])
        with c_t1:
            st.markdown("💡 *Bảng phân cấp chuẩn:* Dòng lệnh chính tô màu vàng, các dòng chi tiết phụ tùng/công việc thụt lề ngay bên dưới.")
        with c_t2:
            show_detail = st.toggle("📂 Bung chi tiết công việc bậc thang (như Excel)", value=True, key=f"toggle_tree_{tab_key}")

        if not show_detail:
            # Chế độ thu gọn: chỉ hiện dòng lệnh cha
            cols_parent_show = ['STT', col_b_ct_lsc, 'Biển số', 'Cố vấn dịch vụ', 'Số mục', 'Tiền xưởng (no VAT)', 'Lý do / Trạng thái']
            df_g_show = df_grouped.copy()
            df_g_show.insert(0, 'STT', range(1, len(df_g_show) + 1))
            cfg_simple = {
                "STT": st.column_config.NumberColumn("STT", width="small"),
                col_b_ct_lsc: st.column_config.TextColumn("Lệnh sửa chữa", width="medium"),
                "Biển số": st.column_config.TextColumn("Biển số", width="small"),
                "Cố vấn dịch vụ": st.column_config.TextColumn("Cố vấn dịch vụ", width="medium"),
                "Số mục": st.column_config.NumberColumn("Số mục", width="small"),
                "Tiền xưởng (no VAT)": st.column_config.NumberColumn(format="%,d đ", width="medium"),
                "Lý do / Trạng thái": st.column_config.TextColumn("Lý do / Trạng thái", width="large")
            }
            st.dataframe(
                df_g_show[[c for c in cols_parent_show if c in df_g_show.columns]].style.map(to_mau_trang_thai, subset=['Lý do / Trạng thái']),
                use_container_width=True,
                hide_index=True,
                column_config=cfg_simple
            )
        else:
            # Chế độ bậc thang (Tree-Table): Văn bản sạch hoàn toàn, không icon tam giác để dễ copy
            rows_tree = []
            stt_p = 1
            for _, r_p in df_grouped.iterrows():
                lsc_val = r_p[col_b_ct_lsc]
                # DÒNG CHA (LỆNH TỔNG): Để nguyên mã LSC sạch, không icon
                rows_tree.append({
                    'STT': str(stt_p),
                    'Cấp': 'CHA',
                    'Lệnh sửa chữa / Chi tiết': str(lsc_val),
                    'Biển số': r_p.get('Biển số', ''),
                    'Cố vấn dịch vụ': r_p.get('Cố vấn dịch vụ', ''),
                    'Mã / Tên công việc & Phụ tùng': f"[TỔNG CỘNG {r_p.get('Số mục', 1)} MỤC]",
                    'Phân loại': 'Lệnh tổng',
                    'SL / Giờ': None,
                    'Tiền xưởng (no VAT)': r_p.get('Tiền xưởng (no VAT)', 0),
                    'Trạng thái': r_p.get('Lý do / Trạng thái', '')
                })

                # CÁC DÒNG CON: Để mã vật tư sạch, thụt lề bằng 4 dấu cách
                df_sub = df_raw_all[df_raw_all[col_b_ct_lsc] == lsc_val].copy()
                stt_c = 1
                for _, r_c in df_sub.iterrows():
                    mavt_c = str(r_c.get(col_b_ct_mavt, '')).strip()
                    mota_c = str(r_c.get(col_b_mota, '')).strip()
                    loai_c = str(r_c.get(col_b_ct_loai, '')).strip()
                    sl_c = r_c.get('SL_DMS', 1)
                    tien_c = r_c.get('Tiền xưởng (DMS)', 0)
                    tt_c = r_c.get('Nhãn Trạng Thái', '')

                    rows_tree.append({
                        'STT': f"{stt_p}.{stt_c}",
                        'Cấp': 'CON',
                        'Lệnh sửa chữa / Chi tiết': f"    {mavt_c}",
                        'Biển số': '',
                        'Cố vấn dịch vụ': '',
                        'Mã / Tên công việc & Phụ tùng': mota_c,
                        'Phân loại': loai_c,
                        'SL / Giờ': sl_c,
                        'Tiền xưởng (no VAT)': tien_c,
                        'Trạng thái': tt_c
                    })
                    stt_c += 1
                stt_p += 1

            df_tree_res = pd.DataFrame(rows_tree)
            cols_show_tree = ['STT', 'Lệnh sửa chữa / Chi tiết', 'Biển số', 'Cố vấn dịch vụ', 'Mã / Tên công việc & Phụ tùng', 'Phân loại', 'SL / Giờ', 'Tiền xưởng (no VAT)', 'Trạng thái']
            cfg_tree = {
                "STT": st.column_config.TextColumn("STT", width="small"),
                "Lệnh sửa chữa / Chi tiết": st.column_config.TextColumn("Lệnh sửa chữa / Mã SP", width="medium"),
                "Biển số": st.column_config.TextColumn("Biển số", width="small"),
                "Cố vấn dịch vụ": st.column_config.TextColumn("Cố vấn dịch vụ", width="medium"),
                "Mã / Tên công việc & Phụ tùng": st.column_config.TextColumn("Mô tả công việc & Phụ tùng", width="large"),
                "Phân loại": st.column_config.TextColumn("Loại", width="small"),
                "SL / Giờ": st.column_config.NumberColumn("SL / Giờ", width="small"),
                "Tiền xưởng (no VAT)": st.column_config.NumberColumn(format="%,d đ", width="medium"),
                "Trạng thái": st.column_config.TextColumn("Trạng thái", width="large")
            }

            st_styled_tree = df_tree_res[cols_show_tree].style.apply(style_bac_thang, axis=1)
            st.dataframe(
                st_styled_tree,
                use_container_width=True,
                hide_index=True,
                column_config=cfg_tree
            )

    # TAB 1: BỊ TRẢ VỀ (BẢNG BẬC THANG SẠCH)
    with t_tab_do:
        st.subheader("🔴 Danh Sách Lệnh Bị Nhà Máy Trả Về (Bảng Phân Cấp Bậc Thang)")
        df_do_only = df_b_filtered[df_b_filtered['Nhãn Trạng Thái'].apply(is_loi_can_sua_gap)].copy()
        if not df_do_only.empty:
            df_do_grouped = df_do_only.groupby(['lsc_norm', col_b_ct_lsc]).agg({
                'Biển số': 'first',
                'Cố vấn dịch vụ': 'first',
                'Trạng thái DMS': 'first',
                'Thời gian đóng LSC': 'first',
                col_b_ct_mavt: 'count',
                'Tiền xưởng (DMS)': 'sum',
                'Tien_DXBH': 'sum',
                'Nhãn Trạng Thái': lambda x: " | ".join(sorted(set(x)))
            }).reset_index().rename(columns={
                col_b_ct_mavt: 'Số mục',
                'Tiền xưởng (DMS)': 'Tiền xưởng (no VAT)',
                'Tien_DXBH': 'Tiền ĐXBH (no VAT)',
                'Nhãn Trạng Thái': 'Lý do / Trạng thái'
            })
            render_tree_table(df_do_grouped, df_bh_m, "🎉 Không có lệnh nào bị trả về trong kỳ lọc này.", "tab_do")
        else:
            st.success("🎉 Không có lệnh nào bị trả về trong kỳ lọc này.")

    # TAB 2: CHƯA UPLOAD ĐỀ XUẤT (BẢNG BẬC THANG SẠCH)
    with t_tab_xanh:
        st.subheader("🔵 Danh Sách Xe Hoàn Thành Chưa Upload Đề Xuất (Bảng Phân Cấp Bậc Thang)")
        df_xanh_only = df_b_filtered[df_b_filtered['Nhãn Trạng Thái'].apply(is_chua_up)].copy()
        ro_co_muc_claim_roi = set(df_bh_m[df_bh_m['Nhãn Trạng Thái'].apply(lambda x: is_da_duyet(x) or ('chờ' in str(x).lower()))]['lsc_norm'])
        df_xanh_that_su = df_xanh_only[~df_xanh_only['lsc_norm'].isin(ro_co_muc_claim_roi)].copy()

        if not df_xanh_that_su.empty:
            df_xanh_grouped = df_xanh_that_su.groupby(['lsc_norm', col_b_ct_lsc]).agg({
                'Biển số': 'first',
                'Cố vấn dịch vụ': 'first',
                'Trạng thái DMS': 'first',
                'Thời gian đóng LSC': 'first',
                col_b_ct_mavt: 'count',
                'Tiền xưởng (DMS)': 'sum'
            }).reset_index().rename(columns={
                col_b_ct_mavt: 'Số mục',
                'Tiền xưởng (DMS)': 'Tiền xưởng (no VAT)'
            })
            df_xanh_grouped['Lý do / Trạng thái'] = "🔴 CHƯA TẠO ĐỀ XUẤT PORTAL"
            render_tree_table(df_xanh_grouped, df_bh_m, "🎉 Toàn bộ xe đã được tạo đề xuất claim lên cổng VinFast!", "tab_xanh")
        else:
            st.success("🎉 Toàn bộ xe đã được tạo đề xuất claim lên cổng VinFast!")

    # TAB 3: TOÀN BỘ DANH SÁCH LSC (GIỮ BẢNG LỚN TỔNG QUAN)
    cfg_bang_chuan = {
        "STT": st.column_config.NumberColumn("STT", width="small"),
        col_b_ct_lsc: st.column_config.TextColumn("Lệnh sửa chữa", width="medium"),
        "Biển số": st.column_config.TextColumn("Biển số", width="small"),
        "Cố vấn dịch vụ": st.column_config.TextColumn("Cố vấn dịch vụ", width="medium"),
        "Trạng thái DMS": st.column_config.TextColumn("Trạng thái", width="small"),
        "Thời gian đóng LSC": st.column_config.TextColumn("Thời gian đóng", width="medium"),
        "Số mục (Part/Công)": st.column_config.NumberColumn("Số mục", width="small"),
        "Tổng tiền xưởng (no VAT)": st.column_config.NumberColumn(format="%,d đ", width="medium"),
        "Tiền ĐXBH (no VAT)": st.column_config.NumberColumn(format="%,d đ", width="medium"),
        "Tiền WCS duyệt (no VAT)": st.column_config.NumberColumn(format="%,d đ", width="medium"),
        "Ngày duyệt lệnh": st.column_config.TextColumn("Ngày duyệt", width="medium"),
        "WCS Số Kỳ Cụ Thể": st.column_config.TextColumn("Bảng Kê / Kỳ Duyệt", width="large"),
        "Nhãn Trạng Thái": st.column_config.TextColumn("Trạng Thái Phê Duyệt", width="large")
    }

    with t_b_ro:
        st.subheader(f"📋 Toàn Bộ Lệnh Bảo Hành ({sel_b_tt} | {sel_b_ky})")
        if not df_b_filtered.empty:
            df_b_ro_grouped = df_b_filtered.groupby(['lsc_norm', col_b_ct_lsc]).agg({
                'Biển số': 'first',
                'Cố vấn dịch vụ': 'first',
                'Trạng thái DMS': 'first',
                'Thời gian đóng LSC': 'first',
                'Ngày_Duyệt_Format': lambda x: ", ".join(sorted(set(str(v) for v in x if v != 'Chưa duyệt'))) if any(v != 'Chưa duyệt' for v in x) else 'Chưa duyệt',
                'WCS Số Kỳ Cụ Thể': lambda x: " | ".join(sorted(set(str(v) for v in x if pd.notna(v)))),
                col_b_ct_mavt: 'count',
                'Tiền xưởng (DMS)': 'sum',
                'Tien_DXBH': 'sum',
                'Tien_WCS': lambda x: x.sum() if x.notna().any() else 0,
                'Nhãn Trạng Thái': lambda x: " | ".join(sorted(set(x)))
            }).reset_index().rename(columns={
                col_b_ct_mavt: 'Số mục (Part/Công)',
                'Tiền xưởng (DMS)': 'Tổng tiền xưởng (no VAT)',
                'Tien_DXBH': 'Tiền ĐXBH (no VAT)',
                'Tien_WCS': 'Tiền WCS duyệt (no VAT)',
                'Ngày_Duyệt_Format': 'Ngày duyệt lệnh'
            })

            df_b_ro_grouped.insert(0, 'STT', range(1, len(df_b_ro_grouped) + 1))
            cols_b_show_ro = [
                'STT', col_b_ct_lsc, 'Biển số', 'Cố vấn dịch vụ', 'Trạng thái DMS', 'Thời gian đóng LSC',
                'Số mục (Part/Công)', 'Tổng tiền xưởng (no VAT)', 'Tiền ĐXBH (no VAT)',
                'Tiền WCS duyệt (no VAT)', 'Ngày duyệt lệnh', 'WCS Số Kỳ Cụ Thể', 'Nhãn Trạng Thái'
            ]
            df_display_ro = df_b_ro_grouped[[c for c in cols_b_show_ro if c in df_b_ro_grouped.columns]].copy()
            st_styled_ro = df_display_ro.style.map(to_mau_trang_thai, subset=['Nhãn Trạng Thái'])
            st.dataframe(st_styled_ro, use_container_width=True, hide_index=True, column_config=cfg_bang_chuan)
        else:
            st.info("Không có dữ liệu phù hợp với bộ lọc.")

    # TAB 4: DUYỆT TÁCH ĐỢT
    with t_b_split:
        st.subheader("🔄 Theo Dõi Tiến Trình Duyệt Tách Đợt (Chờ Cấp 2 / Cấp 1)")
        ro_b_split = df_bh_m.groupby(['lsc_norm', col_b_ct_lsc]).agg({
            'Biển số': 'first',
            'Cố vấn dịch vụ': 'first',
            col_b_ct_mavt: 'count',
            'Tien_WCS': lambda x: sum(1 for v in x if clean_num_bh(v) > 0),
            'Nhãn Trạng Thái': lambda x: sum(1 for v in x if "chờ" in str(v).lower() and not any(k in str(v).lower() for k in ['bỏ qua', 'không duyệt', 'từ chối', 'không claim'])),
            'Tien_DXBH': 'sum'
        }).reset_index().rename(columns={
            col_b_ct_mavt: 'Tổng số mục',
            'Tien_WCS': 'Số mục đã chốt WCS',
            'Nhãn Trạng Thái': 'Số mục đang chờ duyệt',
            'Tien_DXBH': 'Tổng tiền đề xuất'
        })

        df_b_split_show = ro_b_split[(ro_b_split['Số mục đã chốt WCS'] > 0) & (ro_b_split['Số mục đang chờ duyệt'] > 0)].copy()
        if not df_b_split_show.empty:
            df_b_split_show['Tiến độ chi tiết'] = df_b_split_show.apply(
                lambda r: f"Đã duyệt Lần 1 ({r['Số mục đã chốt WCS']} mục) - Chờ duyệt Lần 2 ({r['Số mục đang chờ duyệt']} mục)", axis=1
            )
            df_b_split_show.insert(0, 'STT', range(1, len(df_b_split_show) + 1))
            cfg_b_sp = {
                "STT": st.column_config.NumberColumn("STT", width="small"),
                col_b_ct_lsc: st.column_config.TextColumn("Lệnh sửa chữa", width="medium"),
                "Biển số": st.column_config.TextColumn("Biển số", width="small"),
                "Cố vấn dịch vụ": st.column_config.TextColumn("Cố vấn dịch vụ", width="medium"),
                "Tiến độ chi tiết": st.column_config.TextColumn("Tiến độ chi tiết", width="large"),
                "Tổng tiền đề xuất": st.column_config.NumberColumn(format="%,d đ", width="medium")
            }
            st.dataframe(df_b_split_show[['STT', col_b_ct_lsc, 'Biển số', 'Cố vấn dịch vụ', 'Tiến độ chi tiết', 'Tổng tiền đề xuất']], use_container_width=True, hide_index=True, column_config=cfg_b_sp)

            sel_b_split_ro = st.selectbox("🔍 Bấm chọn xe xem chi tiết vật tư & giờ công từng đợt:", df_b_split_show[col_b_ct_lsc].unique(), key="sb_split_ro")
            df_b_ro_sp_detail = df_bh_m[df_bh_m[col_b_ct_lsc] == sel_b_split_ro].copy()
            df_b_ro_sp_detail.insert(0, 'STT', range(1, len(df_b_ro_sp_detail) + 1))
            cols_b_sp_dt = ['STT', col_b_ct_mavt, col_b_mota, col_b_ct_loai, 'SL_DMS', 'TienCong_DMS', 'Tien_WCS', 'Ngày_Duyệt_Format', 'Nhãn Trạng Thái', 'WCS Số Kỳ Cụ Thể']
            cfg_sub_sp = {
                "Nhãn Trạng Thái": st.column_config.TextColumn("Nhãn Trạng Thái", width="large"),
                "WCS Số Kỳ Cụ Thể": st.column_config.TextColumn("Bảng Kê / Kỳ", width="large")
            }
            st_styled_sub = df_b_ro_sp_detail[[c for c in cols_b_sp_dt if c in df_b_ro_sp_detail.columns]].style.map(to_mau_trang_thai, subset=['Nhãn Trạng Thái'])
            st.dataframe(st_styled_sub, use_container_width=True, hide_index=True, column_config=cfg_sub_sp)
        else:
            st.success("Không có lệnh nào bị treo duyệt tách đợt.")

    # TAB 5: CHI TIẾT TỪNG HẠNG MỤC
    with t_b_detail:
        st.subheader("🔎 Chi Tiết Hạng Mục Linh Kiện / Giờ Công")
        if not df_b_filtered.empty:
            df_b_dt_show = df_b_filtered.copy()
            df_b_dt_show.insert(0, 'STT', range(1, len(df_b_dt_show) + 1))
            cols_b_dt_view = [
                'STT', col_b_ct_lsc, 'Biển số', 'Cố vấn dịch vụ', col_b_ct_mavt, col_b_mota, col_b_ct_loai,
                'SL_DMS', 'SL_WCS', 'Tiền xưởng (DMS)', 'Tien_WCS', 'Ngày_Duyệt_Format', 'WCS Số Kỳ Cụ Thể', 'Nhãn Trạng Thái'
            ]
            cfg_b_dt = {
                "STT": st.column_config.NumberColumn("STT", width="small"),
                col_b_ct_lsc: st.column_config.TextColumn("Lệnh sửa chữa", width="medium"),
                "Biển số": st.column_config.TextColumn("Biển số", width="small"),
                "Cố vấn dịch vụ": st.column_config.TextColumn("Cố vấn dịch vụ", width="medium"),
                col_b_ct_mavt: st.column_config.TextColumn("Mã sản phẩm", width="medium"),
                col_b_mota: st.column_config.TextColumn("Mô tả sản phẩm", width="large"),
                "Tiền xưởng (DMS)": st.column_config.NumberColumn(format="%,d đ", width="medium"),
                "Tien_WCS": st.column_config.NumberColumn(format="%,d đ", width="medium"),
                "WCS Số Kỳ Cụ Thể": st.column_config.TextColumn("Bảng Kê / Kỳ Duyệt", width="large"),
                "Nhãn Trạng Thái": st.column_config.TextColumn("Trạng Thái Phê Duyệt", width="large")
            }
            st_styled_dt = df_b_dt_show[[c for c in cols_b_dt_view if c in df_b_dt_show.columns]].style.map(to_mau_trang_thai, subset=['Nhãn Trạng Thái'])
            st.dataframe(st_styled_dt, use_container_width=True, hide_index=True, column_config=cfg_b_dt)
        else:
            st.info("Không có dữ liệu chi tiết.")

    # TAB 6: CHI TIẾT TỪNG DÒNG BỎ QUA
    with t_b_ignore:
        st.subheader("⚪ Chi Tiết Các Hạng Mục Bỏ Qua & Hãng Không Duyệt (Xưởng Tự Chịu)")
        
        st.markdown("##### ⚡ Nhập nhanh Lệnh sửa chữa muốn gán nhãn BỎ QUA:")
        c_in_ig, c_btn_ig = st.columns([7.5, 2.5])
        with c_in_ig:
            nhap_lsc_bo_qua = st.text_input(
                "Nhập mã LSC (hoặc số đuôi lệnh):",
                placeholder="VD: C23401-WO-260909-0008 hoặc chỉ cần gõ 260909-0008",
                key="input_nhanh_bo_qua"
            )
        with c_btn_ig:
            st.write("")
            st.write("")
            if st.button("➕ Thêm vào danh sách Bỏ Qua", type="primary", use_container_width=True):
                if nhap_lsc_bo_qua.strip():
                    key_nhap = norm_lsc_key(nhap_lsc_bo_qua)
                    matched_norms = [norm for norm in df_bh_m['lsc_norm'].unique() if key_nhap in norm]
                    if matched_norms:
                        for m_norm in matched_norms:
                            st.session_state.ds_bo_qua.add(m_norm)
                            for idx_r, r_val in df_bh_m.iterrows():
                                if r_val.get('lsc_norm') == m_norm:
                                    df_bh_m.at[idx_r, 'Nhãn Trạng Thái'] = 'BỎ QUA (Không theo dõi)'
                        
                        try:
                            conn.update(worksheet="ChiTiet_BaoHanh", data=df_bh_m)
                            st.toast("✅ Đã lưu trạng thái Bỏ Qua lên Google Sheets!", icon="🚀")
                        except Exception as e_save:
                            st.warning(f"Lưu bộ nhớ tạm: {e_save}")
                        st.success(f"Đã gán nhãn BỎ QUA thành công cho: {nhap_lsc_bo_qua}")
                        st.rerun()
                    else:
                        st.error(f"❌ Không tìm thấy lệnh '{nhap_lsc_bo_qua}' trong hệ thống dữ liệu!")
                else:
                    st.warning("Vui lòng nhập mã lệnh sửa chữa!")

        df_b_ignored = df_bh_m[
            df_bh_m['lsc_norm'].isin(st.session_state.ds_bo_qua) | 
            df_bh_m['Nhãn Trạng Thái'].apply(is_bo_qua_hoac_khong_duyet)
        ].copy()

        if not df_b_ignored.empty:
            df_b_ignored.insert(0, 'STT', range(1, len(df_b_ignored) + 1))
            cols_ig_view = [
                'STT', col_b_ct_lsc, 'Biển số', 'Cố vấn dịch vụ',
                col_b_ct_loai, col_b_ct_mavt, col_b_mota, 'SL_DMS',
                'Tiền xưởng (DMS)', 'Nhãn Trạng Thái'
            ]
            cfg_ig_view = {
                "STT": st.column_config.NumberColumn("STT", width="small"),
                col_b_ct_lsc: st.column_config.TextColumn("Lệnh sửa chữa", width="medium"),
                "Biển số": st.column_config.TextColumn("Biển số", width="small"),
                "Cố vấn dịch vụ": st.column_config.TextColumn("Cố vấn dịch vụ", width="medium"),
                col_b_ct_loai: st.column_config.TextColumn("Phân loại", width="small"),
                col_b_ct_mavt: st.column_config.TextColumn("Mã sản phẩm", width="medium"),
                col_b_mota: st.column_config.TextColumn("Tên công việc / Mô tả phụ tùng", width="large"),
                "SL_DMS": st.column_config.NumberColumn("Số lượng / Giờ công", width="small"),
                "Tiền xưởng (DMS)": st.column_config.NumberColumn("Tiền xưởng (no VAT)", format="%,d đ", width="medium"),
                "Nhãn Trạng Thái": st.column_config.TextColumn("Lý do / Trạng thái", width="large")
            }
            st.dataframe(
                df_b_ignored[[c for c in cols_ig_view if c in df_b_ignored.columns]],
                use_container_width=True,
                hide_index=True,
                column_config=cfg_ig_view
            )
        else:
            st.info("Chưa có lệnh nào được gán nhãn Bỏ qua hoặc Không duyệt.")

    # TAB 7: HÓA ĐƠN XUẤT NHÀ MÁY
    with t_b_inv:
        st.subheader(f"🧾 Hóa Đơn Xuất Nhà Máy ({sel_b_ky})")
        df_b_inv_target = df_b_filtered[df_b_filtered['Kỳ_Tính_Toán_DXBH'] != 'Chưa duyệt'].copy()
        
        if not df_b_inv_target.empty:
            if 'Số HĐ' not in df_b_inv_target.columns:
                df_b_inv_target['Số HĐ'] = ""
            if 'Ngày HĐ' not in df_b_inv_target.columns:
                df_b_inv_target['Ngày HĐ'] = ""

            df_b_inv_summary = df_b_inv_target.groupby(['lsc_norm', col_b_ct_lsc]).agg({
                'Biển số': 'first',
                'Cố vấn dịch vụ': 'first',
                'WCS Số Kỳ Cụ Thể': 'first',
                col_b_ct_mavt: 'count',
                'Tien_WCS': lambda x: x.sum() if x.notna().any() else 0,
                'Tien_DXBH': 'sum',
                'Số HĐ': 'first',
                'Ngày HĐ': 'first'
            }).reset_index()

            df_b_inv_summary['Tiền Quyết Toán'] = df_b_inv_summary.apply(
                lambda r: r['Tien_WCS'] if r['Tien_WCS'] > 0 else r['Tien_DXBH'], axis=1
            )
            df_b_inv_summary.insert(0, 'STT', range(1, len(df_b_inv_summary) + 1))

            st.markdown("##### ⚡ Áp dụng nhanh cho toàn bộ danh sách bên dưới:")
            c_in1, c_in2, c_in3, c_btn = st.columns([2.5, 2.5, 3.5, 2.5])
            with c_in1:
                so_hd_b = st.text_input("Số HĐ xuất:", placeholder="VD: 0001234", key="b_in_shd")
            with c_in2:
                ngay_hd_b = st.text_input("Ngày xuất HĐ:", placeholder="VD: 25/09/2026", key="b_in_nhd")
            with c_in3:
                noidung_hd_b = st.text_input("Nội dung HĐ:", value=f"Chi phí bảo hành Ô tô {sel_b_ky}", key="b_in_nd")

            if so_hd_b:
                df_b_inv_summary['Số HĐ'] = so_hd_b
            if ngay_hd_b:
                df_b_inv_summary['Ngày HĐ'] = ngay_hd_b
            df_b_inv_summary['Nội Dung'] = noidung_hd_b

            cfg_b_inv = {
                "STT": st.column_config.NumberColumn("STT", width="small", disabled=True),
                col_b_ct_lsc: st.column_config.TextColumn("Lệnh sửa chữa", width="medium", disabled=True),
                "Biển số": st.column_config.TextColumn("Biển số", width="small", disabled=True),
                "Cố vấn dịch vụ": st.column_config.TextColumn("Cố vấn dịch vụ", width="medium", disabled=True),
                "WCS Số Kỳ Cụ Thể": st.column_config.TextColumn("Bảng kê / Kỳ", width="large", disabled=True),
                col_b_ct_mavt: st.column_config.NumberColumn("Số mục", width="small", disabled=True),
                "Tiền Quyết Toán": st.column_config.NumberColumn("Tiền Quyết Toán", format="%,d đ", width="medium", disabled=True),
                "Số HĐ": st.column_config.TextColumn("Số HĐ (Gõ sửa trực tiếp)", width="medium"),
                "Ngày HĐ": st.column_config.TextColumn("Ngày HĐ (Gõ sửa trực tiếp)", width="medium"),
                "Nội Dung": st.column_config.TextColumn("Nội Dung", width="large", disabled=True)
            }
            cols_b_inv_show = ['STT', col_b_ct_lsc, 'Biển số', 'Cố vấn dịch vụ', 'WCS Số Kỳ Cụ Thể', col_b_ct_mavt, 'Tiền Quyết Toán', 'Số HĐ', 'Ngày HĐ', 'Nội Dung']
            
            edited_inv_df = st.data_editor(
                df_b_inv_summary[[c for c in cols_b_inv_show if c in df_b_inv_summary.columns]],
                use_container_width=True,
                hide_index=True,
                column_config=cfg_b_inv,
                key="editor_hoa_don_bh"
            )

            with c_btn:
                st.write("")
                st.write("")
                if st.button("☁️ Lưu Hóa Đơn Lên GG Sheets", type="primary", use_container_width=True):
                    with st.spinner("⏳ Đang lưu số HĐ vào Google Sheets..."):
                        map_shd = edited_inv_df.set_index(col_b_ct_lsc)['Số HĐ'].to_dict()
                        map_nhd = edited_inv_df.set_index(col_b_ct_lsc)['Ngày HĐ'].to_dict()
                        
                        for idx_r, r_val in df_bh_m.iterrows():
                            lsc_k = r_val.get(col_b_ct_lsc)
                            if lsc_k in map_shd and map_shd[lsc_k]:
                                df_bh_m.at[idx_r, 'Số HĐ'] = map_shd[lsc_k]
                                df_bh_m.at[idx_r, 'Ngày HĐ'] = map_nhd.get(lsc_k, '')

                        conn.update(worksheet="ChiTiet_BaoHanh", data=df_bh_m)
                        st.success("✅ Đã lưu thành công Số HĐ vào Google Sheets!")
                        st.rerun()
        else:
            st.info("Không có dữ liệu hóa đơn bảo hành cho kỳ này.")
else:
    st.info("💡 Chưa có dữ liệu đối soát bảo hành. Hãy mở mục **'Nạp Tệp Dữ Liệu Đối Soát'** phía trên để tải 2 tệp (Chi tiết lệnh + Cổng ĐXBH) lên.")
