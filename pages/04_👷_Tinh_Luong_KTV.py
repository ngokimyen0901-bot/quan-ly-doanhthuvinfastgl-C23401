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

st.title("👷 Tính Lương & Đối Soát Tiền Công Kỹ Thuật Viên")
st.caption("☁️ Đồng bộ và lưu trữ trực tiếp lên Google Sheets: **Luong_KTV**.")

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

def doc_file_5114_thong_minh(file_obj):
    df_raw = pd.read_excel(file_obj, header=None)
    header_idx = 8
    for idx, row in df_raw.head(20).iterrows():
        row_str = " ".join([str(v).lower() for v in row.values if pd.notna(v)])
        if 'số c.từ' in row_str or 'số hóa đơn' in row_str or 'số r/o' in row_str:
            header_idx = idx
            break
    df = pd.read_excel(file_obj, header=header_idx)
    df.columns = [str(c).strip() for c in df.columns]
    return df

def doc_file_ktv_thong_minh(file_obj):
    df_raw = pd.read_excel(file_obj, header=None)
    header_idx = 7
    for idx, row in df_raw.head(20).iterrows():
        row_str = " ".join([str(v).lower() for v in row.values if pd.notna(v)])
        if 'số ro' in row_str or 'mã ktv' in row_str or 'tên ktv' in row_str:
            header_idx = idx
            break
    df = pd.read_excel(file_obj, header=header_idx)
    df.columns = [str(c).strip() for c in df.columns]
    return df

def doc_file_warranty_thong_minh(file_obj):
    xls = pd.ExcelFile(file_obj)
    sheet_target = xls.sheet_names[0]
    for s in xls.sheet_names:
        if 'claim' in s.lower() or 'active' in s.lower():
            sheet_target = s
            break
    df = pd.read_excel(xls, sheet_name=sheet_target)
    df.columns = [str(c).strip() for c in df.columns]
    return df

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

st.markdown("### 📥 Nạp Tệp Dữ Liệu Tính Lương & Đối Soát")
c1, c2, c3 = st.columns(3)
with c1:
    up_5114 = st.file_uploader("1. Sổ chi tiết TK 5114 (5114.xlsx):", type=['xlsx', 'xls', 'csv'], key="p4_5114")
with c2:
    up_ktv = st.file_uploader("2. Báo cáo doanh thu KTV (BÁO CÁO DOANH THU KTV.xlsx):", type=['xlsx', 'xls', 'csv'], key="p4_ktv")
with c3:
    up_bh = st.file_uploader("3. Chi tiết bảo hành hãng (Active Warranty...xlsx):", type=['xlsx', 'xls'], key="p4_bh")

if up_5114 and up_ktv:
    df_5114 = doc_file_5114_thong_minh(up_5114)
    df_ktv = doc_file_ktv_thong_minh(up_ktv)

    # 1. Chuẩn hóa TK 5114
    col_5114_ro = next((c for c in df_5114.columns if any(k in c.lower() for k in ['số r/o hãng', 'r/o hãng', 'lệnh hãng'])), None)
    if not col_5114_ro:
        col_5114_ro = next((c for c in df_5114.columns if any(k in c.lower() for k in ['số r/o', 'số ro'])), 'Số R/O')
    col_5114_shd = next((c for c in df_5114.columns if 'hóa đơn' in c.lower()), 'Số hóa đơn')
    col_5114_ngay = next((c for c in df_5114.columns if 'ngày' in c.lower()), 'Ngày C.từ')
    col_5114_tien = next((c for c in df_5114.columns if c.strip().lower() in ['có', 'co', 'phát sinh có']), 'Có')
    col_5114_kh = next((c for c in df_5114.columns if 'tên khách' in c.lower() or 'khách hàng' in c.lower()), 'Tên khách')
    col_5114_dg = next((c for c in df_5114.columns if 'diễn giải' in c.lower()), 'Diễn giải')

    df_5114['wo_norm'] = df_5114[col_5114_ro].apply(norm_wo_key)
    df_5114['Tien_5114'] = df_5114[col_5114_tien].apply(clean_num)
    df_5114['Thang'] = pd.to_datetime(df_5114[col_5114_ngay], errors='coerce', dayfirst=True).dt.strftime('Tháng %m/%Y').fillna('Chưa rõ')

    df_5114_clean = df_5114[(df_5114['Tien_5114'] > 0) & (df_5114['wo_norm'] != '')].copy()

    # 2. Chuẩn hóa File KTV
    col_ktv_ro = next((c for c in df_ktv.columns if 'số ro hãng' in c.lower() or 'ro hãng' in c.lower()), 'Số RO hãng')
    col_ktv_ma = next((c for c in df_ktv.columns if 'mã ktv' in c.lower()), 'Mã KTV')
    col_ktv_ten = next((c for c in df_ktv.columns if 'tên ktv' in c.lower()), 'Tên KTV')
    col_ktv_cv = next((c for c in df_ktv.columns if 'mã cv' in c.lower() or 'cố vấn' in c.lower()), 'Mã CV')
    col_ktv_tien = next((c for c in df_ktv.columns if 'thành tiền' in c.lower() or 'tiền dịch vụ' in c.lower()), 'Thành tiền')
    col_ktv_bs = next((c for c in df_ktv.columns if 'biển' in c.lower()), 'Biển kiểm soát')
    col_ktv_nd = next((c for c in df_ktv.columns if 'nội dung' in c.lower() or 'hạng mục' in c.lower()), 'Nội dung')

    df_ktv['wo_norm'] = df_ktv[col_ktv_ro].apply(norm_wo_key)
    df_ktv['Tien_KTV'] = df_ktv[col_ktv_tien].apply(clean_num)
    df_ktv['Ma_KTV_Clean'] = df_ktv[col_ktv_ma].fillna('').astype(str).str.strip().str.replace('.0', '', regex=False)
    df_ktv['Ten_KTV_Clean'] = df_ktv[col_ktv_ten].fillna('').astype(str).str.strip()

    ktv_per_ro = df_ktv[df_ktv['wo_norm'] != ''].groupby('wo_norm').agg({
        col_ktv_ro: 'first',
        col_ktv_bs: 'first',
        col_ktv_cv: 'first',
        'Tien_KTV': 'sum',
        'Ma_KTV_Clean': lambda x: ", ".join(sorted(set([str(v) for v in x if str(v).strip() and str(v) != 'nan']))),
        'Ten_KTV_Clean': lambda x: ", ".join(sorted(set([str(v) for v in x if str(v).strip() and str(v) != 'nan']))),
        col_ktv_nd: lambda x: " | ".join(sorted(set([str(v) for v in x if str(v).strip() and str(v) != 'nan'])))[:150]
    }).reset_index()

    if not df_master_ref.empty and 'Cố vấn dịch vụ' in df_master_ref.columns:
        map_cv = df_master_ref.set_index('wo_norm')['Cố vấn dịch vụ'].to_dict()
        ktv_per_ro['Ten_CVDV'] = ktv_per_ro['wo_norm'].map(map_cv).fillna(ktv_per_ro[col_ktv_cv])
    else:
        ktv_per_ro['Ten_CVDV'] = ktv_per_ro[col_ktv_cv]

    df_merged = pd.merge(
        df_5114_clean[['wo_norm', col_5114_ro, col_5114_shd, col_5114_ngay, 'Thang', 'Tien_5114', col_5114_kh, col_5114_dg]],
        ktv_per_ro,
        on='wo_norm',
        how='left'
    )
    df_merged['Tien_KTV'] = df_merged['Tien_KTV'].fillna(0)
    df_merged['Chenh_Lech'] = df_merged['Tien_5114'] - df_merged['Tien_KTV']

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

    # 1. Cột lựa chọn nguồn: Khớp thì tự động, lệch thì chờ bạn chọn
    df_merged['Quyet_Dinh_Nguon'] = df_merged['Ghi_Chu_Doi_Soat'].apply(
        lambda x: "Tự động (Khớp chuẩn)" if "Khớp 100%" in x else "Lấy theo KTV (Hóa đơn)"
    )

    # 2. Cột giá trị chốt: Khớp thì lấy tự động, lệch thì người dùng quyết định
    df_merged['Tien_Cong_Chot_Luong'] = df_merged.apply(
        lambda r: r['Tien_5114'] if "Khớp 100%" in r['Ghi_Chu_Doi_Soat'] else r['Tien_KTV'], axis=1
    )
    df_merged['Ghi_Chu_Thu_Cong'] = ""

    # Khôi phục các dòng bạn đã từng điều chỉnh trên Google Sheets
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

    # Giao diện Tabs
    tab_sc, tab_bh, tab_ktv_sum = st.tabs([
        "🔧 1. Đối Soát & Quyết Định Giá Trị Tính Lương (5114 vs KTV)",
        "🛡️ 2. Chia Công Bảo Hành Cho KTV (Chia Đều)",
        "👷 3. Bảng Lương Tổng Hợp Kỹ Thuật Viên"
    ])

    # ==================== TAB 1 ====================
    with tab_sc:
        st.subheader("🔧 Đối Soát Doanh Thu & Quyết Định Giá Trị Tính Lương")
        st.info("💡 **Quy tắc:** Lệnh chuẩn **🟢 Khớp 100%** sẽ chạy tự động. Lệnh lệch tiền do thuê ngoài, bảo hiểm khấu trừ... được **bôi đỏ** và bạn là người trực tiếp quyết định giá trị ở cột cuối cùng!")

        cnt_khop = (df_merged['Ghi_Chu_Doi_Soat'] == "🟢 Khớp 100%").sum()
        cnt_lech = len(df_merged) - cnt_khop
        
        m1, m2, m3 = st.columns(3)
        m1.metric("Tổng Doanh Thu 5114", f"{df_merged['Tien_5114'].sum():,.0f} đ", f"{len(df_merged)} lệnh")
        m2.metric("Lệnh Chuẩn (Tự động chạy)", f"{cnt_khop} lệnh")
        m3.metric("Lệnh Lệch (Bạn quyết định)", f"{cnt_lech} lệnh", delta=f"{cnt_lech} lệnh cần duyệt" if cnt_lech > 0 else "0", delta_color="inverse")

        # Tô màu nổi bật dòng bị lệch
        def highlight_lech(row):
            status = str(row.get('Ghi_Chu_Doi_Soat', ''))
            if '🔴' in status or 'Lệch' in status:
                return ['background-color: #ffebee; color: #c62828;' for _ in row]
            return ['' for _ in row]

        cols_edit = [
            'Thang', col_5114_ro, col_5114_shd, 'Tien_5114', 'Tien_KTV', 'Chenh_Lech',
            'Ghi_Chu_Doi_Soat', 'Ten_CVDV', 'Ma_KTV_Clean', 'Ten_KTV_Clean', col_ktv_bs,
            'Quyet_Dinh_Nguon', 'Tien_Cong_Chot_Luong', 'Ghi_Chu_Thu_Cong', 'wo_norm'
        ]
        cols_valid = [c for c in cols_edit if c in df_merged.columns]

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
            df_merged[cols_valid].style.apply(highlight_lech, axis=1),
            use_container_width=True,
            height=560,
            column_config=cfg_edit,
            hide_index=True,
            key="editor_doi_soat_luong_v2"
        )

        c_save1, c_save2 = st.columns([3.5, 6.5])
        with c_save1:
            if st.button("☁️ Lưu Quyết Định Của Bạn Lên Google Sheets", type="primary", use_container_width=True):
                with st.spinner("⏳ Đang lưu vào sheet Luong_KTV..."):
                    df_to_save = edited_sc_df.copy()
                    try:
                        conn.update(worksheet="Luong_KTV", data=df_to_save)
                        st.success("✅ Đã lưu quyết định của bạn vĩnh viễn lên Google Sheets!")
                        st.rerun()
                    except Exception as e_s:
                        st.error(f"Lỗi khi lưu: {e_s}")

    # ==================== TAB 2 ====================
    with tab_bh:
        st.subheader("🛡️ Phân Bổ Tiền Công Bảo Hành Cho KTV (Chỉ Tiền Công & Chia Đều)")
        if up_bh:
            df_bh_raw = doc_file_warranty_thong_minh(up_bh)
            col_bh_wo = next((c for c in df_bh_raw.columns if any(k in c.lower() for k in ['lệnh sửa chữa', 'lsc', 'wo'])), 'Lệnh sửa chữa')
            col_bh_cong = next((c for c in df_bh_raw.columns if any(k in c.lower() for k in ['tiền công yêu cầu', 'approved service', 'tiền công'])), 'Tiền công yêu cầu')
            col_bh_cv = next((c for c in df_bh_raw.columns if any(k in c.lower() for k in ['tên công việc chính', 'công việc chính', 'mô tả'])), 'Tên công việc chính')

            df_bh_raw['wo_norm'] = df_bh_raw[col_bh_wo].apply(norm_wo_key)
            df_bh_raw['Tien_Cong_BH'] = df_bh_raw[col_bh_cong].apply(clean_num)

            df_bh_cong_only = df_bh_raw[(df_bh_raw['Tien_Cong_BH'] > 0) & (df_bh_raw['wo_norm'] != '')].copy()

            bh_wo_sum = df_bh_cong_only.groupby('wo_norm').agg({
                col_bh_wo: 'first',
                'Tien_Cong_BH': 'sum',
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

            b_m1, b_m2 = st.columns(2)
            b_m1.metric("Tổng Tiền Công Bảo Hành (No VAT)", f"{bh_wo_sum['Tien_Cong_BH'].sum():,.0f} đ", f"{len(bh_wo_sum)} lệnh WO")
            b_m2.metric("Số Lượt Chia Công KTV", f"{len(df_bh_split)} lượt công")

            cfg_bh_split = {
                col_bh_wo: st.column_config.TextColumn("Lệnh Bảo Hành (WO)", width="medium"),
                col_ktv_bs: st.column_config.TextColumn("Biển số", width="small"),
                col_bh_cv: st.column_config.TextColumn("Công việc bảo hành", width="large"),
                "Tien_Cong_BH": st.column_config.NumberColumn("Tổng tiền công", format="%,d đ", width="medium"),
                "So_KTV_Lam_Chung": st.column_config.NumberColumn("Số KTV chia đều", width="small"),
                "Ma_KTV_Clean": st.column_config.TextColumn("Mã KTV", width="small"),
                "Ten_KTV_Clean": st.column_config.TextColumn("Họ tên KTV", width="medium"),
                "Cong_BH_Thuc_Nhan": st.column_config.NumberColumn("Công thực nhận", format="%,d đ", width="medium")
            }
            cols_show_bh = [col_bh_wo, col_ktv_bs, col_bh_cv, 'Tien_Cong_BH', 'So_KTV_Lam_Chung', 'Ma_KTV_Clean', 'Ten_KTV_Clean', 'Cong_BH_Thuc_Nhan']
            st.dataframe(df_bh_split[cols_show_bh], use_container_width=True, hide_index=True, column_config=cfg_bh_split)
        else:
            st.warning("⚠️ Vui lòng nạp file '3. Chi tiết bảo hành hãng' để chia công bảo hành.")

    # ==================== TAB 3 ====================
    with tab_ktv_sum:
        st.subheader("👷 Bảng Lương Tổng Hợp Kỹ Thuật Viên")
        st.caption("Dữ liệu công sửa chữa được tính tự động từ **Cột giá trị bạn quyết định** ở Tab 1.")
        
        sc_source = edited_sc_df if 'edited_sc_df' in locals() else df_merged
        
        ktv_sc_list = []
        for _, r_m in sc_source.iterrows():
            ma_ktvs = [x.strip() for x in str(r_m['Ma_KTV_Clean']).split(',') if x.strip() and x.strip() != 'nan']
            ten_ktvs = [x.strip() for x in str(r_m['Ten_KTV_Clean']).split(',') if x.strip() and x.strip() != 'nan']
            so_nguoi = max(len(ma_ktvs), 1)
            tien_chia = clean_num(r_m['Tien_Cong_Chot_Luong']) / so_nguoi
            
            for idx_k, m_k in enumerate(ma_ktvs):
                t_k = ten_ktvs[idx_k] if idx_k < len(ten_ktvs) else m_k
                ktv_sc_list.append({'Ma_KTV': m_k, 'Ten_KTV': t_k, 'Cong_SC': tien_chia})

        if ktv_sc_list:
            df_sc_ktv_sum = pd.DataFrame(ktv_sc_list).groupby(['Ma_KTV', 'Ten_KTV'])['Cong_SC'].sum().reset_index()
        else:
            df_sc_ktv_sum = pd.DataFrame(columns=['Ma_KTV', 'Ten_KTV', 'Cong_SC'])

        if up_bh and 'df_bh_split' in locals() and not df_bh_split.empty:
            df_bh_ktv_sum = df_bh_split.groupby(['Ma_KTV_Clean', 'Ten_KTV_Clean'])['Cong_BH_Thuc_Nhan'].sum().reset_index().rename(
                columns={'Ma_KTV_Clean': 'Ma_KTV', 'Ten_KTV_Clean': 'Ten_KTV', 'Cong_BH_Thuc_Nhan': 'Cong_BH'}
            )
        else:
            df_bh_ktv_sum = pd.DataFrame(columns=['Ma_KTV', 'Ten_KTV', 'Cong_BH'])

        df_luong_final = pd.merge(df_sc_ktv_sum, df_bh_ktv_sum, on=['Ma_KTV', 'Ten_KTV'], how='outer').fillna(0)
        df_luong_final['Tong_Luong_Nhan'] = df_luong_final['Cong_SC'] + df_luong_final['Cong_BH']
        df_luong_final = df_luong_final.sort_values(by='Tong_Luong_Nhan', ascending=False).reset_index(drop=True)
        df_luong_final.insert(0, 'STT', range(1, len(df_luong_final) + 1))

        cfg_luong = {
            "STT": st.column_config.NumberColumn("STT", width="small"),
            "Ma_KTV": st.column_config.TextColumn("Mã KTV", width="small"),
            "Ten_KTV": st.column_config.TextColumn("Họ và Tên Kỹ Thuật Viên", width="large"),
            "Cong_SC": st.column_config.NumberColumn("Công Sửa Chữa (Đã Chốt)", format="%,d đ", width="medium"),
            "Cong_BH": st.column_config.NumberColumn("Công Bảo Hành (Đã Chia Đều)", format="%,d đ", width="medium"),
            "Tong_Luong_Nhan": st.column_config.NumberColumn("TỔNG TIỀN CÔNG TÍNH LƯƠNG", format="%,d đ", width="large")
        }
        st.dataframe(df_luong_final, use_container_width=True, hide_index=True, column_config=cfg_luong)
else:
    st.info("💡 Vui lòng nạp tệp **'1. Sổ chi tiết TK 5114'** và **'2. Báo cáo doanh thu KTV'** phía trên để bắt đầu đối soát và tính lương.")
