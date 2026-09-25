# ==================== TAB 4: ĐỐI SOÁT & ĐẨY CHI TIẾT BẢO HÀNH ====================
with tab_bh_wcs:
    st.subheader("🛡️ Đối Soát & Quản Lý Chi Tiết Bảo Hành VinFast")
    st.caption("Tải lên 2 file để hệ thống tự động lọc các lệnh bảo hành (W), tìm lệnh chưa claim và đẩy toàn bộ chi tiết lên Google Sheets.")

    c_up1, c_up2 = st.columns(2)
    with c_up1:
        up_dx = st.file_uploader("📥 1. Tải lên file DEXUATBAOHANH (Cổng Nhà Máy)", type=['csv', 'xlsx'], key="up_dx_duo_fix")
    with c_up2:
        up_ct = st.file_uploader("📥 2. Tải lên file CHITIETLENHSUACHUA (DMS Xưởng)", type=['csv', 'xlsx'], key="up_ct_duo_fix")

    if up_dx and up_ct:
        df_dx_in = pd.read_csv(up_dx, low_memory=False) if up_dx.name.endswith('.csv') else pd.read_excel(up_dx)
        df_ct_in = pd.read_csv(up_ct, low_memory=False) if up_ct.name.endswith('.csv') else pd.read_excel(up_ct)
        
        df_dx_in.columns = [str(c).strip() for c in df_dx_in.columns]
        df_ct_in.columns = [str(c).strip() for c in df_ct_in.columns]

        # Khóa chuẩn hóa LSC
        col_dx_lsc = next((c for c in df_dx_in.columns if 'lệnh sửa chữa' in c.lower()), 'Lệnh sửa chữa')
        col_ct_lsc = next((c for c in df_ct_in.columns if 'lệnh sửa chữa' in c.lower()), 'Lệnh sửa chữa')
        
        df_dx_in['lsc_norm'] = df_dx_in[col_dx_lsc].apply(norm_lsc_key)
        df_ct_in['lsc_norm'] = df_ct_in[col_ct_lsc].apply(norm_lsc_key)

        # LỌC TRỰC TIẾP TỪ FILE CHI TIẾT (LẤY TẤT CẢ DÒNG BẢO HÀNH W, KHÔNG BỊ PHỤ THUỘC MASTERDATA)
        col_pb = next((c for c in df_ct_in.columns if 'classification' in c.lower() or 'p/bill' in c.lower()), 'P/bill classification')
        df_w_all = df_ct_in[df_ct_in[col_pb].astype(str).str.strip().str.upper() == 'W'].copy()

        # Bộ lọc trạng thái đóng nếu có cột 'Đóng dòng'
        col_dong_dong = next((c for c in df_ct_in.columns if 'đóng' in c.lower()), None)
        
        c_filter1, c_filter2 = st.columns([3, 5])
        with c_filter1:
            opt_filter = st.selectbox("📌 Lọc trạng thái lệnh:", ["Tất cả lệnh bảo hành (W)", "Chỉ lấy dòng Đã đóng"], index=0)
        
        if opt_filter == "Chỉ lấy dòng Đã đóng" and col_dong_dong:
            df_w_target = df_w_all[df_w_all[col_dong_dong].astype(str).str.strip().isin(['Đã đóng', 'Đóng'])].copy()
        else:
            df_w_target = df_w_all.copy()

        # Tập hợp mã WO giữa 2 bên
        set_dx_lsc = set(df_dx_in['lsc_norm'].unique())
        set_w_lsc = set(df_w_target['lsc_norm'].unique())

        lsc_chua_claim = sorted(list(set_w_lsc - set_dx_lsc))
        lsc_da_claim = set_w_lsc.intersection(set_dx_lsc)

        st.markdown("---")
        # THẺ KPI
        k1, k2, k3 = st.columns(3)
        k1.metric("📌 Tổng Lệnh Bảo Hành (W)", f"{len(set_w_lsc)} Lệnh", f"{len(df_w_target)} dòng linh kiện/công")
        k2.metric("🟢 Đã Gửi Đề Xuất Nhà Máy", f"{len(lsc_da_claim)} Lệnh")
        k3.metric("🚨 Lệnh Chưa Tạo ĐXBH", f"{len(lsc_chua_claim)} Lệnh BỊ SÓT", delta=f"-{len(lsc_chua_claim)} lệnh", delta_color="inverse")

        # KHỐI 1: CẢNH BÁO LỆNH CHƯA TẠO ĐXBH
        if len(lsc_chua_claim) > 0:
            st.error(f"🚨 **DANH SÁCH {len(lsc_chua_claim)} LỆNH BẢO HÀNH CHƯA TẠO ĐỀ XUẤT CLAIM TRÊN CỔNG NHÀ MÁY:**")
            df_miss_summary = df_w_target[df_w_target['lsc_norm'].isin(lsc_chua_claim)].groupby(col_ct_lsc).agg({
                'Số khung': 'first' if 'Số khung' in df_w_target.columns else lambda x: '',
                'Cố vấn dịch vụ': 'first' if 'Cố vấn dịch vụ' in df_w_target.columns else lambda x: '',
                'Mô tả sản phẩm': 'count'
            }).reset_index()
            df_miss_summary.columns = ['Số Lệnh Sửa Chữa', 'Số Khung', 'Cố Vấn Dịch Vụ', 'Số Mục Chưa Claim']
            st.dataframe(df_miss_summary, use_container_width=True, hide_index=True)

        st.markdown("---")
        st.subheader("📋 Bảng Tổng Hợp Chi Tiết Từng Hạng Mục (Sẵn sàng đẩy lên Google Sheets)")

        # Khớp thông tin phê duyệt từ ĐXBH vào từng dòng chi tiết DMS
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

        # Merge thông tin duyệt
        df_export_gsheet = pd.merge(
            df_w_target,
            df_dx_in[['lsc_norm', 'mavt_norm', col_dx_dx, col_dx_ngay, col_dx_tt, col_dx_tien]].drop_duplicates(subset=['lsc_norm', 'mavt_norm']),
            left_on=['lsc_norm', 'masp_norm'],
            right_on=['lsc_norm', 'mavt_norm'],
            how='left'
        )

        # Tính tiền xưởng
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

        cfg_final = {
            "Tiền xưởng (DMS)": st.column_config.NumberColumn(format="%,d đ"),
            "Tiền claim (ĐXBH)": st.column_config.NumberColumn(format="%,d đ")
        }
        st.dataframe(df_final, use_container_width=True, hide_index=True, column_config=cfg_final)

        # NÚT ĐẨY THẲNG LÊN GOOGLE SHEETS
        if is_admin:
            if st.button("☁️ ĐẨY TOÀN BỘ BẢNG CHI TIẾT NÀY LÊN GOOGLE SHEET (Sheet: ChiTiet_BaoHanh)", type="primary", use_container_width=True):
                p_bar = st.progress(0)
                txt_bar = st.empty()
                txt_bar.write("⏳ Đang chuẩn bị dữ liệu và kết nối Google Sheets...")
                
                df_to_push = df_final.copy()
                for c in df_to_push.columns:
                    if c in ['Tiền xưởng (DMS)', 'Tiền claim (ĐXBH)', 'Số lượng/Nhân công']:
                        df_to_push[c] = pd.to_numeric(df_to_push[c], errors='coerce').fillna(0)
                    else:
                        df_to_push[c] = df_to_push[c].fillna('').astype(str)

                p_bar.progress(50)
                txt_bar.write("⏳ Đang ghi dữ liệu vào sheet 'ChiTiet_BaoHanh'...")
                
                conn.update(worksheet="ChiTiet_BaoHanh", data=df_to_push)
                
                p_bar.progress(100)
                txt_bar.empty()
                st.success(f"✅ ĐÃ ĐẨY THÀNH CÔNG {len(df_to_push)} DÒNG CHI TIẾT LÊN GOOGLE SHEETS TẠI SHEET 'ChiTiet_BaoHanh'!")
                st.balloons()
