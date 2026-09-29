with t_b_inv:
        st.subheader(f"🧾 Hóa Đơn Xuất Nhà Máy ({sel_b_ky})")
        df_b_inv_target = df_b_filtered[df_b_filtered['Kỳ_Tính_Toán_DXBH'] != 'Chưa duyệt'].copy()
        
        if not df_b_inv_target.empty:
            # BỔ SUNG CỘT AN TOÀN TRÁNH LỖI KEYERROR
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
            with c_in1: so_hd_b = st.text_input("Số HĐ xuất:", placeholder="VD: 0001234", key="b_in_shd")
            with c_in2: ngay_hd_b = st.text_input("Ngày xuất HĐ:", placeholder="VD: 25/09/2026", key="b_in_nhd")
            with c_in3: noidung_hd_b = st.text_input("Nội dung HĐ:", value=f"Chi phí bảo hành Ô tô {sel_b_ky}", key="b_in_nd")

            # Áp dụng giá trị nhập nhanh nếu người dùng có gõ
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
