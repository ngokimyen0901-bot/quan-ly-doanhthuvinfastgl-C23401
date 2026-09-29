# TAB 6: HIỂN THỊ CHI TIẾT TỪNG CÔNG VIỆC / VẬT TƯ BỊ BỎ QUA
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

        # LẤY CHI TIẾT TỪNG DÒNG (KHÔNG GOM NHÓM ĐỂ XEM RÕ TỪNG CÔNG VIỆC)
        df_b_ignored = df_bh_m[
            df_bh_m['lsc_norm'].isin(st.session_state.ds_bo_qua) | 
            df_bh_m['Nhãn Trạng Thái'].apply(is_bo_qua_hoac_khong_duyet)
        ].copy()

        if not df_b_ignored.empty:
            df_b_ignored.insert(0, 'STT', range(1, len(df_b_ignored) + 1))
            
            # Các cột hiển thị chi tiết tên vật tư và công việc
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
