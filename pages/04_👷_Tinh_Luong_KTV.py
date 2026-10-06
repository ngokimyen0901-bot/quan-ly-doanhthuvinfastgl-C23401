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

    # Nhận diện dòng dịch vụ/công lao động trong bảng kê claim
    col_loai_sp = next((c for c in df_out.columns if any(k in c.lower() for k in ['loại sản phẩm', 'material3'])), None)
    if col_loai_sp and df_out[col_loai_sp].dropna().astype(str).str.lower().str.contains('dịch vụ|service').any():
        df_out = df_out[df_out[col_loai_sp].astype(str).str.lower().str.contains('dịch vụ|service')].copy()
    
    col_wo = next((c for c in df_out.columns if any(k in c.lower() for k in ['lệnh sửa chữa', 'lsc', 'wo'])), df_out.columns[0])
    col_cong = next((c for c in df_out.columns if any(k in c.lower() for k in ['approved service', 'tiền công yêu cầu', 'tổng số tiền', 'amount', 'thành tiền', 'tiền'])), df_out.columns[-1])
    col_cv = next((c for c in df_out.columns if any(k in c.lower() for k in ['tên công việc', 'mô tả', 'công việc'])), df_out.columns[0])

    df_out['wo_norm'] = df_out[col_wo].apply(norm_wo_key)
    df_out['Tien_Cong_BH'] = df_out[col_cong].apply(clean_num)
    df_out['Ten_CV_BH'] = df_out[col_cv].fillna('').astype(str) if col_cv else ''
    
    # Gán thẳng kỳ bảo hành theo kỳ làm việc hiện tại thay vì suy đoán ngày cũ
    df_out['Thang_BH'] = ky_mac_dinh

    return df_out[(df_out['Tien_Cong_BH'] > 0) & (df_out['wo_norm'] != '')].copy()
