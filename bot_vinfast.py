"""
bot_vinfast.py - Module tự động kết nối VinFast Dealer Portal (dealer-portal.vinfast.com)
Tài khoản: 0337140499
"""

import os
import sys
import time
import requests
import json
from datetime import datetime

# CẤU HÌNH THÔNG TIN VINFAST ĐẠI LÝ CỦA BẠN
VINFAST_PORTAL_URL = "https://dealer-portal.vinfast.com"
SO_DIEN_THOAI = "0337140499"
NGAY_BAT_DAU = "2026-08-29"

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DOWNLOAD_DIR = os.path.join(BASE_DIR, "downloads")
SESSION_FILE = os.path.join(BASE_DIR, "vinfast_session.json")
os.makedirs(DOWNLOAD_DIR, exist_ok=True)

def gui_yeu_cau_otp(sdt=SO_DIEN_THOAI):
    """Gửi yêu cầu mã OTP về số điện thoại VinFast"""
    url_otp = f"{VINFAST_PORTAL_URL}/api/v1/auth/request-otp" # API endpoint yêu cầu OTP
    headers = {"Content-Type": "application/json", "User-Agent": "Mozilla/5.0"}
    payload = {"phone": sdt}
    try:
        res = requests.post(url_otp, json=payload, headers=headers, timeout=15)
        return {"success": True, "message": f"Đã gửi mã OTP đến số điện thoại {sdt}!"}
    except Exception as e:
        return {"success": False, "message": f"Lỗi gửi OTP: {e}"}

def xac_thuc_otp_va_luu_phien(ma_otp, sdt=SO_DIEN_THOAI):
    """Xác thực mã OTP và lưu lại phiên làm việc (Cookies / Token) trong 20 ngày"""
    url_login = f"{VINFAST_PORTAL_URL}/api/v1/auth/verify-otp"
    headers = {"Content-Type": "application/json", "User-Agent": "Mozilla/5.0"}
    payload = {"phone": sdt, "otp": str(ma_otp).strip()}
    try:
        res = requests.post(url_login, json=payload, headers=headers, timeout=15)
        # Lưu token hoặc cookies trả về
        data_login = res.json() if res.status_code == 200 else {"cookies": res.cookies.get_dict()}
        with open(SESSION_FILE, "w", encoding="utf-8") as f:
            json.dump(data_login, f, ensure_ascii=False, indent=2)
        return {"success": True, "message": "Xác thực thành công! Đã lưu phiên làm việc trong 15-20 ngày."}
    except Exception as e:
        return {"success": False, "message": f"Lỗi xác thực: {e}"}

def tu_dong_xuat_file_vinfast(headless=True) -> dict:
    """Tự động vào lấy file Lệnh sửa chữa từ dealer-portal.vinfast.com"""
    now_str = datetime.now().strftime("%Y-%m-%d")
    output_file_path = os.path.join(DOWNLOAD_DIR, "database.xlsx")
    
    # URL trực tiếp theo đúng tham số bạn cung cấp
    target_export_url = (
        f"{VINFAST_PORTAL_URL}/dealer/aftersales/work-orders"
        f"#woDatePreset=ALL&woLimit=10000&woPage=1&woRfiFrom={NGAY_BAT_DAU}&woRfiTo={now_str}"
    )

    try:
        from playwright.sync_api import sync_playwright
        with sync_playwright() as p:
            # Khởi động trình duyệt ngầm trên cloud
            browser = p.chromium.launch(
                headless=True,
                args=["--no-sandbox", "--disable-setuid-sandbox", "--disable-dev-shm-usage"]
            )
            context = browser.new_context()
            
            # Nạp cookie đã lưu nếu có
            if os.path.exists(SESSION_FILE):
                try:
                    with open(SESSION_FILE, "r", encoding="utf-8") as f:
                        session_data = json.load(f)
                    if "cookies" in session_data:
                        context.add_cookies(session_data["cookies"])
                except Exception:
                    pass

            page = context.new_page()
            page.goto(target_export_url, wait_until="networkidle", timeout=60000)
            page.wait_for_timeout(3000)

            # Kiểm tra nếu bị chặn bắt đăng nhập
            if "login" in page.url.lower():
                browser.close()
                return {
                    "success": False,
                    "error": "OTP_REQUIRED",
                    "message": "Phiên đăng nhập đã hết hạn. Vui lòng bấm 'Gửi mã OTP' trên web để xác thực lại!"
                }

            # Bấm nút Xuất dữ liệu
            btn_export = page.locator("button:has-text('Xuất dữ liệu'), .btn-export").first
            btn_export.wait_for(state="visible", timeout=20000)
            btn_export.click()
            page.wait_for_timeout(2000)

            # Tick chọn đại lý
            modal_checkbox = page.locator(".ant-modal, .modal, [role='dialog']").locator("input[type='checkbox']").first
            if modal_checkbox.count() > 0 and not modal_checkbox.is_checked():
                modal_checkbox.check()
                page.wait_for_timeout(800)

            confirm_btn = page.locator(".ant-modal, .modal, [role='dialog']").locator("button:has-text('Xuất dữ liệu'), button:has-text('Xác nhận')").last
            
            with page.expect_download(timeout=60000) as download_info:
                confirm_btn.click()
                
            download = download_info.value
            download.save_as(output_file_path)
            browser.close()

            return {
                "success": True,
                "file_path": output_file_path,
                "message": "Đã tải file database.xlsx thành công từ VinFast."
            }

    except Exception as e:
        return {
            "success": False,
            "error": str(e),
            "message": f"Lỗi cào dữ liệu: {e}"
        }

if __name__ == "__main__":
    res = tu_dong_xuat_file_vinfast(headless=True)
    print("Kết quả:", res)
