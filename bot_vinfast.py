"""
bot_vinfast.py - Module tự động đăng nhập, giữ session 15-20 ngày và xuất dữ liệu từ VinFast DMS
"""

import os
import sys
import time
from datetime import datetime
from playwright.sync_api import sync_playwright

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
PROFILE_DIR = os.path.join(BASE_DIR, "browser_profile_vinfast")
DOWNLOAD_DIR = os.path.join(BASE_DIR, "downloads")
os.makedirs(PROFILE_DIR, exist_ok=True)
os.makedirs(DOWNLOAD_DIR, exist_ok=True)

# URL Cổng thông tin VinFast Hậu Mãi
VINFAST_PORTAL_URL = "https://dms.vinfast.vn"
NGAY_BAT_DAU_MAC_DINH = "29/08/2026"

def dang_nhap_lan_dau_hoac_nhap_otp():
    """
    Chạy hàm này khi bạn cần đăng nhập lần đầu hoặc khi nhận thông báo hết hạn OTP.
    Trình duyệt sẽ hiện lên màn hình để bạn gõ SĐT và mã OTP một lần duy nhất.
    Phiên đăng nhập sẽ được lưu trong 15-20 ngày.
    """
    print("[*] Đang mở trình duyệt để xác thực SĐT & OTP VinFast...")
    with sync_playwright() as p:
        browser = p.chromium.launch_persistent_context(
            user_data_dir=PROFILE_DIR,
            headless=False,
            channel="chrome",
            args=["--start-maximized"]
        )
        page = browser.pages[0] if browser.pages else browser.new_page()
        page.goto(VINFAST_PORTAL_URL)
        
        print("\n👉 HÃY NHẬP SỐ ĐIỆN THOẠI & MÃ OTP TRÊN MÀN HÌNH CHROME.")
        print("👉 KHI ĐÃ VÀO ĐƯỢC TRANG CHÍNH VINFAST, HÃY ĐÓNG TRÌNH DUYỆT LẠI.\n")
        
        try:
            page.wait_for_url("**/aftersale/**", timeout=180000)
            print("[+] Đăng nhập thành công! Phiên đã được lưu 15-20 ngày.")
        except Exception:
            pass
            
        browser.close()

def tu_dong_xuat_file_vinfast(headless=True) -> dict:
    """
    Tự động vào Hậu mãi -> Lệnh sửa chữa -> Lọc 29/08 đến nay -> Tải database.xlsx.
    """
    now_str = datetime.now().strftime("%d/%m/%Y")
    output_file_path = os.path.join(DOWNLOAD_DIR, "database.xlsx")
    
    with sync_playwright() as p:
        try:
            browser = p.chromium.launch_persistent_context(
                user_data_dir=PROFILE_DIR,
                headless=headless,
                channel="chrome",
                args=["--start-maximized", "--disable-blink-features=AutomationControlled"]
            )
            page = browser.pages[0] if browser.pages else browser.new_page()
            
            # 1. Truy cập trang Lệnh sửa chữa VinFast
            page.goto(f"{VINFAST_PORTAL_URL}/aftersale/work-orders", wait_until="networkidle", timeout=45000)
            page.wait_for_timeout(3000)
            
            # 2. Kiểm tra nếu bị yêu cầu OTP
            current_url = page.url.lower()
            if "login" in current_url or page.locator("input[type='tel'], input[placeholder*='điện thoại']").count() > 0:
                browser.close()
                return {
                    "success": False,
                    "error": "OTP_REQUIRED",
                    "message": "Phiên đăng nhập VinFast đã hết hạn. Vui lòng chạy đăng nhập OTP một lần!"
                }

            # 3. Điền dải ngày: 29/08/2026 -> ngày hôm nay
            date_input = page.locator("input[placeholder*='Ngày tạo'], .ant-picker-range, [class*='date-picker']").first
            if date_input.count() > 0:
                date_input.click()
                page.wait_for_timeout(1000)
                page.keyboard.type(f"{NGAY_BAT_DAU_MAC_DINH}")
                page.keyboard.press("Enter")
                page.keyboard.type(f"{now_str}")
                page.keyboard.press("Enter")
            
            page.wait_for_timeout(1500)

            # 4. Bấm nút Xuất dữ liệu
            btn_export = page.locator("button:has-text('Xuất dữ liệu'), .btn-export, [title*='Xuất dữ liệu']").first
            btn_export.wait_for(state="visible", timeout=15000)
            btn_export.click()
            page.wait_for_timeout(2000)

            # 5. Tick chọn đại lý duy nhất trong popup
            modal_checkbox = page.locator(".ant-modal, .modal, [role='dialog']").locator("input[type='checkbox']").first
            if modal_checkbox.count() > 0 and not modal_checkbox.is_checked():
                modal_checkbox.check()
                page.wait_for_timeout(800)

            confirm_btn = page.locator(".ant-modal, .modal, [role='dialog']").locator("button:has-text('Xuất dữ liệu'), button:has-text('Đồng ý'), button:has-text('Xác nhận')").last
            
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
            if 'browser' in locals():
                browser.close()
            return {
                "success": False,
                "error": str(e),
                "message": f"Lỗi thao tác trên web VinFast: {e}"
            }

if __name__ == "__main__":
    if len(sys.argv) > 1 and sys.argv[1] == "--login":
        dang_nhap_lan_dau_hoac_nhap_otp()
    else:
        res = tu_dong_xuat_file_vinfast(headless=False)
        print("Kết quả:", res)
