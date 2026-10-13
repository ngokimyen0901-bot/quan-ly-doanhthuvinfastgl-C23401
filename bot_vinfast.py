"""
bot_vinfast.py - Module tự động xuất dữ liệu VinFast DMS (chọn đúng ô Ngày tạo)
"""
import os
import sys
import time
from datetime import datetime
from playwright.sync_api import sync_playwright
import notifier

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
PROFILE_DIR = os.path.join(BASE_DIR, "browser_profile_vinfast")
DOWNLOAD_DIR = os.path.join(BASE_DIR, "downloads")
os.makedirs(PROFILE_DIR, exist_ok=True)
os.makedirs(DOWNLOAD_DIR, exist_ok=True)

VINFAST_PORTAL_URL = "https://dms.vinfast.vn"
WORK_ORDERS_URL = f"{VINFAST_PORTAL_URL}/aftersale/work-orders"
NGAY_BAT_DAU_MAC_DINH = "29/08/2026"

def chon_chinh_xac_ngay_tao(page, tu_ngay, den_ngay):
    print("[*] Đang tìm ô 'Ngày tạo'...")
    page.wait_for_timeout(2000)
    start_input = None
    end_input = None

    try:
        label_tao = page.locator("xpath=//*[text()[normalize-space()='Ngày tạo'] or contains(., 'Ngày tạo')]").first
        if label_tao.count() > 0:
            start_input = label_tao.locator("xpath=following::input[1]")
            end_input = label_tao.locator("xpath=following::input[2]")
    except Exception:
        pass

    if not start_input or start_input.count() == 0:
        container = page.locator(".ant-form-item, .filter-item, div").filter(
            has_text=lambda t: "ngày tạo" in t.lower() and "ngày quyết toán" not in t.lower()
        ).first
        if container.count() > 0:
            inputs = container.locator("input")
            if inputs.count() >= 2:
                start_input = inputs.nth(0)
                end_input = inputs.nth(1)

    if not start_input or start_input.count() == 0:
        pickers = page.locator(".ant-picker-range, [class*='range-picker']")
        for i in range(pickers.count()):
            p_elem = pickers.nth(i)
            ctx = p_elem.locator("xpath=ancestor::*[position()<=3]").text_content() or ""
            if "tạo" in ctx.lower() and "quyết toán" not in ctx.lower():
                inputs = p_elem.locator("input")
                start_input = inputs.nth(0)
                end_input = inputs.nth(1)
                break

    if not start_input or start_input.count() == 0:
        raise Exception("Không tìm thấy ô 'Ngày tạo' trên giao diện VinFast!")

    # Điền Ngày bắt đầu
    start_input.scroll_into_view_if_needed()
    start_input.click()
    page.wait_for_timeout(300)
    start_input.press("Control+a")
    start_input.press("Backspace")
    page.wait_for_timeout(200)
    start_input.type(tu_ngay, delay=60)
    start_input.press("Enter")
    page.wait_for_timeout(500)

    # Điền Ngày kết thúc
    end_input.click()
    page.wait_for_timeout(300)
    end_input.press("Control+a")
    end_input.press("Backspace")
    page.wait_for_timeout(200)
    end_input.type(den_ngay, delay=60)
    end_input.press("Enter")
    page.wait_for_timeout(600)

    page.keyboard.press("Escape")
    page.wait_for_timeout(500)
    print(f"[+] Đã chọn xong Ngày tạo: {tu_ngay} -> {den_ngay}")

def tu_dong_xuat_file_vinfast(headless=False) -> dict:
    now_str = datetime.now().strftime("%d/%m/%Y")
    output_file_path = os.path.join(DOWNLOAD_DIR, "database.xlsx")
    pw = sync_playwright().start()
    browser = None
    try:
        browser = pw.chromium.launch_persistent_context(
            user_data_dir=PROFILE_DIR,
            headless=headless,
            channel="chrome",
            args=["--start-maximized", "--disable-blink-features=AutomationControlled"]
        )
        page = browser.pages[0] if browser.pages else browser.new_page()
        try:
            page.goto(WORK_ORDERS_URL, wait_until="load", timeout=60000)
        except Exception:
            pass
        page.wait_for_timeout(3000)

        # Kiểm tra OTP
        current_url = page.url.lower()
        if "login" in current_url or page.locator("input[type='tel'], input[placeholder*='điện thoại']").count() > 0:
            notifier.thong_bao_canh_bao_otp()
            return {"success": False, "error": "OTP_REQUIRED", "message": "VinFast yêu cầu nhập OTP. Vui lòng mở Chrome để nhập OTP!"}

        chon_chinh_xac_ngay_tao(page, NGAY_BAT_DAU_MAC_DINH, now_str)

        btn_export = page.locator("button:has-text('Xuất dữ liệu'), .btn-export, [title*='Xuất dữ liệu']").first
        btn_export.wait_for(state="visible", timeout=20000)
        btn_export.click()
        page.wait_for_timeout(2000)

        modal = page.locator(".ant-modal, .modal, [role='dialog']").first
        modal.wait_for(state="visible", timeout=15000)
        checkbox = modal.locator(".ant-checkbox, input[type='checkbox']").first
        if checkbox.count() > 0 and modal.locator(".ant-checkbox-checked").count() == 0:
            checkbox.click()
            page.wait_for_timeout(800)

        confirm_btn = modal.locator("button:has-text('Xuất dữ liệu'), button:has-text('Đồng ý'), button:has-text('Xác nhận')").last
        with page.expect_download(timeout=90000) as download_info:
            confirm_btn.click()

        download = download_info.value
        download.save_as(output_file_path)
        return {"success": True, "file_path": output_file_path, "message": "Tải file thành công!"}
    except Exception as e:
        return {"success": False, "error": str(e), "message": f"Lỗi VinFast: {e}"}
    finally:
        if browser:
            try: browser.close()
            except Exception: pass
        try: pw.stop()
        except Exception: pass
