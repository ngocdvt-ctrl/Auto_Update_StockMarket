import time
import os
import re
from datetime import datetime, timedelta

import numpy as np
import pandas as pd
from openpyxl import load_workbook

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC


# =========================================================================
# 0. PATH / BASE DIR (QUAN TRỌNG)
#    => Ép Excel luôn nằm "cùng folder với loadVNIndex.py"
# =========================================================================

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

VNDIRECT_URL = "https://banggia.vndirect.com.vn/chung-khoan/hose"
EXCEL_FILE_NAME = "VNDirect_data.xlsx"
EXCEL_FILE_PATH = os.path.join(BASE_DIR, EXCEL_FILE_NAME)

TIMEOUT = 20

# ĐƯỜNG DẪN USER PROFILE: Thay thế bằng đường dẫn thư mục profile Chrome của bạn
USER_DATA_DIR = r"C:\Users\A22M\Programming\Python\Chrome VPS Profile"

XPATH_SELECTORS = {
    "VNIndex": '//*[@id="charts-wrapper"]/div/div/div[1]/div[2]/p[1]/span[3]',
    "Spread_Icon": '//*[@id="charts-wrapper"]/div/div/div[1]/div[2]/p[1]/span[2]',  # icon tăng/giảm
    "Spread": '//*[@id="charts-wrapper"]/div/div/div[1]/div[2]/p[1]/span[4]',       # cả Spread và Spread%
    "Value": '//*[@id="charts-wrapper"]/div/div/div[1]/div[2]/p[2]/span[3]',
    "Volume": '//*[@id="charts-wrapper"]/div/div/div[1]/div[2]/p[2]/span[1]',
    "CP_Tang": '//*[@id="charts-wrapper"]/div/div/div[1]/div[2]/p[3]/span[2]',
    "CP_Giam": '//*[@id="charts-wrapper"]/div/div/div[1]/div[2]/p[3]/span[7]',
    "CP_KhongDoi": '//*[@id="charts-wrapper"]/div/div/div[1]/div[2]/p[3]/span[5]',
}

FINAL_COLUMN_ORDER = [
    "ThoiGian",
    "VNIndex",
    "Spread",
    "Spread%",
    "Value",
    "Volume",
    "CP_Tang",
    "CP_Giam",
    "CP_KhongDoi",
]

# Các cột dùng để so sánh (Loại bỏ 'ThoiGian')
COMPARE_COLUMNS = [col for col in FINAL_COLUMN_ORDER if col != "ThoiGian"]

# Ép kiểu tất cả các cột so sánh thành chuỗi (str) khi đọc Excel để tránh lỗi kiểu hỗn hợp
DTYPE_CONVERTERS = {col: str for col in COMPARE_COLUMNS}


# =========================================================================
# 1. HÀM HỖ TRỢ: NGÀY GIAO DỊCH + CHUẨN HÓA + ĐỌC DÒNG CUỐI EXCEL
# =========================================================================

def get_trading_date() -> str:
    """Xác định ngày giao dịch dựa trên thời gian hiện tại (trước/sau 9:00 sáng).
    - Trước 9:00 hoặc T7/CN => lùi về ngày gần nhất T2-T6.
    - Sau 9:00 và T2-T6 => dùng ngày hiện tại.
    """
    now = datetime.now()
    opening_time = now.replace(hour=9, minute=0, second=0, microsecond=0)
    weekday = now.weekday()  # 0=Mon ... 6=Sun

    if now < opening_time or weekday >= 5:
        current_date = now.date()
        while True:
            current_date -= timedelta(days=1)
            if 0 <= current_date.weekday() <= 4:
                return current_date.strftime("%d/%m/%Y")
    else:
        return now.strftime("%d/%m/%Y")


def normalize_value_for_comparison(value) -> str:
    """Chuyển đổi giá trị sang định dạng chuỗi chuẩn để so sánh."""
    if value is None:
        return "N/A"

    # NaN
    if isinstance(value, (float, np.number)) and np.isnan(value):
        return "N/A"

    if isinstance(value, str):
        return value.strip().replace(",", "")

    try:
        if isinstance(value, (float, int)):
            if isinstance(value, float) and value.is_integer():
                return str(int(value))
            if isinstance(value, int):
                return str(value)
            return "{:.3f}".format(float(value))
    except Exception:
        pass

    return str(value).strip().replace(",", "")


def get_last_excel_data():
    """Đọc và trả về dữ liệu của dòng cuối cùng trong file Excel (ĐÃ CHUẨN HÓA)."""
    if not os.path.isfile(EXCEL_FILE_PATH):
        return None

    try:
        df = pd.read_excel(
            EXCEL_FILE_PATH,
            usecols=COMPARE_COLUMNS,
            dtype=DTYPE_CONVERTERS
        )

        if df.empty:
            return None

        last_row = df.iloc[-1].to_dict()
        normalized = {}

        for col in COMPARE_COLUMNS:
            normalized[col] = normalize_value_for_comparison(last_row.get(col))

        return normalized

    except Exception as e:
        print(f"⚠️ Lỗi khi đọc/chuẩn hóa Excel: {e}. Bỏ qua kiểm tra trùng lặp.")
        return None


# =========================================================================
# 2. HÀM CHÍNH: QUÉT DATA + CHECK TRÙNG + GHI EXCEL
# =========================================================================

def get_market_data_and_save():
    # LOG để anh biết chắc chắn đang ghi file vào đâu
    print("📌 Current Working Directory (CWD):", os.getcwd())
    print("📌 Script folder (BASE_DIR):       ", BASE_DIR)
    print("📌 Excel path (EXCEL_FILE_PATH):   ", EXCEL_FILE_PATH)

    print("\n🚀 Đang khởi động trình duyệt ảo...")
    chrome_options = Options()
    chrome_options.add_argument(f"user-data-dir={USER_DATA_DIR}")
    chrome_options.add_argument("--window-size=1920,1080")

    driver = None
    try:
        driver = webdriver.Chrome(options=chrome_options)
    except Exception as e:
        print(f"❌ Lỗi khởi tạo WebDriver: {e}")
        return

    # default row
    data_row = {key: "N/A" for key in COMPARE_COLUMNS}
    is_spread_negative = False

    try:
        print(f"🌐 Truy cập website: {VNDIRECT_URL}")
        driver.get(VNDIRECT_URL)

        WebDriverWait(driver, TIMEOUT).until(
            EC.presence_of_element_located((By.XPATH, XPATH_SELECTORS["VNIndex"]))
        )
        print("✅ VNIndex đã sẵn sàng.")

        # --- BƯỚC 1: XÁC ĐỊNH XU HƯỚNG SPREAD (DỰA TRÊN ICON) ---
        try:
            icon_element = driver.find_element(By.XPATH, XPATH_SELECTORS["Spread_Icon"])
            icon_class = (icon_element.get_attribute("class") or "").lower()

            if "icon-arrowdown" in icon_class:
                is_spread_negative = True
                print("⬇️ Xu hướng Spread: GIẢM (sẽ thêm dấu âm '-').")
            else:
                is_spread_negative = False
                print("⬆️ Xu hướng Spread: TĂNG/KHÔNG ĐỔI (giữ nguyên).")

        except Exception as e:
            msg = str(e).split("\n")[0].replace("Message: ", "")
            print(f"⚠️ Cảnh báo: Không tìm thấy icon Spread ({msg}). Mặc định Spread TĂNG.")

        # --- BƯỚC 2: LẤY DỮ LIỆU ---
        for name, selector in XPATH_SELECTORS.items():
            if name == "Spread_Icon":
                continue

            try:
                element = driver.find_element(By.XPATH, selector)
                value = (element.text or "").strip()

                # ===== Spread + Spread% =====
                if name == "Spread":
                    raw_spread = "N/A"
                    raw_spread_percent = "N/A"

                    # dạng "1.23 0.45%" hoặc tương tự
                    match = re.search(r"([\d\.\,\-]+)\s+([\d\.\,\-]+%)", value)
                    if match:
                        raw_spread = match.group(1).strip().replace(",", "")
                        raw_spread_percent = match.group(2).strip().replace("%", "")
                    elif "/" in value:
                        parts = value.split("/")
                        if len(parts) >= 2:
                            raw_spread = parts[0].strip().replace(",", "")
                            raw_spread_percent = parts[1].strip().replace("%", "")

                    if is_spread_negative:
                        # thêm dấu âm nếu chưa có
                        if raw_spread != "N/A" and not raw_spread.startswith("-"):
                            data_row["Spread"] = "-" + raw_spread
                        else:
                            data_row["Spread"] = raw_spread

                        if raw_spread_percent != "N/A" and not raw_spread_percent.startswith("-"):
                            data_row["Spread%"] = "-" + raw_spread_percent
                        else:
                            data_row["Spread%"] = raw_spread_percent
                    else:
                        data_row["Spread"] = raw_spread
                        data_row["Spread%"] = raw_spread_percent

                    print(f"   -> Spread (điểm): {data_row['Spread']}")
                    print(f"   -> Spread% (chỉ số): {data_row['Spread%']}")
                    continue

                # ===== Value: bỏ 'tỷ' + format 3 chữ số thập phân =====
                if name == "Value":
                    temp = value.replace(" tỷ", "").strip()
                    temp = temp.replace(",", "")
                    m = re.search(r"([\d.]+)", temp)
                    if m:
                        raw_number_str = m.group(1)
                        try:
                            num_value = float(raw_number_str)
                            value = "{:,.3f}".format(num_value)
                        except ValueError:
                            value = raw_number_str
                    else:
                        value = "N/A"

                # các field còn lại
                data_row[name] = value if value else "N/A"

                if name != "Spread":
                    print(f"   -> {name}: {data_row[name]}")

            except Exception as e:
                msg = str(e).split("\n")[0].replace("Message: ", "")
                print(f"❌ Lỗi: Không tìm thấy phần tử {name} | Chi tiết: {msg}")
                data_row[name] = "N/A"
                if name == "Spread":
                    data_row["Spread%"] = "N/A"

        # --- BƯỚC 3: CHECK TRÙNG LẶP ---
        last_data_normalized = get_last_excel_data()
        current_data_normalized = {
            col: normalize_value_for_comparison(data_row.get(col))
            for col in COMPARE_COLUMNS
        }

        is_duplicate = False
        if last_data_normalized:
            is_duplicate = all(
                current_data_normalized.get(col) == last_data_normalized.get(col)
                for col in COMPARE_COLUMNS
            )

        if is_duplicate:
            print("\n=======================================================")
            print("🚫 Dữ liệu hiện tại GIỐNG HỆT dữ liệu cuối cùng trong Excel.")
            print("➡️ **Dữ liệu đã có là dữ liệu mới nhất** (Phiên giao dịch có thể đã kết thúc).")
            print("=======================================================\n")
            return

        print("\n=======================================================")
        print("✅ Dữ liệu mới vừa được thu thập!")
        print("➡️ **Dữ liệu thu thập được lần này là dữ liệu mới nhất**.")
        print("=======================================================\n")

        # --- BƯỚC 4: GHI EXCEL ---
        data_row["ThoiGian"] = get_trading_date()
        df = pd.DataFrame([data_row])[FINAL_COLUMN_ORDER]

        print(f"💾 Ghi dữ liệu vào: {EXCEL_FILE_PATH}")
        file_exists = os.path.isfile(EXCEL_FILE_PATH)

        if file_exists:
            try:
                # append vào sheet đang active
                book = load_workbook(EXCEL_FILE_PATH)
                with pd.ExcelWriter(
                    EXCEL_FILE_PATH,
                    engine="openpyxl",
                    mode="a",
                    if_sheet_exists="overlay"
                ) as writer:
                    sheet = writer.book.active
                    start_row = sheet.max_row
                    df.to_excel(
                        writer,
                        sheet_name=sheet.title,
                        startrow=start_row,
                        index=False,
                        header=False
                    )
            except Exception as e:
                print(f"⚠️ Lỗi khi nối thêm dữ liệu ({e}), sẽ ghi đè file.")
                df.to_excel(EXCEL_FILE_PATH, index=False, header=True, engine="openpyxl")
        else:
            df.to_excel(EXCEL_FILE_PATH, index=False, header=True, engine="openpyxl")

        print("🎉 Hoàn tất ghi file!")

    except Exception as e:
        print(f"❌ Lỗi khi quét dữ liệu tổng thể: {e}")

    finally:
        if driver:
            driver.quit()
            print("🔒 Đóng trình duyệt.")


# =========================================================================
# 3. MAIN
# =========================================================================

if __name__ == "__main__":
    get_market_data_and_save()
