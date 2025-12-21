import time
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
from openpyxl import load_workbook 
import os
from datetime import datetime, timedelta
import re 
import numpy as np 

# =========================================================================
# 1. CẤU HÌNH VÀ XPATH
# =========================================================================

VNDIRECT_URL = "https://banggia.vndirect.com.vn/chung-khoan/hose"
EXCEL_FILE_NAME = "VNDirect_data.xlsx"
TIMEOUT = 20
# ĐƯỜNG DẪN USER PROFILE: Thay thế bằng đường dẫn thư mục profile Chrome của bạn
USER_DATA_DIR = r"C:\Users\A22M\Programming\Python\Chrome VPS Profile" 

XPATH_SELECTORS = {
    "VNIndex":'//*[@id="charts-wrapper"]/div/div/div[1]/div[2]/p[1]/span[3]', 
    "Spread_Icon":'//*[@id="charts-wrapper"]/div/div/div[1]/div[2]/p[1]/span[2]', # XPATH: Mũi tên tăng/giảm
    "Spread":'//*[@id="charts-wrapper"]/div/div/div[1]/div[2]/p[1]/span[4]', # Xpath: Lấy cả 2 giá trị: Spread và Spread%
    "Value":'//*[@id="charts-wrapper"]/div/div/div[1]/div[2]/p[2]/span[3]', 
    "Volume":'//*[@id="charts-wrapper"]/div/div/div[1]/div[2]/p[2]/span[1]',
    "CP_Tang":'//*[@id="charts-wrapper"]/div/div/div[1]/div[2]/p[3]/span[2]',
    "CP_Giam":'//*[@id="charts-wrapper"]/div/div/div[1]/div[2]/p[3]/span[7]',
    "CP_KhongDoi":'//*[@id="charts-wrapper"]/div/div/div[1]/div[2]/p[3]/span[5]',
}

FINAL_COLUMN_ORDER = [
    'ThoiGian',
    'VNIndex',
    'Spread',
    'Spread%', 
    'Value', 
    'Volume', 
    'CP_Tang', 
    'CP_Giam', 
    'CP_KhongDoi',
]
# Các cột dùng để so sánh (Loại bỏ 'ThoiGian')
COMPARE_COLUMNS = [col for col in FINAL_COLUMN_ORDER if col != 'ThoiGian']

# Ép kiểu tất cả các cột so sánh thành chuỗi (str) khi đọc Excel để tránh lỗi kiểu hỗn hợp
DTYPE_CONVERTERS = {col: str for col in COMPARE_COLUMNS}

# =========================================================================
# 2. HÀM HỖ TRỢ KIỂM TRA DỮ LIỆU VÀ XÁC ĐỊNH NGÀY GIAO DỊCH
# =========================================================================

def get_trading_date():
    """Xác định ngày giao dịch dựa trên thời gian hiện tại (trước/sau 9:00 sáng)."""
    
    now = datetime.now()
    # Xác định mốc 9:00 sáng của ngày hiện tại
    opening_time = now.replace(hour=9, minute=0, second=0, microsecond=0)
    
    # 0 = Thứ Hai, 6 = Chủ Nhật
    weekday = now.weekday() 
    
    if now < opening_time or weekday >= 5: # Nếu trước 9:00 sáng HOẶC là T7/CN
        # Cần tìm ngày giao dịch cuối cùng: lùi ngày cho đến khi gặp T2-T6
        current_date = now.date()
        
        # Bắt đầu lùi 1 ngày
        while True:
            current_date -= timedelta(days=1)
            trading_weekday = current_date.weekday()
            
            # Nếu là ngày giao dịch hợp lệ (T2-T6) thì dùng ngày này
            if trading_weekday >= 0 and trading_weekday <= 4: 
                return current_date.strftime("%d/%m/%Y")
            
    else:
        # Nếu đã >= 9:00 sáng VÀ là T2-T6, dùng ngày hiện tại
        return now.strftime("%d/%m/%Y")


def normalize_value_for_comparison(value):
    """Chuyển đổi giá trị sang định dạng chuỗi chuẩn để so sánh."""
    if value is None or (isinstance(value, (float, np.number)) and np.isnan(value)):
        return "N/A"
    
    if isinstance(value, str):
        return value.strip().replace(',', '')
    
    try:
        if isinstance(value, (float, int)):
            if value.is_integer():
                return str(int(value))
            return "{:.3f}".format(value)
    except:
        pass 
        
    return str(value).strip().replace(',', '')


def get_last_excel_data():
    """Đọc và trả về dữ liệu của dòng cuối cùng trong file Excel (ĐÃ CHUẨN HÓA)."""
    if not os.path.isfile(EXCEL_FILE_NAME):
        return None
    try:
        # CHỈ ĐỌC CÁC CỘT CẦN SO SÁNH (BỎ CỘT THOIGIAN) VÀ ÉP KIỂU VỀ STR
        df = pd.read_excel(EXCEL_FILE_NAME, usecols=COMPARE_COLUMNS, dtype=DTYPE_CONVERTERS) 
        
        if df.empty:
            return None
            
        last_row = df.iloc[-1].to_dict()
        normalized_data = {}
        
        # CHUẨN HÓA CÁC CỘT SO SÁNH
        for col in COMPARE_COLUMNS:
            normalized_data[col] = normalize_value_for_comparison(last_row.get(col))
            
        return normalized_data
        
    except Exception as e:
        print(f"⚠️ Lỗi khi đọc và chuẩn hóa file Excel cuối cùng: {e}. Bỏ qua kiểm tra trùng lặp.")
        return None

# =========================================================================
# 3. HÀM LẤY DỮ LIỆU CHÍNH (ĐÃ CẬP NHẬT LOGIC NGÀY)
# =========================================================================

def get_market_data_and_save():
    print("🚀 Đang khởi động trình duyệt ảo...")
    chrome_options = Options()
    chrome_options.add_argument(f"user-data-dir={USER_DATA_DIR}") 
    chrome_options.add_argument("--window-size=1920,1080")

    driver = None
    try:
        driver = webdriver.Chrome(options=chrome_options)
    except Exception as e:
        print(f"❌ Lỗi khởi tạo WebDriver: {e}")
        return

    data_row = {key: "N/A" for key in COMPARE_COLUMNS}
    is_spread_negative = False 

    try:
        print(f"🌐 Truy cập website: {VNDIRECT_URL}")
        driver.get(VNDIRECT_URL)

        WebDriverWait(driver, TIMEOUT).until(
            EC.presence_of_element_located((By.XPATH, XPATH_SELECTORS['VNIndex']))
        )
        print("✅ VNIndex đã sẵn sàng.")

        # --- BƯỚC 1: XÁC ĐỊNH XU HƯỚNG TĂNG/GIẢM CỦA SPREAD DỰA TRÊN ICON ---
        try:
            icon_element = driver.find_element(By.XPATH, XPATH_SELECTORS['Spread_Icon'])
            icon_class = icon_element.get_attribute("class")
            
            if "icon-arrowdown" in icon_class.lower():
                is_spread_negative = True
                print("⬇️ Xu hướng Spread: GIẢM (sẽ thêm dấu âm '-').")
            else:
                is_spread_negative = False
                print("⬆️ Xu hướng Spread: TĂNG/KHÔNG ĐỔI (giữ nguyên).")
                
        except Exception as e:
            print(f"⚠️ Cảnh báo: Không tìm thấy icon Spread ({str(e).split('\n')[0].replace('Message: ', '')}). Mặc định Spread TĂNG.")


        # --- BƯỚC 2: LẤY DỮ LIỆU VÀ XỬ LÝ (Áp dụng logic Spread) ---
        for name, selector in XPATH_SELECTORS.items():
            if name == "Spread_Icon":
                continue 

            try:
                element = driver.find_element(By.XPATH, selector) 
                value = element.text.strip()
                
                # *** LOGIC XỬ LÝ SPREAD VÀ SPREAD% ***
                if name == "Spread":
                    raw_spread = "N/A"
                    raw_spread_percent = "N/A"
                    
                    match = re.search(r'([\d\.\,\-]+)\s+([\d\.\,\-]+%)', value)
                    
                    if match:
                        raw_spread = match.group(1).strip().replace(',', '')
                        raw_spread_percent = match.group(2).strip().replace('%', '') 
                    
                    elif '/' in value:
                         parts = value.split('/')
                         raw_spread = parts[0].strip().replace(',', '')
                         raw_spread_percent = parts[1].strip().replace('%', '')

                    # ÁP DỤNG DẤU ÂM NẾU XU HƯỚNG LÀ GIẢM (cho cả 2 cột)
                    if is_spread_negative:
                        if raw_spread != "N/A" and not raw_spread.startswith('-'):
                            data_row['Spread'] = "-" + raw_spread
                        else:
                            data_row['Spread'] = raw_spread
                        
                        if raw_spread_percent != "N/A" and not raw_spread_percent.startswith('-'):
                            data_row['Spread%'] = "-" + raw_spread_percent
                        else:
                            data_row['Spread%'] = raw_spread_percent

                    else:
                        data_row['Spread'] = raw_spread
                        data_row['Spread%'] = raw_spread_percent
                        
                    print(f"   -> Spread (điểm): {data_row['Spread']}")
                    print(f"   -> Spread% (chỉ số): {data_row['Spread%']}")
                    continue 
                # ***********************************

                # *** LOGIC XỬ LÝ VALUE (ĐỊNH DẠNG DẤU PHẨY) ***
                if name == "Value":
                    temp_value = value.replace(' tỷ', '').strip() 
                    temp_value = temp_value.replace(',', '')
                    match_final = re.search(r'([\d.]+)', temp_value)
                    
                    if match_final:
                        raw_number_str = match_final.group(1)
                        try:
                            num_value = float(raw_number_str)
                            value = "{:,.3f}".format(num_value)
                        except ValueError:
                            value = raw_number_str
                    else:
                        value = "N/A"
                # ***********************************
                
                # 3. Cập nhật dữ liệu cho các chỉ số còn lại
                data_row[name] = value 
                if name != "Spread": 
                    print(f"   -> {name}: {value}")
            
            except Exception as e:
                error_msg = str(e).split('\n')[0].replace('Message: ', '')
                print(f"❌ Lỗi: Không tìm thấy phần tử {name} | Chi tiết: {error_msg}")
                data_row[name] = "N/A" 
                if name == "Spread":
                    data_row['Spread%'] = "N/A"

        # --- BƯỚC 3: KIỂM TRA TRÙNG LẶP VÀ GHI FILE ---
        
        # 3a. Lấy bản ghi cuối cùng trong file Excel (ĐÃ CHUẨN HÓA)
        last_data_normalized = get_last_excel_data() 
        
        # Chuẩn hóa dữ liệu thu thập được để so sánh
        current_data_normalized = {col: normalize_value_for_comparison(data_row.get(col)) for col in COMPARE_COLUMNS}
        
        is_duplicate = False
        if last_data_normalized:
            # So sánh các giá trị đã được chuẩn hóa
            is_duplicate = all(current_data_normalized.get(col) == last_data_normalized.get(col) for col in COMPARE_COLUMNS)

        if is_duplicate:
            # THÔNG BÁO 1: Dữ liệu TRÙNG LẶP
            print("\n=======================================================")
            print("🚫 Dữ liệu hiện tại GIỐNG HỆT dữ liệu cuối cùng trong Excel.")
            print("➡️ **Dữ liệu đã có là dữ liệu mới nhất** (Phiên giao dịch có thể đã kết thúc).")
            print("=======================================================\n")
            # Kết thúc hàm nếu dữ liệu trùng lặp
            return 
        
        # THÔNG BÁO 2: Dữ liệu MỚI (và tiến hành ghi file)
        print("\n=======================================================")
        print("✅ Dữ liệu mới vừa được thu thập!")
        print("➡️ **Dữ liệu thu thập được lần này là dữ liệu mới nhất**.")
        print("=======================================================\n")
        
        # 3b. Ghi file Excel nếu dữ liệu KHÔNG trùng lặp
        
        # LẤY NGÀY GIAO DỊCH ĐÃ ĐIỀU CHỈNH
        data_row['ThoiGian'] = get_trading_date() 
        
        df = pd.DataFrame([data_row])[FINAL_COLUMN_ORDER]

        print(f"💾 Ghi dữ liệu vào {EXCEL_FILE_NAME}")
        file_exists = os.path.isfile(EXCEL_FILE_NAME)
        
        if file_exists:
            try:
                book = load_workbook(EXCEL_FILE_NAME)
                with pd.ExcelWriter(EXCEL_FILE_NAME, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
                    sheet = writer.book.active 
                    start_row = sheet.max_row
                    df.to_excel(writer, sheet_name=sheet.title, startrow=start_row, index=False, header=False)
            except Exception as e:
                print(f"⚠️ Lỗi khi nối thêm dữ liệu ({e}), ghi đè file.")
                df.to_excel(EXCEL_FILE_NAME, index=False, header=True, engine='openpyxl')
        else:
            df.to_excel(EXCEL_FILE_NAME, index=False, header=True, engine='openpyxl')

        print("🎉 Hoàn tất ghi file!")

    except Exception as e:
        print(f"❌ Lỗi khi quét dữ liệu tổng thể: {e}")

    finally:
        if driver:
            driver.quit()
            print("🔒 Đóng trình duyệt.")

# =========================================================================
# 4. MAIN
# =========================================================================

if __name__ == "__main__":
    get_market_data_and_save()