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
# 0. PATH / BASE DIR
#    => Excel を必ず loadVNIndex.py と同じフォルダに保存
# =========================================================================

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

VNDIRECT_URL = "https://banggia.vndirect.com.vn/chung-khoan/hose"
EXCEL_FILE_NAME = "VNDirect_data.xlsx"
EXCEL_FILE_PATH = os.path.join(BASE_DIR, EXCEL_FILE_NAME)

TIMEOUT = 20

# Chrome ユーザープロファイル（必要に応じて変更）
USER_DATA_DIR = r"C:\Users\A22M\Programming\Python\Chrome VPS Profile"


# =========================================================================
# 1. 取得する要素（XPATH）
# =========================================================================

XPATH_SELECTORS = {
    "VNIndex": '//*[@id="charts-wrapper"]/div/div/div[1]/div[2]/p[1]/span[3]',
    "Spread_Icon": '//*[@id="charts-wrapper"]/div/div/div[1]/div[2]/p[1]/span[2]',  # 上下矢印アイコン
    "Spread": '//*[@id="charts-wrapper"]/div/div/div[1]/div[2]/p[1]/span[4]',       # Spread と Spread% が入る
    "Value": '//*[@id="charts-wrapper"]/div/div/div[1]/div[2]/p[2]/span[3]',
    "Volume": '//*[@id="charts-wrapper"]/div/div/div[1]/div[2]/p[2]/span[1]',
    "CP_Tang": '//*[@id="charts-wrapper"]/div/div/div[1]/div[2]/p[3]/span[2]',
    "CP_Giam": '//*[@id="charts-wrapper"]/div/div/div[1]/div[2]/p[3]/span[7]',
    "CP_KhongDoi": '//*[@id="charts-wrapper"]/div/div/div[1]/div[2]/p[3]/span[5]',
}

# =========================================================================
# 2. カラム名（日本語）定義
# =========================================================================

# 内部キー（英語） -> Excel出力用（日本語）
COLUMN_JP = {
    "ThoiGian": "取引日",
    "VNIndex": "VN指数",
    "Spread": "前日比(ポイント)",
    "Spread%": "前日比(%)",
    "Value": "売買代金",
    "Volume": "出来高",
    "CP_Tang": "上昇銘柄数",
    "CP_Giam": "下落銘柄数",
    "CP_KhongDoi": "変わらず銘柄数",
}

FINAL_COLUMN_ORDER_INTERNAL = [
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

FINAL_COLUMN_ORDER_JP = [COLUMN_JP[c] for c in FINAL_COLUMN_ORDER_INTERNAL]

# 比較用（取引日は除外）
COMPARE_COLUMNS_INTERNAL = [c for c in FINAL_COLUMN_ORDER_INTERNAL if c != "ThoiGian"]
COMPARE_COLUMNS_JP = [COLUMN_JP[c] for c in COMPARE_COLUMNS_INTERNAL]

# Excel を読むときは比較カラムを全部 str に
DTYPE_CONVERTERS_JP = {col: str for col in COMPARE_COLUMNS_JP}

# ログ表示用（内部キー -> 日本語ラベル）
LOG_LABEL = {
    "VNIndex": "VN指数",
    "Spread": "前日比",
    "Spread%": "前日比(%)",
    "Value": "売買代金",
    "Volume": "出来高",
    "CP_Tang": "上昇銘柄数",
    "CP_Giam": "下落銘柄数",
    "CP_KhongDoi": "変わらず銘柄数",
}


# =========================================================================
# 3. 補助関数：取引日判定、正規化、Excel最後行取得
# =========================================================================

def get_trading_date() -> str:
    """現在時刻に基づいて取引日を判定する（9:00 前 / 土日なら直近営業日）。"""
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
    """比較用に値を文字列へ正規化。"""
    if value is None:
        return "N/A"

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
    """Excel の最終行（比較対象カラムのみ）を読み、正規化して返す。"""
    if not os.path.isfile(EXCEL_FILE_PATH):
        return None

    try:
        df = pd.read_excel(
            EXCEL_FILE_PATH,
            usecols=COMPARE_COLUMNS_JP,
            dtype=DTYPE_CONVERTERS_JP
        )

        if df.empty:
            return None

        last_row = df.iloc[-1].to_dict()
        normalized = {}

        for col in COMPARE_COLUMNS_JP:
            normalized[col] = normalize_value_for_comparison(last_row.get(col))

        return normalized

    except Exception as e:
        print(f"⚠️ Excel 読み込み/正規化でエラー: {e}。重複チェックをスキップします。")
        return None


# =========================================================================
# 4. メイン処理：取得 + 重複チェック + Excel追記
# =========================================================================

def get_market_data_and_save():
    # ログ：保存先の確認
    print("📌 実行ディレクトリ(CWD):", os.getcwd())
    print("📌 スクリプトのフォルダ:", BASE_DIR)
    print("📌 Excel 保存パス:", EXCEL_FILE_PATH)

    print("\n🚀 ブラウザを起動中...")
    chrome_options = Options()
    chrome_options.add_argument(f"user-data-dir={USER_DATA_DIR}")
    chrome_options.add_argument("--window-size=1920,1080")

    driver = None
    try:
        driver = webdriver.Chrome(options=chrome_options)
    except Exception as e:
        print(f"❌ WebDriver 初期化エラー: {e}")
        return

    # 取得データ（内部キー）
    data_row_internal = {key: "N/A" for key in COMPARE_COLUMNS_INTERNAL}
    is_spread_negative = False

    try:
        print(f"🌐 サイトへアクセス: {VNDIRECT_URL}")
        driver.get(VNDIRECT_URL)

        WebDriverWait(driver, TIMEOUT).until(
            EC.presence_of_element_located((By.XPATH, XPATH_SELECTORS["VNIndex"]))
        )
        print("✅ VN指数の要素を検出しました。")

        # --- 1) Spread の増減方向をアイコンで判定 ---
        try:
            icon_element = driver.find_element(By.XPATH, XPATH_SELECTORS["Spread_Icon"])
            icon_class = (icon_element.get_attribute("class") or "").lower()

            if "icon-arrowdown" in icon_class:
                is_spread_negative = True
                print("⬇️ 前日比: 下落（マイナスを付与）")
            else:
                is_spread_negative = False
                print("⬆️ 前日比: 上昇/変わらず")

        except Exception as e:
            msg = str(e).split("\n")[0].replace("Message: ", "")
            print(f"⚠️ 前日比アイコン未検出 ({msg})。デフォルトは上昇扱い。")

        # --- 2) データ取得 ---
        for name, selector in XPATH_SELECTORS.items():
            if name == "Spread_Icon":
                continue

            try:
                element = driver.find_element(By.XPATH, selector)
                value = (element.text or "").strip()

                # Spread / Spread%
                if name == "Spread":
                    raw_spread = "N/A"
                    raw_spread_percent = "N/A"

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
                        if raw_spread != "N/A" and not raw_spread.startswith("-"):
                            data_row_internal["Spread"] = "-" + raw_spread
                        else:
                            data_row_internal["Spread"] = raw_spread

                        if raw_spread_percent != "N/A" and not raw_spread_percent.startswith("-"):
                            data_row_internal["Spread%"] = "-" + raw_spread_percent
                        else:
                            data_row_internal["Spread%"] = raw_spread_percent
                    else:
                        data_row_internal["Spread"] = raw_spread
                        data_row_internal["Spread%"] = raw_spread_percent

                    print(f"   -> 前日比(ポイント): {data_row_internal['Spread']}")
                    print(f"   -> 前日比(%): {data_row_internal['Spread%']}")
                    continue

                # Value（'tỷ' 除去 + 小数3桁整形）
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

                # その他
                data_row_internal[name] = value if value else "N/A"

                # ログ（日本語ラベル）
                if name != "Spread":
                    label = LOG_LABEL.get(name, name)
                    print(f"   -> {label}: {data_row_internal[name]}")

            except Exception as e:
                msg = str(e).split("\n")[0].replace("Message: ", "")
                label = LOG_LABEL.get(name, name)
                print(f"❌ 要素未検出: {label} | 詳細: {msg}")
                data_row_internal[name] = "N/A"
                if name == "Spread":
                    data_row_internal["Spread%"] = "N/A"

        # --- 3) 重複チェック（Excel の日本語カラムで比較）---
        last_data_normalized = get_last_excel_data()

        # 現在データ（比較用）を “日本語カラム名” に変換して正規化
        current_data_jp = {}
        for internal_key in COMPARE_COLUMNS_INTERNAL:
            jp_col = COLUMN_JP[internal_key]
            current_data_jp[jp_col] = normalize_value_for_comparison(data_row_internal.get(internal_key))

        is_duplicate = False
        if last_data_normalized:
            is_duplicate = all(
                current_data_jp.get(col) == last_data_normalized.get(col)
                for col in COMPARE_COLUMNS_JP
            )

        if is_duplicate:
            print("\n=======================================================")
            print("🚫 現在データは Excel の最終行と同一です。")
            print("➡️ 既に最新データが保存されています（取引終了の可能性あり）。")
            print("=======================================================\n")
            return

        print("\n=======================================================")
        print("✅ 新しいデータを取得しました！")
        print("➡️ 今回取得したデータを Excel に追記します。")
        print("=======================================================\n")

        # --- 4) Excel へ保存（日本語ヘッダー）---
        trading_date = get_trading_date()

        # 内部データ -> 日本語カラムへ変換
        data_row_jp = {
            COLUMN_JP["ThoiGian"]: trading_date,
            COLUMN_JP["VNIndex"]: data_row_internal.get("VNIndex", "N/A"),
            COLUMN_JP["Spread"]: data_row_internal.get("Spread", "N/A"),
            COLUMN_JP["Spread%"]: data_row_internal.get("Spread%", "N/A"),
            COLUMN_JP["Value"]: data_row_internal.get("Value", "N/A"),
            COLUMN_JP["Volume"]: data_row_internal.get("Volume", "N/A"),
            COLUMN_JP["CP_Tang"]: data_row_internal.get("CP_Tang", "N/A"),
            COLUMN_JP["CP_Giam"]: data_row_internal.get("CP_Giam", "N/A"),
            COLUMN_JP["CP_KhongDoi"]: data_row_internal.get("CP_KhongDoi", "N/A"),
        }

        df_out = pd.DataFrame([data_row_jp])[FINAL_COLUMN_ORDER_JP]

        print(f"💾 Excel に保存: {EXCEL_FILE_PATH}")
        file_exists = os.path.isfile(EXCEL_FILE_PATH)

        if file_exists:
            try:
                # 既存ファイルの最初のシートへ追記
                book = load_workbook(EXCEL_FILE_PATH)
                sheet = book.active

                # 既存ヘッダー確認（日本語じゃなければ作り直し）
                existing_header = [cell.value for cell in next(sheet.iter_rows(min_row=1, max_row=1))]
                if existing_header != FINAL_COLUMN_ORDER_JP:
                    raise ValueError("既存Excelのヘッダーが日本語カラムと一致しません（作り直しを実行）。")

                with pd.ExcelWriter(
                    EXCEL_FILE_PATH,
                    engine="openpyxl",
                    mode="a",
                    if_sheet_exists="overlay"
                ) as writer:
                    sheet2 = writer.book.active
                    start_row = sheet2.max_row
                    df_out.to_excel(
                        writer,
                        sheet_name=sheet2.title,
                        startrow=start_row,
                        index=False,
                        header=False
                    )

            except Exception as e:
                print(f"⚠️ 追記に失敗: {e}")
                print("➡️ 日本語カラムで新規作成（上書き）します。")
                df_out.to_excel(EXCEL_FILE_PATH, index=False, header=True, engine="openpyxl")
        else:
            df_out.to_excel(EXCEL_FILE_PATH, index=False, header=True, engine="openpyxl")

        print("🎉 保存完了！")

    except Exception as e:
        print(f"❌ 全体処理エラー: {e}")

    finally:
        if driver:
            driver.quit()
            print("🔒 ブラウザを終了しました。")


# =========================================================================
# MAIN
# =========================================================================

if __name__ == "__main__":
    get_market_data_and_save()
