import os
import time
import traceback
from datetime import datetime, timedelta
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager

# === 【0】定数定義 ===
ID = 'admin'
PASSWORD = 'fhC7UPJiforgKTJ8'
BASE_SAVE_DIR = os.path.join(os.path.dirname(__file__), "アップデート×ビズコラボ日次CSV")
SEARCH_KEYWORD = "AX-AD"
CSV_FILENAME_PREFIX = "AX-ADレポート_"
LOGIN_URL = "https://admin.fam-8.net/report/index.php"

# 前日付取得（例：20250331）
yesterday_str = (datetime.today() - timedelta(days=1)).strftime("%Y%m%d")
target_dir = os.path.join(BASE_SAVE_DIR, yesterday_str)
os.makedirs(target_dir, exist_ok=True)

# === 【1】表示項目チェックボックス定義 ===
DISPLAY_ITEMS_XPATHS = {
    "メディアオーナーID": '//*[@id="display_itemsseller_id"]',
    "メディアオーナー名": '//*[@id="display_itemsseller_name"]',
    "メディアID": '//*[@id="display_itemssite_id"]',
    "CPC(グロス)": '//*[@id="display_itemscpc_gross"]',
    "CPA(グロス)": '//*[@id="display_itemscpa_gross"]',
    "eCPM(グロス)": '//*[@id="display_itemscpm_gross"]',
    "メディア名": '//*[@id="display_itemssite_name"]',
    "特別保証期間(From)": '//*[@id="display_itemsfixed_cost_from"]',
}

# === 【2】ChromeDriverセットアップ ===
def setup_driver():
    print("[INFO] WebDriver を起動します")
    options = webdriver.ChromeOptions()
    prefs = {"download.default_directory": target_dir}
    options.add_experimental_option("prefs", prefs)
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=options)
    driver.get(LOGIN_URL)
    print("[INFO] WebDriver の起動完了")
    return driver

# === 【3】ログイン処理 ===
def login(driver):
    try:
        print("[INFO] ログイン処理を開始します")
        wait = WebDriverWait(driver, 10)
        user_input = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="topmenu"]/tbody/tr[2]/td/div[1]/form/div/table/tbody/tr[1]/td/input')))
        user_input.clear()
        user_input.send_keys(ID)
        pass_input = driver.find_element(By.XPATH, '//*[@id="topmenu"]/tbody/tr[2]/td/div[1]/form/div/table/tbody/tr[2]/td/input')
        pass_input.clear()
        pass_input.send_keys(PASSWORD)
        driver.find_element(By.XPATH, '//*[@id="topmenu"]/tbody/tr[2]/td/div[1]/form/div/table/tbody/tr[3]/td/input[2]').click()
        time.sleep(2)
        print("[INFO] ログイン完了")
    except Exception as e:
        print("[ERROR] ログイン処理失敗:")
        print(traceback.format_exc())
        driver.quit()
        exit(1)

# === 【4】ブラウザ操作とCSVダウンロード ===
def operate_and_download(driver):
    try:
        wait = WebDriverWait(driver, 10)

        print("[INFO] レポートメニューをクリック")
        wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="sidemenu"]/div[3]/a[4]/div'))).click()
        time.sleep(1)

        print("[INFO] 表示モード切り替えをクリック")
        wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="display_modesummary_mode"]'))).click()
        time.sleep(1)

        print(f"[INFO] キーワード『{SEARCH_KEYWORD}』を入力")
        input_box = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="main_area"]/form/div[1]/input[7]')))
        input_box.clear()
        input_box.send_keys(SEARCH_KEYWORD)
        time.sleep(1)

        print("[INFO] 表示項目設定をクリック")
        display_link = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="main_area"]/form/div[1]/a[3]')))
        display_link.click()
        time.sleep(1)

        print("[INFO] 表示項目チェックを開始")
        for name, xpath in DISPLAY_ITEMS_XPATHS.items():
            checkbox = wait.until(EC.presence_of_element_located((By.XPATH, xpath)))
            if not checkbox.is_selected():
                checkbox.click()
                print(f"[INFO] [CHECKED] {name} → チェックON")
            else:
                print(f"[INFO] [CHECKED] {name} → 既にON")

        print("[INFO] 検索ボタンをクリック")
        wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="main_area"]/form/div[1]/input[10]'))).click()

        print("[INFO] 検索結果を待機中（5秒）")
        time.sleep(5)

        print("[INFO] CSVダウンロードボタンをクリック")
        csv_button = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="topmenu"]/table/tbody/tr/td[4]/table/tbody/tr[2]/td[2]/input[1]')))
        csv_button.click()
        print("[SUCCESS] CSVダウンロードを実行しました")

        # 6秒間待機してファイルダウンロード完了を想定
        time.sleep(6)

    except Exception as e:
        print("[ERROR] 操作中にエラー発生:")
        print(traceback.format_exc())
        driver.quit()
        exit(1)

# === 【5】CSVリネーム処理（最新のcsvファイルを対象） ===
def rename_downloaded_csv():
    files = [f for f in os.listdir(target_dir) if f.lower().endswith(".csv")]
    if not files:
        print("[ERROR] ダウンロードCSVが見つかりません")
        exit(1)

    latest_file = max(files, key=lambda f: os.path.getmtime(os.path.join(target_dir, f)))
    old_path = os.path.join(target_dir, latest_file)
    new_filename = f"{CSV_FILENAME_PREFIX}{yesterday_str}.csv"
    new_path = os.path.join(target_dir, new_filename)

    if os.path.exists(new_path):
        print(f"[WARNING] 同名のCSVが既に存在 → 上書き: {new_filename}")
        os.remove(new_path)

    os.rename(old_path, new_path)
    print(f"[INFO] CSVファイルを {new_filename} にリネーム完了")

# === 【6】メイン ===
if __name__ == "__main__":
    driver = setup_driver()
    login(driver)
    operate_and_download(driver)
    driver.quit()
    rename_downloaded_csv()
    print("[INFO] スクリプト正常終了")
