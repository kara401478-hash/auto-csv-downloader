"""
案件実績CSV自動ダウンロードスクリプト
社内採用管理システムから複数条件で案件データをCSV取得し、
指定フォルダに保存する。タスクスケジューラでの定期実行を想定。
"""

import os
import sys
import time
import logging
import traceback
from dotenv import load_dotenv
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

# ============================================================
# 設定読み込み
# ============================================================
load_dotenv()

LOGIN_ID          = os.getenv("LOGIN_ID")
LOGIN_PASSWORD    = os.getenv("LOGIN_PASSWORD")
DOWNLOAD_FOLDER   = os.getenv("DOWNLOAD_FOLDER")
LOG_FILE          = os.getenv("LOG_FILE")
TARGET_URL        = os.getenv("TARGET_URL")
NAV_LABEL         = os.getenv("NAV_LABEL")
NAV_SUBLABEL      = os.getenv("NAV_SUBLABEL")
TEMPLATE_DISPATCH = os.getenv("TEMPLATE_DISPATCH")
TEMPLATE_REFERRAL = os.getenv("TEMPLATE_REFERRAL")

# ============================================================
# ログ設定
# ============================================================
logging.basicConfig(
    filename=LOG_FILE,
    level=logging.INFO,
    format="%(asctime)s %(levelname)s %(message)s",
    encoding="utf-8"
)

class StreamToLogger:
    """print/stderr をログファイルにリダイレクト"""
    def __init__(self, logger, level):
        self.logger = logger
        self.level = level
    def write(self, message):
        if message.strip():
            self.logger.log(self.level, message.strip())
    def flush(self):
        pass

sys.stdout = StreamToLogger(logging.getLogger(), logging.INFO)
sys.stderr = StreamToLogger(logging.getLogger(), logging.ERROR)

# ============================================================
# 各ダウンロード条件の定義
# （ファイル名, 雇用形態ID, ステータスIDリスト, 追加オプション）
# ============================================================
# テンプレート名は .env で設定（TEMPLATE_DISPATCH / TEMPLATE_REFERRAL）
# ファイル名も任意に変更してください
DOWNLOAD_TARGETS = [
    {
        "file_name": "dispatch_recruiting_only.csv",
        "statuses":  ["q_recruiting_status_in_recruiting"],
        "emp_types": ["q_with_employment_type_dispatch"],
        "template":  TEMPLATE_DISPATCH,
    },
    {
        "file_name": "dispatch_all_status.csv",
        "statuses":  ["q_recruiting_status_in_recruiting",
                      "q_recruiting_status_in_completed",
                      "q_recruiting_status_in_lost_order"],
        "emp_types": ["q_with_employment_type_dispatch"],
        "template":  TEMPLATE_DISPATCH,
    },
    {
        "file_name": "referral_recruiting_only.csv",
        "statuses":  ["q_recruiting_status_in_recruiting"],
        "emp_types": ["q_with_employment_type_employment_agency"],
        "template":  TEMPLATE_REFERRAL,
    },
    {
        "file_name": "scheduled_temp_recruiting_only.csv",
        "statuses":  ["q_recruiting_status_in_recruiting"],
        "emp_types": ["q_with_employment_type_scheduled_temp"],
        "template":  TEMPLATE_REFERRAL,
    },
    {
        "file_name": "daily_referral_recruiting_only.csv",
        "statuses":  ["q_recruiting_status_in_recruiting"],
        "emp_types": ["q_with_employment_type_daily_referral"],
        "template":  TEMPLATE_REFERRAL,
    },
    {
        "file_name": "training_referral_recruiting_only.csv",
        "statuses":  ["q_recruiting_status_in_recruiting"],
        "emp_types": [],
        "extra_check": "q_introduction_training_eq",
        "template":  TEMPLATE_REFERRAL,
    },
    {
        "file_name": "referral_all_status.csv",
        "statuses":  ["q_recruiting_status_in_recruiting",
                      "q_recruiting_status_in_completed",
                      "q_recruiting_status_in_lost_order",
                      "q_recruiting_status_in_consultation"],
        "emp_types": ["q_with_employment_type_employment_agency"],
        "template":  TEMPLATE_DISPATCH,
    },
    {
        "file_name": "daily_referral_all_status.csv",
        "statuses":  ["q_recruiting_status_in_recruiting",
                      "q_recruiting_status_in_completed",
                      "q_recruiting_status_in_lost_order",
                      "q_recruiting_status_in_consultation"],
        "emp_types": ["q_with_employment_type_daily_referral"],
        "template":  TEMPLATE_DISPATCH,
    },
    {
        "file_name": "training_referral_all_status.csv",
        "statuses":  ["q_recruiting_status_in_recruiting",
                      "q_recruiting_status_in_completed",
                      "q_recruiting_status_in_lost_order",
                      "q_recruiting_status_in_consultation"],
        "emp_types": [],
        "extra_check": "q_introduction_training_eq",
        "template":  TEMPLATE_DISPATCH,
    },
    {
        "file_name": "scheduled_temp_all_status.csv",
        "statuses":  ["q_recruiting_status_in_recruiting",
                      "q_recruiting_status_in_completed",
                      "q_recruiting_status_in_lost_order",
                      "q_recruiting_status_in_consultation"],
        "emp_types": ["q_with_employment_type_scheduled_temp"],
        "template":  TEMPLATE_DISPATCH,
    },
    {
        "file_name": "placement_all.csv",
        "statuses":  ["q_recruiting_status_in_recruiting"],
        "emp_types": ["q_with_employment_type_employment_agency"],
        "template":  TEMPLATE_DISPATCH,
    },
    {
        "file_name": "placement_training.csv",
        "statuses":  ["q_recruiting_status_in_recruiting"],
        "emp_types": [],
        "extra_check": "q_introduction_training_eq",
        "template":  TEMPLATE_DISPATCH,
    },
    {
        "file_name": "placement_inexperienced.csv",
        "statuses":  ["q_recruiting_status_in_recruiting"],
        "emp_types": [],
        "extra_check": "q_introduction_inexperienced_eq",
        "template":  TEMPLATE_DISPATCH,
    },
    {
        "file_name": "dispatch_senior.csv",
        "statuses":  ["q_recruiting_status_in_recruiting"],
        "emp_types": ["q_with_employment_type_dispatch",
                      "q_with_employment_type_scheduled_temp"],
        "extra_check": "q_is_senior_eq",
        "template":  TEMPLATE_DISPATCH,
    },
    {
        "file_name": "referral_tentative_order.csv",
        "statuses":  ["q_recruiting_status_in_tentative_order"],
        "emp_types": ["q_with_employment_type_dispatch",
                      "q_with_employment_type_scheduled_temp",
                      "q_with_employment_type_employment_agency"],
        "template":  TEMPLATE_DISPATCH,
    },
]

# ============================================================
# ユーティリティ関数
# ============================================================

def js_click(driver, wait, by, value):
    """JavaScriptクリック（通常クリックが効かない要素用）"""
    element = wait.until(EC.presence_of_element_located((by, value)))
    driver.execute_script("arguments[0].click();", element)


def wait_for_downloads(download_folder, timeout=60):
    """ダウンロード完了まで待機（.crdownloadが消えるまで）"""
    for _ in range(timeout):
        time.sleep(1)
        if not any(f.endswith(".crdownload") for f in os.listdir(download_folder)):
            print("Download complete.")
            return True
    print("Timeout: Download did not complete.")
    return False


def save_latest_csv(download_folder, new_file_name):
    """最新のCSVを指定ファイル名で保存（上書き）"""
    import shutil
    new_path = os.path.join(download_folder, new_file_name)
    csv_files = [f for f in os.listdir(download_folder) if f.endswith(".csv")]
    if not csv_files:
        print(f"CSV not found for: {new_file_name}")
        return
    latest = max(
        [os.path.join(download_folder, f) for f in csv_files],
        key=os.path.getmtime
    )
    if os.path.exists(new_path):
        os.remove(new_path)
    shutil.move(latest, new_path)
    print(f"Saved: {new_path}")


def close_modal_if_open(driver):
    """モーダルが開いていれば閉じる"""
    try:
        btn = WebDriverWait(driver, 5).until(
            EC.element_to_be_clickable((
                By.XPATH,
                "//div[@id='recruiter-contract-search-modal']//button[contains(@class,'close')]"
            ))
        )
        driver.execute_script("arguments[0].click();", btn)
        print("✅ モーダルを閉じました")
    except Exception:
        pass


# ============================================================
# メイン処理
# ============================================================

def setup_driver(download_folder):
    """ChromeDriverのセットアップ"""
    options = Options()
    prefs = {
        "download.default_directory": download_folder,
        "download.prompt_for_download": False,
        "directory_upgrade": True,
        "safebrowsing.enabled": True,
        "profile.default_content_setting_values.notifications": 2,
        "profile.default_content_setting_values.automatic_downloads": 1,
    }
    options.add_experimental_option("prefs", prefs)
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=options)
    driver.maximize_window()
    return driver


def login(driver, wait):
    """ログイン処理"""
    driver.get(TARGET_URL)
    wait.until(EC.presence_of_element_located(
        (By.XPATH, '/html/body/div/form/div[1]/input')
    )).send_keys(LOGIN_ID)
    driver.find_element(By.XPATH, '/html/body/div/form/div[2]/input').send_keys(LOGIN_PASSWORD)
    wait.until(EC.element_to_be_clickable(
        (By.XPATH, '/html/body/div/form/input[2]')
    )).click()
    logging.info("✅ ログイン完了")


def download_one(driver, wait, target, download_folder):
    """1件分のCSVをダウンロード"""
    close_modal_if_open(driver)
    time.sleep(3)

    # ナビ → サブメニュー
    wait.until(EC.element_to_be_clickable(
        (By.XPATH, f"//span[contains(text(),'{NAV_LABEL}')]")
    )).click()
    wait.until(EC.element_to_be_clickable(
        (By.XPATH, f"//span[@class='nav-label' and text()='{NAV_SUBLABEL}']")
    )).click()

    # 条件を指定して検索
    wait.until(EC.element_to_be_clickable(
        (By.XPATH, "/html/body/div[2]/div/div[3]/div/div/div/div[1]/div/button")
    )).click()
    time.sleep(2)

    # ステータスチェック
    for status_id in target["statuses"]:
        js_click(driver, wait, By.ID, status_id)
        time.sleep(1)

    # 雇用形態チェック
    for emp_id in target.get("emp_types", []):
        wait.until(EC.element_to_be_clickable((By.ID, emp_id))).click()
        time.sleep(1)

    # 追加オプションチェック（育成紹介・未経験・シニアなど）
    if "extra_check" in target:
        element = driver.find_element(By.ID, target["extra_check"])
        if not element.is_selected():
            element.click()
        time.sleep(1)

    # 検索
    wait.until(EC.element_to_be_clickable(
        (By.CSS_SELECTOR, "input[type='submit'][value='検索']")
    )).click()
    time.sleep(10)

    # CSVダウンロードボタン
    wait.until(EC.element_to_be_clickable(
        (By.XPATH, "/html/body/div[2]/div/div[3]/div/div/div/div[2]/div[1]/div[1]/div/button")
    )).click()

    # テンプレート選択
    wait.until(EC.element_to_be_clickable(
        (By.XPATH, f"//a[contains(text(),'{target['template']}')]")
    )).click()
    time.sleep(70)

    # サイドバーCSVダウンロード
    WebDriverWait(driver, 20).until(
        EC.element_to_be_clickable((By.XPATH, '//span[text()="CSVダウンロード"]'))
    ).click()
    wait.until(EC.element_to_be_clickable(
        (By.XPATH, "/html/body/div[2]/div/div[3]/div/div/div/div[2]/div[2]/table/tbody/tr[1]/td[1]/ul/li[1]")
    )).click()

    wait_for_downloads(download_folder)
    save_latest_csv(download_folder, target["file_name"])
    print(f"✅ 完了: {target['file_name']}")


def main():
    logging.info("====== 処理開始 ======")
    driver = setup_driver(DOWNLOAD_FOLDER)
    wait = WebDriverWait(driver, 30)

    try:
        login(driver, wait)
        for target in DOWNLOAD_TARGETS:
            try:
                download_one(driver, wait, target, DOWNLOAD_FOLDER)
            except Exception:
                logging.error(f"❌ エラー ({target['file_name']}): {traceback.format_exc()}")
    finally:
        driver.quit()
        logging.info("====== 処理終了 ======")


if __name__ == "__main__":
    main()
