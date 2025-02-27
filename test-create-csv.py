import os
import time
import csv
from datetime import datetime, timedelta

from dotenv import load_dotenv
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

# 環境変数の読み込み
load_dotenv()

# AMBIログイン情報
AMBI_LOGIN_URL = os.getenv("AMBI_LOGIN_URL")
AMBI_LOGIN_ID = os.getenv("AMBI_LOGIN_ID")
AMBI_PASSWORD = os.getenv("AMBI_PASSWORD")

# 基本URLとデータ種別ごとのエンドポイント
BASE_URL = "https://en-ambi.com/company/effect_ma"
ENDPOINTS = {
    "platinum": "/acc_scout/platinum/",
    "regular": "/acc_scout/",
    "interested": "/acc_interests/",
}
COMMON_PARAMS = "PK=CA19C6"

def setup_driver():
    """
    Chromeウェブドライバーを設定する
    """
    chrome_options = Options()
    # デバッグ時はヘッドレスモードをコメントアウト
    # chrome_options.add_argument("--headless")
    chrome_options.add_argument("--start-maximized")
    chrome_options.add_argument("--disable-extensions")
    
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=chrome_options)
    return driver

def login_to_ambi(driver):
    """
    AMBIサイトにログイン
    """
    driver.get(AMBI_LOGIN_URL)
    wait = WebDriverWait(driver, 10)

    username_field = wait.until(EC.presence_of_element_located((By.NAME, "accLoginID")))
    password_field = wait.until(EC.presence_of_element_located((By.NAME, "accLoginPW")))
    login_button = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "button[type='submit']")))

    username_field.send_keys(AMBI_LOGIN_ID)
    password_field.send_keys(AMBI_PASSWORD)
    login_button.click()

    # ログイン後の画面が読み込まれるまで待機
    WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.TAG_NAME, "body"))
    )
    print("AMBIサイトにログインしました。")
    return True

def fetch_data_by_job_names(driver, date, data_type, target_job_names):
    """
    指定されたjobNameに対応するデータを取得する
    """
    query_params = f"?_pp_=date_from%3D{date}%7Cdate_to%3D{date}&{COMMON_PARAMS}"
    url = f"{BASE_URL}{ENDPOINTS[data_type]}{query_params}"
    
    driver.get(url)
    time.sleep(2)  # ページロード待機
    
    results = []

    for target_job_name in target_job_names:
        try:
            # jobNameに一致する行を特定
            row = driver.find_element(By.XPATH, f"//div[@class='jobName' and text()='{target_job_name}']/ancestor::tr")
            
            # その行に含まれるdataクラスを持つtd要素を取得
            data_tds = row.find_elements(By.CLASS_NAME, "data")
            
            # 各td要素のテキストを取得
            data_values = [td.text for td in data_tds]
            
            # 結果を保存
            results.append({
                "jobName": target_job_name,
                "data": data_values
            })
            print(f"{target_job_name}のデータ:")
            print(data_values)
        
        except Exception as e:
            # 指定したjobNameが見つからない場合のエラーハンドリング
            print(f"{target_job_name}のデータ取得中にエラーが発生しました: {e}")
    
    return results

def save_to_csv(data_list, filename='ambi_data.csv'):
    """
    データをCSVファイルに保存
    """
    keys = ['date', 'data_type', 'jobName', 'data']
    
    with open(filename, 'w', newline='', encoding='utf-8') as csvfile:
        writer = csv.DictWriter(csvfile, fieldnames=keys)
        writer.writeheader()
        
        for data in data_list:
            writer.writerow(data)
    
    print(f"データを{filename}に保存しました。")

def main():
    driver = None
    all_data = []

    try:
        # Webドライバーの設定
        driver = setup_driver()

        # AMBIにログイン
        if not login_to_ambi(driver):
            return

        # データ収集
        today = datetime.today()
        start_date = today - timedelta(days=3)  # 過去60日分

        # 指定するjobNameのリスト
        target_job_names = ["山中沙矢", "橘萌生", "奥野翔子"]

        for single_date in (start_date + timedelta(n) for n in range(3)):
            formatted_date = single_date.strftime("%Y-%m-%d")
            print(f"\n{formatted_date}のデータ収集を開始:")

            for data_type in ENDPOINTS.keys():
                data = fetch_data_by_job_names(driver, formatted_date, data_type, target_job_names)
                
                # 収集したデータを結果リストに追加
                for entry in data:
                    all_data.append({
                        "date": formatted_date,
                        "data_type": data_type,
                        "jobName": entry["jobName"],
                        "data": entry["data"]
                    })

                time.sleep(1)  # サーバー負荷軽減のため

        # データをCSVに保存
        save_to_csv(all_data)
    except Exception as e:
        print(f"スクリプト実行中に致命的なエラーが発生しました: {e}")
    finally:
        if driver:
            driver.quit()

if __name__ == "__main__":
    main()
