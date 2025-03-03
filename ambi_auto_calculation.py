import os
import time
import string
from datetime import datetime, timedelta
from dotenv import load_dotenv
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
import gspread
from google.oauth2.service_account import Credentials

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

# Google Sheets設定
SHEET_NAME = "テスト"  # スプレッドシート名
SERVICE_ACCOUNT_FILE = "service_account.json"  # サービスアカウントのJSONファイル

def setup_driver():
    """
    Chromeウェブドライバーを設定する
    """
    chrome_options = Options()
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

def fetch_data_by_contact_names(driver, date, data_type, contact_names):
    """
    指定されたjobNameに対応するデータを取得する
    """
    query_params = f"?_pp_=date_from%3D{date}%7Cdate_to%3D{date}&{COMMON_PARAMS}"
    url = f"{BASE_URL}{ENDPOINTS[data_type]}{query_params}"
    
    driver.get(url)
    time.sleep(2)  # ページロード待機
    
    results = []

    for contact_name in contact_names:
        try:
            # jobNameに一致する行を特定
            row = driver.find_element(By.XPATH, f"//div[@class='jobName' and text()='{contact_name}']/ancestor::tr")
            
            # その行に含まれるdataクラスを持つtd要素を取得
            data_tds = row.find_elements(By.CLASS_NAME, "data")
            
            # 各td要素のテキストを取得
            scout_mail_stats = [td.text for td in data_tds]
            if len(scout_mail_stats) == 8:
                scout_mail_stats_dict = dict(
                    contact_name = scout_mail_stats[0],
                    interested_count = scout_mail_stats[1],
                    passed_judgement_count = scout_mail_stats[2],
                    passed_judgement_rate = scout_mail_stats[3],
                    entry_count = scout_mail_stats[4],
                    entry_rate = scout_mail_stats[5],
                    interview_req_count = scout_mail_stats[6],
                    interview_req_rate = scout_mail_stats[7],
                )
            else:
                scout_mail_stats_dict = dict(
                    contact_name = scout_mail_stats[0],
                    send_count = scout_mail_stats[1],
                    opens_count = scout_mail_stats[2],
                    open_rate = scout_mail_stats[3],
                    refusals_count = scout_mail_stats[4],
                    entry_count = scout_mail_stats[5],
                    post_opening_entry_rate = scout_mail_stats[6],
                    entry_rate = scout_mail_stats[7],
                    interview_req_count = scout_mail_stats[8],
                    interview_req_rate = scout_mail_stats[9],
                )

        except Exception as e:
            # 指定したjobNameが見つからない場合のエラーハンドリング
            print(f"{contact_name}のデータ取得中にエラーが発生しました: {e}")
            scout_mail_stats_dict = dict(
                contact_name = contact_name,
                send_count = 0,
                opens_count = 0,
                open_rate = 0,
                refusals_count = 0,
                entry_count = 0,
                post_opening_entry_rate = 0,
                entry_rate = 0,
                interview_req_count = 0,
                interview_req_rate = 0,
                interested_count = 0,
                passed_judgement_count = 0,
                passed_judgement_rate = 0,
            )

        results.append({
            "contact_name": contact_name,
            "scout_mail_stats_dict": scout_mail_stats_dict
        })
        print(f"{contact_name}のデータ:{scout_mail_stats_dict}")
    
    return results

def data_entry_position(contact_name, scout_type):
    """
    データを書き込む位置を取得
    """
    if scout_type == "platinum":
        if contact_name == "山中沙矢":
            return 20
        elif contact_name == "橘萌生":
            return 37
        elif contact_name == "奥野翔子":
            return 71
    elif scout_type == "regular":
        if contact_name == "山中沙矢":
            return 25
        elif contact_name == "橘萌生":
            return 42
        elif contact_name == "奥野翔子":
            return 76
    elif scout_type == "interested":
        if contact_name == "山中沙矢":
            return 30
        elif contact_name == "橘萌生":
            return 47
        elif contact_name == "奥野翔子":
            return 81


def get_column_letter(column_number):
    """
    列番号をExcelの列文字に変換する
    例: 1 -> A, 2 -> B, ..., 27 -> AA
    """
    column_letter = ""
    while column_number > 0:
        column_number, remainder = divmod(column_number - 1, 26)
        column_letter = chr(65 + remainder) + column_letter
    return column_letter

def get_column_from_date(date_str):
    """
    YYYY-MM-DD形式の日付を受け取り、対応する列番号（整数）を返す。
    例: '2025-01-01' → 7 (G列), '2025-01-02' → 8 (H列)
         一週間ごとに1列空ける。例: '2025-01-08' → 15 (O列)
    """
    import datetime

    base_column_index = ord("G") - ord("A") + 1  # 'G'列はExcel上で7列目
    date = datetime.datetime.strptime(date_str, "%Y-%m-%d")
    day_of_month = date.day

    # 一週間ごとに1列空ける。7日ごとに1列追加。
    extra_columns = (day_of_month - 1) // 7

    # ベース列 + 日数 - 1 + 空ける列数
    column_number = base_column_index + day_of_month - 1 + extra_columns
    return get_column_letter(column_number)

def write_to_google_sheets(all_scout_data):
    """
    データをGoogleスプレッドシートに書き込む
    """
    # Google Sheets API認証
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    credentials = Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE,
        scopes=scope
    )
    gc = gspread.authorize(credentials)

    # バッチ書き込みのためのリクエストリスト
    requests_by_sheet = {}

    for entry in all_scout_data:
        date = entry["date"] # 日付
        scout_type = entry["data_type"] # スカウトタイプ
        contact_name = entry["contact_name"] # エージェント名
        scout_mail_stats_dict = entry["scout_mail_stats_dict"] # 集計データの辞書
        send_count = int(scout_mail_stats_dict["send_count"]) if scout_type in ["platinum", "regular"] else 0 # 送信数
        opens_count = int(scout_mail_stats_dict["opens_count"]) if scout_type in ["platinum", "regular"] else 0 # 開封数
        entry_count = int(scout_mail_stats_dict["entry_count"]) #エントリー数
        interested_count = int(scout_mail_stats_dict["interested_count"]) if scout_type == "interested" else 0 # 興味あり数
        interested_entry_count = int(scout_mail_stats_dict["entry_count"]) if scout_type == "interested" else 0 # 興味ありエントリー数

        # 該当セルにデータを書き込むリクエストを追加
        sheet_name = date[:7].replace("-", ".")  # YYYY-MM形式のシート名
        if sheet_name not in requests_by_sheet:
            requests_by_sheet[sheet_name] = []

        if scout_type in ["platinum", "regular"]:
            requests_by_sheet[sheet_name].append({
                'range': f"{get_column_from_date(date)}{data_entry_position(contact_name, scout_type)}",
                'values': [[send_count]]
            })
            requests_by_sheet[sheet_name].append({
                'range': f"{get_column_from_date(date)}{data_entry_position(contact_name, scout_type) + 1}",
                'values': [[opens_count]]
            })
            requests_by_sheet[sheet_name].append({
                'range': f"{get_column_from_date(date)}{data_entry_position(contact_name, scout_type) + 3}",
                'values': [[entry_count]]
            })
        elif scout_type == "interested":
            requests_by_sheet[sheet_name].append({
                'range': f"{get_column_from_date(date)}{data_entry_position(contact_name, scout_type)}",
                'values': [[interested_count]]
            })
            requests_by_sheet[sheet_name].append({
                'range': f"{get_column_from_date(date)}{data_entry_position(contact_name, scout_type) + 1}",
                'values': [[interested_entry_count]]
            })

    # 各シートに対してバッチ書き込みを実行
    for sheet_name, requests in requests_by_sheet.items():
        sheet = gc.open("AMBIエントリー分析_エージェント_2024").worksheet(sheet_name)
        sheet.batch_update(requests)

    print("データの更新が完了しました！")

def main():
    driver = None
    all_scout_data = []

    try:
        # Webドライバーの設定
        driver = setup_driver()

        # AMBIにログイン
        if not login_to_ambi(driver):
            return

        # データ収集
        today = datetime.today()
        start_date = today - timedelta(days=14)  # 過去14日分
        # 指定するjobNameのリスト
        contact_names = ["橘萌生", "奥野翔子"]

        for single_date in (start_date + timedelta(n) for n in range(14)):
            formatted_date = single_date.strftime("%Y-%m-%d")
            print(f"\n{formatted_date}のデータ収集を開始:")

            for data_type in ENDPOINTS.keys():
                data = fetch_data_by_contact_names(driver, formatted_date, data_type, contact_names)
                
                for entry in data:
                    all_scout_data.append({
                        "date": formatted_date,
                        "data_type": data_type,
                        "contact_name": entry["contact_name"],
                        "scout_mail_stats_dict": entry["scout_mail_stats_dict"]
                    })

                time.sleep(1)
        # データをGoogleスプレッドシートに書き込む
        write_to_google_sheets(all_scout_data)

    except Exception as e:
        print(f"スクリプト実行中に致命的なエラーが発生しました: {e}")
    finally:
        if driver:
            driver.quit()

if __name__ == "__main__":
    main()
