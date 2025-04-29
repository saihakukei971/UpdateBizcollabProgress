import os
import pandas as pd
import chardet
from datetime import datetime, timedelta
from google.oauth2.service_account import Credentials
import gspread

# === 【1】定数設定 ===
SPREADSHEET_ID = "1BD0YswHVbCxFqTEWO9uGTlV64pyJDGgM7qhK89p_Lks"
SHEET_NAME = "集計用シート（AX-AD）"
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]
CREDENTIALS_FILE = "updatebizcollabreport-18c214963c69.json"  # 同階層に配置されていること

# === 【2】CSVファイルのパス構築 ===
yesterday = (datetime.today() - timedelta(days=1)).strftime("%Y%m%d")
csv_path = os.path.join("アップデート×ビズコラボ日次CSV", yesterday, f"AX-ADレポート_{yesterday}.csv")

# === 【3】文字コード検出 ===
def detect_encoding(file_path):
    with open(file_path, "rb") as f:
        return chardet.detect(f.read())["encoding"]

# === 【4】CSV読み込み・整形（skiprows + total除外 + 空白除外）===
def preprocess_csv(csv_path):
    if not os.path.exists(csv_path):
        print(f"[WARNING] CSVが見つかりません: {csv_path}")
        return None
    print(f"[DEBUG] 読込CSV: {csv_path}")
    encoding = detect_encoding(csv_path)
    try:
        df = pd.read_csv(csv_path, skiprows=2, encoding=encoding)
    except Exception:
        df = pd.read_csv(csv_path, skiprows=2, encoding="cp932")
    # 広告枠IDが空白の行除去 + total 行除去
    df = df[df.iloc[:, 0].notna()]
    df = df[~df.iloc[:, 0].astype(str).str.contains("total", case=False, na=False)]
    return df

# === 【5】列番号 → アルファベット列名変換（例: 1 → A, 27 → AA）===
def get_column_letter(col_index):
    result = ''
    while col_index > 0:
        col_index, rem = divmod(col_index - 1, 26)
        result = chr(65 + rem) + result
    return result

# === 【6】Google Sheets へアップロード（A2 から、カラム名含む）===
def upload_to_sheet(client, df):
    if df is None or df.empty:
        print("[SKIP] CSVデータが空です")
        return

    sheet = client.open_by_key(SPREADSHEET_ID).worksheet(SHEET_NAME)
    df = df.fillna("")

    # A2 から開始 → A2 にカラム行、A3 以降にデータ
    num_cols = len(df.columns)
    num_rows = len(df)
    end_col_letter = get_column_letter(num_cols)
    end_cell = f"{end_col_letter}{2 + num_rows}"  # A2〜最終データ行
    range_str = f"A2:{end_cell}"

    try:
        sheet.update(range_name=range_str, values=[df.columns.tolist()] + df.values.tolist())
        print(f"[INFO] アップロード完了 → {range_str}")
    except Exception as e:
        print(f"[ERROR] アップロード失敗: {e}")

# === 【7】Google認証処理 ===
def authenticate_google():
    creds = Credentials.from_service_account_file(CREDENTIALS_FILE, scopes=SCOPES)
    return gspread.authorize(creds)

# === 【8】メイン処理 ===
def main():
    print("[INFO] アップロード処理を開始")
    print(f"[CHECK] CSV パス: {csv_path}")
    print(f"[CHECK] 存在: {os.path.exists(csv_path)}")

    client = authenticate_google()
    df = preprocess_csv(csv_path)
    upload_to_sheet(client, df)
    print("[INFO] 全処理完了")

# === 【9】実行 ===
if __name__ == "__main__":
    main()
