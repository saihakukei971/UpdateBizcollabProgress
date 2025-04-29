import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime
import re
import math

# ▼認証設定
CREDENTIAL_FILE = 'updatebizcollabreport-18c214963c69.json'
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
creds = Credentials.from_service_account_file(CREDENTIAL_FILE, scopes=SCOPES)
client = gspread.authorize(creds)

# ▼スプレッドシートキーとシート名
REPORT_KEY = '1I09qhpkRx0JC8jJaXzkFj5rpAeMlSm1-c2_YtUUFxOw'
PROGRESS_KEY = '1BD0YswHVbCxFqTEWO9uGTlV64pyJDGgM7qhK89p_Lks'
ID_SHEET_NAME = 'マイム合計値ID検索シート'

REPORT_SHEET_NAME = '日時レポート'

# ▼列変換
def col_letter_to_index(letter):
    index = 0
    for c in letter:
        index = index * 26 + (ord(c.upper()) - ord('A') + 1)
    return index

def col_index_to_letter(index):
    result = ''
    while index > 0:
        index, rem = divmod(index - 1, 26)
        result = chr(65 + rem) + result
    return result

# ▼今月シート取得（名前順で最大の年月を持つシートを右端と見なす）
def get_progress_sheet():
    ss = client.open_by_key(PROGRESS_KEY)
    all_sheets = ss.worksheets()
    current_name = f"{datetime.now().year}年{datetime.now().month}月"

    # ▼年と月の形式にマッチするシートを抽出
    date_pattern = re.compile(r"(\d{4})年(\d{1,2})月")
    matched_sheets = []

    for sheet in all_sheets:
        title = sheet.title.strip()
        match = date_pattern.fullmatch(title)
        if match:
            year, month = int(match.group(1)), int(match.group(2))
            matched_sheets.append((year, month, sheet))

    if not matched_sheets:
        print(f"[ERROR] 『〇〇年〇〇月』形式のシートが見つかりません。")
        return None

    # ▼年と月で降順ソート → 一番新しいものを取得
    matched_sheets.sort(reverse=True)
    latest_year, latest_month, rightmost_sheet = matched_sheets[0]
    rightmost_title = rightmost_sheet.title.strip()

    if rightmost_title != current_name:
        print(f"[ERROR] 進捗シート『{current_name}』が見つかりません。")
        print(f"        ※一番右端の年月シートは『{rightmost_title}』です。")
        print(f"        →『{current_name}』が右端に存在しないため、処理を中止します。")
        return None

    return rightmost_sheet

# ▼広告枠設定の取得（ID・開始セル・行範囲）
def get_table_config_from_sheet():
    sheet = client.open_by_key(PROGRESS_KEY).worksheet(ID_SHEET_NAME)
    ad_ids = sheet.col_values(2)[2:]
    progress_cells = sheet.col_values(3)[2:]
    config = {}
    for ad_id, progress_cell in zip(ad_ids, progress_cells):
        ad_id = ad_id.strip()
        if not ad_id or not ad_id.isdigit():
            continue
        match = re.match(r"([A-Z]+)(\d+)", progress_cell.strip())
        if not match:
            continue
        col_letter, row_num = match.groups()
        row_start = int(row_num)
        config[ad_id] = {
            'report_col': ad_id,
            'start_cell': f"{col_letter}3",
            'row_range': (row_start, row_start + 45)
        }
    print("[DEBUG] 自動生成された table_config:")
    for k, v in config.items():
        print(f"  {k}: {v}")
    return config

# ▼日時レポートからfam8/mime/補填フラグを取得
def get_values_with補填判定(sheet, col_idx):
    col_letter = col_index_to_letter(col_idx + 1)
    max_row = 300  # 検索範囲の最大行（必要なら拡大）
    for row in range(100, max_row + 1):
        base_cell = f"{col_letter}{row}"
    # ▼FORMULA形式の列データを一括取得（429対策）
    formula_col_range = f"{col_letter}100:{col_letter}{max_row}"
    try:
        formula_column = sheet.batch_get(
            [formula_col_range],
            value_render_option='FORMULA'
        )[0]

    except Exception as e:
        print(f"[DEBUG] 式列の一括取得に失敗：{type(e).__name__} - {e}")
        return None, None, False, None, None, None, None

    for i, row in enumerate(formula_column):
        cell_val = row[0] if len(row) > 0 else ''
        if isinstance(cell_val, str) and (cell_val.startswith("=") or cell_val.strip() != ""):
            data_row = 100 + i - 1
            fam8_cell = f"{col_letter}{data_row}"
            mime_col_letter = col_index_to_letter(col_idx + 2)
            mime_cell = f"{mime_col_letter}{data_row}"
            補填_col_letter = col_index_to_letter(col_idx + 4)
            補填_cell = f"{補填_col_letter}{data_row}"


            # ▼一括取得で値を得る（429対策強化）
            try:
                full_range = sheet.get(f"{fam8_cell}:{補填_cell}")
                if not full_range or len(full_range) == 0:
                    fam8_val = mime_val = 補填_val = None
                else:
                    row_vals = full_range[0]
                    fam8_val = row_vals[0] if len(row_vals) > 0 else None
                    mime_val = row_vals[1] if len(row_vals) > 1 else None
                    補填_val = row_vals[2] if len(row_vals) > 2 else None

                if fam8_val is None:
                    print(f"[DEBUG] fam8セルが取得できませんでした: {fam8_cell}")

            except Exception as e:
                print(f"[DEBUG] 一括取得での値取得失敗：{type(e).__name__} - {e}")
                continue


            safe_fam8 = str(fam8_val).encode('ascii', 'ignore').decode('ascii') if fam8_val is not None else ''
            print(f"[DEBUG] fam8取得 - {fam8_cell} = '{safe_fam8}'")


            safe_mime = str(mime_val).encode('ascii', 'ignore').decode('ascii') if mime_val is not None else ''
            print(f"[DEBUG] mime取得 - {mime_cell} = '{safe_mime}'")



            safe_補填 = str(補填_val).encode('ascii', 'ignore').decode('ascii') if 補填_val is not None else ''
            print(f"[DEBUG] 補填判定取得 - {補填_cell} = '{safe_補填}'")



            # ▼補填対象の判定：文字列が「補填対象」または5%以上の%値ならTrue
            is_補填 = False
            if isinstance(補填_val, str):
                if 補填_val.strip() == "補填対象":
                    is_補填 = True
                elif 補填_val.strip().endswith('%'):
                    try:
                        percent_val = float(補填_val.strip().replace('%', ''))
                        if percent_val >= 5.0:
                            is_補填 = True
                    except ValueError:
                        pass


            return fam8_val, mime_val, is_補填, data_row, col_letter, data_row, 補填_col_letter


    return None, None, False, None, None, None, None

# ▼書き込み行を見つける
def find_write_row(progress_data, col_idx, row_range):
    col_data = []
    for i in range(row_range[0] - 1, row_range[1]):
        if i < len(progress_data):
            row = progress_data[i]
            val = row[col_idx - 1] if col_idx - 1 < len(row) else ''
            col_data.append(val)
        else:
            col_data.append('')
    for i in reversed(range(len(col_data))):
        if col_data[i].strip():
            return row_range[0] + i + 1
    return row_range[0]

# ▼メイン処理
def main():
    report_sheet = client.open_by_key(REPORT_KEY).worksheet(REPORT_SHEET_NAME)

    progress_sheet = get_progress_sheet()
    if progress_sheet is None:
        print("[STOP] 今月の進捗シートが右端にないため、処理を中止しました。")
        return  # 処理を完全停止

    progress_sheet_name = progress_sheet.title.strip()

    progress_data = progress_sheet.get_all_values()
    header_row = report_sheet.row_values(1)
    table_config = get_table_config_from_sheet()
    logs = []
    updates = []

    for ad_id, config in table_config.items():
        print(f"[INFO] 処理中 → ID: {ad_id}")
        report_col = config['report_col']
        start_cell = config['start_cell']
        row_range = config['row_range']
        base_col = col_letter_to_index(''.join(filter(str.isalpha, start_cell)))
        fam8_col = base_col + 5
        mime_col = base_col + 1

        if report_col not in header_row:
            missing_col_index = header_row.index(report_col) if report_col in header_row else "（該当なし）"
            logs.append(
                f"● 広告枠ID: {ad_id}\n"
                f"　┗ 日時レポ列 '{report_col}' がヘッダーに見つかりません（列インデックス: {missing_col_index}）\n"
                f"　┗ ヘッダー最初の10列: {header_row[:10]}\n"
                f"　┗ 処理結果：スキップ"
            )
            continue


        report_col_idx = header_row.index(report_col)
        fam8_val, mime_val, is_補填, base_row, col_letter, fam8_row, 補填_col_letter = get_values_with補填判定(report_sheet, report_col_idx)
        fam8 = fam8_val




        if fam8_val is None or fam8 == '':
            fam8_val_raw = fam8_val if fam8_val is not None else '(None)'
            logs.append(
                f"● 広告枠ID: {ad_id}\n"
                f"　┗ 値取得失敗（fam8セルが空 or 取得できず）\n"
                f"　┗ fam8取得元セル：{col_letter}{fam8_row} ／ 値：'{fam8_val_raw}'\n"
                f"　┗ 処理結果：スキップ"
            )
            continue



        try:
            if fam8_val.strip() == '-' or fam8_val.strip() == '':
                fam8_val_f = None
            else:
                # カンマ・¥・空白・その他非数値文字を除去してfloat化
                fam8_val_f = float(re.sub(r"[^\d.\-]", "", fam8_val))


        except:
            fam8_val_f = None

        try:
            if mime_val and mime_val.strip() != '-':
                mime_val_f = float(mime_val.replace(',', ''))
            else:
                mime_val_f = None

        except:
            mime_val_f = None

        if fam8_val_f is None:
            logs.append(f"● 広告枠ID: {ad_id}\n　┗ fam8値が数値変換できません\n　┗ 処理結果：スキップ")
            continue

        write_row = find_write_row(progress_data, fam8_col, row_range)
        fam8_cell = f"{col_index_to_letter(fam8_col)}{write_row}"
        mime_cell = f"{col_index_to_letter(mime_col)}{write_row}"
        fam8_書き込み = {'range': fam8_cell, 'values': [[fam8_val_f]]}

        if is_補填:
            mime_calc = round(fam8_val_f * 0.95)
            updates.append(fam8_書き込み)
            updates.append({'range': mime_cell, 'values': [[mime_calc]]})

            補填セル = f"{補填_col_letter}{base_row}"
            fam8セル = f"{col_letter}{fam8_row}"

            logs.append(
                f"● 補填処理：{ad_id}\n"
                f"　┗ fam8（{fam8セル}）: {fam8_val_f}\n"
                f"　┗ mime(補填) = ROUND({fam8_val_f} * 0.95) → {mime_calc}\n"
                f"　┗ 判定セル：{補填セル} = '補填対象'\n"
                f"　┗ 書き込み先：進捗シート『{progress_sheet_name}』 → fam8 → {fam8_cell}, mime → {mime_cell}\n"
                f"　┗ 処理結果：OK"
            )
        else:
            mime_output = "" if mime_val_f in [None, ""] else mime_val_f
            updates.append(fam8_書き込み)
            updates.append({'range': mime_cell, 'values': [[mime_output]]})

            logs.append(
                f"● 通常処理：{ad_id}\n"
                f"　┗ fam8（{col_letter}{fam8_row}）: {fam8_val_f}\n"
                f"　┗ mime（{col_index_to_letter(report_col_idx + 2)}{base_row}）: {mime_val_f}\n"
                f"　┗ 書き込み先：進捗シート『{progress_sheet_name}』 → fam8 → {fam8_cell}, mime → {mime_cell}\n"
                f"　┗ 処理結果：OK"
            )



    if updates:
        progress_sheet.batch_update(updates, value_input_option='USER_ENTERED')

    print("\n".join(logs))
    print(f"[INFO] 処理件数：{len(table_config)} 件")
    print("[INFO] 全処理完了。")

# ▼実行
if __name__ == "__main__":
    main()