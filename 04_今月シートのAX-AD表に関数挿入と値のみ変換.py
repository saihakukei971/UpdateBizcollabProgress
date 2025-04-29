
import gspread
from google.oauth2.service_account import Credentials
import requests
import datetime
from google.auth.transport.requests import Request
import json
import time

# **Googleスプレッドシートに接続**
def get_google_sheet():
    scope = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
    creds = Credentials.from_service_account_file("updatebizcollabreport-18c214963c69.json", scopes=scope)
    client = gspread.authorize(creds)
    sheet_id = "1BD0YswHVbCxFqTEWO9uGTlV64pyJDGgM7qhK89p_Lks"
    return client.open_by_key(sheet_id), creds

def convert_to_date(value):
    if isinstance(value, (datetime.datetime, datetime.date)):
        return value
    try:
        return datetime.datetime.strptime(value, "%Y/%m/%d").date()
    except ValueError:
        return None

def column_letter(n):
    result = ""
    while n > 0:
        n -= 1
        result = chr(65 + (n % 26)) + result
        n //= 26
    return result

def update_with_user_entered_force(sheet, creds, all_requests):
    if not creds.valid:
        creds.refresh(Request())
    access_token = creds.token
    spreadsheet_id = sheet.spreadsheet.id
    url = f"https://sheets.googleapis.com/v4/spreadsheets/{spreadsheet_id}:batchUpdate"
    headers = {"Authorization": f"Bearer {access_token}", "Content-Type": "application/json"}
    response = requests.post(url, headers=headers, data=json.dumps({"requests": all_requests}))
    if response.status_code != 200:
        raise Exception(f"[API ERROR] batchUpdate失敗: {response.status_code}: {response.text}")
    else:
        print(f"[INFO] 関数挿入 batchUpdate 成功：{len(all_requests)} リクエスト実行済")

def set_right_alignment(sheet, start_col, end_col, row):
    sheet_id = sheet._properties["sheetId"]
    requests_payload = [{
        "repeatCell": {
            "range": {
                "sheetId": sheet_id,
                "startRowIndex": row - 1,
                "endRowIndex": row,
                "startColumnIndex": start_col - 1,
                "endColumnIndex": end_col
            },
            "cell": {
                "userEnteredFormat": {
                    "horizontalAlignment": "RIGHT"
                }
            },
            "fields": "userEnteredFormat.horizontalAlignment"
        }
    }]
    access_token = sheet.spreadsheet.client.auth.token
    url = f"https://sheets.googleapis.com/v4/spreadsheets/{sheet.spreadsheet.id}:batchUpdate"
    headers = {"Authorization": f"Bearer {access_token}", "Content-Type": "application/json"}
    response = requests.post(url, headers=headers, data=json.dumps({"requests": requests_payload}))
    if response.status_code != 200:
        raise Exception(f"[API ERROR] 右寄せ適用失敗: {response.status_code}: {response.text}")

def batch_update_values(sheet, creds, batch_cells):
    """
    batch_cells: List[dict]
        [
            {
                "range": {
                    "sheetId": SHEET_ID,
                    "startRowIndex": 88,
                    "endRowIndex": 89,
                    "startColumnIndex": 9,
                    "endColumnIndex": 13
                },
                "values": ["¥55459", "¥433", "¥305", "¥128"]
            },
            ...
        ]
    """
    if not creds.valid:
        creds.refresh(Request())
    access_token = creds.token
    spreadsheet_id = sheet.spreadsheet.id

    requests_payload = []

    for item in batch_cells:
        formatted_values = []
        for val in item["values"]:
            try:
                num = float(val.replace(",", "").replace("¥", ""))
                # 小数第1位で四捨五入（整数なら小数なし、少数ありなら1位含め下の位は切り捨てた）
                formatted = f"¥{round(num):,}"
            except:
                formatted = val  # エラー時はそのまま
            formatted_values.append({
                "userEnteredValue": {"stringValue": formatted},
                "userEnteredFormat": {
                    "textFormat": {
                        "foregroundColor": {"red": 0, "green": 0, "blue": 0}
                    }
                }
            })

        row_values = formatted_values


        requests_payload.append({
            "updateCells": {
                "rows": [{"values": row_values}],
                "fields": "userEnteredValue,userEnteredFormat.textFormat.foregroundColor",
                "range": item["range"]
            }
        })



    url = f"https://sheets.googleapis.com/v4/spreadsheets/{spreadsheet_id}:batchUpdate"
    headers = {"Authorization": f"Bearer {access_token}", "Content-Type": "application/json"}
    response = requests.post(url, headers=headers, data=json.dumps({"requests": requests_payload}))

    if response.status_code != 200:
        raise Exception(f"[API ERROR] 値貼付一括失敗: {response.status_code}: {response.text}")
    else:
        print(f"[INFO] 値のみ貼付一括完了: {len(requests_payload)}リクエスト")


def get_valid_number_from_cell(sheet, cell):
    cell_value = sheet.acell(cell).value
    # 空でない場合に変換を試みる
    if not cell_value or cell_value.strip() == "":
        raise ValueError(f"[ERROR] {cell} は空または無効です")

    # 余計なスペースやカンマを削除してから数値として扱う
    cell_value_cleaned = cell_value.replace(",", "").strip()  # カンマを削除し、前後のスペースも除去

    try:
        # 文字列を数値に変換
        return float(cell_value_cleaned)  # 数値として扱う
    except ValueError:
        raise ValueError(f"[ERROR] {cell} の値が数値に変換できません: {cell_value}")

def generate_axad_formulas(id_cell_value, cpm_cell, imp_col, sales_col, pay_col, profit_col, row_num):
    id_cell_value_int = int(id_cell_value)  # 整数化することで完全に一致させる
    print(f"[INFO] 数式に使用する ID（整数化済）: {id_cell_value_int}")

    # ★★★ ここをシングルクォートで囲むよう修正（これが重要！） ★★★
    sheet_name = "'集計用シート（AX-AD）'"

    return [
        f'=IFERROR(VLOOKUP({id_cell_value_int}, {sheet_name}!$C:$L, 10, FALSE), 0)',
        f'=IF(OR({imp_col}{row_num}=0, {cpm_cell}=""), 0, {imp_col}{row_num}/1000*{cpm_cell})',
        f'=IFERROR(VLOOKUP({id_cell_value_int}, {sheet_name}!$C:$R, 16, FALSE), 0)',
        f'=IF({sales_col}{row_num}=0, 0, {sales_col}{row_num}-{pay_col}{row_num})'
    ]



def main():
    book, creds = get_google_sheet()
    # 対象シート名を「YYYY年M月」形式で動的に取得
    today = datetime.date.today()
    sheet_name = f"{today.year}年{today.month}月"
    axad_sheet = book.worksheet(sheet_name)
    print(f"[INFO] 処理対象シート：{sheet_name}")

    meta_sheet = book.worksheet("マイム合計値ID検索シート")

    ref_cell = meta_sheet.acell("E3").value
    if not ref_cell:
        print("[ERROR] マイム合計値ID検索シート E3 にセル番地がありません")
        return

    col_letter = ''.join(filter(str.isalpha, ref_cell))
    start_row = int(''.join(filter(str.isdigit, ref_cell))) + 2
    date_col_index = gspread.utils.a1_to_rowcol(f"{col_letter}1")[1]
    date_cells = axad_sheet.col_values(date_col_index)[start_row - 1:start_row - 1 + 31]

    today = datetime.date.today()
    yesterday = today - datetime.timedelta(days=1)

    today_row = yesterday_row = None
    for idx, val in enumerate(date_cells):
        date = convert_to_date(val)
        if date == today and not today_row:
            today_row = start_row + idx
        elif date == yesterday and not yesterday_row:
            yesterday_row = start_row + idx

    if not today_row or not yesterday_row:
        print("[ERROR] 本日または昨日の日付が見つかりません")
        return

    print(f"[INFO] 処理対象行：今日={today_row} / 昨日={yesterday_row}")
    print(f"\n[INFO] 処理対象行\n(今日){today.strftime('%Y/%m/%d')}→{today_row}行目に関数挿入\n(昨日){yesterday.strftime('%Y/%m/%d')}→{yesterday_row}行目に値のみ変換\n")

    all_requests = []
    batch_data = []
    batch_right_alignments = []  # ← 右寄せ整形用バッチ

    id_cells_range = []
    id_values = []

    for i in range(0, (78 - 8) // 7):
        base_col = 9 + i * 7
        id_col = column_letter(base_col)
        id_cell = f"{id_col}86"
        id_cells_range.append(id_cell)

    # 個別に取得し、確実に変換する
    for cell in id_cells_range:
        try:
            value = axad_sheet.acell(cell).value
            if not value or value.strip() == "":
                print(f"[警告] {cell} は空")
                id_values.append(0)
                continue
            value_cleaned = value.replace(",", "").strip()
            id_values.append(float(value_cleaned))
        except Exception as e:
            print(f"[エラー] セル{cell}取得中エラー: {e}")
            id_values.append(0)

    print(f"[INFO] 取得後のIDセルの値（数値変換済）: {id_values}")




    for i, id_value in enumerate(id_values):
        base_col = 9 + i * 7
        imp_col = column_letter(base_col + 1)
        sales_col = column_letter(base_col + 2)
        pay_col = column_letter(base_col + 3)
        profit_col = column_letter(base_col + 4)
        cpm_col = column_letter(base_col + 2) + "85"

        formulas = generate_axad_formulas(id_value, cpm_col, imp_col, sales_col, pay_col, profit_col, today_row)

        formula_cells = [{
            "userEnteredValue": {"formulaValue": f},
            "userEnteredFormat": {
                "textFormat": {"foregroundColor": {"red": 1, "green": 1, "blue": 1}},
                "borders": {
                    "top": {"style": "SOLID", "color": {"red": 0, "green": 0, "blue": 0}},
                    "bottom": {"style": "SOLID", "color": {"red": 0, "green": 0, "blue": 0}},
                    "left": {"style": "SOLID", "color": {"red": 0, "green": 0, "blue": 0}},
                    "right": {"style": "SOLID", "color": {"red": 0, "green": 0, "blue": 0}}
                }
            }
        } for f in formulas]



        all_requests.append({
            "updateCells": {
                "rows": [{"values": formula_cells}],
                "fields": "userEnteredValue,userEnteredFormat",
                "range": {
                    "sheetId": axad_sheet._properties["sheetId"],
                    "startRowIndex": today_row - 1,
                    "endRowIndex": today_row,
                    "startColumnIndex": base_col,
                    "endColumnIndex": base_col + 4
                }
            }
        })

        print(f"[INFO] 数式挿入の範囲: 行 {today_row}, 列 {base_col} - {base_col + 4}")
        print(f"[INFO] 挿入した関数: {formulas}")

    # 関数一括挿入
    update_with_user_entered_force(axad_sheet, creds, all_requests)

    # 評価待ち時間を増やす（2秒→5秒）
    time.sleep(5)

    for i in range(0, (78 - 8) // 7):
        base_col = 9 + i * 7
        start_col_letter = column_letter(base_col + 1)
        end_col_letter = column_letter(base_col + 4)
        range_name = f"{start_col_letter}{today_row}:{end_col_letter}{today_row}"

        try:
            values = axad_sheet.get(range_name, value_render_option='FORMATTED_VALUE')
            if not values:
                cell_list = axad_sheet.range(range_name)
                flat_values = [cell.value for cell in cell_list]
                if not any(flat_values):
                    flat_values = ["0", "0", "0", "0"]
            else:
                flat_values = values[0] if values else ["0", "0", "0", "0"]

            for j in range(len(flat_values)):
                if flat_values[j] == "#ERROR!" or not flat_values[j]:
                    flat_values[j] = "0"
                    # 右寄せ範囲に含める
                    batch_right_alignments.append({
                        "repeatCell": {
                            "range": {
                                "sheetId": axad_sheet._properties["sheetId"],
                                "startRowIndex": yesterday_row - 1,
                                "endRowIndex": yesterday_row,
                                "startColumnIndex": base_col + 1 + j - 1,
                                "endColumnIndex": base_col + 1 + j
                            },
                            "cell": {
                                "userEnteredFormat": {
                                    "horizontalAlignment": "RIGHT"
                                }
                            },
                            "fields": "userEnteredFormat.horizontalAlignment"
                        }
                    })
        except Exception as e:
            print(f"[警告] 範囲 {range_name} の取得中にエラー: {str(e)}")
            flat_values = ["0", "0", "0", "0"]

        print(f"[INFO] 評価結果を取得 → 範囲: {range_name} / 値: {flat_values}")
        # batch_data に追加しておく（あとでまとめて貼付）
        batch_data.append({
            "range": {
                "sheetId": axad_sheet._properties["sheetId"],
                "startRowIndex": yesterday_row - 1,
                "endRowIndex": yesterday_row,
                "startColumnIndex": base_col + 1 - 1,
                "endColumnIndex": base_col + 1 - 1 + 4
            },
            "values": flat_values
        })

        # batch_right_alignments に追加（あとでまとめて右寄せ一括実行）
        batch_right_alignments.append({
            "repeatCell": {
                "range": {
                    "sheetId": axad_sheet._properties["sheetId"],
                    "startRowIndex": yesterday_row - 1,
                    "endRowIndex": yesterday_row,
                    "startColumnIndex": base_col + 1 - 1,
                    "endColumnIndex": base_col + 1 - 1 + 4
                },
                "cell": {
                    "userEnteredFormat": {
                        "horizontalAlignment": "RIGHT"
                    }
                },
                "fields": "userEnteredFormat.horizontalAlignment"
            }
        })

    # 値の一括貼付もここで実行（これが抜けていたため、値が貼られなかった）
    if batch_data:
        batch_update_values(axad_sheet, creds, batch_data)
        print(f"[INFO] 値のみ貼付一括実行成功: {len(batch_data)} リクエスト")

    # 右寄せ整形も一括実行（省略せずに全てまとめる）
    if batch_right_alignments:
        url = f"https://sheets.googleapis.com/v4/spreadsheets/{axad_sheet.spreadsheet.id}:batchUpdate"
        headers = {
            "Authorization": f"Bearer {creds.token}",
            "Content-Type": "application/json"
        }
        response = requests.post(url, headers=headers, data=json.dumps({"requests": batch_right_alignments}))
        if response.status_code != 200:
            raise Exception(f"[API ERROR] 右寄せ一括整形失敗: {response.status_code}: {response.text}")
        else:
            print(f"[INFO] 右寄せ一括整形成功: {len(batch_right_alignments)} リクエスト")



    print("[正常終了] 対象の広告枠IDの表に対して関数の挿入と前日行への値のみ貼付が完了しました")
    print(f"[確認用] 関数を挿入した行（本日）: {today_row} 行目 / 値を貼り付けた行（前日）: {yesterday_row} 行目")


if __name__ == "__main__":
    main()