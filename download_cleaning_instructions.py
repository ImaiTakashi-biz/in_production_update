import sqlite3
import openpyxl
from openpyxl import Workbook
from datetime import datetime
import os
import pytz

def download_cleaning_instructions():
    # --- 設定 ---
    db_path = r'\\192.168.1.200\共有\製造課\ロボパット\python app\cleaning_instructions.db'
    excel_path = r'C:\Users\SEIZOU-20\Desktop\洗浄指示書.xlsx'
    
    # --- 東京時間の今日の日付を取得 ---
    try:
        tokyo_tz = pytz.timezone('Asia/Tokyo')
        today = datetime.now(tokyo_tz).strftime('%Y-%m-%d')
        sheet_name = datetime.now(tokyo_tz).strftime('%m%d')
    except pytz.UnknownTimeZoneError:
        print("タイムゾーン 'Asia/Tokyo' が見つかりません。pytzライブラリが正しくインストールされているか確認してください。")
        return

    print(f"本日の日付 ({tokyo_tz}): {today}")
    print(f"シート名: {sheet_name}")

    # --- データベースからデータを取得 ---
    data_to_write = []
    try:
        conn = sqlite3.connect(db_path)
        cursor = conn.cursor()
        print("データベースに接続しました。")

        # 正しいテーブル名 'production_plan' を使用
        table_name = 'production_plan'

        query = f"SELECT cleaning_instruction FROM {table_name} WHERE acquisition_date = ? ORDER BY machine_no ASC"
        print(f"実行するクエリ: {query}")
        
        cursor.execute(query, (today,))
        rows = cursor.fetchall()
        
        if rows:
            data_to_write = [row[0] for row in rows]
            print(f"{len(data_to_write)} 件のデータを取得しました。")
        else:
            print("本日の日付に一致するデータは見つかりませんでした。")

    except sqlite3.Error as e:
        print(f"データベースエラーが発生しました: {e}")
        return
    finally:
        if 'conn' in locals() and conn:
            conn.close()
            print("データベース接続を閉じました。")

    if not data_to_write:
        print("書き込むデータがないため、処理を終了します。")
        return

    # --- Excelファイルにデータを書き込む ---
    try:
        if os.path.exists(excel_path):
            workbook = openpyxl.load_workbook(excel_path)
            print(f"既存のExcelファイルを開きました: {excel_path}")
        else:
            workbook = Workbook()
            print(f"新しいExcelファイルを作成します: {excel_path}")

        # 同じ名前のシートが存在する場合は削除
        if sheet_name in workbook.sheetnames:
            del workbook[sheet_name]
            print(f"既存のシート '{sheet_name}' を削除しました。")
            
        sheet = workbook.create_sheet(title=sheet_name)
        print(f"新しいシート '{sheet_name}' を作成しました。")

        # データをAK列の2行目から書き込む
        for index, value in enumerate(data_to_write, start=2):
            sheet.cell(row=index, column=37, value=value) # AK列は37番目

        # 不要なデフォルトシートを削除 (ファイル新規作成時)
        if "Sheet" in workbook.sheetnames and len(workbook.sheetnames) > 1:
             if workbook["Sheet"].cell(row=1, column=1).value is None:
                del workbook["Sheet"]
                print("デフォルトの空のシート 'Sheet' を削除しました。")

        workbook.save(excel_path)
        print(f"Excelファイルにデータを保存しました: {excel_path}")

    except openpyxl.utils.exceptions.InvalidFileException:
        print(f"エラー: {excel_path} は有効なExcelファイルではありません。")
    except Exception as e:
        print(f"Excel処理中にエラーが発生しました: {e}")

if __name__ == "__main__":
    download_cleaning_instructions()
