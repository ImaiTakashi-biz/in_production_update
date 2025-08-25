import pandas as pd
import sqlite3
import numpy as np
from datetime import datetime, timedelta
from zoneinfo import ZoneInfo
import os
import sys
import smtplib
from email.mime.text import MIMEText
import traceback

# --- メール通知用の設定 ---
# これらの設定値は、ご自身の環境に合わせて変更してください。
# パスワードを直接コードに書くことは推奨しません。
EMAIL_SENDER = "imai@araiseimitsu.onmicrosoft.com"
EMAIL_PASSWORD = "Arai267786"
EMAIL_RECEIVERS = [
    "takada@araiseimitsu.onmicrosoft.com",
    "imai@araiseimitsu.onmicrosoft.com",
    "n.kizaki@araiseimitsu.onmicrosoft.com"
]
SMTP_SERVER = "smtp.office365.com"
SMTP_PORT = 587

def send_error_email(error_info, program_name, program_path, subject_prefix="【エラー通知】"):
    """
    エラー発生時に指定されたアカウントへメールを送信する関数
    """
    try:
        subject = f"{subject_prefix}Pythonスクリプト実行中にエラーが発生しました"
        body = f"""
お疲れ様です。

Pythonスクリプトの実行中にエラーが発生しました。
下記に詳細を記載します。

---
日時: {datetime.now().strftime('%Y/%m/%d %H:%M:%S')}
プログラム名: {program_name}
ファイルパス: {program_path}
エラー詳細:
{error_info}
---

お手数ですが、ご確認をお願いします。
"""
        msg = MIMEText(body, "plain", "utf-8")
        msg["Subject"] = subject
        msg["From"] = EMAIL_SENDER
        msg["To"] = ", ".join(EMAIL_RECEIVERS)

        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
            server.starttls()
            server.login(EMAIL_SENDER, EMAIL_PASSWORD)
            server.sendmail(EMAIL_SENDER, EMAIL_RECEIVERS, msg.as_string())
        print("エラー通知メールを送信しました。")

    except Exception as e:
        print(f"メール送信中にエラーが発生しました: {e}", file=sys.stderr)

def main():
    # プログラム名とファイルパスを取得
    program_name = os.path.basename(__file__)
    program_path = os.path.abspath(__file__)
    
    # === 設定 ===
    EXCEL_PATH = r"\\192.168.1.200\共有\生産管理課\セット予定表.xlsx"
    SHEET_NAME = "生産中"
    DB_PATH = r"\\192.168.1.200\共有\製造課\ロボパット\python app\Cleaning_instructions.db"
    TABLE_NAME = "production_plan"

    # Excel列名 → DBカラム名 マッピング
    COLUMN_MAPPING = {
        "セット予定日": "set_date",
        "機械NO": "machine_no",
        "機種": "model",
        "客先名": "customer_name",
        "品番": "part_number",
        "製品名": "product_name",
        "数量": "quantity",
        "材質＆材料径": "material_info",
        "次工程": "next_process",
        "取扱注意事項": "handling_notes",
        "加工終了日": "completion_date",
        "前回   日産": "prev_daily_output",
        "必要数": "required_quantity",
        "材料　　　　識別": "material_id"
    }

    try:
        # --- 曜日チェックと実行回数設定 ---
        print("曜日をチェックし、実行回数を設定しています...")
        tokyo_tz = ZoneInfo("Asia/Tokyo")
        today = datetime.now(tokyo_tz).date()
        weekday_map = {0: "月曜日", 1: "火曜日", 2: "水曜日", 3: "木曜日", 4: "金曜日", 5: "土曜日", 6: "日曜日"}
        day_of_week_str = weekday_map[today.weekday()]

        if today.weekday() == 4:
            num_runs = 3
            print(f"本日は{day_of_week_str}です。3日分のデータを作成します。")
        else:
            num_runs = 1
            print(f"本日は{day_of_week_str}です。1日分のデータを作成します。")
    except Exception as e:
        error_detail = traceback.format_exc()
        send_error_email(error_detail, program_name, program_path)
        print(f"日付/タイムゾーン設定中にエラーが発生しました。\nプログラム名: {program_name}\nファイルパス: {program_path}\nエラー詳細: {e}", file=sys.stderr)
        sys.exit(1)

    try:
        # === データ読込と前処理 ===
        print("Excelファイルを読み込んでいます...")
        df_base = pd.read_excel(EXCEL_PATH, sheet_name=SHEET_NAME, engine="openpyxl")
        df_base = df_base.rename(columns=COLUMN_MAPPING)

        # 機械NOが空または"0"の行を削除
        df_base['machine_no'] = df_base['machine_no'].replace(r'^\s*$', np.nan, regex=True)
        df_base.loc[pd.to_numeric(df_base['machine_no'], errors='coerce') == 0, 'machine_no'] = np.nan
        df_base.dropna(subset=['machine_no'], inplace=True)
        print("データの前処理が完了しました。")

    except Exception as e:
        error_detail = traceback.format_exc()
        send_error_email(error_detail, program_name, program_path)
        print(f"Excelデータの読み込みまたは前処理中にエラーが発生しました。\nプログラム名: {program_name}\nファイルパス: {program_path}\nエラー詳細: {e}", file=sys.stderr)
        sys.exit(1)

    # 必要なカラムだけ抽出
    try:
        existing_columns = [col for col in COLUMN_MAPPING.values() if col in df_base.columns]
        df_base = df_base[existing_columns].copy()
    except Exception as e:
        error_detail = traceback.format_exc()
        send_error_email(error_detail, program_name, program_path)
        print(f"必要な列の抽出中にエラーが発生しました。\nプログラム名: {program_name}\nファイルパス: {program_path}\nエラー詳細: {e}", file=sys.stderr)
        sys.exit(1)

    # --- データ型を整数に変換 ---
    try:
        integer_columns = [
            'quantity',
            'prev_daily_output',
            'required_quantity',
            'material_id'
        ]
        for col in integer_columns:
            if col in df_base.columns:
                df_base[col] = pd.to_numeric(df_base[col], errors='coerce').fillna(0).astype(int)
        print("データ型の変換が完了しました。")

    except Exception as e:
        error_detail = traceback.format_exc()
        send_error_email(error_detail, program_name, program_path)
        print(f"データ型の変換中にエラーが発生しました。\nプログラム名: {program_name}\nファイルパス: {program_path}\nエラー詳細: {e}", file=sys.stderr)
        sys.exit(1)

    # --- DB専用カラムを追加（取得日以外） ---
    df_base["manufacturing_check"] = 0
    df_base["cleaning_check"] = 0
    df_base["cleaning_instruction"] = 0

    # === データベース登録 ===
    conn = None
    try:
        print("データベースに接続し、テーブルを作成しています...")
        conn = sqlite3.connect(DB_PATH)
        cur = conn.cursor()

        # テーブル作成
        cur.execute(f"""
        CREATE TABLE IF NOT EXISTS {TABLE_NAME} (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            set_date TEXT,
            manufacturing_check INTEGER DEFAULT 0,
            cleaning_check INTEGER DEFAULT 0,
            machine_no TEXT,
            model TEXT,
            customer_name TEXT,
            part_number TEXT,
            product_name TEXT,
            quantity INTEGER,
            material_info TEXT,
            next_process TEXT,
            handling_notes TEXT,
            completion_date TEXT,
            prev_daily_output INTEGER,
            required_quantity INTEGER,
            material_id INTEGER,
            cleaning_instruction INTEGER DEFAULT 0,
            acquisition_date TEXT,
            previous_day_set INTEGER DEFAULT 0,
            notes TEXT
        )
        """)
        print("テーブル作成が完了しました。")

    except sqlite3.Error as e:
        error_detail = traceback.format_exc()
        send_error_email(error_detail, program_name, program_path)
        print(f"データベース接続またはテーブル作成中にエラーが発生しました。\nプログラム名: {program_name}\nファイルパス: {program_path}\nエラー詳細: {e}", file=sys.stderr)
        if conn:
            conn.close()
        sys.exit(1)

    # カラム順を定義
    column_order = [
        "set_date",
        "manufacturing_check", 
        "cleaning_check",
        "machine_no",
        "model",
        "customer_name",
        "part_number",
        "product_name",
        "quantity",
        "material_info",
        "next_process",
        "handling_notes",
        "completion_date",
        "prev_daily_output",
        "required_quantity",
        "material_id",
        "cleaning_instruction",
        "acquisition_date",
        "previous_day_set",
        "notes TEXT"
    ]

    # --- ループ処理でDBに登録 ---
    try:
        for i in range(1, num_runs + 1):
            df = df_base.copy()
            
            # 取得日を計算して追加
            acquisition_date = today + timedelta(days=i)
            df["acquisition_date"] = acquisition_date.strftime("%Y-%m-%d")
            
            # dfに存在する列のみでカラム順を再設定
            final_columns = [col for col in column_order if col in df.columns]
            df = df[final_columns]

            # 新規データを挿入（追記）
            df.to_sql(TABLE_NAME, conn, if_exists="append", index=False)
            print(f"-> {acquisition_date.strftime('%Y-%m-%d')} 分のデータを挿入しました。")
        
        conn.commit()
        print("\n✅ データベースへの登録が完了しました。")

    except Exception as e:
        error_detail = traceback.format_exc()
        send_error_email(error_detail, program_name, program_path)
        print(f"データベースへのデータ挿入中にエラーが発生しました。\nプログラム名: {program_name}\nファイルパス: {program_path}\nエラー詳細: {e}", file=sys.stderr)
        conn.rollback()
        sys.exit(1)
    
    finally:
        if conn:
            conn.close()
            print("データベース接続を閉じました。")

if __name__ == "__main__":
    main()
