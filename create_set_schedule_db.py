import pandas as pd
import sqlite3
import numpy as np
from datetime import datetime, timedelta
from zoneinfo import ZoneInfo

# === 設定 ===
EXCEL_PATH = r"\\192.168.1.200\共有\生産管理課\セット予定表.xlsx"
SHEET_NAME = "生産中"
DB_PATH = "set_schedule.db"
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
    "前回   日産": "prev_daily_output",
    "必要数": "required_quantity",
    "材料　　　　識別": "material_id"
}

# --- 曜日チェックと実行回数設定 ---
tokyo_tz = ZoneInfo("Asia/Tokyo")
today = datetime.now(tokyo_tz).date()
weekday_map = {0: "月曜日", 1: "火曜日", 2: "水曜日", 3: "木曜日", 4: "金曜日", 5: "土曜日", 6: "日曜日"}
day_of_week_str = weekday_map[today.weekday()]

# 本番用に金曜日(weekday()==4)に設定。
# 0:月, 1:火, 2:水, 3:木, 4:金, 5:土, 6:日
if today.weekday() == 4:
    num_runs = 3
    print(f"本日は{day_of_week_str}です。3日分のデータを作成します。")
else:
    num_runs = 1
    print(f"本日は{day_of_week_str}です。1日分のデータを作成します。")


# === データ読込と前処理 ===
df_base = pd.read_excel(EXCEL_PATH, sheet_name=SHEET_NAME, engine="openpyxl")
df_base = df_base.rename(columns=COLUMN_MAPPING)

# 機械NOが空または"0"の行を削除
df_base['machine_no'] = df_base['machine_no'].replace(r'^\s*$', np.nan, regex=True)
df_base.loc[pd.to_numeric(df_base['machine_no'], errors='coerce') == 0, 'machine_no'] = np.nan
df_base.dropna(subset=['machine_no'], inplace=True)


# 必要なカラムだけ抽出
# 抽出前にカラムが存在するか確認
existing_columns = [col for col in COLUMN_MAPPING.values() if col in df_base.columns]
df_base = df_base[existing_columns].copy()

# --- データ型を整数に変換 ---
integer_columns = [
    'quantity',
    'prev_daily_output',
    'required_quantity',
    'material_id'
]
for col in integer_columns:
    if col in df_base.columns:
        df_base[col] = pd.to_numeric(df_base[col], errors='coerce').fillna(0).astype(int)

# --- DB専用カラムを追加（取得日以外） ---
df_base["manufacturing_check"] = 0
df_base["cleaning_check"] = 0
df_base["cleaning_instruction"] = 0


# === データベース登録 ===
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
    acquisition_date TEXT
)
""")

# カラム順を定義
column_order = [
    "set_date",
    "manufacturing_check",   # DB専用
    "cleaning_check",        # DB専用
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
    "cleaning_instruction",   # DB専用
    "acquisition_date"      # DB専用
]

# --- ループ処理でDBに登録 ---
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
conn.close()

print(f"\n✅ データベースへの登録が完了しました。({DB_PATH})")