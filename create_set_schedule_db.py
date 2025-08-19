import pandas as pd
import sqlite3
import numpy as np
from datetime import date

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

# === データ読込 ===
df = pd.read_excel(EXCEL_PATH, sheet_name=SHEET_NAME, engine="openpyxl")
df = df.rename(columns=COLUMN_MAPPING)

# 機械NOが空または"0"の行を削除
# 1. 空白やスペースのみのセルをNaN（Not a Number）に置換
df['machine_no'] = df['machine_no'].replace(r'^\s*$', np.nan, regex=True)

# 2. '0' や 0.0 といった値をNaNに置換
df.loc[pd.to_numeric(df['machine_no'], errors='coerce') == 0, 'machine_no'] = np.nan

# 3. NaN（空または0）の行をまとめて削除
df.dropna(subset=['machine_no'], inplace=True)


# 必要なカラムだけ抽出
# 抽出前にカラムが存在するか確認
existing_columns = [col for col in COLUMN_MAPPING.values() if col in df.columns]
df = df[existing_columns].copy()

# --- データ型を整数に変換 ---
integer_columns = [
    'quantity',
    'prev_daily_output',
    'required_quantity',
    'material_id'
]
for col in integer_columns:
    if col in df.columns:
        # 文字列や小数を含む可能性のある列を数値に変換し、変換不能な値はNaNにする
        # その後、NaNを0で埋めてから整数型に変換する
        df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0).astype(int)

# --- DB専用カラムを追加（初期値は0） ---
df["manufacturing_check"] = 0
df["cleaning_check"] = 0
df["cleaning_instruction"] = 0
df["acquisition_date"] = date.today().strftime("%Y-%m-%d")

# カラム順を整理（指定順でDBに格納）
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
# dfに存在する列のみでカラム順を再設定
final_columns = [col for col in column_order if col in df.columns]
df = df[final_columns]

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

# 新規データを挿入（追記）
df.to_sql(TABLE_NAME, conn, if_exists="append", index=False)

conn.commit()
conn.close()

print("✅ Data inserted into SQLite database:", DB_PATH)