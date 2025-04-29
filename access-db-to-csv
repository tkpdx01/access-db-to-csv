#!/usr/bin/env python
"""
导出 Access 数据表：
1. 全表保存为 <db_stem>_<table>.csv
2. 每个字段再拆分为 <db_stem>_<列名>.csv
"""

import pathlib as pl
import pyodbc
import pandas as pd
from datetime import datetime

# ====== ✏️  根据需要修改这几项 ======
MDB_PATH   = pl.Path(r"xx.mdb")        # Access 文件
TABLE_NAME = "None"                # 要导出的表名；填 None 则导出所有用户表
PWD        = "123456"            # 无密码留空即可
# ===================================

def connect_access(mdb: pl.Path, pwd: str = "") -> pyodbc.Connection:
    conn_str = (
        r"DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};"
        rf"DBQ={mdb};"
        + (rf"PWD={pwd};" if pwd else "")
    )
    return pyodbc.connect(conn_str, autocommit=True)

def all_user_tables(cur: pyodbc.Cursor):
    for (name,) in cur.tables(tableType="TABLE"):
        # 过滤系统表 / 临时表
        if not name.startswith("MSys"):
            yield name

def export_table(conn: pyodbc.Connection, table: str, prefix: str):
    print(f"▶ 正在导出 {table} ...")
    df = pd.read_sql(f"SELECT * FROM [{table}]", conn)

    # 1️⃣ 整表导出
    full_csv = f"{prefix}_{table}.csv"
    df.to_csv(full_csv, index=False, encoding="utf-8-sig")
    print(f"   ✔ 全表 => {full_csv}")

    # 2️⃣ 各列单独导出
    for col in df.columns:
        single_csv = f"{prefix}_{col}.csv"
        df[[col]].to_csv(single_csv, index=False, encoding="utf-8-sig")
        print(f"   ✔ 列  {col} => {single_csv}")

def main():
    prefix = MDB_PATH.stem            # 文件名前缀
    with connect_access(MDB_PATH, PWD) as conn:
        cur = conn.cursor()

        tables = [TABLE_NAME] if TABLE_NAME else list(all_user_tables(cur))
        if not tables:
            print("❌ 没找到可导出的表")
            return

        for t in tables:
            export_table(conn, t, prefix)

    print("\n🎉 全部完成", datetime.now().strftime("%F %T"))

if __name__ == "__main__":
    main()
