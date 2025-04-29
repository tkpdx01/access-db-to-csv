# access-db-to-csv / Access 数据库 CSV 导出工具

> **Chinese | English**

---

## 项目简介 Project Overview

**access-db-to-csv** 是一个用 Python 编写的轻量级工具，能够从 Microsoft Access（*.mdb / *.accdb*）数据库提取表数据并导出为 CSV，同时提供后续处理 CSV 的辅助脚本。

**access-db-to-csv** is a lightweight Python utility that extracts table data from Microsoft Access (*.mdb / *.accdb*) files, converts them to CSV, and ships helper scripts for post‑processing.

---

## 功能特性 Features

| 中文 | English |
|------|---------|
| ODBC 连接 Access 数据库 | ODBC connectivity to Access files |
| 按 SQL 查询提取数据 | Extract data via custom SQL |
| **整表导出**：`<库名>_<表名>.csv` | **Whole‑table export**: `<db>_<table>.csv` |
| **按列拆分**：`<库名>_<列名>.csv` | **Per‑column split**: `<db>_<column>.csv` |
| 从 CSV 生成 `TYPE_CODE → TYPE_NAME` 映射 | Build `TYPE_CODE → TYPE_NAME` mappings from CSV |

---

## 环境依赖 Requirements

- Python 3.x  
- **Microsoft Access Database Engine (ACE)** 或任意 Access ODBC 驱动  
- Python 3.x  
- **Microsoft Access Database Engine (ACE)** or another Access‑compatible ODBC driver

---

## 安装 Installation

```bash
# Clone the repository
git clone https://github.com/tkpdx01/access-db-to-csv.git
cd access-db-to-csv

# 创建并激活虚拟环境（推荐）
python -m venv .venv
.venv\Scripts\activate  # Windows
# source .venv/bin/activate  # Linux/Mac

# 安装依赖 Install deps with uv
uv pip install pyodbc pandas
```

---

## 使用 Usage

### 1️⃣ 脚本一：`access_db_to_csv.py`  —— 单表→单 CSV  +  映射提取

适合只关心 **单张表** 且需要生成键值映射的场景。

```bash
# 编辑脚本顶部 mdb_path / conn_str / 表名
python access_db_to_csv.py
# 生成：table.csv  &  mapping.txt
```

### 2️⃣ 脚本二：`mdb2multi_csv.py`  —— 整库批量导出

```bash
# 在脚本顶部设置
# MDB_PATH   = r"xx.mdb"
# TABLE_NAME = None           # None = 所有用户表
# PWD        = "your_pwd"

python mdb2multi_csv.py
```

输出示例  (Access 库名 `sample.mdb`, 表 `codetype`)：

```
sample_codetype.csv      # 整表
sample_TYPE_CODE.csv     # 单列
sample_TYPE_NAME.csv     # …
```

编码默认 `UTF‑8‑SIG`，Windows/Excel 直接可读。

---

## 自定义查询 Customization

- 在 `access_db_to_csv.py` 里把 `pd.read_sql("SELECT * FROM codetype", conn)` 改为自定义 SQL。  
- `mdb2multi_csv.py` 默认 `SELECT *`；如需筛选列可在 `export_table()` 中改写。

---

## 许可证 License

本项目采用 **MIT License** ；详情见 LICENSE 文件。

Released under the **MIT License** – see LICENSE for details.

