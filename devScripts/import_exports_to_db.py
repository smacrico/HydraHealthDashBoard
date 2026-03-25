#!/usr/bin/env python3
"""
Import exported worksheet CSV files into a database.

Default target: SQLite
Also supports any SQLAlchemy database URL if SQLAlchemy + driver are installed.

Examples
--------
SQLite:
    python import_exports_to_db.py \
      --manifest /path/to/manifest.json \
      --db-url sqlite:////path/to/my_database.sqlite

PostgreSQL:
    python import_exports_to_db.py \
      --manifest /path/to/manifest.json \
      --db-url postgresql+psycopg2://user:password@host:5432/dbname \
      --schema health_raw

Behavior
--------
- Creates one table per worksheet
- Table name format:
    <workbook_key>__<sheet_name_sanitized>
- Imports the worksheet exactly as exported
- Uses generic text columns: col_1, col_2, ...
- Adds metadata columns:
    source_workbook, source_sheet
"""

from __future__ import annotations

import argparse
import json
from pathlib import Path
import re

import pandas as pd

def sanitize(name: str) -> str:
    name = re.sub(r"[^A-Za-z0-9]+", "_", name).strip("_").lower()
    return name

def load_manifest(path: str):
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)

def read_csv_preserve(csv_path: str) -> pd.DataFrame:
    df = pd.read_csv(csv_path, header=None, dtype=str, keep_default_na=False, na_values=[])
    df.columns = [f"col_{i+1}" for i in range(df.shape[1])]
    return df

def import_sqlite(manifest, db_path: str):
    import sqlite3

    conn = sqlite3.connect(db_path)
    try:
        for item in manifest:
            df = read_csv_preserve(item["csv_path"])
            df["source_workbook"] = item["workbook"]
            df["source_sheet"] = item["sheet_name"]
            table_name = f'{item["workbook_key"]}__{sanitize(item["sheet_name"])}'
            df.to_sql(table_name, conn, if_exists="replace", index=False)
            print(f"Imported {table_name} ({len(df)} rows)")
    finally:
        conn.close()

def import_sqlalchemy(manifest, db_url: str, schema: str | None = None):
    from sqlalchemy import create_engine
    engine = create_engine(db_url)

    with engine.begin() as conn:
        for item in manifest:
            df = read_csv_preserve(item["csv_path"])
            df["source_workbook"] = item["workbook"]
            df["source_sheet"] = item["sheet_name"]
            table_name = f'{item["workbook_key"]}__{sanitize(item["sheet_name"])}'
            df.to_sql(table_name, conn, schema=schema, if_exists="replace", index=False, method="multi")
            print(f"Imported {table_name} ({len(df)} rows)")

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--manifest", required=True, help="Path to manifest.json")
    parser.add_argument("--db-url", required=True, help="Database URL. Example sqlite:////tmp/my.db")
    parser.add_argument("--schema", default=None, help="Optional schema for SQLAlchemy-supported databases")
    args = parser.parse_args()

    manifest = load_manifest(args.manifest)

    if args.db_url.startswith("sqlite:///"):
        db_path = args.db_url.replace("sqlite:///", "", 1)
        import_sqlite(manifest, db_path)
    else:
        import_sqlalchemy(manifest, args.db_url, args.schema)

if __name__ == "__main__":
    main()
