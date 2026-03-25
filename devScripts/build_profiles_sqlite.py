#!/usr/bin/env python3
"""
Build a SQLite database from the raw profile sheets in the 3 latest Excel workbooks.

Tables created:
- raw_lipid_data
- raw_endocrine_data
- cbc_hematology_raw
- iron_profile_raw
- liver_profile_raw
- inflammation_immun_raw
- vitamin_nutritiona_raw
- liver_monitoring_raw
- cbc_treatment_moni_raw

Usage:
    python build_profiles_sqlite.py

Optional custom paths:
    python build_profiles_sqlite.py \
        --lipid "Lipid_Derived_Markers_Trends_March2026.xlsx" \
        --endocrine "Endocrinology_Derived_Markers_Trends_March2026.xlsx" \
        --additional "Additional_Profiles_Merged_Dashboard_March2026.xlsx" \
        --db "profiles_raw_data.sqlite"
"""

from __future__ import annotations

import argparse
import re
import sqlite3
from pathlib import Path
from typing import Dict, List

import pandas as pd


DEFAULT_LIPID = "/mnt/data/Lipid_Derived_Markers_Trends_March2026.xlsx"
DEFAULT_ENDOCRINE = "/mnt/data/Endocrinology_Derived_Markers_Trends_March2026.xlsx"
DEFAULT_ADDITIONAL = "/mnt/data/Additional_Profiles_Merged_Dashboard_March2026.xlsx"
DEFAULT_DB = "/mnt/data/profiles_raw_data.sqlite"


SOURCE_MAP = {
    "raw_lipid_data": {
        "file_arg": "lipid",
        "sheet": "Raw_Lipid_Data",
        "header_row_excel": 3,   # Excel row number, 1-based
    },
    "raw_endocrine_data": {
        "file_arg": "endocrine",
        "sheet": "Raw_Endocrine_Data",
        "header_row_excel": 3,
    },
    "cbc_hematology_raw": {
        "file_arg": "additional",
        "sheet": "CBC_Hematology_Raw",
        "header_row_excel": 1,
    },
    "iron_profile_raw": {
        "file_arg": "additional",
        "sheet": "Iron_Profile_Raw",
        "header_row_excel": 1,
    },
    "liver_profile_raw": {
        "file_arg": "additional",
        "sheet": "Liver_Profile_Raw",
        "header_row_excel": 1,
    },
    "inflammation_immun_raw": {
        "file_arg": "additional",
        "sheet": "Inflammation_Immun_Raw",
        "header_row_excel": 1,
    },
    "vitamin_nutritiona_raw": {
        "file_arg": "additional",
        "sheet": "Vitamin_Nutritiona_Raw",
        "header_row_excel": 1,
    },
    "liver_monitoring_raw": {
        "file_arg": "additional",
        "sheet": "Liver_Monitoring_Raw",
        "header_row_excel": 1,
    },
    "cbc_treatment_moni_raw": {
        "file_arg": "additional",
        "sheet": "CBC_Treatment_Moni_Raw",
        "header_row_excel": 1,
    },
}


def snake_case(value: str) -> str:
    """Convert Excel headers to SQLite-safe snake_case column names."""
    value = str(value).strip()
    value = value.replace("%", " pct ")
    value = value.replace("/", " ")
    value = value.replace("(", " ")
    value = value.replace(")", " ")
    value = value.replace("[", " ")
    value = value.replace("]", " ")
    value = value.replace("-", " ")
    value = value.replace(".", " ")
    value = value.replace("μ", "u")
    value = re.sub(r"[^0-9a-zA-Z_ ]+", "", value)
    value = re.sub(r"\s+", "_", value).strip("_").lower()
    if not value:
        value = "col"
    if value[0].isdigit():
        value = f"c_{value}"
    return value


def make_unique(names: List[str]) -> List[str]:
    """Ensure duplicate column names become unique."""
    counts: Dict[str, int] = {}
    out: List[str] = []

    for name in names:
        if name not in counts:
            counts[name] = 0
            out.append(name)
        else:
            counts[name] += 1
            out.append(f"{name}_{counts[name]}")
    return out


def read_excel_table(path: str, sheet_name: str, header_row_excel: int) -> pd.DataFrame:
    """
    Read one raw sheet from Excel.

    header_row_excel is the 1-based row number containing the header in Excel.
    pandas header is 0-based, so we subtract 1.
    """
    df = pd.read_excel(
        path,
        sheet_name=sheet_name,
        header=header_row_excel - 1,
        engine="openpyxl",
    )

    # Drop fully empty rows and columns
    df = df.dropna(axis=0, how="all").dropna(axis=1, how="all")

    # Normalize column names
    cleaned_cols = make_unique([snake_case(c) for c in df.columns])
    df.columns = cleaned_cols

    # Normalize datetimes to ISO strings so SQLite stays predictable
    for col in df.columns:
        if pd.api.types.is_datetime64_any_dtype(df[col]):
            df[col] = pd.to_datetime(df[col], errors="coerce").dt.strftime("%Y-%m-%d")

    # Add source metadata
    df["source_workbook"] = Path(path).name
    df["source_sheet"] = sheet_name

    return df


def create_database(db_path: str, input_files: Dict[str, str]) -> None:
    conn = sqlite3.connect(db_path)

    try:
        for table_name, spec in SOURCE_MAP.items():
            source_path = input_files[spec["file_arg"]]
            df = read_excel_table(
                path=source_path,
                sheet_name=spec["sheet"],
                header_row_excel=spec["header_row_excel"],
            )

            # Replace NaN with None so SQLite stores NULL
            df = df.where(pd.notnull(df), None)

            # Write table
            df.to_sql(table_name, conn, if_exists="replace", index=False)

            # Create a date index if exam_date exists
            cols = set(df.columns)
            if "exam_date" in cols:
                conn.execute(
                    f'CREATE INDEX IF NOT EXISTS idx_{table_name}_exam_date '
                    f'ON "{table_name}" ("exam_date")'
                )

        conn.commit()
    finally:
        conn.close()


def main() -> None:
    parser = argparse.ArgumentParser(description="Build SQLite DB from Excel raw profile sheets.")
    parser.add_argument("--lipid", default=DEFAULT_LIPID, help="Path to lipid workbook")
    parser.add_argument("--endocrine", default=DEFAULT_ENDOCRINE, help="Path to endocrine workbook")
    parser.add_argument("--additional", default=DEFAULT_ADDITIONAL, help="Path to merged/additional workbook")
    parser.add_argument("--db", default=DEFAULT_DB, help="Output SQLite database path")
    args = parser.parse_args()

    input_files = {
        "lipid": args.lipid,
        "endocrine": args.endocrine,
        "additional": args.additional,
    }

    for label, path in input_files.items():
        if not Path(path).exists():
            raise FileNotFoundError(f"{label} workbook not found: {path}")

    create_database(args.db, input_files)

    print(f"SQLite database created: {args.db}")
    print("Tables:")
    for table_name in SOURCE_MAP:
        print(f" - {table_name}")


if __name__ == "__main__":
    main()
