#!/usr/bin/env python3
from __future__ import annotations

import json
from pathlib import Path

import pandas as pd


def clean_value(value):
    if pd.isna(value):
        return ""
    if isinstance(value, str):
        return value.strip()
    return value


def dataframe_to_rows(df: pd.DataFrame) -> list[dict]:
    records: list[dict] = []
    for row in df.to_dict(orient="records"):
        clean_row = {str(k).strip(): clean_value(v) for k, v in row.items()}
        if any(str(v).strip() for v in clean_row.values()):
            records.append(clean_row)
    return records


def main():
    project_dir = Path(__file__).resolve().parent.parent
    excel_path = project_dir / "课程专业评价预设题库_薄弱点匹配提升建议版.xlsx"
    output_path = project_dir / "question-bank.json"

    if not excel_path.exists():
        raise FileNotFoundError(f"找不到 Excel 文件: {excel_path}")

    workbook = pd.ExcelFile(excel_path)
    output: dict[str, list[dict]] = {}

    for sheet_name in workbook.sheet_names:
        df = pd.read_excel(workbook, sheet_name=sheet_name, dtype=object)
        output[sheet_name] = dataframe_to_rows(df)

    output_path.write_text(
        json.dumps(output, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )

    print(f"转换完成: {output_path}")
    print(f"共读取 {len(workbook.sheet_names)} 个 Sheet")


if __name__ == "__main__":
    main()
