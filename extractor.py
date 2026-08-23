from pathlib import Path

import openpyxl
import pandas as pd
import yaml


def load_config(config_path):
    with open(config_path, "r", encoding="utf-8") as f:
        raw = yaml.safe_load(f) or {}
    return {
        "kpis": raw.get("kpis", []),
        "charts": raw.get("charts", []),
        "tables": raw.get("tables", []),
    }


def _resolve_file(uploads_dir, source_file):
    file_path = Path(uploads_dir) / source_file
    if not file_path.exists():
        raise ValueError(f"文件不存在: {source_file}")
    return file_path


def _load_worksheet(uploads_dir, source_file, sheet_name):
    file_path = _resolve_file(uploads_dir, source_file)
    wb = openpyxl.load_workbook(file_path, data_only=True)
    if sheet_name not in wb.sheetnames:
        raise ValueError(f"找不到工作表: {sheet_name}")
    return wb[sheet_name]


def _read_fixed_range_values(worksheet, range_str):
    result = worksheet[range_str]
    if hasattr(result, "value"):
        return [result.value]
    values = []
    for row in result:
        for cell in row:
            values.append(cell.value)
    return values


def _read_dataframe(uploads_dir, source_file, sheet_name):
    file_path = _resolve_file(uploads_dir, source_file)
    xls = pd.ExcelFile(file_path)
    if sheet_name not in xls.sheet_names:
        raise ValueError(f"找不到工作表: {sheet_name}")
    return xls.parse(sheet_name)


def _clean_value(v):
    if pd.isna(v):
        return None
    return v


def extract_kpi(item, uploads_dir):
    key = item["key"]
    label = item["label"]
    try:
        if item["mode"] == "fixed_range":
            worksheet = _load_worksheet(uploads_dir, item["source_file"], item["sheet"])
            values = _read_fixed_range_values(worksheet, item["range"])
        elif item["mode"] == "header_match":
            df = _read_dataframe(uploads_dir, item["source_file"], item["sheet"])
            header = item["header"]
            if header not in df.columns:
                raise ValueError(f"找不到表头: '{header}'")
            values = [_clean_value(v) for v in df[header].tolist()]
        else:
            raise ValueError(f"未知的取数模式: {item['mode']}")
        numeric_values = [v for v in values if isinstance(v, (int, float))]
        value = sum(numeric_values) if numeric_values else None
        return {"key": key, "label": label, "value": value, "error": None}
    except ValueError as exc:
        return {"key": key, "label": label, "value": None, "error": str(exc)}
