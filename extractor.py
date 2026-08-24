from pathlib import Path
import datetime
import json

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


def _clean_cell_value(v):
    if v is None:
        return None
    if isinstance(v, (datetime.datetime, datetime.date)):
        return v.strftime("%Y-%m-%d")
    return v


def _read_fixed_range_column(worksheet, range_str):
    return [_clean_cell_value(v) for v in _read_fixed_range_values(worksheet, range_str)]


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
    except Exception as exc:
        return {"key": key, "label": label, "value": None, "error": str(exc)}


def extract_chart(item, uploads_dir):
    key = item["key"]
    try:
        mode = item.get("mode", "header_match")
        if mode == "fixed_range":
            worksheet = _load_worksheet(uploads_dir, item["source_file"], item["sheet"])
            x = _read_fixed_range_column(worksheet, item["x_range"])
            result = {
                "key": key,
                "type": item["type"],
                "title": item["title"],
                "x": x,
                "error": None,
            }
            if "series" in item:
                result["series"] = [
                    {"name": s["name"], "data": _read_fixed_range_column(worksheet, s["range"])}
                    for s in item["series"]
                ]
            else:
                result["y"] = _read_fixed_range_column(worksheet, item["y_range"])
            return result
        elif mode == "header_match":
            df = _read_dataframe(uploads_dir, item["source_file"], item["sheet"])
            x_header = item["x_header"]
            y_header = item["y_header"]
            missing = [h for h in (x_header, y_header) if h not in df.columns]
            if missing:
                raise ValueError(f"找不到表头: '{missing[0]}'")
            x = [_clean_value(v) for v in df[x_header].tolist()]
            y = [_clean_value(v) for v in df[y_header].tolist()]
            return {
                "key": key,
                "type": item["type"],
                "title": item["title"],
                "x": x,
                "y": y,
                "error": None,
            }
        else:
            raise ValueError(f"未知的取数模式: {mode}")
    except Exception as exc:
        return {
            "key": key,
            "type": item.get("type"),
            "title": item.get("title"),
            "x": [],
            "y": [],
            "error": str(exc),
        }


def extract_table(item, uploads_dir):
    key = item["key"]
    try:
        mode = item["mode"]
        if mode == "fixed_range":
            worksheet = _load_worksheet(uploads_dir, item["source_file"], item["sheet"])
            columns = [c["label"] for c in item["columns"]]
            column_values = [_read_fixed_range_column(worksheet, c["range"]) for c in item["columns"]]
            rows = [list(row) for row in zip(*column_values)]
            return {
                "key": key,
                "title": item["title"],
                "columns": columns,
                "rows": rows,
                "error": None,
                "view_group": item.get("view_group"),
                "view_label": item.get("view_label"),
            }
        df = _read_dataframe(uploads_dir, item["source_file"], item["sheet"])
        if mode == "header_match":
            columns = item["columns"]
            missing = [c for c in columns if c not in df.columns]
            if missing:
                raise ValueError(f"找不到表头: '{missing[0]}'")
            rows = [[_clean_value(v) for v in row] for row in df[columns].values.tolist()]
        elif mode == "group_by_sum":
            group_by_header = item["group_by_header"]
            sum_header = item["sum_header"]
            missing = [h for h in (group_by_header, sum_header) if h not in df.columns]
            if missing:
                raise ValueError(f"找不到表头: '{missing[0]}'")
            grouped = df.groupby(group_by_header)[sum_header].sum().sort_values(ascending=False)
            columns = [group_by_header, sum_header]
            rows = [[_clean_value(k), _clean_value(v)] for k, v in grouped.items()]
        else:
            raise ValueError(f"未知的取数模式: {mode}")
        return {
            "key": key,
            "title": item["title"],
            "columns": columns,
            "rows": rows,
            "error": None,
            "view_group": item.get("view_group"),
            "view_label": item.get("view_label"),
        }
    except Exception as exc:
        raw_columns = item.get("columns", [])
        columns = [c["label"] if isinstance(c, dict) else c for c in raw_columns]
        return {
            "key": key,
            "title": item.get("title"),
            "columns": columns,
            "rows": [],
            "error": str(exc),
            "view_group": item.get("view_group"),
            "view_label": item.get("view_label"),
        }


def _resolve_computed_kpi(item, resolved_values):
    key = item["key"]
    label = item["label"]
    operation = item.get("operation")
    if operation != "subtract":
        return {"key": key, "label": label, "value": None, "error": f"未知的运算类型: {operation}"}
    from_key = item.get("from")
    minus_key = item.get("minus")
    from_value = resolved_values.get(from_key)
    minus_value = resolved_values.get(minus_key)
    if from_value is None or minus_value is None:
        missing = from_key if from_value is None else minus_key
        return {"key": key, "label": label, "value": None, "error": f"引用的指标不存在或取值失败: {missing}"}
    return {"key": key, "label": label, "value": from_value - minus_value, "error": None}


def build_dashboard_data(config, uploads_dir):
    resolved_values = {}
    raw_results = {}
    for item in config["kpis"]:
        if item.get("mode") != "computed":
            result = extract_kpi(item, uploads_dir)
            raw_results[item["key"]] = result
            resolved_values[item["key"]] = result["value"]

    kpi_results = []
    for item in config["kpis"]:
        if item.get("mode") == "computed":
            result = _resolve_computed_kpi(item, resolved_values)
        else:
            result = raw_results[item["key"]]
        kpi_results.append(result)

    return {
        "kpis": kpi_results,
        "charts": [extract_chart(item, uploads_dir) for item in config["charts"]],
        "tables": [extract_table(item, uploads_dir) for item in config["tables"]],
    }


def save_data_json(data, data_path):
    with open(data_path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def load_data_json(data_path):
    with open(data_path, "r", encoding="utf-8") as f:
        return json.load(f)
