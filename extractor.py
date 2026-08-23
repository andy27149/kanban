import yaml


def load_config(config_path):
    with open(config_path, "r", encoding="utf-8") as f:
        raw = yaml.safe_load(f) or {}
    return {
        "kpis": raw.get("kpis", []),
        "charts": raw.get("charts", []),
        "tables": raw.get("tables", []),
    }
