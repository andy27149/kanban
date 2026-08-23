# 老板看板（Boss Dashboard）Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 构建一个 Flask Web 应用，让数据维护人通过密码保护的 `/upload` 页面上传 Excel 文件，按 `config.yaml` 规则解析为标准化的 `data.json`，并通过无密码、响应式的 `/dashboard` 页面用 ECharts 展示 KPI、图表与明细表格。

**Architecture:** 单进程 Flask 应用同时承担上传接收与看板展示；`extractor.py` 按 `config.yaml` 中每一项的取数模式（`fixed_range`/`header_match`/`computed`/`group_by_sum`）从 `uploads/` 目录下的 Excel 文件中取数或跨指标运算，产出并覆盖写入 `data.json`；共享同一 `view_group` 的多个表格条目在 `data.json` 中仍是独立数组元素，由前端归并为一张卡片 + tab 切换；前端为静态 HTML/CSS/JS + ECharts，通过 `/api/data` 拉取数据渲染。

**Tech Stack:** Python 3.9+ / Flask / pandas / openpyxl / PyYAML / pytest（后端）；原生 HTML/CSS/JS + ECharts（前端，CDN 引入）。

**Spec:** `docs/superpowers/specs/2026-08-20-boss-dashboard-design.md`

## Global Constraints

- Python 版本以本机现有解释器为准：3.9.6（`/usr/bin/python3`），所有代码需兼容 Python 3.9 语法。
- 依赖统一通过项目本地虚拟环境安装（`venv/`），不污染系统 Python；依赖清单固定写入 `requirements.txt`。
- `/dashboard` 与 `/api/data` 不做任何身份校验，任何人拿到链接即可访问（spec 明确要求）。
- `/upload`（GET 与 POST）与 `/login` 必须要求登录态（基于 Flask session 的共享密码校验）。
- 上传文件保存**不得**使用 `werkzeug.utils.secure_filename()`——它会转写/剥离非 ASCII 字符，而 `config.yaml` 里的 `source_file` 引用的是原始中文文件名（如 `经营数据.xlsx`），一旦被转写就对不上。改用手写校验：`os.path.basename` 防目录穿越 + 扩展名白名单（`.xlsx`/`.xls`）。
- 每次解析必须读取**整个** `uploads/` 目录，而不仅是本次上传的文件，保证 `data.json` 始终是完整快照（spec 明确要求）。
- 单个取数项失败（找不到文件/sheet/表头）不得让整个解析流程崩溃；失败项需带清晰错误信息，其余项正常产出（spec 明确要求）。
- 读取 Excel 单元格时必须使用 `data_only=True` 打开 workbook，否则公式单元格取到的是公式文本而不是计算结果值。
- `data.json` 写入时使用 `ensure_ascii=False, indent=2`，保证中文标签可读。
- 不引入数据库；`data.json` 是唯一的派生数据存储，每次上传后整体覆盖。

## File Structure

```
kanban/
├── app.py                      # Flask 应用：路由、鉴权、上传处理
├── extractor.py                # 配置驱动的 Excel 解析
├── config.yaml                 # 指标/图表/表格的数据来源映射配置
├── requirements.txt
├── .gitignore
├── data.json                   # 运行时生成，不提交仓库
├── uploads/
│   └── .gitkeep
├── templates/
│   ├── login.html
│   ├── upload.html
│   └── dashboard.html
├── static/
│   ├── css/
│   │   └── dashboard.css
│   └── js/
│       └── dashboard.js
└── tests/
    ├── conftest.py             # 共享 fixture：临时 uploads 目录 + 生成测试用 xlsx 的辅助函数
    ├── test_extractor.py
    └── test_app.py
```

**extractor.py 对外接口（供后续任务与前端契约参考）：**

```python
def load_config(config_path) -> dict:
    """返回 {'kpis': [...], 'charts': [...], 'tables': [...]}，缺失的 key 默认为 []"""

def extract_kpi(item: dict, uploads_dir) -> dict:
    """item 是 config['kpis'] 中的一条。返回 {'key', 'label', 'value', 'error'}"""

def extract_chart(item: dict, uploads_dir) -> dict:
    """返回 {'key', 'type', 'title', 'x': [...], 'y': [...], 'error'}"""

def extract_table(item: dict, uploads_dir) -> dict:
    """返回 {'key', 'title', 'columns': [...], 'rows': [[...]], 'error'}"""

def build_dashboard_data(config: dict, uploads_dir) -> dict:
    """返回 {'kpis': [...], 'charts': [...], 'tables': [...]}，单项失败不抛异常"""

def save_data_json(data: dict, data_path) -> None: ...
def load_data_json(data_path) -> dict: ...
```

**app.py 对外接口：**

```python
def create_app(test_config: dict = None) -> Flask:
    """工厂函数。test_config 可覆盖 UPLOAD_PASSWORD / SECRET_KEY / UPLOAD_DIR / CONFIG_PATH / DATA_PATH"""
```

---

### Task 1: 项目脚手架 + 配置加载

**Files:**
- Create: `requirements.txt`
- Create: `.gitignore`
- Create: `config.yaml`
- Create: `uploads/.gitkeep`
- Create: `extractor.py`
- Test: `tests/test_extractor.py`

**Interfaces:**
- Produces: `extractor.load_config(config_path) -> dict`（后续所有任务都要用它读取配置）

- [ ] **Step 1: 创建项目脚手架文件**

`requirements.txt`：
```
Flask>=2.3
pandas>=2.0
openpyxl>=3.1
PyYAML>=6.0
pytest>=7.4
```

`.gitignore`：
```
venv/
__pycache__/
*.pyc
.pytest_cache/
uploads/*
!uploads/.gitkeep
data.json
```

`config.yaml`：
```yaml
kpis:
  - key: total_revenue
    label: "总营收"
    source_file: "经营数据.xlsx"
    sheet: "汇总"
    mode: fixed_range
    range: "B2"
  - key: total_profit
    label: "总利润"
    source_file: "经营数据.xlsx"
    sheet: "汇总"
    mode: header_match
    header: "利润"

charts:
  - key: monthly_sales_trend
    type: line
    title: "月度销售趋势"
    source_file: "销售明细.xlsx"
    sheet: "月度"
    mode: header_match
    x_header: "月份"
    y_header: "销售额"

tables:
  - key: sales_detail
    title: "销售明细"
    source_file: "销售明细.xlsx"
    sheet: "明细"
    mode: header_match
    columns: ["日期", "客户", "产品", "金额"]
```

`uploads/.gitkeep`：空文件。

创建虚拟环境并安装依赖：
```bash
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt
```

- [ ] **Step 2: 写失败的测试**

```python
# tests/test_extractor.py
import textwrap

from extractor import load_config


def test_load_config_reads_all_sections(tmp_path):
    config_path = tmp_path / "config.yaml"
    config_path.write_text(
        textwrap.dedent(
            """
            kpis:
              - key: total_revenue
                label: "总营收"
                source_file: "经营数据.xlsx"
                sheet: "汇总"
                mode: fixed_range
                range: "B2"
            charts:
              - key: monthly_sales_trend
                type: line
                title: "月度销售趋势"
                source_file: "销售明细.xlsx"
                sheet: "月度"
                mode: header_match
                x_header: "月份"
                y_header: "销售额"
            tables:
              - key: sales_detail
                title: "销售明细"
                source_file: "销售明细.xlsx"
                sheet: "明细"
                mode: header_match
                columns: ["日期", "客户", "产品", "金额"]
            """
        ),
        encoding="utf-8",
    )

    config = load_config(config_path)

    assert config["kpis"][0]["key"] == "total_revenue"
    assert config["charts"][0]["key"] == "monthly_sales_trend"
    assert config["tables"][0]["key"] == "sales_detail"


def test_load_config_defaults_missing_sections_to_empty_list(tmp_path):
    config_path = tmp_path / "config.yaml"
    config_path.write_text("kpis: []\n", encoding="utf-8")

    config = load_config(config_path)

    assert config["kpis"] == []
    assert config["charts"] == []
    assert config["tables"] == []
```

- [ ] **Step 3: 运行测试确认失败**

Run: `pytest tests/test_extractor.py -v`
Expected: FAIL，报 `ModuleNotFoundError: No module named 'extractor'`（文件还不存在）

- [ ] **Step 4: 写最小实现**

```python
# extractor.py
import yaml


def load_config(config_path):
    with open(config_path, "r", encoding="utf-8") as f:
        raw = yaml.safe_load(f) or {}
    return {
        "kpis": raw.get("kpis", []),
        "charts": raw.get("charts", []),
        "tables": raw.get("tables", []),
    }
```

- [ ] **Step 5: 运行测试确认通过**

Run: `pytest tests/test_extractor.py -v`
Expected: PASS（2 passed）

- [ ] **Step 6: 提交**

```bash
git add requirements.txt .gitignore config.yaml uploads/.gitkeep extractor.py tests/test_extractor.py
git commit -m "feat: add project scaffolding and config loader"
```

---

### Task 2: KPI 固定范围（fixed_range）取数

**Files:**
- Create: `tests/conftest.py`
- Modify: `extractor.py`
- Modify: `tests/test_extractor.py`

**Interfaces:**
- Consumes: `load_config()` 产出的 kpi item 结构（`key`/`label`/`source_file`/`sheet`/`mode`/`range`）
- Produces: `extractor.extract_kpi(item: dict, uploads_dir) -> dict`，返回 `{'key', 'label', 'value', 'error'}`（后续任务 build_dashboard_data 依赖此签名）

- [ ] **Step 1: 创建共享测试 fixture**

```python
# tests/conftest.py
import openpyxl
import pytest


@pytest.fixture
def uploads_dir(tmp_path):
    d = tmp_path / "uploads"
    d.mkdir()
    return d


@pytest.fixture
def make_workbook(uploads_dir):
    def _make(filename, sheet_data):
        """sheet_data: {sheet_name: [[row1_values], [row2_values], ...]}"""
        wb = openpyxl.Workbook()
        wb.remove(wb.active)
        for sheet_name, rows in sheet_data.items():
            ws = wb.create_sheet(sheet_name)
            for row in rows:
                ws.append(row)
        path = uploads_dir / filename
        wb.save(path)
        return path
    return _make
```

- [ ] **Step 2: 写失败的测试**

```python
# 追加到 tests/test_extractor.py
from extractor import extract_kpi


def test_extract_kpi_fixed_range_single_cell(uploads_dir, make_workbook):
    make_workbook("经营数据.xlsx", {"汇总": [["指标", "数值"], ["总营收", 1000000]]})
    item = {
        "key": "total_revenue",
        "label": "总营收",
        "source_file": "经营数据.xlsx",
        "sheet": "汇总",
        "mode": "fixed_range",
        "range": "B2",
    }

    result = extract_kpi(item, uploads_dir)

    assert result == {
        "key": "total_revenue",
        "label": "总营收",
        "value": 1000000,
        "error": None,
    }


def test_extract_kpi_fixed_range_missing_file_reports_error(uploads_dir, make_workbook):
    item = {
        "key": "total_revenue",
        "label": "总营收",
        "source_file": "不存在.xlsx",
        "sheet": "汇总",
        "mode": "fixed_range",
        "range": "B2",
    }

    result = extract_kpi(item, uploads_dir)

    assert result["value"] is None
    assert "不存在.xlsx" in result["error"]


def test_extract_kpi_fixed_range_missing_sheet_reports_error(uploads_dir, make_workbook):
    make_workbook("经营数据.xlsx", {"汇总": [["指标", "数值"], ["总营收", 1000000]]})
    item = {
        "key": "total_revenue",
        "label": "总营收",
        "source_file": "经营数据.xlsx",
        "sheet": "不存在的表",
        "mode": "fixed_range",
        "range": "B2",
    }

    result = extract_kpi(item, uploads_dir)

    assert result["value"] is None
    assert "不存在的表" in result["error"]
```

- [ ] **Step 3: 运行测试确认失败**

Run: `pytest tests/test_extractor.py -v`
Expected: FAIL，报 `ImportError: cannot import name 'extract_kpi' from 'extractor'`

- [ ] **Step 4: 写最小实现**

```python
# 追加到 extractor.py
from pathlib import Path

import openpyxl


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


def extract_kpi(item, uploads_dir):
    key = item["key"]
    label = item["label"]
    try:
        if item["mode"] == "fixed_range":
            worksheet = _load_worksheet(uploads_dir, item["source_file"], item["sheet"])
            values = _read_fixed_range_values(worksheet, item["range"])
        else:
            raise ValueError(f"未知的取数模式: {item['mode']}")
        numeric_values = [v for v in values if isinstance(v, (int, float))]
        value = sum(numeric_values) if numeric_values else None
        return {"key": key, "label": label, "value": value, "error": None}
    except ValueError as exc:
        return {"key": key, "label": label, "value": None, "error": str(exc)}
```

- [ ] **Step 5: 运行测试确认通过**

Run: `pytest tests/test_extractor.py -v`
Expected: PASS（5 passed）

- [ ] **Step 6: 提交**

```bash
git add extractor.py tests/conftest.py tests/test_extractor.py
git commit -m "feat: add fixed_range KPI extraction"
```

---

### Task 3: KPI 按表头匹配（header_match）取数

**Files:**
- Modify: `extractor.py`
- Modify: `tests/test_extractor.py`

**Interfaces:**
- Consumes: `_resolve_file()`（Task 2）
- Produces: `extractor._read_dataframe(uploads_dir, source_file, sheet_name) -> pandas.DataFrame`、`extractor._clean_value(v)`（后续 Task 4/5 的 chart/table 提取复用这两个函数）
- `extract_kpi` 新增支持 `mode: header_match`，行为：对匹配到的列求数值和作为 KPI 值（因为 header_match 通常对应"明细列"而非单一数字；若需要精确单值，配置应使用 `fixed_range`）

- [ ] **Step 1: 写失败的测试**

```python
# 追加到 tests/test_extractor.py
def test_extract_kpi_header_match_sums_numeric_column(uploads_dir, make_workbook):
    make_workbook(
        "经营数据.xlsx",
        {"汇总": [["月份", "利润"], ["1月", 100], ["2月", 200], ["3月", 300]]},
    )
    item = {
        "key": "total_profit",
        "label": "总利润",
        "source_file": "经营数据.xlsx",
        "sheet": "汇总",
        "mode": "header_match",
        "header": "利润",
    }

    result = extract_kpi(item, uploads_dir)

    assert result == {"key": "total_profit", "label": "总利润", "value": 600, "error": None}


def test_extract_kpi_header_match_missing_header_reports_error(uploads_dir, make_workbook):
    make_workbook("经营数据.xlsx", {"汇总": [["月份", "利润"], ["1月", 100]]})
    item = {
        "key": "total_profit",
        "label": "总利润",
        "source_file": "经营数据.xlsx",
        "sheet": "汇总",
        "mode": "header_match",
        "header": "不存在的表头",
    }

    result = extract_kpi(item, uploads_dir)

    assert result["value"] is None
    assert "不存在的表头" in result["error"]
```

- [ ] **Step 2: 运行测试确认失败**

Run: `pytest tests/test_extractor.py -v`
Expected: FAIL，`test_extract_kpi_header_match_sums_numeric_column` 报 `ValueError: 未知的取数模式: header_match` 被捕获后返回的 `result["error"]` 不为 `None`，断言 `result == {...}` 失败

- [ ] **Step 3: 写最小实现**

```python
# 在 extractor.py 顶部新增 import
import pandas as pd


# 新增函数
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


# 修改 extract_kpi，在 if/else 中插入 header_match 分支
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
```

- [ ] **Step 4: 运行测试确认通过**

Run: `pytest tests/test_extractor.py -v`
Expected: PASS（7 passed）

- [ ] **Step 5: 提交**

```bash
git add extractor.py tests/test_extractor.py
git commit -m "feat: add header_match KPI extraction"
```

---

### Task 4: 图表（chart）取数

**Files:**
- Modify: `extractor.py`
- Modify: `tests/test_extractor.py`

**Interfaces:**
- Consumes: `_read_dataframe()`、`_clean_value()`（Task 3）
- Produces: `extractor.extract_chart(item: dict, uploads_dir) -> dict`，返回 `{'key', 'type', 'title', 'x': [...], 'y': [...], 'error'}`

- [ ] **Step 1: 写失败的测试**

```python
# 追加到 tests/test_extractor.py
from extractor import extract_chart


def test_extract_chart_header_match_reads_x_and_y(uploads_dir, make_workbook):
    make_workbook(
        "销售明细.xlsx",
        {"月度": [["月份", "销售额"], ["1月", 100], ["2月", 200]]},
    )
    item = {
        "key": "monthly_sales_trend",
        "type": "line",
        "title": "月度销售趋势",
        "source_file": "销售明细.xlsx",
        "sheet": "月度",
        "mode": "header_match",
        "x_header": "月份",
        "y_header": "销售额",
    }

    result = extract_chart(item, uploads_dir)

    assert result == {
        "key": "monthly_sales_trend",
        "type": "line",
        "title": "月度销售趋势",
        "x": ["1月", "2月"],
        "y": [100, 200],
        "error": None,
    }


def test_extract_chart_missing_header_reports_error(uploads_dir, make_workbook):
    make_workbook("销售明细.xlsx", {"月度": [["月份", "销售额"], ["1月", 100]]})
    item = {
        "key": "monthly_sales_trend",
        "type": "line",
        "title": "月度销售趋势",
        "source_file": "销售明细.xlsx",
        "sheet": "月度",
        "mode": "header_match",
        "x_header": "月份",
        "y_header": "不存在的列",
    }

    result = extract_chart(item, uploads_dir)

    assert result["x"] == []
    assert result["y"] == []
    assert "不存在的列" in result["error"]
```

- [ ] **Step 2: 运行测试确认失败**

Run: `pytest tests/test_extractor.py -v`
Expected: FAIL，`ImportError: cannot import name 'extract_chart' from 'extractor'`

- [ ] **Step 3: 写最小实现**

```python
# 追加到 extractor.py
def extract_chart(item, uploads_dir):
    key = item["key"]
    try:
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
    except ValueError as exc:
        return {
            "key": key,
            "type": item.get("type"),
            "title": item.get("title"),
            "x": [],
            "y": [],
            "error": str(exc),
        }
```

- [ ] **Step 4: 运行测试确认通过**

Run: `pytest tests/test_extractor.py -v`
Expected: PASS（9 passed）

- [ ] **Step 5: 提交**

```bash
git add extractor.py tests/test_extractor.py
git commit -m "feat: add chart extraction"
```

---

### Task 5: 明细表格（table）取数

**Files:**
- Modify: `extractor.py`
- Modify: `tests/test_extractor.py`

**Interfaces:**
- Consumes: `_read_dataframe()`、`_clean_value()`（Task 3）
- Produces: `extractor.extract_table(item: dict, uploads_dir) -> dict`，返回 `{'key', 'title', 'columns': [...], 'rows': [[...]], 'error', 'view_group', 'view_label'}`；`view_group`/`view_label` 原样透传自 `item`（未配置时为 `None`），供 Task 12 前端按 `view_group` 归并渲染 tab

- [ ] **Step 1: 写失败的测试**

```python
# 追加到 tests/test_extractor.py
from extractor import extract_table


def test_extract_table_header_match_reads_selected_columns(uploads_dir, make_workbook):
    make_workbook(
        "销售明细.xlsx",
        {
            "明细": [
                ["日期", "客户", "产品", "金额", "备注"],
                ["2026-01-01", "客户A", "产品X", 1000, "无"],
                ["2026-01-02", "客户B", "产品Y", 2000, "无"],
            ]
        },
    )
    item = {
        "key": "sales_detail",
        "title": "销售明细",
        "source_file": "销售明细.xlsx",
        "sheet": "明细",
        "mode": "header_match",
        "columns": ["日期", "客户", "产品", "金额"],
    }

    result = extract_table(item, uploads_dir)

    assert result["columns"] == ["日期", "客户", "产品", "金额"]
    assert result["rows"] == [
        ["2026-01-01", "客户A", "产品X", 1000],
        ["2026-01-02", "客户B", "产品Y", 2000],
    ]
    assert result["error"] is None


def test_extract_table_missing_column_reports_error(uploads_dir, make_workbook):
    make_workbook(
        "销售明细.xlsx",
        {"明细": [["日期", "客户"], ["2026-01-01", "客户A"]]},
    )
    item = {
        "key": "sales_detail",
        "title": "销售明细",
        "source_file": "销售明细.xlsx",
        "sheet": "明细",
        "mode": "header_match",
        "columns": ["日期", "客户", "产品"],
    }

    result = extract_table(item, uploads_dir)

    assert result["rows"] == []
    assert "产品" in result["error"]
```

- [ ] **Step 2: 运行测试确认失败**

Run: `pytest tests/test_extractor.py -v`
Expected: FAIL，`ImportError: cannot import name 'extract_table' from 'extractor'`

- [ ] **Step 3: 写最小实现**

```python
# 追加到 extractor.py
def extract_table(item, uploads_dir):
    key = item["key"]
    try:
        df = _read_dataframe(uploads_dir, item["source_file"], item["sheet"])
        columns = item["columns"]
        missing = [c for c in columns if c not in df.columns]
        if missing:
            raise ValueError(f"找不到表头: '{missing[0]}'")
        rows = [[_clean_value(v) for v in row] for row in df[columns].values.tolist()]
        return {
            "key": key,
            "title": item["title"],
            "columns": columns,
            "rows": rows,
            "error": None,
            "view_group": item.get("view_group"),
            "view_label": item.get("view_label"),
        }
    except ValueError as exc:
        return {
            "key": key,
            "title": item.get("title"),
            "columns": item.get("columns", []),
            "rows": [],
            "error": str(exc),
            "view_group": item.get("view_group"),
            "view_label": item.get("view_label"),
        }
```

- [ ] **Step 4: 运行测试确认通过**

Run: `pytest tests/test_extractor.py -v`
Expected: PASS（11 passed）

- [ ] **Step 5: 提交**

```bash
git add extractor.py tests/test_extractor.py
git commit -m "feat: add table extraction"
```

---

### Task 5.1: 明细表格 group_by_sum（分组汇总）取数

**Files:**
- Modify: `extractor.py`
- Modify: `tests/test_extractor.py`

**Interfaces:**
- Consumes: `_read_dataframe()`、`_clean_value()`（Task 3）
- Produces: `extractor.extract_table` 新增支持 `mode: group_by_sum`：按 `group_by_header` 分组、对 `sum_header` 求和，并按汇总值降序排列，返回的 `columns` 固定为 `[group_by_header, sum_header]`；与 `header_match` 分支互斥，由 `item["mode"]` 决定走哪条路径；本任务整体重写 `extract_table`，继续保留 Task 5 引入的 `view_group`/`view_label` 透传字段

- [ ] **Step 1: 写失败的测试**

```python
# 追加到 tests/test_extractor.py
def test_extract_table_group_by_sum_ranks_descending(uploads_dir, make_workbook):
    make_workbook(
        "销售明细.xlsx",
        {
            "明细": [
                ["日期", "客户", "产品", "金额"],
                ["2026-01-01", "客户A", "产品X", 1000],
                ["2026-01-02", "客户B", "产品Y", 2000],
                ["2026-01-03", "客户A", "产品Z", 500],
            ]
        },
    )
    item = {
        "key": "sales_by_customer",
        "title": "客户销售排名",
        "source_file": "销售明细.xlsx",
        "sheet": "明细",
        "mode": "group_by_sum",
        "group_by_header": "客户",
        "sum_header": "金额",
    }

    result = extract_table(item, uploads_dir)

    assert result["columns"] == ["客户", "金额"]
    assert result["rows"] == [["客户B", 2000], ["客户A", 1500]]
    assert result["error"] is None


def test_extract_table_group_by_sum_missing_header_reports_error(uploads_dir, make_workbook):
    make_workbook(
        "销售明细.xlsx",
        {"明细": [["日期", "客户"], ["2026-01-01", "客户A"]]},
    )
    item = {
        "key": "sales_by_customer",
        "title": "客户销售排名",
        "source_file": "销售明细.xlsx",
        "sheet": "明细",
        "mode": "group_by_sum",
        "group_by_header": "客户",
        "sum_header": "金额",
    }

    result = extract_table(item, uploads_dir)

    assert result["rows"] == []
    assert "金额" in result["error"]
```

- [ ] **Step 2: 运行测试确认失败**

Run: `pytest tests/test_extractor.py -v`
Expected: FAIL，`test_extract_table_group_by_sum_ranks_descending` 报 `KeyError: 'columns'`（当前实现完全忽略 `item["mode"]`，仍按 header_match 逻辑去读 `item["columns"]`，而 group_by_sum 配置项没有这个字段）

- [ ] **Step 3: 写最小实现**

```python
# 修改 extractor.py 中的 extract_table，按 item["mode"] 分支
def extract_table(item, uploads_dir):
    key = item["key"]
    try:
        df = _read_dataframe(uploads_dir, item["source_file"], item["sheet"])
        mode = item["mode"]
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
    except ValueError as exc:
        return {
            "key": key,
            "title": item.get("title"),
            "columns": item.get("columns", []),
            "rows": [],
            "error": str(exc),
            "view_group": item.get("view_group"),
            "view_label": item.get("view_label"),
        }
```

- [ ] **Step 4: 运行测试确认通过**

Run: `pytest tests/test_extractor.py -v`
Expected: PASS（13 passed）

- [ ] **Step 5: 提交**

```bash
git add extractor.py tests/test_extractor.py
git commit -m "feat: add group_by_sum table extraction mode"
```

---

### Task 6: 编排（build_dashboard_data）与 data.json 落盘

**Files:**
- Modify: `extractor.py`
- Modify: `tests/test_extractor.py`

**Interfaces:**
- Consumes: `extract_kpi`/`extract_chart`/`extract_table`（Task 2-5）
- Produces: `extractor.build_dashboard_data(config: dict, uploads_dir) -> dict`（返回 `{'kpis', 'charts', 'tables'}`，单项失败不影响其余项，供 Task 9 上传处理调用）、`extractor.save_data_json(data, data_path) -> None`、`extractor.load_data_json(data_path) -> dict`（供 Task 7 的 `/api/data` 调用）
- `build_dashboard_data` 始终基于传入的 `uploads_dir` 当前全部内容重新计算每一项，天然满足"每次解析读取整个 uploads 目录"的要求——调用方（Task 9）无需自己做目录遍历

- [ ] **Step 1: 写失败的测试**

```python
# 追加到 tests/test_extractor.py
from extractor import build_dashboard_data, save_data_json, load_data_json


def test_build_dashboard_data_combines_all_sections_and_isolates_failures(uploads_dir, make_workbook):
    make_workbook("经营数据.xlsx", {"汇总": [["指标", "数值"], ["总营收", 1000000]]})
    config = {
        "kpis": [
            {
                "key": "total_revenue",
                "label": "总营收",
                "source_file": "经营数据.xlsx",
                "sheet": "汇总",
                "mode": "fixed_range",
                "range": "B2",
            },
            {
                "key": "missing_kpi",
                "label": "缺失指标",
                "source_file": "不存在.xlsx",
                "sheet": "汇总",
                "mode": "fixed_range",
                "range": "B2",
            },
        ],
        "charts": [],
        "tables": [],
    }

    data = build_dashboard_data(config, uploads_dir)

    assert data["kpis"][0]["value"] == 1000000
    assert data["kpis"][0]["error"] is None
    assert data["kpis"][1]["value"] is None
    assert data["kpis"][1]["error"] is not None
    assert data["charts"] == []
    assert data["tables"] == []


def test_save_and_load_data_json_roundtrip(tmp_path):
    data = {
        "kpis": [{"key": "k", "label": "中文标签", "value": 1, "error": None}],
        "charts": [],
        "tables": [],
    }
    data_path = tmp_path / "data.json"

    save_data_json(data, data_path)
    loaded = load_data_json(data_path)

    assert loaded == data
    assert "中文标签" in data_path.read_text(encoding="utf-8")
```

- [ ] **Step 2: 运行测试确认失败**

Run: `pytest tests/test_extractor.py -v`
Expected: FAIL，`ImportError: cannot import name 'build_dashboard_data' from 'extractor'`

- [ ] **Step 3: 写最小实现**

```python
# 在 extractor.py 顶部新增 import
import json


# 追加到 extractor.py
def build_dashboard_data(config, uploads_dir):
    return {
        "kpis": [extract_kpi(item, uploads_dir) for item in config["kpis"]],
        "charts": [extract_chart(item, uploads_dir) for item in config["charts"]],
        "tables": [extract_table(item, uploads_dir) for item in config["tables"]],
    }


def save_data_json(data, data_path):
    with open(data_path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def load_data_json(data_path):
    with open(data_path, "r", encoding="utf-8") as f:
        return json.load(f)
```

- [ ] **Step 4: 运行测试确认通过**

Run: `pytest tests/test_extractor.py -v`
Expected: PASS（15 passed）

- [ ] **Step 5: 提交**

```bash
git add extractor.py tests/test_extractor.py
git commit -m "feat: add dashboard data orchestration and JSON persistence"
```

---

### Task 6.1: KPI 计算型指标（computed）跨指标运算

**Files:**
- Modify: `extractor.py`
- Modify: `tests/test_extractor.py`

**Interfaces:**
- Consumes: `extract_kpi`（Task 2-3）、`build_dashboard_data`（Task 6，本任务在其基础上改为两遍扫描）
- Produces: `extractor._resolve_computed_kpi(item: dict, resolved_values: dict) -> dict`（返回 `{'key', 'label', 'value', 'error'}`）；`build_dashboard_data` 新增支持 `kpis` 配置项中 `mode: computed`——其值不来自 Excel，而是引用其余已解析 KPI 的 `value`，按 `operation`/`from`/`minus` 字段计算，结果仍出现在 `data["kpis"]` 数组中且保持 config 中声明的原始顺序

- [ ] **Step 1: 写失败的测试**

```python
# 追加到 tests/test_extractor.py
def test_build_dashboard_data_resolves_computed_kpi_by_subtracting_two_kpis(uploads_dir, make_workbook):
    make_workbook(
        "库存表.xlsx",
        {"汇总": [["指标", "数值"], ["库存余额", 500], ["港口实际库存", 320]]},
    )
    config = {
        "kpis": [
            {
                "key": "stock_balance",
                "label": "库存余额",
                "source_file": "库存表.xlsx",
                "sheet": "汇总",
                "mode": "fixed_range",
                "range": "B2",
            },
            {
                "key": "port_actual_stock",
                "label": "港口实际库存",
                "source_file": "库存表.xlsx",
                "sheet": "汇总",
                "mode": "fixed_range",
                "range": "B3",
            },
            {
                "key": "stock_discrepancy",
                "label": "库存账实差异",
                "mode": "computed",
                "operation": "subtract",
                "from": "stock_balance",
                "minus": "port_actual_stock",
            },
        ],
        "charts": [],
        "tables": [],
    }

    data = build_dashboard_data(config, uploads_dir)

    assert [k["key"] for k in data["kpis"]] == ["stock_balance", "port_actual_stock", "stock_discrepancy"]
    assert data["kpis"][2]["value"] == 180
    assert data["kpis"][2]["error"] is None


def test_build_dashboard_data_computed_kpi_reports_error_when_reference_missing(uploads_dir, make_workbook):
    make_workbook("库存表.xlsx", {"汇总": [["指标", "数值"], ["库存余额", 500]]})
    config = {
        "kpis": [
            {
                "key": "stock_balance",
                "label": "库存余额",
                "source_file": "库存表.xlsx",
                "sheet": "汇总",
                "mode": "fixed_range",
                "range": "B2",
            },
            {
                "key": "stock_discrepancy",
                "label": "库存账实差异",
                "mode": "computed",
                "operation": "subtract",
                "from": "stock_balance",
                "minus": "port_actual_stock",
            },
        ],
        "charts": [],
        "tables": [],
    }

    data = build_dashboard_data(config, uploads_dir)

    assert data["kpis"][1]["value"] is None
    assert "port_actual_stock" in data["kpis"][1]["error"]
```

- [ ] **Step 2: 运行测试确认失败**

Run: `pytest tests/test_extractor.py -v`
Expected: FAIL，两个新测试均失败——当前 `build_dashboard_data` 对每个 `kpis` 条目都直接调用 `extract_kpi`，而 `extract_kpi` 遇到 `mode: computed` 时会走到"未知的取数模式"分支：第一个测试断言 `data["kpis"][2]["value"] == 180` 失败（实际为 `None`）；第二个测试断言 `"port_actual_stock" in data["kpis"][1]["error"]` 失败（实际 `error` 是 `"未知的取数模式: computed"`）

- [ ] **Step 3: 写最小实现**

```python
# 追加到 extractor.py
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


# 修改 build_dashboard_data，改为两遍扫描：先解析所有非 computed 的 KPI，再解析 computed 的 KPI
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
```

- [ ] **Step 4: 运行测试确认通过**

Run: `pytest tests/test_extractor.py -v`
Expected: PASS（17 passed）

- [ ] **Step 5: 提交**

```bash
git add extractor.py tests/test_extractor.py
git commit -m "feat: add computed KPI resolution across other KPIs"
```

---

### Task 7: Flask 应用骨架（/dashboard、/api/data，无需密码）

**Files:**
- Create: `app.py`
- Create: `templates/dashboard.html`
- Test: `tests/test_app.py`

**Interfaces:**
- Consumes: `extractor.load_data_json`（Task 6）
- Produces: `app.create_app(test_config: dict = None) -> Flask`（Task 8/9 在此基础上添加 `/login`、`/upload` 路由）
- `/dashboard`、`/api/data` 不做任何鉴权（spec 明确要求）

- [ ] **Step 1: 写失败的测试**

```python
# tests/test_app.py
import pytest

from app import create_app


@pytest.fixture
def client(tmp_path):
    uploads_dir = tmp_path / "uploads"
    uploads_dir.mkdir()
    app = create_app(
        {
            "TESTING": True,
            "UPLOAD_PASSWORD": "test-pass",
            "SECRET_KEY": "test-secret",
            "UPLOAD_DIR": uploads_dir,
            "CONFIG_PATH": tmp_path / "config.yaml",
            "DATA_PATH": tmp_path / "data.json",
        }
    )
    return app.test_client()


def test_index_redirects_to_dashboard(client):
    response = client.get("/")
    assert response.status_code == 302
    assert response.headers["Location"].endswith("/dashboard")


def test_dashboard_is_publicly_accessible(client):
    response = client.get("/dashboard")
    assert response.status_code == 200


def test_api_data_returns_empty_state_when_no_data_json(client):
    response = client.get("/api/data")
    assert response.status_code == 200
    body = response.get_json()
    assert body["kpis"] == []
    assert body["charts"] == []
    assert body["tables"] == []


def test_api_data_returns_saved_data(client, tmp_path):
    from extractor import save_data_json

    save_data_json(
        {"kpis": [{"key": "k"}], "charts": [], "tables": []},
        tmp_path / "data.json",
    )

    response = client.get("/api/data")

    assert response.get_json()["kpis"] == [{"key": "k"}]
```

- [ ] **Step 2: 运行测试确认失败**

Run: `pytest tests/test_app.py -v`
Expected: FAIL，`ModuleNotFoundError: No module named 'app'`

- [ ] **Step 3: 写最小实现**

```python
# app.py
import os
from pathlib import Path

from flask import Flask, jsonify, redirect, render_template, url_for

from extractor import load_data_json


def create_app(test_config=None):
    app = Flask(__name__)
    app.config.update(
        UPLOAD_PASSWORD=os.environ.get("UPLOAD_PASSWORD", "changeme"),
        SECRET_KEY=os.environ.get("FLASK_SECRET_KEY", "dev-secret-key-change-me"),
        UPLOAD_DIR=Path("uploads"),
        CONFIG_PATH=Path("config.yaml"),
        DATA_PATH=Path("data.json"),
    )
    if test_config:
        app.config.update(test_config)

    @app.route("/")
    def index():
        return redirect(url_for("dashboard"))

    @app.route("/dashboard")
    def dashboard():
        return render_template("dashboard.html")

    @app.route("/api/data")
    def api_data():
        data_path = Path(app.config["DATA_PATH"])
        if not data_path.exists():
            return jsonify({"kpis": [], "charts": [], "tables": []})
        return jsonify(load_data_json(data_path))

    return app


if __name__ == "__main__":
    if os.environ.get("UPLOAD_PASSWORD") is None:
        print("警告: 未设置环境变量 UPLOAD_PASSWORD，使用默认密码 'changeme'，请在生产环境中修改。")
    flask_app = create_app()
    flask_app.run(host="0.0.0.0", port=5000, debug=False)
```

```html
<!-- templates/dashboard.html -->
<!DOCTYPE html>
<html lang="zh-CN">
<head>
  <meta charset="UTF-8">
  <title>经营看板</title>
</head>
<body>
  <div id="app">看板加载中...</div>
</body>
</html>
```

- [ ] **Step 4: 运行测试确认通过**

Run: `pytest tests/test_app.py -v`
Expected: PASS（4 passed）

- [ ] **Step 5: 提交**

```bash
git add app.py templates/dashboard.html tests/test_app.py
git commit -m "feat: add Flask app skeleton with public dashboard and data API"
```

---

### Task 8: 密码登录与 /upload 访问保护

**Files:**
- Modify: `app.py`
- Create: `templates/login.html`
- Create: `templates/upload.html`
- Modify: `tests/test_app.py`

**Interfaces:**
- Consumes: `create_app()`（Task 7）
- Produces: `login_required` 装饰器、`/login`（GET/POST）、`/upload`（当前仅 GET，Task 9 补充 POST）

- [ ] **Step 1: 写失败的测试**

```python
# 追加到 tests/test_app.py
def test_login_page_loads(client):
    response = client.get("/login")
    assert response.status_code == 200


def test_login_with_wrong_password_shows_error(client):
    response = client.post("/login", data={"password": "wrong"})
    assert response.status_code == 200
    assert "密码错误".encode() in response.data


def test_login_with_correct_password_redirects_to_upload(client):
    response = client.post("/login", data={"password": "test-pass"})
    assert response.status_code == 302
    assert response.headers["Location"].endswith("/upload")


def test_upload_page_requires_login(client):
    response = client.get("/upload")
    assert response.status_code == 302
    assert "/login" in response.headers["Location"]


def test_upload_page_accessible_after_login(client):
    client.post("/login", data={"password": "test-pass"})
    response = client.get("/upload")
    assert response.status_code == 200
```

- [ ] **Step 2: 运行测试确认失败**

Run: `pytest tests/test_app.py -v`
Expected: FAIL，`404 Not Found`（`/login`、`/upload` 路由还不存在）

- [ ] **Step 3: 写最小实现**

```python
# app.py 顶部新增 import
import secrets
from functools import wraps

from flask import flash, request, session
```

```python
# 追加到 create_app() 内部，在 return app 之前
def login_required(view):
    @wraps(view)
    def wrapped(*args, **kwargs):
        if not session.get("logged_in"):
            return redirect(url_for("login"))
        return view(*args, **kwargs)
    return wrapped

@app.route("/login", methods=["GET", "POST"])
def login():
    error = None
    if request.method == "POST":
        password = request.form.get("password", "")
        if secrets.compare_digest(password, app.config["UPLOAD_PASSWORD"]):
            session["logged_in"] = True
            return redirect(url_for("upload"))
        error = "密码错误"
    return render_template("login.html", error=error)

@app.route("/upload", methods=["GET"])
@login_required
def upload():
    return render_template("upload.html")
```

```html
<!-- templates/login.html -->
<!DOCTYPE html>
<html lang="zh-CN">
<head>
  <meta charset="UTF-8">
  <title>登录 - 数据维护</title>
</head>
<body>
  <h1>数据维护登录</h1>
  {% if error %}<p style="color:red;">{{ error }}</p>{% endif %}
  <form method="post">
    <input type="password" name="password" placeholder="请输入密码" required>
    <button type="submit">登录</button>
  </form>
</body>
</html>
```

```html
<!-- templates/upload.html -->
<!DOCTYPE html>
<html lang="zh-CN">
<head>
  <meta charset="UTF-8">
  <title>上传 Excel 数据</title>
</head>
<body>
  <h1>上传 Excel 数据</h1>
</body>
</html>
```

- [ ] **Step 4: 运行测试确认通过**

Run: `pytest tests/test_app.py -v`
Expected: PASS（9 passed）

- [ ] **Step 5: 提交**

```bash
git add app.py templates/login.html templates/upload.html tests/test_app.py
git commit -m "feat: add password login and protect /upload"
```

---

### Task 9: 上传处理（多文件保存/覆盖 + 触发解析 + 结果反馈）

**Files:**
- Modify: `app.py`
- Modify: `templates/upload.html`
- Modify: `tests/test_app.py`

**Interfaces:**
- Consumes: `extractor.load_config`、`extractor.build_dashboard_data`、`extractor.save_data_json`（Task 1/6）
- Produces: `/upload` 的 POST 分支；`_is_safe_filename(filename) -> bool`

- [ ] **Step 1: 写失败的测试**

先替换 `client` fixture（新增默认空 `config.yaml`，供解析流程使用），并追加新测试：

```python
# 替换 tests/test_app.py 顶部的 import 与 client fixture
import io

import pytest

from app import create_app


@pytest.fixture
def client(tmp_path):
    uploads_dir = tmp_path / "uploads"
    uploads_dir.mkdir()
    config_path = tmp_path / "config.yaml"
    config_path.write_text("kpis: []\ncharts: []\ntables: []\n", encoding="utf-8")
    app = create_app(
        {
            "TESTING": True,
            "UPLOAD_PASSWORD": "test-pass",
            "SECRET_KEY": "test-secret",
            "UPLOAD_DIR": uploads_dir,
            "CONFIG_PATH": config_path,
            "DATA_PATH": tmp_path / "data.json",
        }
    )
    return app.test_client()
```

```python
# 追加到 tests/test_app.py
def test_upload_saves_file_and_generates_data_json(client, tmp_path):
    import openpyxl

    client.post("/login", data={"password": "test-pass"})

    (tmp_path / "config.yaml").write_text(
        "kpis:\n"
        "  - key: total_revenue\n"
        "    label: \"总营收\"\n"
        "    source_file: \"经营数据.xlsx\"\n"
        "    sheet: \"汇总\"\n"
        "    mode: fixed_range\n"
        "    range: \"B2\"\n"
        "charts: []\n"
        "tables: []\n",
        encoding="utf-8",
    )

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "汇总"
    ws.append(["指标", "数值"])
    ws.append(["总营收", 500])
    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)

    response = client.post(
        "/upload",
        data={"files": (buffer, "经营数据.xlsx")},
        content_type="multipart/form-data",
        follow_redirects=True,
    )

    assert response.status_code == 200
    assert (tmp_path / "uploads" / "经营数据.xlsx").exists()

    from extractor import load_data_json

    data = load_data_json(tmp_path / "data.json")
    assert data["kpis"][0]["value"] == 500


def test_upload_overwrites_existing_file_with_same_name(client, tmp_path):
    (tmp_path / "uploads" / "经营数据.xlsx").write_bytes(b"old content")
    client.post("/login", data={"password": "test-pass"})

    response = client.post(
        "/upload",
        data={"files": (io.BytesIO(b"new content"), "经营数据.xlsx")},
        content_type="multipart/form-data",
        follow_redirects=True,
    )

    assert response.status_code == 200
    assert (tmp_path / "uploads" / "经营数据.xlsx").read_bytes() == b"new content"


def test_upload_accepts_multiple_files_in_one_request(client, tmp_path):
    client.post("/login", data={"password": "test-pass"})

    response = client.post(
        "/upload",
        data={
            "files": [
                (io.BytesIO(b"a"), "经营数据.xlsx"),
                (io.BytesIO(b"b"), "销售明细.xlsx"),
            ]
        },
        content_type="multipart/form-data",
        follow_redirects=True,
    )

    assert response.status_code == 200
    assert (tmp_path / "uploads" / "经营数据.xlsx").exists()
    assert (tmp_path / "uploads" / "销售明细.xlsx").exists()


def test_upload_rejects_path_traversal_filename(client, tmp_path):
    client.post("/login", data={"password": "test-pass"})

    response = client.post(
        "/upload",
        data={"files": (io.BytesIO(b"data"), "../evil.xlsx")},
        content_type="multipart/form-data",
        follow_redirects=True,
    )

    assert response.status_code == 200
    assert not (tmp_path / "evil.xlsx").exists()
    assert not (tmp_path / "uploads" / "../evil.xlsx").resolve().exists()


def test_upload_rejects_non_excel_extension(client, tmp_path):
    client.post("/login", data={"password": "test-pass"})

    response = client.post(
        "/upload",
        data={"files": (io.BytesIO(b"data"), "malware.exe")},
        content_type="multipart/form-data",
        follow_redirects=True,
    )

    assert response.status_code == 200
    assert not (tmp_path / "uploads" / "malware.exe").exists()
```

- [ ] **Step 2: 运行测试确认失败**

Run: `pytest tests/test_app.py -v`
Expected: FAIL，上传的文件不会被保存（POST 分支还不存在，`/upload` 只处理了 GET）

- [ ] **Step 3: 写最小实现**

```python
# app.py 顶部的 extractor import 改为
from extractor import build_dashboard_data, load_config, load_data_json, save_data_json
```

```python
# 追加到 app.py（模块级）
ALLOWED_EXTENSIONS = {".xlsx", ".xls"}


def _is_safe_filename(filename):
    if not filename:
        return False
    if os.path.basename(filename) != filename:
        return False
    return Path(filename).suffix.lower() in ALLOWED_EXTENSIONS
```

```python
# 替换 create_app() 内的 upload() 视图函数
@app.route("/upload", methods=["GET", "POST"])
@login_required
def upload():
    if request.method == "POST":
        upload_dir = Path(app.config["UPLOAD_DIR"])
        upload_dir.mkdir(parents=True, exist_ok=True)

        saved, rejected = [], []
        for file in request.files.getlist("files"):
            filename = file.filename
            if not _is_safe_filename(filename):
                rejected.append(filename or "(空文件名)")
                continue
            file.save(upload_dir / filename)
            saved.append(filename)

        if saved:
            config = load_config(app.config["CONFIG_PATH"])
            data = build_dashboard_data(config, upload_dir)
            save_data_json(data, app.config["DATA_PATH"])
            errors = [
                item["error"]
                for item in data["kpis"] + data["charts"] + data["tables"]
                if item["error"]
            ]
            if errors:
                flash(f"已保存 {len(saved)} 个文件，但部分指标解析失败：" + "；".join(errors), "warning")
            else:
                flash(f"已成功保存并解析 {len(saved)} 个文件", "success")

        if rejected:
            flash(
                "以下文件被拒绝（仅支持 .xlsx/.xls，文件名不能包含路径）：" + "；".join(rejected),
                "error",
            )

        return redirect(url_for("upload"))

    return render_template("upload.html")
```

```html
<!-- templates/upload.html -->
<!DOCTYPE html>
<html lang="zh-CN">
<head>
  <meta charset="UTF-8">
  <title>上传 Excel 数据</title>
</head>
<body>
  <h1>上传 Excel 数据</h1>
  {% with messages = get_flashed_messages(with_categories=true) %}
    {% if messages %}
      <ul>
        {% for category, message in messages %}
          <li class="flash-{{ category }}">{{ message }}</li>
        {% endfor %}
      </ul>
    {% endif %}
  {% endwith %}
  <form method="post" enctype="multipart/form-data">
    <input type="file" name="files" multiple accept=".xlsx,.xls" required>
    <button type="submit">上传</button>
  </form>
</body>
</html>
```

- [ ] **Step 4: 运行测试确认通过**

Run: `pytest tests/test_app.py -v`
Expected: PASS（14 passed）

- [ ] **Step 5: 提交**

```bash
git add app.py templates/upload.html tests/test_app.py
git commit -m "feat: handle multi-file upload with overwrite and error feedback"
```

---

### Task 10: 前端骨架 + KPI 卡片渲染

**Files:**
- Modify: `templates/dashboard.html`
- Create: `static/css/dashboard.css`
- Create: `static/js/dashboard.js`
- Modify: `tests/test_app.py`

**Interfaces:**
- Consumes: `GET /api/data` 返回的 `{'kpis': [{'key','label','value','error'}], 'charts': [...], 'tables': [...]}`（Task 6/7）
- Produces: `dashboard.html` 中的容器 `#kpi-section`/`#chart-section`/`#table-section`（Task 11/12 分别在 chart-section/table-section 中渲染）；`dashboard.js` 中的 `loadDashboardData()`、`renderKpis(kpis)` 全局函数

- [ ] **Step 1: 写失败的测试**

```python
# 追加到 tests/test_app.py
def test_dashboard_page_includes_section_containers(client):
    response = client.get("/dashboard")
    html = response.data.decode("utf-8")

    assert 'id="kpi-section"' in html
    assert 'id="chart-section"' in html
    assert 'id="table-section"' in html
    assert "dashboard.css" in html
    assert "dashboard.js" in html


def test_static_dashboard_css_is_served(client):
    response = client.get("/static/css/dashboard.css")
    assert response.status_code == 200


def test_static_dashboard_js_is_served(client):
    response = client.get("/static/js/dashboard.js")
    assert response.status_code == 200
```

- [ ] **Step 2: 运行测试确认失败**

Run: `pytest tests/test_app.py -v`
Expected: FAIL，`dashboard.html` 还没有这些容器，`static/css/dashboard.css`、`static/js/dashboard.js` 还不存在（404）

- [ ] **Step 3: 写最小实现**

```html
<!-- 替换 templates/dashboard.html -->
<!DOCTYPE html>
<html lang="zh-CN">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>经营看板</title>
  <link rel="stylesheet" href="{{ url_for('static', filename='css/dashboard.css') }}">
</head>
<body>
  <header class="dashboard-header">
    <h1>经营看板</h1>
  </header>
  <main id="app">
    <section id="kpi-section" class="kpi-grid" aria-label="KPI 指标"></section>
    <section id="chart-section" class="chart-grid" aria-label="图表"></section>
    <section id="table-section" class="table-grid" aria-label="明细表格"></section>
  </main>
  <script src="{{ url_for('static', filename='js/dashboard.js') }}"></script>
</body>
</html>
```

```css
/* static/css/dashboard.css */
* {
  box-sizing: border-box;
}

body {
  margin: 0;
  font-family: -apple-system, "PingFang SC", "Microsoft YaHei", sans-serif;
  background-color: #f5f6f8;
  color: #1f2329;
}

.dashboard-header {
  padding: 16px 24px;
  background-color: #ffffff;
  border-bottom: 1px solid #e5e6eb;
}

.dashboard-header h1 {
  margin: 0;
  font-size: 20px;
  font-weight: 600;
}

#app {
  padding: 16px;
  display: flex;
  flex-direction: column;
  gap: 16px;
}

.kpi-grid {
  display: grid;
  grid-template-columns: repeat(4, 1fr);
  gap: 16px;
}

.kpi-card {
  background-color: #ffffff;
  border-radius: 8px;
  padding: 16px;
  box-shadow: 0 1px 2px rgba(0, 0, 0, 0.06);
}

.kpi-label {
  font-size: 13px;
  color: #6b7280;
  margin-bottom: 8px;
}

.kpi-value {
  font-size: 28px;
  font-weight: 700;
}

.kpi-error {
  font-size: 13px;
  color: #d93026;
}

.chart-grid {
  display: grid;
  grid-template-columns: repeat(2, 1fr);
  gap: 16px;
}

.table-grid {
  display: grid;
  grid-template-columns: 1fr;
  gap: 16px;
}

@media (max-width: 1024px) {
  .kpi-grid {
    grid-template-columns: repeat(2, 1fr);
  }
  .chart-grid {
    grid-template-columns: 1fr;
  }
}

@media (max-width: 640px) {
  .kpi-grid {
    grid-template-columns: 1fr;
  }
}
```

```js
// static/js/dashboard.js
async function loadDashboardData() {
  const response = await fetch("/api/data");
  const data = await response.json();
  renderKpis(data.kpis);
}

function renderKpis(kpis) {
  const container = document.getElementById("kpi-section");
  container.innerHTML = "";
  kpis.forEach((kpi) => {
    const card = document.createElement("div");
    card.className = "kpi-card";
    if (kpi.error) {
      card.innerHTML =
        `<div class="kpi-label">${kpi.label}</div>` +
        `<div class="kpi-error">数据异常：${kpi.error}</div>`;
    } else {
      card.innerHTML =
        `<div class="kpi-label">${kpi.label}</div>` +
        `<div class="kpi-value">${kpi.value}</div>`;
    }
    container.appendChild(card);
  });
}

document.addEventListener("DOMContentLoaded", loadDashboardData);
```

- [ ] **Step 4: 运行测试确认通过**

Run: `pytest tests/test_app.py -v`
Expected: PASS（17 passed）

- [ ] **Step 5: 手动浏览器验证**

```bash
source venv/bin/activate
UPLOAD_PASSWORD=test123 python app.py
```

1. 浏览器打开 `http://localhost:5000/upload`，输入密码 `test123` 登录，上传一份按 `config.yaml` 配置命名的测试 Excel（如 `经营数据.xlsx`，`汇总` sheet 的 `B2` 单元格填任意数字）。
2. 打开 `http://localhost:5000/dashboard`，确认 KPI 卡片区显示对应数值；调整浏览器窗口宽度到手机尺寸（约 375px），确认卡片自动收窄为单列，无横向滚动条。

- [ ] **Step 6: 提交**

```bash
git add templates/dashboard.html static/css/dashboard.css static/js/dashboard.js tests/test_app.py
git commit -m "feat: add dashboard shell and KPI card rendering"
```

---

### Task 11: 图表渲染（ECharts：折线/柱状/饼图 + 自适应重绘）

**Files:**
- Modify: `templates/dashboard.html`
- Modify: `static/css/dashboard.css`
- Modify: `static/js/dashboard.js`
- Modify: `tests/test_app.py`

**Interfaces:**
- Consumes: `data.charts` 中每一项 `{'key','type','title','x':[...],'y':[...],'error'}`（Task 4/6）；`#chart-section` 容器（Task 10）
- Produces: `dashboard.js` 中的 `renderCharts(charts)`、`buildChartOption(chart)`、模块级 `chartInstances` 字典（key -> ECharts 实例，供 resize 监听复用）

- [ ] **Step 1: 写失败的测试**

```python
# 追加到 tests/test_app.py
def test_dashboard_page_includes_echarts_cdn(client):
    response = client.get("/dashboard")
    html = response.data.decode("utf-8")
    assert "echarts" in html.lower()
```

- [ ] **Step 2: 运行测试确认失败**

Run: `pytest tests/test_app.py -v`
Expected: FAIL，`dashboard.html` 还没有引入 ECharts

- [ ] **Step 3: 写最小实现**

```html
<!-- 修改 templates/dashboard.html，在引入 dashboard.js 之前新增一行 -->
  <script src="https://cdn.jsdelivr.net/npm/echarts@5/dist/echarts.min.js"></script>
  <script src="{{ url_for('static', filename='js/dashboard.js') }}"></script>
```

```js
// 替换 static/js/dashboard.js 中的 loadDashboardData，并在文件顶部新增 chartInstances
const chartInstances = {};

async function loadDashboardData() {
  const response = await fetch("/api/data");
  const data = await response.json();
  renderKpis(data.kpis);
  renderCharts(data.charts);
}
```

```js
// 追加到 static/js/dashboard.js
function renderCharts(charts) {
  const container = document.getElementById("chart-section");
  container.innerHTML = "";
  charts.forEach((chart) => {
    const wrapper = document.createElement("div");
    wrapper.className = "chart-card";

    const titleEl = document.createElement("div");
    titleEl.className = "chart-title";
    titleEl.textContent = chart.title || chart.key;
    wrapper.appendChild(titleEl);

    if (chart.error) {
      const errorEl = document.createElement("div");
      errorEl.className = "chart-error";
      errorEl.textContent = `数据异常：${chart.error}`;
      wrapper.appendChild(errorEl);
      container.appendChild(wrapper);
      return;
    }

    const chartEl = document.createElement("div");
    chartEl.className = "chart-canvas";
    chartEl.id = `chart-${chart.key}`;
    wrapper.appendChild(chartEl);
    container.appendChild(wrapper);

    const instance = echarts.init(chartEl);
    instance.setOption(buildChartOption(chart));
    chartInstances[chart.key] = instance;
  });
}

function buildChartOption(chart) {
  if (chart.type === "pie") {
    return {
      title: { text: chart.title, left: "center" },
      tooltip: { trigger: "item" },
      series: [
        {
          type: "pie",
          radius: "60%",
          data: chart.x.map((name, i) => ({ name, value: chart.y[i] })),
        },
      ],
    };
  }
  return {
    tooltip: { trigger: "axis" },
    xAxis: { type: "category", data: chart.x },
    yAxis: { type: "value" },
    series: [{ type: chart.type === "bar" ? "bar" : "line", data: chart.y }],
  };
}

window.addEventListener("resize", () => {
  Object.values(chartInstances).forEach((instance) => instance.resize());
});
```

```css
/* 追加到 static/css/dashboard.css */
.chart-card {
  background-color: #ffffff;
  border-radius: 8px;
  padding: 16px;
  box-shadow: 0 1px 2px rgba(0, 0, 0, 0.06);
}

.chart-title {
  font-size: 14px;
  font-weight: 600;
  margin-bottom: 8px;
}

.chart-canvas {
  width: 100%;
  height: 320px;
}

.chart-error {
  font-size: 13px;
  color: #d93026;
}
```

- [ ] **Step 4: 运行测试确认通过**

Run: `pytest tests/test_app.py -v`
Expected: PASS（18 passed）

- [ ] **Step 5: 手动浏览器验证**

```bash
source venv/bin/activate
UPLOAD_PASSWORD=test123 python app.py
```

在 `config.yaml` 的 `charts` 中确认有一项 `type: line`（如脚手架自带的 `monthly_sales_trend`），上传对应的 `销售明细.xlsx`（`月度` sheet 含"月份"/"销售额"两列）。打开 `http://localhost:5000/dashboard`：
1. 确认折线图正确显示月份为横轴、销售额为纵轴的趋势线。
2. 缩放浏览器窗口，确认图表跟随容器宽度自适应重绘，不出现压扁或溢出。

- [ ] **Step 6: 提交**

```bash
git add templates/dashboard.html static/css/dashboard.css static/js/dashboard.js tests/test_app.py
git commit -m "feat: render charts with ECharts and resize handling"
```

---

### Task 12: 明细表格渲染（响应式：小屏横向滚动）

**Files:**
- Modify: `static/css/dashboard.css`
- Modify: `static/js/dashboard.js`
- Modify: `tests/test_app.py`

**Interfaces:**
- Consumes: `data.tables` 中每一项 `{'key','title','columns':[...],'rows':[[...]],'error','view_group'?,'view_label'?}`（Task 5/5.1/6）；`#table-section` 容器（Task 10）
- Produces: `dashboard.js` 中的 `renderTables(tables)`、`renderTableGroup(group)`、`buildTableBody(table)`。多个 `tables` 条目共享同一个 `view_group` 时，`renderTables` 会把它们归并为一组传给同一次 `renderTableGroup` 调用，渲染成一张卡片 + tab 切换，默认显示分组中第一项；没有 `view_group` 的条目各自单独成组（长度为 1 的分组不渲染 tab）

- [ ] **Step 1: 写失败的测试**

```python
# 追加到 tests/test_app.py
def test_dashboard_js_includes_table_rendering(client):
    response = client.get("/static/js/dashboard.js")
    body = response.data.decode("utf-8")
    assert "renderTables" in body


def test_dashboard_js_includes_table_view_group_tabs(client):
    response = client.get("/static/js/dashboard.js")
    body = response.data.decode("utf-8")
    assert "renderTableGroup" in body
    assert "view_group" in body
```

- [ ] **Step 2: 运行测试确认失败**

Run: `pytest tests/test_app.py -v`
Expected: FAIL，`static/js/dashboard.js` 中还没有 `renderTables`/`renderTableGroup`

- [ ] **Step 3: 写最小实现**

```js
// 替换 static/js/dashboard.js 中的 loadDashboardData
async function loadDashboardData() {
  const response = await fetch("/api/data");
  const data = await response.json();
  renderKpis(data.kpis);
  renderCharts(data.charts);
  renderTables(data.tables);
}
```

```js
// 追加到 static/js/dashboard.js
function renderTables(tables) {
  const container = document.getElementById("table-section");
  container.innerHTML = "";

  const groups = [];
  const groupIndexByKey = {};
  tables.forEach((table) => {
    if (table.view_group) {
      if (!(table.view_group in groupIndexByKey)) {
        groupIndexByKey[table.view_group] = groups.length;
        groups.push([]);
      }
      groups[groupIndexByKey[table.view_group]].push(table);
    } else {
      groups.push([table]);
    }
  });

  groups.forEach((group) => {
    container.appendChild(renderTableGroup(group));
  });
}

function renderTableGroup(group) {
  const wrapper = document.createElement("div");
  wrapper.className = "table-card";

  const titleEl = document.createElement("div");
  titleEl.className = "table-title";
  titleEl.textContent = group[0].title || group[0].key;
  wrapper.appendChild(titleEl);

  const body = document.createElement("div");

  if (group.length > 1) {
    const tabsEl = document.createElement("div");
    tabsEl.className = "table-tabs";
    group.forEach((table, index) => {
      const tabEl = document.createElement("button");
      tabEl.type = "button";
      tabEl.className = "table-tab" + (index === 0 ? " active" : "");
      tabEl.textContent = table.view_label || table.title || table.key;
      tabEl.addEventListener("click", () => {
        tabsEl.querySelectorAll(".table-tab").forEach((el) => el.classList.remove("active"));
        tabEl.classList.add("active");
        body.innerHTML = "";
        body.appendChild(buildTableBody(table));
      });
      tabsEl.appendChild(tabEl);
    });
    wrapper.appendChild(tabsEl);
  }

  body.appendChild(buildTableBody(group[0]));
  wrapper.appendChild(body);
  return wrapper;
}

function buildTableBody(table) {
  if (table.error) {
    const errorEl = document.createElement("div");
    errorEl.className = "table-error";
    errorEl.textContent = `数据异常：${table.error}`;
    return errorEl;
  }

  const scroll = document.createElement("div");
  scroll.className = "table-scroll";

  const tableEl = document.createElement("table");
  tableEl.className = "data-table";

  const thead = document.createElement("thead");
  const headRow = document.createElement("tr");
  table.columns.forEach((col) => {
    const th = document.createElement("th");
    th.textContent = col;
    headRow.appendChild(th);
  });
  thead.appendChild(headRow);
  tableEl.appendChild(thead);

  const tbody = document.createElement("tbody");
  table.rows.forEach((row) => {
    const tr = document.createElement("tr");
    row.forEach((cell) => {
      const td = document.createElement("td");
      td.textContent = cell === null || cell === undefined ? "" : cell;
      tr.appendChild(td);
    });
    tbody.appendChild(tr);
  });
  tableEl.appendChild(tbody);

  scroll.appendChild(tableEl);
  return scroll;
}
```

```css
/* 追加到 static/css/dashboard.css */
.table-card {
  background-color: #ffffff;
  border-radius: 8px;
  padding: 16px;
  box-shadow: 0 1px 2px rgba(0, 0, 0, 0.06);
}

.table-title {
  font-size: 14px;
  font-weight: 600;
  margin-bottom: 8px;
}

.table-scroll {
  overflow-x: auto;
}

.data-table {
  width: 100%;
  border-collapse: collapse;
  white-space: nowrap;
}

.data-table th,
.data-table td {
  padding: 8px 12px;
  border-bottom: 1px solid #e5e6eb;
  text-align: left;
  font-size: 13px;
}

.data-table th {
  color: #6b7280;
  font-weight: 600;
}

.table-error {
  font-size: 13px;
  color: #d93026;
}

.table-tabs {
  display: flex;
  gap: 8px;
  margin-bottom: 12px;
  border-bottom: 1px solid #e5e6eb;
}

.table-tab {
  appearance: none;
  border: none;
  background: none;
  padding: 6px 4px;
  font-size: 13px;
  color: #6b7280;
  cursor: pointer;
  border-bottom: 2px solid transparent;
}

.table-tab.active {
  color: #1f2937;
  font-weight: 600;
  border-bottom-color: #2f54eb;
}
```

- [ ] **Step 4: 运行测试确认通过**

Run: `pytest tests/test_app.py -v`
Expected: PASS（20 passed）

- [ ] **Step 5: 手动浏览器验证**

```bash
source venv/bin/activate
UPLOAD_PASSWORD=test123 python app.py
```

上传对应 `销售明细.xlsx`（`明细` sheet 含"日期"/"客户"/"产品"/"金额"列）。打开 `http://localhost:5000/dashboard`：
1. 确认明细表格正确显示表头与数据行。
2. 缩小浏览器窗口到手机尺寸（约 375px），确认表格区域出现横向滚动条、可以左右滑动查看被遮挡的列，页面本身不出现整体横向滚动。

- [ ] **Step 6: 提交**

```bash
git add static/css/dashboard.css static/js/dashboard.js tests/test_app.py
git commit -m "feat: render detail tables with horizontal-scroll responsiveness"
```

---

### Task 13: README（部署与使用说明）

**Files:**
- Create: `README.md`

**Interfaces:**
- 不产出代码接口；文档需与 `app.py`/`config.yaml`/`requirements.txt`（Task 1、7、8、9）的实际行为一致

- [ ] **Step 1: 编写 README.md**

```markdown
# 老板看板（Boss Dashboard）

数据维护人通过密码保护的上传页面提交 Excel 文件，系统按 `config.yaml` 配置解析生成看板数据，供老板通过公开链接查看 KPI、图表与明细表格。

## 环境要求

- Python 3.9+

## 安装

```bash
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt
```

## 配置

1. 设置数据维护人登录密码（默认是 `changeme`，务必修改）：

   ```bash
   export UPLOAD_PASSWORD="你的密码"
   ```

2. 设置 Flask session 密钥（默认是开发用的固定字符串，生产环境务必修改）：

   ```bash
   export FLASK_SECRET_KEY="一个随机字符串"
   ```

3. 编辑 `config.yaml`，为每个 KPI/图表/表格指定数据来源文件名（`source_file`，需与上传时的 Excel 文件名完全一致，支持中文）、sheet 名，以及取数方式：
   - `mode: fixed_range` — 指定固定单元格范围（如 `B2`），适合位置稳定的单一数值。
   - `mode: header_match` — 指定表头文字（如 `header: "利润"`），程序自动定位对应列，能容忍列顺序变化；用在 KPI 上时会对该列所有数值求和。

## 运行

```bash
source venv/bin/activate
python app.py
```

默认监听 `http://localhost:5000`。

## 使用流程

1. 数据维护人打开 `http://localhost:5000/upload`，输入 `UPLOAD_PASSWORD` 登录。
2. 选择一个或多个本地 Excel 文件上传（文件名需与 `config.yaml` 中的 `source_file` 一致）。上传后会覆盖 `uploads/` 目录下的同名文件，并立即基于**整个** `uploads/` 目录重新生成 `data.json`。
3. 页面会提示成功或具体的失败原因（如"找不到表头 '销售额'"），失败的单个指标不影响其余指标正常生成。
4. 老板或任何人打开 `http://localhost:5000/dashboard` 即可查看最新数据，无需密码。

## 运行测试

```bash
source venv/bin/activate
pytest -v
```

## 部署说明

当前架构本地运行与服务器部署无需改动核心代码——只需将 `python app.py` 换成生产级 WSGI 服务器（如 `gunicorn app:create_app()`），并确保 `UPLOAD_PASSWORD`/`FLASK_SECRET_KEY` 通过环境变量正确设置。
```

- [ ] **Step 2: 手动验证**

在一个干净的 shell 中，从零按 README 的步骤操作一遍（创建虚拟环境、安装依赖、设置环境变量、启动应用、上传测试文件、访问看板），确认每一步命令都能按文档描述成功执行，看板能正确显示数据。

- [ ] **Step 3: 提交**

```bash
git add README.md
git commit -m "docs: add setup and usage instructions"
```

---

## Self-Review

**Spec 覆盖检查：**

| 设计文档条目 | 对应任务 |
|---|---|
| `/upload` 密码保护，GET 返回表单、POST 接收并解析 | Task 8（GET+鉴权）、Task 9（POST 处理） |
| `/dashboard` 只读、无密码 | Task 7 |
| `/api/data` 返回 `data.json` 内容 | Task 7 |
| 支持单个/多个文件一次性上传 | Task 9（`request.files.getlist("files")` + `multiple` 属性 + 多文件测试） |
| 同名文件覆盖、不存在则新建 | Task 9 |
| 解析读取整个 `uploads/` 目录而非仅本次上传文件 | Task 6（`build_dashboard_data` 始终基于传入的 `uploads_dir` 全量重算）+ Task 9 调用方式 |
| `config.yaml` 支持 `fixed_range`/`header_match`/`computed`/`group_by_sum` 四种取数方式 | Task 2（fixed_range）、Task 3（header_match）、Task 5.1（group_by_sum）、Task 6.1（computed） |
| 单项取数失败不影响其余项，报清晰错误 | Task 2-6、Task 5.1、Task 6.1（每个 `extract_*`/`_resolve_computed_kpi` 独立 try/except 或显式错误分支，`build_dashboard_data` 逐项调用） |
| 库存账实差异等衍生 KPI 直接展示在首屏 KPI 区 | Task 6.1（`computed` 结果并入 `data["kpis"]`，与其他 KPI 同数组）+ Task 10（KPI 卡片渲染无需区分来源） |
| 表格支持多视图切换（如出库明细列表 / 按收货人排名） | Task 5.1（`group_by_sum` 提供排名视图数据）+ Task 12（`view_group`/`view_label` 归并渲染 tab，默认首项） |
| `data.json` 单文件、覆盖式落盘 | Task 6 |
| ECharts 渲染折线/柱状/饼图 | Task 11 |
| KPI 卡片、图表区、明细表格区 | Task 10（骨架）、Task 10（KPI）、Task 11（图表）、Task 12（表格） |
| 响应式多端适配（大屏多列、小屏单列堆叠） | Task 10（CSS Grid + 媒体查询断点） |
| 图表容器随尺寸变化自适应重绘 | Task 11（`resize` 监听 + `chartInstances`） |
| 明细表格小屏适配（横向可滚动） | Task 12（`.table-scroll { overflow-x: auto }`） |
| 文件名不可用 `secure_filename`，需保留中文原名 | Task 9（`_is_safe_filename` 手写校验，写入 Global Constraints） |
| Excel 读取需 `data_only=True` | Task 2（`_load_worksheet`），写入 Global Constraints |
| 部署环境决定可推迟 | README 部署说明段落 + Global Constraints/Architecture 中未绑定具体部署方式 |
| 不做历史趋势、不做在线文档接入、不做多用户账号、不做自动定时拉取（v1 范围之外） | 全计划均未实现这些功能，符合"明确不做" |

无遗漏项。

**Placeholder 扫描：** 全文搜索未发现 "TBD"/"TODO"/"implement later" 等占位符；所有代码块均为可直接使用的完整实现，无"参考 Task N"式的省略。

**类型/签名一致性检查：**
- `extract_kpi`/`extract_chart`/`extract_table` 在 Task 2-6 中签名与返回字段（`key`/`label`/`value`/`error`，`key`/`type`/`title`/`x`/`y`/`error`，`key`/`title`/`columns`/`rows`/`error`）全程一致，`build_dashboard_data`（Task 6）、`app.py` 的 `/upload`（Task 9）、`dashboard.js` 的 `renderKpis`/`renderCharts`/`renderTables`（Task 10-12）均按同一字段名读取，未出现改名。
- `create_app(test_config=None)` 签名（Task 7）在 Task 8/9 中原样复用，`test_config` 覆盖的 key（`UPLOAD_PASSWORD`/`SECRET_KEY`/`UPLOAD_DIR`/`CONFIG_PATH`/`DATA_PATH`）全程一致。
- `login_required`（Task 8）在 Task 9 的 `/upload` 视图上原样复用，未重复定义或改名。
- Task 9 中 `app.py` 顶部的 `extractor` import 语句在 Task 9 Step 3 中做了完整替换（而非追加），避免与 Task 7 的旧 import 行重复。
- `extract_table` 在 Task 5 引入 `view_group`/`view_label` 透传字段后，Task 5.1 整体重写该函数时延续了这两个字段（否则 Task 12 的 `renderTables` 会拿不到分组依据）；`build_dashboard_data`（Task 6）在 Task 6.1 改为两遍扫描后，仍保持原有 `{'kpis', 'charts', 'tables'}` 返回结构与 `kpis` 数组的 config 声明顺序，未破坏 Task 7/9 对该函数的调用方式。
- extractor 侧测试计数链路（`test_extractor.py`）：Task 1（2）→ Task 2（5）→ Task 3（7）→ Task 4（9）→ Task 5（11）→ Task 5.1（13）→ Task 6（15）→ Task 6.1（17），与 app 侧测试计数链路（`test_app.py`）：Task 7（4）→ Task 8（9）→ Task 9（14）→ Task 10（17）→ Task 11（18）→ Task 12（20）均为累计连续递增，两条链路彼此独立、互不干扰。

---

Plan complete and saved to `docs/superpowers/plans/2026-08-20-boss-dashboard.md`. Two execution options:

**1. Subagent-Driven (recommended)** — I dispatch a fresh subagent per task, review between tasks, fast iteration

**2. Inline Execution** — Execute tasks in this session using executing-plans, batch execution with checkpoints

**Which approach？**
