from pathlib import Path

from extractor import build_dashboard_data, extract_chart, extract_table, load_config


def test_extract_table_sheet_table_reads_rows_and_skips_blank_id(uploads_dir, make_workbook):
    make_workbook(
        "上下游结算.xlsx",
        {
            "上游结算": [
                ["采购管理表"],
                ["合同编号", "矿区", "结算数量"],
                ["HT-001", "A矿区", 100],
                [None, "残留行", 999],
                ["HT-002", "B矿区", 200],
            ]
        },
    )
    item = {
        "key": "purchase_contracts",
        "title": "采购管理表",
        "source_file": "上下游结算.xlsx",
        "sheet": "上游结算",
        "mode": "sheet_table",
        "header_row": 2,
        "id_column": "合同编号",
    }

    result = extract_table(item, uploads_dir)

    assert result["error"] is None
    assert result["columns"] == ["合同编号", "矿区", "结算数量"]
    assert result["rows"] == [["HT-001", "A矿区", 100], ["HT-002", "B矿区", 200]]
    assert result["totals"] is None


def test_extract_table_sheet_table_missing_id_column_returns_error(uploads_dir, make_workbook):
    make_workbook(
        "上下游结算.xlsx",
        {"上游结算": [["采购管理表"], ["矿区", "结算数量"], ["A矿区", 100]]},
    )
    item = {
        "key": "purchase_contracts",
        "title": "采购管理表",
        "source_file": "上下游结算.xlsx",
        "sheet": "上游结算",
        "mode": "sheet_table",
        "header_row": 2,
        "id_column": "合同编号",
    }

    result = extract_table(item, uploads_dir)

    assert result["error"] == "找不到列: 合同编号"
    assert result["rows"] == []


def test_extract_table_sheet_table_show_totals_sums_numeric_columns(uploads_dir, make_workbook):
    make_workbook(
        "上下游结算.xlsx",
        {
            "上游结算": [
                ["采购管理表"],
                ["合同编号", "矿区", "结算数量"],
                ["HT-001", "A矿区", 100],
                ["HT-002", "B矿区", 200],
            ]
        },
    )
    item = {
        "key": "purchase_contracts",
        "title": "采购管理表",
        "source_file": "上下游结算.xlsx",
        "sheet": "上游结算",
        "mode": "sheet_table",
        "header_row": 2,
        "id_column": "合同编号",
        "show_totals": True,
    }

    result = extract_table(item, uploads_dir)

    assert result["totals"] == ["合计（2 份合同）", None, 300]


def test_extract_chart_group_by_sum_aggregates_numeric_column(uploads_dir, make_workbook):
    make_workbook(
        "上下游结算.xlsx",
        {
            "上游结算": [
                ["采购管理表"],
                ["合同编号", "矿区", "结算数量"],
                ["HT-001", "A矿区", 100],
                ["HT-002", "A矿区", 50],
                ["HT-003", "B矿区", 200],
            ]
        },
    )
    item = {
        "key": "purchase_volume_by_region",
        "type": "bar",
        "title": "采购按矿区汇总实发数量（吨）",
        "source_file": "上下游结算.xlsx",
        "sheet": "上游结算",
        "mode": "group_by",
        "agg": "sum",
        "header_row": 2,
        "id_column": "合同编号",
        "group_column": "矿区",
        "value_column": "结算数量",
    }

    result = extract_chart(item, uploads_dir)

    assert result["error"] is None
    assert result["x"] == ["A矿区", "B矿区"]
    assert result["y"] == [150, 200]


def test_extract_chart_group_by_count_counts_rows_and_skips_none_group(uploads_dir, make_workbook):
    make_workbook(
        "上下游结算.xlsx",
        {
            "上游结算": [
                ["采购管理表"],
                ["合同编号", "进度"],
                ["HT-001", "进行中"],
                ["HT-002", "进行中"],
                ["HT-003", None],
                ["HT-004", "结项"],
            ]
        },
    )
    item = {
        "key": "purchase_progress_distribution",
        "type": "pie",
        "title": "采购合同进度分布",
        "source_file": "上下游结算.xlsx",
        "sheet": "上游结算",
        "mode": "group_by",
        "agg": "count",
        "header_row": 2,
        "id_column": "合同编号",
        "group_column": "进度",
    }

    result = extract_chart(item, uploads_dir)

    assert result["error"] is None
    assert result["x"] == ["进行中", "结项"]
    assert result["y"] == [2, 1]
    assert sum(result["y"]) == 3


def test_extract_chart_sum_bars_multi_series_sums_absolute_ranges(uploads_dir, make_workbook):
    make_workbook(
        "上下游结算.xlsx",
        {
            "上游结算": [
                ["采购管理表"],
                ["合同编号", "矿区", "已付金额", "未付金额", "应退金额"],
                ["HT-001", "A矿区", 1000, 0, "无"],
                ["HT-002", "B矿区", 2000, 0, 500],
            ]
        },
    )
    item = {
        "key": "purchase_payment_status",
        "type": "bar",
        "title": "采购煤款/运费收付情况（元）",
        "source_file": "上下游结算.xlsx",
        "sheet": "上游结算",
        "mode": "sum_bars",
        "x": ["已付金额", "未付金额", "应退金额"],
        "series": [
            {"name": "煤款", "ranges": ["C3:C4", "D3:D4", "E3:E4"]},
        ],
    }

    result = extract_chart(item, uploads_dir)

    assert result["error"] is None
    assert result["x"] == ["已付金额", "未付金额", "应退金额"]
    assert result["series"] == [{"name": "煤款", "data": [3000, 0, 500]}]


def test_extract_chart_sum_bars_flat_ranges_produces_single_series_y(uploads_dir, make_workbook):
    make_workbook(
        "上下游结算.xlsx",
        {
            "下游结算": [
                ["销售管理表"],
                ["合同编号", "结算金额", "回款合计", "待回金额"],
                ["XS-001", 10000, None, None],
                ["XS-002", 20000, None, None],
            ]
        },
    )
    item = {
        "key": "sales_collection_status",
        "type": "bar",
        "title": "销售回款情况（元）",
        "source_file": "上下游结算.xlsx",
        "sheet": "下游结算",
        "mode": "sum_bars",
        "x": ["结算金额合计", "已回金额合计", "待回金额合计"],
        "ranges": ["B3:B4", "C3:C4", "D3:D4"],
    }

    result = extract_chart(item, uploads_dir)

    assert result["error"] is None
    assert result["x"] == ["结算金额合计", "已回金额合计", "待回金额合计"]
    assert result["y"] == [30000, 0, 0]


def test_build_dashboard_data_applies_hint_when_chart_is_all_zero(uploads_dir, make_workbook):
    make_workbook(
        "上下游结算.xlsx",
        {
            "上游结算": [
                ["采购管理表"],
                ["合同编号", "已付金额"],
                ["HT-001", None],
                ["HT-002", 0],
            ]
        },
    )
    config = {
        "kpis": [],
        "charts": [
            {
                "key": "purchase_payment_status",
                "type": "bar",
                "title": "采购煤款/运费收付情况（元）",
                "source_file": "上下游结算.xlsx",
                "sheet": "上游结算",
                "mode": "sum_bars",
                "x": ["已付金额"],
                "ranges": ["B3:B4"],
            }
        ],
        "tables": [],
    }

    data = build_dashboard_data(config, uploads_dir)

    assert data["charts"][0] == {
        "key": "purchase_payment_status",
        "title": "采购煤款/运费收付情况（元）",
        "hint": "该图表依赖尚未回填的业务字段，暂无数据",
    }


def test_build_dashboard_data_keeps_normal_chart_when_not_all_zero(uploads_dir, make_workbook):
    make_workbook(
        "上下游结算.xlsx",
        {
            "上游结算": [
                ["采购管理表"],
                ["合同编号", "已付金额"],
                ["HT-001", 100],
            ]
        },
    )
    config = {
        "kpis": [],
        "charts": [
            {
                "key": "purchase_payment_status",
                "type": "bar",
                "title": "采购煤款/运费收付情况（元）",
                "source_file": "上下游结算.xlsx",
                "sheet": "上游结算",
                "mode": "sum_bars",
                "x": ["已付金额"],
                "ranges": ["B3:B3"],
            }
        ],
        "tables": [],
    }

    data = build_dashboard_data(config, uploads_dir)

    assert data["charts"][0]["y"] == [100]
    assert "hint" not in data["charts"][0]


def test_build_dashboard_data_does_not_hint_a_chart_with_error(uploads_dir):
    config = {
        "kpis": [],
        "charts": [
            {
                "key": "broken_chart",
                "type": "bar",
                "title": "坏图表",
                "source_file": "不存在.xlsx",
                "sheet": "不存在",
                "mode": "sum_bars",
                "x": ["A"],
                "ranges": ["A1:A1"],
            }
        ],
        "tables": [],
    }

    data = build_dashboard_data(config, uploads_dir)

    assert data["charts"][0]["error"] is not None
    assert "hint" not in data["charts"][0]


def test_config_yaml_includes_settlement_tables_and_charts():
    config_path = Path(__file__).resolve().parent.parent / "config.yaml"
    config = load_config(config_path)

    table_keys = {t["key"] for t in config["tables"]}
    assert {"purchase_contracts", "order_management", "sales_contracts"} <= table_keys

    chart_keys = {c["key"] for c in config["charts"]}
    assert {
        "purchase_volume_by_region",
        "purchase_progress_distribution",
        "order_status_distribution",
        "sales_status_distribution",
        "purchase_payment_status",
        "sales_collection_status",
    } <= chart_keys
