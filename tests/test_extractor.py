import textwrap

from extractor import load_config, extract_kpi, extract_chart, extract_table, build_dashboard_data, save_data_json, load_data_json


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


def test_extract_kpi_corrupted_excel_file_reports_error_instead_of_raising(uploads_dir):
    (uploads_dir / "损坏文件.xlsx").write_bytes(b"this is not a real excel file, just garbage bytes")
    item = {
        "key": "total_revenue",
        "label": "总营收",
        "source_file": "损坏文件.xlsx",
        "sheet": "汇总",
        "mode": "fixed_range",
        "range": "B2",
    }

    result = extract_kpi(item, uploads_dir)

    assert result["value"] is None
    assert result["error"]


def test_extract_table_corrupted_excel_file_reports_error_instead_of_raising(uploads_dir):
    (uploads_dir / "损坏文件.xlsx").write_bytes(b"this is not a real excel file, just garbage bytes")
    item = {
        "key": "sales_detail",
        "title": "销售明细",
        "source_file": "损坏文件.xlsx",
        "sheet": "明细",
        "mode": "header_match",
        "columns": ["日期", "客户"],
    }

    result = extract_table(item, uploads_dir)

    assert result["rows"] == []
    assert result["error"]


def test_build_dashboard_data_isolates_corrupted_file_failure(uploads_dir, make_workbook):
    (uploads_dir / "损坏文件.xlsx").write_bytes(b"this is not a real excel file, just garbage bytes")
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
                "key": "broken_kpi",
                "label": "损坏指标",
                "source_file": "损坏文件.xlsx",
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
    assert data["kpis"][1]["error"]


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
