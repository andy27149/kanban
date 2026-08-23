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
