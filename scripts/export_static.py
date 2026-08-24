"""Bundle the live dashboard into a single self-contained HTML file.

The output embeds the current data.json snapshot, dashboard.css,
dashboard.js, and the ECharts library inline, so the recipient can open
it directly (file://) with no server and no internet connection required.

Usage:
    python scripts/export_static.py [output_path]

Re-run this any time data.json is refreshed (i.e. after a new upload) to
produce an up-to-date snapshot to send out.
"""
import json
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent


def build_static_html():
    data_path = ROOT / "data.json"
    if not data_path.exists():
        raise SystemExit("data.json 不存在，请先通过 /upload 上传一次 Excel 生成数据")

    data = json.loads(data_path.read_text(encoding="utf-8"))
    css = (ROOT / "static" / "css" / "dashboard.css").read_text(encoding="utf-8")
    dashboard_js = (ROOT / "static" / "js" / "dashboard.js").read_text(encoding="utf-8")
    echarts_js = (ROOT / "vendor" / "echarts.min.js").read_text(encoding="utf-8")
    data_json_str = json.dumps(data, ensure_ascii=False)

    return f"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>经营看板</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Sora:wght@400;600;700;800&family=Inter:wght@400;500;600&family=JetBrains+Mono:wght@400;500;600;700&display=swap" rel="stylesheet">
<style>
{css}
</style>
<script>
{echarts_js}
</script>
</head>
<body>
<header class="dashboard-header">
  <div>
    <p class="header-eyebrow"><span class="pulse-dot"></span>经营数据快照</p>
    <h1>经营看板</h1>
  </div>
  <span class="updated" id="last-sync"></span>
</header>
<main id="app">
  <p class="section-label"><span class="lead-glyph">▍</span>核心指标</p>
  <section id="kpi-section" class="kpi-grid" aria-label="KPI 指标"></section>
  <p class="section-label"><span class="lead-glyph">▍</span>图表</p>
  <section id="chart-section" class="chart-grid" aria-label="图表"></section>
  <p class="section-label"><span class="lead-glyph">▍</span>明细表格</p>
  <section id="table-section" class="table-grid" aria-label="明细表格"></section>
</main>
<script>
window.__DASHBOARD_DATA__ = {data_json_str};
</script>
<script>
{dashboard_js}
</script>
</body>
</html>
"""


def main():
    output_path = Path(sys.argv[1]) if len(sys.argv) > 1 else ROOT / "exports" / "经营看板.html"
    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text(build_static_html(), encoding="utf-8")
    print(f"已生成静态看板文件: {output_path}")


if __name__ == "__main__":
    main()
