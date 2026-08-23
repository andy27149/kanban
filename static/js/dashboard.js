const chartInstances = {};

let chartResizeObserver = null;
if ("ResizeObserver" in window) {
  chartResizeObserver = new ResizeObserver(() => {
    Object.values(chartInstances).forEach((instance) => instance.resize());
  });
}

function palette() {
  return {
    text: '#2b2620',
    muted: '#8a8071',
    line: '#e6dcc4',
    gold: '#9c7a25',
    goldDeep: '#7a5c1c',
    goldLight: '#c9a94e',
  };
}

function kFormat(v) {
  if (Math.abs(v) >= 1000) return (v / 1000).toFixed(v % 1000 === 0 ? 0 : 1) + 'k';
  return String(v);
}

async function loadDashboardData() {
  try {
    const response = await fetch("/api/data");
    if (!response.ok) {
      throw new Error(`HTTP ${response.status}: ${response.statusText}`);
    }
    const data = await response.json();
    renderKpis(data.kpis);
    renderCharts(data.charts);
    document.getElementById("last-sync").textContent =
      "LAST SYNC " + new Date().toLocaleString("zh-CN", { hour12: false });
  } catch (error) {
    console.error("Failed to load dashboard data:", error);
    const container = document.getElementById("kpi-section");
    container.innerHTML = "";
    const errorCard = document.createElement("div");
    errorCard.className = "kpi-error";
    errorCard.textContent = "数据加载失败：" + error.message;
    container.appendChild(errorCard);
  }
}

function renderKpis(kpis) {
  const container = document.getElementById("kpi-section");
  container.innerHTML = "";
  kpis.forEach((kpi, index) => {
    const card = document.createElement("div");
    card.className = "kpi-card" + (index === 0 ? " hero" : "");
    if (kpi.error) {
      const labelEl = document.createElement("div");
      labelEl.className = "kpi-label";
      labelEl.textContent = kpi.label;
      card.appendChild(labelEl);

      const errorEl = document.createElement("div");
      errorEl.className = "kpi-error";
      errorEl.textContent = "数据异常：" + kpi.error;
      card.appendChild(errorEl);
    } else {
      const labelEl = document.createElement("div");
      labelEl.className = "kpi-label";
      labelEl.textContent = kpi.label;
      card.appendChild(labelEl);

      const valueEl = document.createElement("div");
      valueEl.className = "kpi-value";
      valueEl.textContent = kpi.value;
      card.appendChild(valueEl);
    }
    container.appendChild(card);
  });
}

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
    instance.setOption(buildChartOption(chart, chartEl));
    chartInstances[chart.key] = instance;
    if (chartResizeObserver) chartResizeObserver.observe(chartEl);
  });
}

function buildChartOption(chart, chartEl) {
  const p = palette();

  if (chart.type === "pie") {
    return {
      textStyle: { color: p.text, fontFamily: 'Inter, sans-serif' },
      tooltip: { trigger: "item" },
      legend: { bottom: 0, textStyle: { color: p.muted, fontSize: 12 } },
      color: [p.gold, '#8c3a2e', '#5b6b7a', p.goldLight],
      series: [
        {
          type: "pie",
          radius: ["42%", "66%"],
          center: ["50%", "46%"],
          itemStyle: { borderColor: "#ffffff", borderWidth: 2 },
          label: { color: p.text },
          data: chart.x.map((name, i) => ({ name, value: chart.y[i] })),
        },
      ],
    };
  }

  const isBar = chart.type === "bar";
  const rotate = chartEl && chartEl.clientWidth < 340 ? 28 : 0;

  return {
    textStyle: { color: p.text, fontFamily: 'Inter, sans-serif' },
    grid: { left: 44, right: 16, top: 20, bottom: 28 },
    tooltip: { trigger: "axis" },
    xAxis: {
      type: "category",
      data: chart.x,
      axisLine: { lineStyle: { color: p.line } },
      axisLabel: { color: p.muted, interval: 0, rotate, fontSize: 11 },
    },
    yAxis: {
      type: "value",
      splitLine: { lineStyle: { color: p.line } },
      axisLabel: { color: p.muted, formatter: kFormat },
    },
    series: [
      isBar
        ? {
            type: "bar",
            data: chart.y,
            barWidth: "46%",
            itemStyle: {
              borderRadius: [4, 4, 0, 0],
              color: {
                type: "linear", x: 0, y: 0, x2: 0, y2: 1,
                colorStops: [
                  { offset: 0, color: p.goldDeep },
                  { offset: 1, color: "rgba(122, 92, 28, 0.35)" },
                ],
              },
            },
          }
        : {
            type: "line",
            data: chart.y,
            smooth: true,
            lineStyle: { color: p.gold, width: 2.5 },
            itemStyle: { color: p.gold },
            areaStyle: {
              color: {
                type: "linear", x: 0, y: 0, x2: 0, y2: 1,
                colorStops: [
                  { offset: 0, color: "rgba(156, 122, 37, 0.28)" },
                  { offset: 1, color: "rgba(156, 122, 37, 0)" },
                ],
              },
            },
          },
    ],
  };
}

window.addEventListener("resize", () => {
  Object.values(chartInstances).forEach((instance) => instance.resize());
});

document.addEventListener("DOMContentLoaded", loadDashboardData);
