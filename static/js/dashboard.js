const chartInstances = {};
let dashboardData = null;

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

function formatNumber(v) {
  if (typeof v !== "number") return v;
  return Number.isInteger(v) ? String(v) : v.toFixed(2);
}

async function loadDashboardData() {
  try {
    let data;
    if (window.__DASHBOARD_DATA__) {
      data = window.__DASHBOARD_DATA__;
    } else {
      const response = await fetch("api/data");
      if (!response.ok) {
        throw new Error(`HTTP ${response.status}: ${response.statusText}`);
      }
      data = await response.json();
    }
    dashboardData = data;
    renderKpis(data.kpis);
    renderCharts(data.charts);
    renderTables(data.tables);
    document.getElementById("last-sync").textContent =
      "数据更新时间 " + (data.generated_at || "未知");
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
  let lastGroup = null;
  kpis.forEach((kpi, index) => {
    if (kpi.group && kpi.group !== lastGroup) {
      const groupEl = document.createElement("div");
      groupEl.className = "kpi-group-label";
      groupEl.innerHTML =
        `<span class="kpi-group-icon">${kpi.group_icon || ""}</span>` +
        `<span>${kpi.group_label || kpi.group}</span>`;
      container.appendChild(groupEl);
    }
    lastGroup = kpi.group || null;

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
      const isEmpty = kpi.value === null || kpi.value === undefined;
      valueEl.className = "kpi-value" + (isEmpty ? " kpi-empty" : "");
      valueEl.textContent = isEmpty ? "暂无数据" : formatNumber(kpi.value);
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
    wrapper.className = "chart-card" + (chart.key === "category_in_out" ? " chart-card--wide" : "");

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

    if (chart.hint) {
      const hintEl = document.createElement("div");
      hintEl.className = "chart-hint";
      hintEl.textContent = chart.hint;
      wrapper.appendChild(hintEl);
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
      tooltip: { trigger: "item", valueFormatter: formatNumber },
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
  const pxPerLabel = chartEl && chart.x.length ? chartEl.clientWidth / chart.x.length : Infinity;
  const rotate = pxPerLabel < 60 ? 45 : pxPerLabel < 80 ? 28 : 0;
  const seriesColors = [p.goldDeep, '#5b6b7a', '#8c3a2e', p.goldLight];

  const labeledBarCharts = ["inventory_balance_by_category", "purchase_volume_by_region"];

  const series = chart.series
    ? chart.series.map((s, i) => ({
        name: s.name,
        type: "bar",
        data: s.data,
        barGap: "20%",
        label: chart.key === "category_in_out"
          ? {
              show: true,
              position: "top",
              offset: [0, i === 0 ? -14 : 0],
              color: p.text,
              fontSize: 10,
              formatter: (params) => formatNumber(params.value),
            }
          : undefined,
        itemStyle: {
          borderRadius: [4, 4, 0, 0],
          color: seriesColors[i % seriesColors.length],
        },
      }))
    : [
        isBar
          ? {
              type: "bar",
              data: chart.y,
              barWidth: "46%",
              label: labeledBarCharts.includes(chart.key)
                ? { show: true, position: "top", color: p.text, fontSize: 11, formatter: (params) => formatNumber(params.value) }
                : undefined,
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
      ];

  return {
    textStyle: { color: p.text, fontFamily: 'Inter, sans-serif' },
    grid: { left: 8, right: 16, top: chart.series ? 36 : 20, bottom: rotate > 0 ? 48 : 28, containLabel: true },
    tooltip: { trigger: "axis", valueFormatter: formatNumber },
    legend: chart.series
      ? { top: 0, textStyle: { color: p.muted, fontSize: 12 } }
      : undefined,
    xAxis: {
      type: "category",
      data: chart.x,
      axisLine: { lineStyle: { color: p.line } },
      axisLabel: { color: p.muted, interval: 0, rotate, fontSize: 11 },
    },
    yAxis: {
      type: "value",
      splitLine: { lineStyle: { color: p.line } },
      axisLabel: { color: p.muted, formatter: formatNumber },
    },
    series,
  };
}

function formatCell(cell) {
  if (cell === null || cell === undefined) return "";
  if (typeof cell === "number" && !Number.isInteger(cell)) return cell.toFixed(2);
  return cell;
}

function ensureModal() {
  let modal = document.getElementById("drill-modal");
  if (modal) return modal;
  modal = document.createElement("div");
  modal.id = "drill-modal";
  modal.className = "modal-overlay hidden";
  modal.innerHTML =
    '<div class="modal-box">' +
    '<div class="modal-header"><span class="modal-title"></span><button type="button" class="modal-close" aria-label="关闭">&times;</button></div>' +
    '<div class="modal-body"></div>' +
    "</div>";
  modal.querySelector(".modal-close").addEventListener("click", closeModal);
  modal.addEventListener("click", (event) => {
    if (event.target === modal) closeModal();
  });
  document.body.appendChild(modal);
  return modal;
}

function closeModal() {
  const modal = document.getElementById("drill-modal");
  if (modal) modal.classList.add("hidden");
}

function openModal(title, bodyEl) {
  const modal = ensureModal();
  modal.querySelector(".modal-title").textContent = title;
  const body = modal.querySelector(".modal-body");
  body.innerHTML = "";
  body.appendChild(bodyEl);
  modal.classList.remove("hidden");
}

function buildSimpleTable(columns, rows) {
  const table = document.createElement("table");
  table.className = "data-table";
  const thead = document.createElement("thead");
  const headRow = document.createElement("tr");
  columns.forEach((col) => {
    const th = document.createElement("th");
    th.textContent = col;
    headRow.appendChild(th);
  });
  thead.appendChild(headRow);
  table.appendChild(thead);
  const tbody = document.createElement("tbody");
  rows.forEach((row) => {
    const tr = document.createElement("tr");
    row.forEach((cell) => {
      const td = document.createElement("td");
      td.textContent = formatCell(cell);
      tr.appendChild(td);
    });
    tbody.appendChild(tr);
  });
  table.appendChild(tbody);
  return table;
}

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

function buildDataRow(table, row, rowIndex) {
  const tr = document.createElement("tr");

  const rowExtra = table.key === "outbound_detail" && Array.isArray(table.row_extra)
    ? table.row_extra[rowIndex] || []
    : null;
  if (rowExtra && rowExtra.length > 0) {
    tr.classList.add("row-clickable");
    tr.addEventListener("click", () => {
      openModal(
        "配煤品类明细",
        buildSimpleTable(["品类/指标", "数值（吨）"], rowExtra.map((e) => [e.label, e.value]))
      );
    });
  }

  const categorySources = table.key === "category_summary" && dashboardData && dashboardData.category_sources
    ? dashboardData.category_sources[row[0]]
    : null;
  if (categorySources) {
    tr.classList.add("row-clickable");
    tr.addEventListener("click", () => {
      const wrap = document.createElement("div");
      categorySources.forEach((segment) => {
        const heading = document.createElement("div");
        heading.className = "modal-subtitle";
        heading.textContent = segment.sheet;
        wrap.appendChild(heading);
        wrap.appendChild(buildSimpleTable(segment.headers, segment.rows));
      });
      openModal(`矿区明细：${row[0]}`, wrap);
    });
  }

  row.forEach((cell) => {
    const td = document.createElement("td");
    td.textContent = formatCell(cell);
    tr.appendChild(td);
  });
  return tr;
}

function buildTableBody(table) {
  if (table.error) {
    const errorEl = document.createElement("div");
    errorEl.className = "table-error";
    errorEl.textContent = `数据异常：${table.error}`;
    return errorEl;
  }

  const wrapper = document.createElement("div");

  const statusIndex = table.status_column ? table.columns.indexOf(table.status_column) : -1;
  const scroll = document.createElement("div");
  scroll.className = "table-scroll";

  function renderRows(activeStatus) {
    scroll.innerHTML = "";

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
    table.rows.forEach((row, rowIndex) => {
      if (activeStatus !== null && row[statusIndex] !== activeStatus) return;
      tbody.appendChild(buildDataRow(table, row, rowIndex));
    });

    if (table.totals) {
      const totalsRow = document.createElement("tr");
      totalsRow.className = "totals-row";
      table.totals.forEach((cell) => {
        const td = document.createElement("td");
        td.textContent = formatCell(cell);
        totalsRow.appendChild(td);
      });
      tbody.appendChild(totalsRow);
    }

    tableEl.appendChild(tbody);
    scroll.appendChild(tableEl);
  }

  if (statusIndex !== -1) {
    const counts = {};
    table.rows.forEach((row) => {
      const value = row[statusIndex];
      if (value === null || value === undefined || value === "") return;
      counts[value] = (counts[value] || 0) + 1;
    });

    let activeStatus = null;
    const badgesEl = document.createElement("div");
    badgesEl.className = "status-badges";

    const allBadge = document.createElement("button");
    allBadge.type = "button";
    allBadge.className = "status-badge active";
    allBadge.textContent = `全部（${table.rows.length}）`;
    allBadge.addEventListener("click", () => {
      activeStatus = null;
      badgesEl.querySelectorAll(".status-badge").forEach((el) => el.classList.remove("active"));
      allBadge.classList.add("active");
      renderRows(activeStatus);
    });
    badgesEl.appendChild(allBadge);

    Object.keys(counts).forEach((status) => {
      const badge = document.createElement("button");
      badge.type = "button";
      badge.className = "status-badge";
      badge.textContent = `${status}（${counts[status]}）`;
      badge.addEventListener("click", () => {
        activeStatus = activeStatus === status ? null : status;
        badgesEl.querySelectorAll(".status-badge").forEach((el) => el.classList.remove("active"));
        if (activeStatus === status) {
          badge.classList.add("active");
        } else {
          allBadge.classList.add("active");
        }
        renderRows(activeStatus);
      });
      badgesEl.appendChild(badge);
    });

    wrapper.appendChild(badgesEl);
  }

  wrapper.appendChild(scroll);
  renderRows(null);
  return wrapper;
}

window.addEventListener("resize", () => {
  Object.values(chartInstances).forEach((instance) => instance.resize());
});

document.addEventListener("DOMContentLoaded", loadDashboardData);
