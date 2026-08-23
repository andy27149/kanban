async function loadDashboardData() {
  const response = await fetch("/api/data");
  const data = await response.json();
  renderKpis(data.kpis);
  document.getElementById("last-sync").textContent =
    "LAST SYNC " + new Date().toLocaleString("zh-CN", { hour12: false });
}

function renderKpis(kpis) {
  const container = document.getElementById("kpi-section");
  container.innerHTML = "";
  kpis.forEach((kpi, index) => {
    const card = document.createElement("div");
    card.className = "kpi-card" + (index === 0 ? " hero" : "");
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
