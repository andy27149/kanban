async function loadDashboardData() {
  try {
    const response = await fetch("/api/data");
    if (!response.ok) {
      throw new Error(`HTTP ${response.status}: ${response.statusText}`);
    }
    const data = await response.json();
    renderKpis(data.kpis);
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

document.addEventListener("DOMContentLoaded", loadDashboardData);
