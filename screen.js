const dom = {
  uploadForm: document.getElementById("uploadForm"),
  uploadInput: document.getElementById("uploadInput"),
  uploadBtn: document.getElementById("uploadBtn"),
  status: document.getElementById("statusText"),
  fileMeta: document.getElementById("fileMeta"),
  clockNow: document.getElementById("clockNow"),

  tabOverviewBtn: document.getElementById("tabOverviewBtn"),
  tabOutboundBtn: document.getElementById("tabOutboundBtn"),
  overviewView: document.getElementById("overviewView"),
  outboundView: document.getElementById("outboundView"),

  province: document.getElementById("provinceFilter"),
  distributor: document.getElementById("distributorFilter"),
  category: document.getElementById("categoryFilter"),
  refreshBtn: document.getElementById("refreshBtn"),

  kpiTotal: document.getElementById("kpiTotal"),
  kpiMom: document.getElementById("kpiMom"),
  kpiStore: document.getElementById("kpiStore"),
  kpiDist: document.getElementById("kpiDist"),
  kpiProv: document.getElementById("kpiProv"),
  kpiStock: document.getElementById("kpiStock"),
  kpiScan: document.getElementById("kpiScan"),
  rankTableBody: document.getElementById("rankTableBody"),
  rankTableFoot: document.getElementById("rankTableFoot"),

  outProvince: document.getElementById("outProvinceFilter"),
  outDistributor: document.getElementById("outDistributorFilter"),
  outCategory: document.getElementById("outCategoryFilter"),
  outWeight: document.getElementById("outWeightFilter"),
  outYear: document.getElementById("outYearFilter"),
  outMonth: document.getElementById("outMonthFilter"),
  outRefreshBtn: document.getElementById("outRefreshBtn"),
  outDrillBackBtn: document.getElementById("outDrillBackBtn"),
  outDrillTitle: document.getElementById("outDrillTitle"),

  outTotalBoxes: document.getElementById("outTotalBoxes"),
  outMonthlyBoxes: document.getElementById("outMonthlyBoxes"),
  outLatestDayBoxes: document.getElementById("outLatestDayBoxes"),
  outInventoryBoxes: document.getElementById("outInventoryBoxes"),
  outQ1AvgBoxes: document.getElementById("outQ1AvgBoxes"),
  outSellableMonths: document.getElementById("outSellableMonths"),
  outCurrentNewCustomers: document.getElementById("outCurrentNewCustomers"),
  outCumulativeNewCustomers: document.getElementById("outCumulativeNewCustomers"),
  outMomRate: document.getElementById("outMomRate"),
  outSelectedMonth: document.getElementById("outSelectedMonth"),
  outDetailHead: document.getElementById("outDetailHead"),
  outDetailBody: document.getElementById("outDetailBody"),
  outDetailFoot: document.getElementById("outDetailFoot"),
};

const charts = {
  trend: echarts.init(document.getElementById("trendChart")),
  province: echarts.init(document.getElementById("provinceChart")),
  category: echarts.init(document.getElementById("categoryChart")),
  distributor: echarts.init(document.getElementById("distChart")),
  outMonthly: echarts.init(document.getElementById("outMonthlyTrendChart")),
  outDaily: echarts.init(document.getElementById("outDailyTrendChart")),
  outNewcust: echarts.init(document.getElementById("outNewcustTrendChart")),
};

let currentOutboundDrill = { level: "province", path: { province: null, distributor: null } };
let latestOutboundPayload = null;
let outboundSortState = { key: "total_boxes", direction: "desc" };

function fmtNumber(value) {
  if (value === null || value === undefined || Number.isNaN(Number(value))) return "--";
  return Number(value).toLocaleString("zh-CN", { maximumFractionDigits: 2 });
}

function fmtTableNumber(value) {
  if (value === null || value === undefined || value === "" || Number.isNaN(Number(value))) return "--";
  const n = Number(value);
  if (Number.isInteger(n)) {
    return n.toLocaleString("zh-CN", { maximumFractionDigits: 0 });
  }
  return n.toLocaleString("zh-CN", { minimumFractionDigits: 1, maximumFractionDigits: 1 });
}

function sumNumeric(rows, key) {
  let total = 0;
  let found = false;
  (rows || []).forEach((row) => {
    const raw = row ? row[key] : null;
    if (raw === null || raw === undefined || raw === "") return;
    const n = Number(raw);
    if (Number.isNaN(n)) return;
    total += n;
    found = true;
  });
  return { found, total };
}

function calcSellableSummary(rows) {
  const inv = sumNumeric(rows, "inventory_boxes");
  const q1 = sumNumeric(rows, "q1_avg_boxes");
  if (!inv.found || !q1.found || q1.total <= 0) return null;
  return inv.total / q1.total;
}

function renderRankSummary(rows) {
  if (!dom.rankTableFoot) return;
  dom.rankTableFoot.innerHTML = "";
  if (!rows || !rows.length) return;

  const tr = document.createElement("tr");
  tr.className = "table-summary-row";

  const totalCell = document.createElement("td");
  totalCell.textContent = "\u5408\u8ba1";
  tr.appendChild(totalCell);

  const provinceCell = document.createElement("td");
  provinceCell.textContent = "--";
  tr.appendChild(provinceCell);

  const distributorCell = document.createElement("td");
  distributorCell.textContent = "--";
  tr.appendChild(distributorCell);

  const valueCell = document.createElement("td");
  const sum = sumNumeric(rows, "value");
  valueCell.textContent = sum.found ? fmtNumber(sum.total) : "--";
  tr.appendChild(valueCell);

  dom.rankTableFoot.appendChild(tr);
}

function renderOutboundSummary(columns, rows) {
  if (!dom.outDetailFoot) return;
  dom.outDetailFoot.innerHTML = "";
  if (!columns || !columns.length || !rows || !rows.length) return;

  const tr = document.createElement("tr");
  tr.className = "table-summary-row";

  columns.forEach((col, idx) => {
    const td = document.createElement("td");
    if (idx === 0 || col.key === "name") {
      td.textContent = "\u5408\u8ba1";
    } else if (col.key === "__mini_trend") {
      td.textContent = "--";
    } else if (col.key === "sellable_months") {
      const sellable = calcSellableSummary(rows);
      td.textContent = sellable == null ? "--" : fmtTableNumber(sellable);
    } else {
      const sum = sumNumeric(rows, col.key);
      td.textContent = sum.found ? fmtTableNumber(sum.total) : "--";
    }
    tr.appendChild(td);
  });

  dom.outDetailFoot.appendChild(tr);
}

function getSellableClass(v) {
  const n = Number(v);
  if (Number.isNaN(n)) return "";
  if (n > 2) return "sellable-red";
  if (n >= 1.5) return "sellable-yellow";
  if (n >= 1) return "sellable-green";
  return "sellable-orange";
}

function classifyTrend(values) {
  const seq = (values || []).map((v) => Number(v) || 0);
  if (seq.length < 2) return { color: "#f5c451", label: "波动", score: 1 };
  let isUp = true;
  let isDown = true;
  for (let i = 1; i < seq.length; i += 1) {
    if (!(seq[i] > seq[i - 1])) isUp = false;
    if (!(seq[i] < seq[i - 1])) isDown = false;
  }
  if (isUp) return { color: "#3bd979", label: "上升", score: 2 };
  if (isDown) return { color: "#ff5f6d", label: "下降", score: 0 };
  return { color: "#f5c451", label: "波动", score: 1 };
}

function buildSparklineSvg(values, color = "#f5c451") {
  const seq = (values || []).map((v) => Number(v) || 0);
  if (!seq.length) return "";
  const width = 100;
  const height = 24;
  const pad = 2;
  const minVal = Math.min(...seq);
  const maxVal = Math.max(...seq);
  const range = Math.max(maxVal - minVal, 1);
  const step = seq.length > 1 ? (width - pad * 2) / (seq.length - 1) : 0;
  const points = seq
    .map((v, idx) => {
      const x = pad + idx * step;
      const y = height - pad - ((v - minVal) / range) * (height - pad * 2);
      return `${x.toFixed(2)},${y.toFixed(2)}`;
    })
    .join(" ");
  return `<svg viewBox="0 0 100 24" preserveAspectRatio="none" aria-hidden="true"><polyline fill="none" stroke="${color}" stroke-width="2" points="${points}"></polyline></svg>`;
}

function getDrillTrendYear(monthColumns) {
  const april = monthColumns.find((m) => /^(20\d{2})-04$/.test(String(m.key || "")));
  if (april) return Number(String(april.key).slice(0, 4));
  const years = monthColumns
    .map((m) => {
      const mm = String(m.key || "").match(/^(20\d{2})-\d{2}$/);
      return mm ? Number(mm[1]) : 0;
    })
    .filter((y) => y > 0);
  if (!years.length) return null;
  return Math.max(...years);
}

function getSortValue(row, key) {
  if (key === "__mini_trend") return Number(row.__mini_trend_score);
  const v = row[key];
  if (v === null || v === undefined || v === "") return null;
  const n = Number(v);
  if (!Number.isNaN(n)) return n;
  return String(v);
}

function compareRowsByKey(a, b, key, direction = "asc") {
  const av = getSortValue(a, key);
  const bv = getSortValue(b, key);
  let result = 0;
  if (av === null && bv === null) {
    result = 0;
  } else if (av === null) {
    result = 1;
  } else if (bv === null) {
    result = -1;
  } else if (typeof av === "number" && typeof bv === "number") {
    result = av - bv;
  } else {
    result = String(av).localeCompare(String(bv), "zh-CN");
  }
  if (result === 0) {
    result = String(a.name || "").localeCompare(String(b.name || ""), "zh-CN");
  }
  return direction === "asc" ? result : -result;
}

function setStatus(text, isError = false) {
  dom.status.textContent = text;
  dom.status.style.color = isError ? "#ffd06f" : "#98def3";
}

function fillSelect(el, options, keepValue = "全部") {
  const list = options && options.length ? options : ["全部"];
  el.innerHTML = "";
  list.forEach((item) => {
    const op = document.createElement("option");
    op.value = item;
    op.textContent = item;
    el.appendChild(op);
  });
  el.value = list.includes(keepValue) ? keepValue : list[0];
}

function switchView(view) {
  const isOverview = view === "overview";
  dom.overviewView.classList.toggle("active", isOverview);
  dom.outboundView.classList.toggle("active", !isOverview);
  dom.tabOverviewBtn.classList.toggle("active", isOverview);
  dom.tabOutboundBtn.classList.toggle("active", !isOverview);
  setTimeout(() => Object.values(charts).forEach((c) => c.resize()), 100);
}

function overviewFilters() {
  return {
    province: dom.province.value || "全部",
    distributor: dom.distributor.value || "全部",
    category: dom.category.value || "全部",
  };
}

function outboundFilters() {
  return {
    province: dom.outProvince.value || "全部",
    distributor: dom.outDistributor.value || "全部",
    category: dom.outCategory.value || "全部",
    weight: dom.outWeight.value || "全部",
    year: dom.outYear.value || "全部",
    month: dom.outMonth.value || "全部",
  };
}

function renderSimpleLine(chart, labels, values, color1 = "#46d3c0", color2 = "#4ca5ff") {
  chart.setOption({
    tooltip: { trigger: "axis" },
    grid: { left: 42, right: 16, top: 26, bottom: 34 },
    xAxis: { type: "category", data: labels || [], axisLabel: { color: "#c6e9fb" }, axisLine: { lineStyle: { color: "#6aa9c7" } } },
    yAxis: { type: "value", axisLabel: { color: "#c6e9fb" }, splitLine: { lineStyle: { color: "rgba(120,190,220,0.14)" } } },
    series: [{
      type: "line",
      smooth: true,
      symbolSize: 7,
      data: values || [],
      lineStyle: { width: 3, color: color1 },
      itemStyle: { color: color2 },
      areaStyle: {
        color: {
          type: "linear", x: 0, y: 0, x2: 0, y2: 1,
          colorStops: [{ offset: 0, color: "rgba(70,211,192,0.35)" }, { offset: 1, color: "rgba(70,211,192,0.03)" }],
        },
      },
    }],
  });
}

function renderHBar(chart, labels, values) {
  chart.setOption({
    tooltip: { trigger: "axis", axisPointer: { type: "shadow" } },
    grid: { left: 86, right: 14, top: 26, bottom: 22 },
    xAxis: { type: "value", axisLabel: { color: "#c7eafb" }, splitLine: { lineStyle: { color: "rgba(120,190,220,0.14)" } } },
    yAxis: { type: "category", data: (labels || []).slice().reverse(), axisLabel: { color: "#d6f3ff" } },
    series: [{ type: "bar", data: (values || []).slice().reverse(), barWidth: 12, itemStyle: { color: "#4caeff", borderRadius: [0, 8, 8, 0] } }],
  });
}

function renderVBar(chart, labels, values) {
  chart.setOption({
    tooltip: { trigger: "axis", axisPointer: { type: "shadow" } },
    grid: { left: 46, right: 12, top: 24, bottom: 56 },
    xAxis: { type: "category", data: labels || [], axisLabel: { rotate: 30, color: "#c7eafb", fontSize: 10 } },
    yAxis: { type: "value", axisLabel: { color: "#c7eafb" }, splitLine: { lineStyle: { color: "rgba(120,190,220,0.14)" } } },
    series: [{ type: "bar", data: values || [], barMaxWidth: 18, itemStyle: { color: "#49a7ff", borderRadius: [6, 6, 0, 0] } }],
  });
}

function renderPie(chart, labels, values) {
  const data = (labels || []).map((name, i) => ({ name, value: values[i] || 0 }));
  chart.setOption({
    tooltip: { trigger: "item" },
    legend: { bottom: 0, textStyle: { color: "#bedff0", fontSize: 11 } },
    series: [{ type: "pie", radius: ["32%", "70%"], center: ["50%", "44%"], roseType: "radius", data }],
  });
}

function renderOverview(payload, replaceFilters = false, selected = null) {
  const keep = selected || overviewFilters();
  if (replaceFilters) {
    fillSelect(dom.province, payload.filters?.province_options || ["全部"], keep.province);
    fillSelect(dom.distributor, payload.filters?.distributor_options || ["全部"], keep.distributor);
    fillSelect(dom.category, payload.filters?.category_options || ["全部"], keep.category);
  }

  const k = payload.kpis || {};
  dom.kpiTotal.textContent = fmtNumber(k.total_value);
  dom.kpiStore.textContent = fmtNumber(k.store_count);
  dom.kpiDist.textContent = fmtNumber(k.distributor_count);
  dom.kpiProv.textContent = fmtNumber(k.province_count);
  dom.kpiStock.textContent = fmtNumber(k.inventory_boxes);
  dom.kpiMom.textContent = k.mom_rate == null ? "环比: --" : `环比: ${k.mom_rate >= 0 ? "+" : ""}${fmtNumber(k.mom_rate)}%`;
  dom.kpiScan.textContent = k.scan_rate == null ? "扫码率: --" : `扫码率: ${fmtNumber(k.scan_rate)}%`;

  renderSimpleLine(charts.trend, payload.trend?.labels || [], payload.trend?.values || []);
  renderHBar(charts.province, payload.province_rank?.labels || [], payload.province_rank?.values || []);
  renderPie(charts.category, payload.category_share?.labels || [], payload.category_share?.values || []);
  renderVBar(charts.distributor, payload.distributor_rank?.labels || [], payload.distributor_rank?.values || []);

  dom.rankTableBody.innerHTML = "";
  const rows = payload.table_rows || [];
  if (!rows.length) {
    dom.rankTableBody.innerHTML = "<tr><td colspan='4'>当前筛选下暂无数据</td></tr>";
  } else {
    rows.forEach((row) => {
      const tr = document.createElement("tr");
      tr.innerHTML = `<td>${row.rank}</td><td>${row.province}</td><td>${row.distributor}</td><td>${fmtNumber(row.value)}</td>`;
      dom.rankTableBody.appendChild(tr);
    });
  }

  renderRankSummary(rows);

  const meta = payload.meta || {};
  if (meta.file_name) {
    dom.fileMeta.textContent = `数据文件: ${meta.file_name} | 更新时间: ${meta.updated_at || "--"} | 出库来源: ${meta.outbound_sheet || "--"}`;
  }
}

function renderOutboundDrill(payload) {
  const drill = payload.drill || { level: "province", rows: [], path: {} };
  currentOutboundDrill = drill;

  const levelNameMap = {
    province: "省区",
    distributor: "\u5ba2\u6237\u7b80\u79f0",
    store: "门店",
  };
  const levelLabel = levelNameMap[drill.level] || "维度";
  if (drill.level === "province") {
    dom.outDrillTitle.textContent = "下钻层级: 省区";
  } else if (drill.level === "distributor") {
    dom.outDrillTitle.textContent = `下钻层级: 省区 ${drill.path?.province || ""} -> 客户简称`;
  } else {
    dom.outDrillTitle.textContent = `下钻层级: ${drill.path?.province || ""} / ${drill.path?.distributor || ""} -> 门店`;
  }

  const baseMonthColumns = Array.isArray(drill.month_columns) ? drill.month_columns : [];
  const monthColumns = [...baseMonthColumns];
  const trendYear = getDrillTrendYear(baseMonthColumns);
  const monthKeysForTrend = trendYear
    ? [1, 2, 3, 4].map((m) => `${trendYear}-${String(m).padStart(2, "0")}`)
    : baseMonthColumns.slice(0, 4).map((m) => m.key);
  let miniTrendInsertAt = -1;
  if (trendYear) {
    miniTrendInsertAt = monthColumns.findIndex((m) => m.key === `${trendYear}-04`);
  }
  if (miniTrendInsertAt < 0 && monthColumns.length > 0) {
    miniTrendInsertAt = monthColumns.length - 1;
  }
  if (monthColumns.length > 0) {
    monthColumns.splice(miniTrendInsertAt + 1, 0, { key: "__mini_trend", label: "1-4\u6708\u8d8b\u52bf", type: "spark" });
  }

  let metricsColumns = Array.isArray(drill.metrics_columns) ? drill.metrics_columns : [];
  if (!metricsColumns.length) {
    if (drill.level === "province") {
      metricsColumns = [
        { key: "boxes", label: "\u7d2f\u8ba1\u51fa\u5e93(\u7bb1)" },
        { key: "customer_count", label: "\u5ba2\u6237\u6570" },
        { key: "store_count", label: "\u95e8\u5e97\u6570" },
      ];
    } else if (drill.level === "distributor") {
      metricsColumns = [
        { key: "boxes", label: "\u7d2f\u8ba1\u51fa\u5e93(\u7bb1)" },
        { key: "store_count", label: "\u95e8\u5e97\u6570" },
      ];
    } else {
      metricsColumns = [
        { key: "boxes", label: "\u7d2f\u8ba1\u51fa\u5e93(\u7bb1)" },
        { key: "product_count", label: "\u4ea7\u54c1\u6570" },
      ];
    }
  }

  const columns = [{ key: "name", label: drill.name_label || levelLabel }, ...monthColumns, ...metricsColumns];
  const sortableKeys = new Set(columns.map((c) => c.key));
  if (!sortableKeys.has(outboundSortState.key)) {
    outboundSortState = { key: sortableKeys.has("total_boxes") ? "total_boxes" : "name", direction: "desc" };
  }

  const tableEl = dom.outDetailHead.closest("table");
  if (tableEl) {
    const oldColgroup = tableEl.querySelector("colgroup");
    if (oldColgroup) oldColgroup.remove();
    const monthKeySet = new Set(monthColumns.map((c) => c.key));
    const colWidths = columns.map((col, idx) => {
      if (idx === 0 || col.key === "name") return 220;
      if (col.key === "__mini_trend") return 170;
      if (monthKeySet.has(col.key)) return 108;
      return 126;
    });
    const cg = document.createElement("colgroup");
    colWidths.forEach((widthPx) => {
      const col = document.createElement("col");
      col.style.width = `${widthPx}px`;
      cg.appendChild(col);
    });
    tableEl.insertBefore(cg, tableEl.firstChild);
    const wrapEl = tableEl.closest(".table-wrap");
    const minWidth = colWidths.reduce((sum, w) => sum + w, 0);
    const wrapWidth = wrapEl ? wrapEl.clientWidth : 0;
    tableEl.style.minWidth = `${Math.max(minWidth, wrapWidth)}px`;
  }

  const detailColumns = [...monthColumns, ...metricsColumns];
  const groupTr = document.createElement("tr");
  const groupNameTh = document.createElement("th");
  groupNameTh.rowSpan = 2;
  groupNameTh.classList.add("sortable-th");
  groupNameTh.dataset.sortKey = "name";
  const nameArrow = outboundSortState.key === "name" ? (outboundSortState.direction === "asc" ? " \u25b2" : " \u25bc") : "";
  groupNameTh.textContent = `${drill.name_label || levelLabel}${nameArrow}`;
  groupTr.appendChild(groupNameTh);

  if (monthColumns.length > 0) {
    const monthGroupTh = document.createElement("th");
    monthGroupTh.colSpan = monthColumns.length;
    monthGroupTh.textContent = "月度出库 / 趋势";
    groupTr.appendChild(monthGroupTh);
  }
  if (metricsColumns.length > 0) {
    const metricGroupTh = document.createElement("th");
    metricGroupTh.colSpan = metricsColumns.length;
    metricGroupTh.textContent = "分析指标";
    groupTr.appendChild(metricGroupTh);
  }

  const detailTr = document.createElement("tr");
  detailColumns.forEach((col) => {
    const th = document.createElement("th");
    th.classList.add("sortable-th");
    th.dataset.sortKey = col.key;
    const arrow = outboundSortState.key === col.key ? (outboundSortState.direction === "asc" ? " \u25b2" : " \u25bc") : "";
    th.textContent = `${col.label || col.key}${arrow}`;
    detailTr.appendChild(th);
  });

  dom.outDetailHead.innerHTML = "";
  dom.outDetailHead.appendChild(groupTr);
  dom.outDetailHead.appendChild(detailTr);

  dom.outDetailBody.innerHTML = "";
  const rawRows = Array.isArray(drill.rows) ? drill.rows : [];
  const rows = rawRows.map((row) => {
    const trendValues = monthKeysForTrend.map((k) => Number(row[k]) || 0);
    const trendSignal = classifyTrend(trendValues);
    return {
      ...row,
      __mini_trend_values: trendValues,
      __mini_trend_color: trendSignal.color,
      __mini_trend_label: trendSignal.label,
      __mini_trend_score: trendSignal.score,
    };
  });
  rows.sort((a, b) => compareRowsByKey(a, b, outboundSortState.key, outboundSortState.direction));
  renderOutboundSummary(columns, rows);

  if (!rows.length) {
    const tr = document.createElement("tr");
    const td = document.createElement("td");
    td.colSpan = columns.length;
    td.textContent = "\u5f53\u524d\u7b5b\u9009\u4e0b\u65e0\u6570\u636e";
    tr.appendChild(td);
    dom.outDetailBody.appendChild(tr);
    return;
  }

  rows.forEach((row) => {
    const tr = document.createElement("tr");
    tr.dataset.name = row.name || "";
    if (drill.level !== "store") tr.style.cursor = "pointer";

    columns.forEach((col) => {
      const td = document.createElement("td");
      if (col.key === "name") {
        td.textContent = row.name || "--";
      } else if (col.key === "__mini_trend") {
        td.classList.add("mini-trend-cell");
        const svg = buildSparklineSvg(row.__mini_trend_values || [], row.__mini_trend_color || "#f5c451");
        td.innerHTML = `${svg}<span class="trend-chip" style="color:${row.__mini_trend_color || "#f5c451"}">${row.__mini_trend_label || "波动"}</span>`;
      } else {
        td.textContent = fmtTableNumber(row[col.key]);
        if (col.key === "sellable_months") {
          const cls = getSellableClass(row[col.key]);
          if (cls) td.classList.add("sellable-cell", cls);
        }
      }
      tr.appendChild(td);
    });
    dom.outDetailBody.appendChild(tr);
  });

  dom.outDetailHead.querySelectorAll(".sortable-th").forEach((th) => {
    th.addEventListener("click", () => {
      const key = th.dataset.sortKey || "";
      if (!key) return;
      if (outboundSortState.key === key) {
        outboundSortState.direction = outboundSortState.direction === "asc" ? "desc" : "asc";
      } else {
        outboundSortState = { key, direction: key === "name" ? "asc" : "desc" };
      }
      if (latestOutboundPayload) {
        renderOutboundDrill(latestOutboundPayload);
      }
    });
  });
}

function renderOutbound(payload, replaceFilters = false, selected = null) {
  latestOutboundPayload = payload;
  const keep = selected || outboundFilters();
  if (replaceFilters) {
    fillSelect(dom.outProvince, payload.filters?.province_options || ["全部"], keep.province);
    fillSelect(dom.outDistributor, payload.filters?.distributor_options || ["全部"], keep.distributor);
    fillSelect(dom.outCategory, payload.filters?.category_options || ["全部"], keep.category);
    fillSelect(dom.outWeight, payload.filters?.weight_options || ["全部"], keep.weight);
    fillSelect(dom.outYear, payload.filters?.year_options || ["全部"], keep.year);
    fillSelect(dom.outMonth, payload.filters?.month_options || ["全部"], keep.month);
  }

  const k = payload.kpis || {};
  dom.outTotalBoxes.textContent = fmtNumber(k.total_boxes);
  dom.outMonthlyBoxes.textContent = fmtNumber(k.monthly_boxes);
  dom.outLatestDayBoxes.textContent = fmtNumber(k.latest_day_boxes);
  dom.outInventoryBoxes.textContent = fmtNumber(k.inventory_boxes);
  dom.outQ1AvgBoxes.textContent = `1-3月均出库: ${fmtNumber(k.q1_avg_boxes)}`;
  dom.outSellableMonths.textContent = fmtTableNumber(k.sellable_months);
  dom.outSellableMonths.classList.remove("sellable-red", "sellable-yellow", "sellable-green", "sellable-orange");
  const sellableCls = getSellableClass(k.sellable_months);
  if (sellableCls) dom.outSellableMonths.classList.add(sellableCls);
  dom.outCurrentNewCustomers.textContent = fmtNumber(k.current_new_customers);
  dom.outCumulativeNewCustomers.textContent = `累计: ${fmtNumber(k.cumulative_new_customers)}`;
  dom.outMomRate.textContent = k.mom_rate == null ? "环比: --" : `环比: ${k.mom_rate >= 0 ? "+" : ""}${fmtNumber(k.mom_rate)}%`;
  dom.outSelectedMonth.textContent = `月份: ${k.selected_month_label || "--"}`;

  renderSimpleLine(charts.outMonthly, payload.monthly_trend?.labels || [], payload.monthly_trend?.values || [], "#3bd9c6", "#7ebeff");
  renderVBar(charts.outDaily, payload.daily_trend?.labels || [], payload.daily_trend?.values || []);
  renderSimpleLine(charts.outNewcust, payload.newcust_monthly_trend?.labels || [], payload.newcust_monthly_trend?.values || [], "#f5c451", "#ffd36b");
  renderOutboundDrill(payload);
}

async function readJsonResponse(resp) {
  const raw = await resp.text();
  if (!raw) {
    throw new Error(`\u670d\u52a1\u8fd4\u56de\u7a7a\u54cd\u5e94\uff08HTTP ${resp.status}\uff09`);
  }
  try {
    return JSON.parse(raw);
  } catch (err) {
    const preview = raw.slice(0, 120).replace(/\s+/g, " ");
    throw new Error(`\u670d\u52a1\u8fd4\u56de\u4e86\u975e JSON \u5185\u5bb9\uff08HTTP ${resp.status}\uff09\uff1a${preview || "\u7a7a\u6587\u672c"}`);
  }
}

async function fetchOverview(selected = null, replaceFilters = false) {
  const filters = selected || overviewFilters();
  const query = new URLSearchParams(filters);
  const resp = await fetch(`/api/dashboard?${query.toString()}`);
  const body = await readJsonResponse(resp);
  if (!resp.ok || !body.ok) throw new Error(body.error || "总览刷新失败");
  renderOverview(body.data, replaceFilters, filters);
}

async function fetchOutbound(selected = null, replaceFilters = true) {
  const filters = selected || outboundFilters();
  const query = new URLSearchParams(filters);
  const resp = await fetch(`/api/outbound/trend?${query.toString()}`);
  const body = await readJsonResponse(resp);
  if (!resp.ok || !body.ok) throw new Error(body.error || "出库趋势刷新失败");
  renderOutbound(body.data, replaceFilters, filters);
}

async function uploadAndRefresh(file) {
  const formData = new FormData();
  formData.append("file", file);
  dom.uploadBtn.disabled = true;
  setStatus("\u6b63\u5728\u89e3\u6790\u6587\u4ef6\uff0c\u8bf7\u7a0d\u5019...");
  const resp = await fetch("/api/upload", { method: "POST", body: formData });
  dom.uploadBtn.disabled = false;
  const body = await readJsonResponse(resp);
  if (!resp.ok || !body.ok) throw new Error(body.error || "\u4e0a\u4f20\u5931\u8d25");
  renderOverview(body.dashboard || body.data, true, overviewFilters());
  renderOutbound(body.outbound || {}, true, outboundFilters());
  setStatus("\u6570\u636e\u5df2\u5237\u65b0");
}

function bindEvents() {
  dom.tabOverviewBtn.addEventListener("click", () => switchView("overview"));
  dom.tabOutboundBtn.addEventListener("click", () => switchView("outbound"));

  dom.uploadForm.addEventListener("submit", async (e) => {
    e.preventDefault();
    const f = dom.uploadInput.files && dom.uploadInput.files[0];
    if (!f) {
      setStatus("\u8bf7\u5148\u9009\u62e9 Excel \u6216 CSV \u6587\u4ef6", true);
      return;
    }
    try {
      await uploadAndRefresh(f);
    } catch (err) {
      setStatus(err.message || "\u4e0a\u4f20\u5931\u8d25", true);
    }
  });

  dom.refreshBtn.addEventListener("click", async () => {
    try {
      setStatus("\u6b63\u5728\u5237\u65b0\u603b\u89c8...");
      await fetchOverview();
      setStatus("\u603b\u89c8\u5df2\u66f4\u65b0");
    } catch (err) {
      setStatus(err.message || "\u603b\u89c8\u5237\u65b0\u5931\u8d25", true);
    }
  });

  [dom.province, dom.distributor, dom.category].forEach((el) => {
    el.addEventListener("change", async () => {
      try {
        await fetchOverview();
        setStatus("\u603b\u89c8\u7b5b\u9009\u5df2\u66f4\u65b0");
      } catch (err) {
        setStatus(err.message || "\u7b5b\u9009\u5931\u8d25", true);
      }
    });
  });

  dom.outRefreshBtn.addEventListener("click", async () => {
    try {
      setStatus("\u6b63\u5728\u5237\u65b0\u51fa\u5e93\u8d8b\u52bf...");
      await fetchOutbound();
      setStatus("\u51fa\u5e93\u8d8b\u52bf\u5df2\u66f4\u65b0");
    } catch (err) {
      setStatus(err.message || "\u51fa\u5e93\u8d8b\u52bf\u5237\u65b0\u5931\u8d25", true);
    }
  });

  [dom.outProvince, dom.outDistributor, dom.outCategory, dom.outWeight, dom.outYear, dom.outMonth].forEach((el) => {
    el.addEventListener("change", async () => {
      try {
        await fetchOutbound();
        setStatus("\u51fa\u5e93\u7b5b\u9009\u5df2\u66f4\u65b0");
      } catch (err) {
        setStatus(err.message || "\u7b5b\u9009\u5931\u8d25", true);
      }
    });
  });

  dom.outDetailBody.addEventListener("click", async (e) => {
    const tr = e.target.closest("tr");
    if (!tr || !tr.dataset.name) return;
    const name = tr.dataset.name;

    if (currentOutboundDrill.level === "province") {
      const next = { ...outboundFilters(), province: name, distributor: "全部" };
      try {
        await fetchOutbound(next, true);
      } catch (err) {
        setStatus(err.message || "下钻失败", true);
      }
      return;
    }

    if (currentOutboundDrill.level === "distributor") {
      const next = { ...outboundFilters(), distributor: name };
      try {
        await fetchOutbound(next, true);
      } catch (err) {
        setStatus(err.message || "下钻失败", true);
      }
    }
  });

  dom.outDrillBackBtn.addEventListener("click", async () => {
    const f = outboundFilters();
    if (currentOutboundDrill.level === "store") {
      f.distributor = "全部";
    } else if (currentOutboundDrill.level === "distributor") {
      f.province = "全部";
      f.distributor = "全部";
    } else {
      return;
    }
    try {
      await fetchOutbound(f, true);
    } catch (err) {
      setStatus(err.message || "\u8fd4\u56de\u5931\u8d25", true);
    }
  });

  window.addEventListener("resize", () => Object.values(charts).forEach((c) => c.resize()));
}

function tickClock() {
  dom.clockNow.textContent = new Date().toLocaleString("zh-CN", { hour12: false });
}

function emptyState() {
  fillSelect(dom.province, ["全部"], "全部");
  fillSelect(dom.distributor, ["全部"], "全部");
  fillSelect(dom.category, ["全部"], "全部");
  fillSelect(dom.outProvince, ["全部"], "全部");
  fillSelect(dom.outDistributor, ["全部"], "全部");
  fillSelect(dom.outCategory, ["全部"], "全部");
  fillSelect(dom.outWeight, ["全部"], "全部");
  fillSelect(dom.outYear, ["全部"], "全部");
  fillSelect(dom.outMonth, ["全部"], "全部");

  renderSimpleLine(charts.trend, [], []);
  renderHBar(charts.province, [], []);
  renderPie(charts.category, [], []);
  renderVBar(charts.distributor, [], []);
  renderSimpleLine(charts.outMonthly, [], []);
  renderVBar(charts.outDaily, [], []);
  renderSimpleLine(charts.outNewcust, [], [], "#f5c451", "#ffd36b");
  if (dom.rankTableFoot) dom.rankTableFoot.innerHTML = "";
  if (dom.outDetailFoot) dom.outDetailFoot.innerHTML = "";
  dom.rankTableBody.innerHTML = "<tr><td colspan='4'>\u8bf7\u5148\u4e0a\u4f20\u6570\u636e</td></tr>";
  dom.outDetailHead.innerHTML = "<tr><th>下钻明细</th></tr>";
  dom.outDetailBody.innerHTML = "<tr><td>\u8bf7\u5148\u4e0a\u4f20\u6570\u636e</td></tr>";
}

async function boot() {
  bindEvents();
  tickClock();
  setInterval(tickClock, 1000);
  switchView("overview");

  if (window.__PAGE_BOOTSTRAP__?.loaded) {
    try {
      setStatus("\u6b63\u5728\u52a0\u8f7d\u5df2\u6709\u6570\u636e...");
      await fetchOverview();
      await fetchOutbound();
      setStatus("\u6570\u636e\u5df2\u52a0\u8f7d");
      return;
    } catch (err) {
      setStatus("\u521d\u59cb\u5316\u5931\u8d25\uff0c\u8bf7\u91cd\u65b0\u4e0a\u4f20\u6587\u4ef6", true);
    }
  }
  emptyState();
  setStatus("\u8bf7\u5148\u4e0a\u4f20\u6570\u636e\u6587\u4ef6");
}

boot();

