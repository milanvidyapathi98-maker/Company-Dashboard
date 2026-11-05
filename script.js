/* =====================================================================
   PART 1 - small helpers
   ===================================================================== */
const $ = (id) => document.getElementById(id);
const css = (n) => getComputedStyle(document.documentElement).getPropertyValue(n).trim();
const col = { s1: css("--s1"), s2: css("--s2"), s3: css("--s3"), s4: css("--s4"), s5: css("--s5"),
              grid: "#252b33", muted: css("--muted"), other: "#6b7280" };

const pct = (v, d = 1) => (isFinite(v) ? (v * 100).toFixed(d) + "%" : "n/a");
const bn = (m, d = 1) => (m / 1000).toFixed(d);
let CUR = "$";                                    // currency symbol, set from the Settings sheet
const money = (m) => CUR + (Math.abs(m) >= 1e6 ? (m / 1e6).toFixed(2) + "T" : (m / 1000).toFixed(1) + "B");
const num = (v, d = 1) => (isFinite(v) ? v.toLocaleString("en-US", { minimumFractionDigits: d, maximumFractionDigits: d }) : "n/a");
const arrow = (up) => (up ? "▲" : "▼");
const store = {                                   // localStorage, safe if blocked
  get(k) { try { return localStorage.getItem(k); } catch (e) { return null; } },
  set(k, v) { try { localStorage.setItem(k, v); } catch (e) {} },
};

function savePriceCache(symbol, data) {
  try {
    localStorage.setItem(
      `prices_${symbol}`,
      JSON.stringify({
        timestamp: Date.now(),
        prices: data
      })
    );
  } catch (e) {
    console.error("Could not save price cache", e);
  }
}

function loadPriceCache(symbol) {
  try {
    const cached = localStorage.getItem(`prices_${symbol}`);

    if (!cached) return null;

    return JSON.parse(cached);
  } catch (e) {
    console.error("Could not read price cache", e);
    return null;
  }
}

function cacheAge(timestamp) {
  const hours = Math.floor((Date.now() - timestamp) / 3600000);

  if (hours < 1) return "less than 1 hour old";
  if (hours < 24) return `${hours} hours old`;

  const days = Math.floor(hours / 24);
  return `${days} day(s) old`;
}

/* Companies in the ticker drop-down. Edit this list, or set CONFIG.TICKERS in config.js. */
const DEFAULT_TICKERS = [
  { symbol: "NVDA", name: "NVIDIA" }, { symbol: "AAPL", name: "Apple" }, { symbol: "MSFT", name: "Microsoft" },
  { symbol: "GOOGL", name: "Alphabet", aka: ["Google"] }, { symbol: "AMZN", name: "Amazon" },
  { symbol: "META", name: "Meta Platforms" }, { symbol: "TSLA", name: "Tesla" }, { symbol: "AVGO", name: "Broadcom" },
  { symbol: "AMD", name: "Advanced Micro Devices" }, { symbol: "NFLX", name: "Netflix" },
];
const TICKERS = CONFIG.TICKERS && CONFIG.TICKERS.length ? CONFIG.TICKERS : DEFAULT_TICKERS;
function fillTickers() {
  $("ticker").innerHTML = `<option value="">Select a company…</option>` +
    TICKERS.map((t) => `<option value="${t.symbol}">${t.symbol} – ${esc(t.name)}</option>`).join("");
}
function setTicker(sym) {
  sym = (sym || "").toUpperCase();
  const sel = $("ticker");
  if (sym && ![...sel.options].some((o) => o.value === sym)) sel.add(new Option(sym, sym));  // ticker not in the list (e.g. from the Settings sheet)
  sel.value = sym;
}

/* Chart.js: keep every chart so we can destroy them when a new file is loaded */
const charts = [];
const mk = (el, cfg) => { const c = new Chart(el, cfg); charts.push(c); return c; };
const destroyCharts = () => charts.splice(0).forEach((c) => c.destroy());

Chart.defaults.color = col.muted;
Chart.defaults.borderColor = col.grid;
Chart.defaults.font.size = 11;
Chart.defaults.plugins.legend.labels.boxWidth = 10;
Chart.defaults.plugins.legend.labels.boxHeight = 10;
Chart.defaults.plugins.tooltip.backgroundColor = "#0d1117";
Chart.defaults.plugins.tooltip.borderColor = "#3a424d";
Chart.defaults.plugins.tooltip.borderWidth = 1;
Chart.defaults.plugins.tooltip.padding = 10;

function chartOpts({ fmt = (v) => v, legend = true, stacked = false } = {}) {
  return {
    responsive: true, maintainAspectRatio: false,
    interaction: { mode: "index", intersect: false },
    plugins: { legend: { display: legend, position: "bottom" },
               tooltip: { callbacks: { label: (c) => ` ${c.dataset.label}: ${fmt(c.parsed.y)}` } } },
    scales: { x: { stacked, grid: { display: false } },
              y: { stacked, grid: { color: col.grid }, ticks: { callback: (v) => fmt(v) } } },
  };
}
const bar = (label, data, color) => ({ label, data, backgroundColor: color, borderRadius: 4, maxBarThickness: 26, borderSkipped: false });
const line = (label, data, color) => ({ label, data, borderColor: color, backgroundColor: color, borderWidth: 2, tension: 0.3, pointRadius: 3, pointHoverRadius: 5 });

/* =====================================================================
   PART 2 - read the Excel workbook into one DATA object
   Row labels must match the NVIDIA template. Required rows stop the
   import with a clear message; optional rows default to 0.
   ===================================================================== */
const SPEC = {
  income: [
    ["revenue", ["Revenue", "Total revenue", "Total revenues", "Revenues"], true],
    ["costOfRevenue", ["Cost of revenue", "Cost of revenues", "Cost of sales"]],
    ["grossProfit", ["Gross profit"]],
    ["sga", ["Selling, general & administrative", "Selling, general and administrative"]],
    ["rnd", ["Research & development", "Research and development"]],
    ["operatingIncome", ["Operating income", "Income from operations"], true],
    ["interestExpense", ["Interest expense"]],
    ["investmentGain", ["Gain (loss) on sale of investments"]],
    ["netIncome", ["Net income", "Net income (loss)"], true],
    ["dilutedShares", ["Weighted avg. shares - diluted (millions)"], true],
    ["epsDiluted", ["EPS - diluted ($)", "Diluted EPS", "EPS (diluted)"], true],
    ["dps", ["Dividends per share ($)", "Dividends per share"]],
  ],
  balance: [
    ["cash", ["Cash & cash equivalents", "Cash and cash equivalents"], true],
    ["securities", ["Marketable securities (short-term investments)", "Marketable securities", "Short-term investments"]],
    ["receivables", ["Accounts receivable, net", "Accounts receivable"]],
    ["inventory", ["Inventories", "Inventory"]],
    ["currentAssets", ["Total current assets"], true],
    ["ppe", ["Property, plant & equipment, net", "Property and equipment, net", "Property, plant and equipment, net"]],
    ["ltInvestments", ["Long-term investments (incl. non-marketable equity)", "Long-term investments", "Non-marketable securities"]],
    ["goodwill", ["Goodwill"]],
    ["totalAssets", ["Total assets"], true],
    ["payables", ["Accounts payable"]],
    ["currentLiabilities", ["Total current liabilities"], true],
    ["equity", ["Total shareholders' equity", "Total stockholders' equity", "Total equity", "Total Equity Gross Minority Interest", "Stockholders' equity", "Shareholders' equity"], true],
    ["_d1", ["Current portion of long-term debt", "Current debt"]],
    ["_d2", ["Current portion of leases", "Current portion of operating lease liabilities"]],
    ["_d3", ["Long-term debt"]],
    ["_d4", ["Long-term leases", "Long-term operating lease liabilities"]],
  ],
  cashflow: [
    ["ocf", ["Net cash from operating activities", "Net cash provided by operating activities", "Operating cash flow"], true],
    ["capex", ["Capital expenditures (PP&E and intangibles)", "Capital expenditures", "Capital expenditure", "Purchases of property and equipment", "Purchase of property and equipment", "Purchases of property, plant and equipment", "Purchase of PP&E"], true],
    ["buybacks", ["Repurchase of common stock", "Repurchases of stock", "Repurchase of capital stock", "Common stock repurchased"]],
    ["dividends", ["Common dividends paid", "Dividends paid", "Cash dividends paid"]],
  ],
};
// Row names are compared in a simplified form (lower case, "&" = "and", no punctuation).
const norm = (t) => String(t).toLowerCase().replace(/&/g, " and ").replace(/[^a-z0-9]+/g, " ").trim();
const SYMBOLS = { USD: "$", EUR: "€", GBP: "£", JPY: "¥", CNY: "¥", INR: "₹", CHF: "CHF ", CAD: "C$", AUD: "A$", LKR: "Rs " };
function currencyOf(v) {
  v = String(v || "USD").trim();
  const code = v.toUpperCase();
  return SYMBOLS[code] ? { sym: SYMBOLS[code], code } : { sym: v, code: v };
}
const esc = (t) => String(t).replace(/[&<>]/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;" }[c]));

// Optional sheet "Settings": two columns, name in A and value in B (Company, Ticker, Currency, Subtitle, Note)
function readSettings(wb) {
  const res = { notes: [] };
  const name = wb.SheetNames.find((n) => norm(n) === "settings");
  if (!name) return res;
  XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, defval: null }).forEach((r) => {
    if (r[0] == null || r[1] == null || r[1] === "") return;
    const k = norm(r[0]), v = String(r[1]).trim();
    if (k === "note" || k === "notes") res.notes.push(v);
    else if (k.startsWith("company")) res.company = v;
    else if (k.startsWith("ticker")) res.ticker = v.toUpperCase();
    else if (k.startsWith("currency")) res.currency = v;
    else if (k === "subtitle") res.subtitle = v;
  });
  return res;
}

const SHEET_HINT = { income: "income", balance: "balance", cashflow: "cash flow" };

function readSheet(wb, kind) {
  const name = wb.SheetNames.find((n) => n.toLowerCase().includes(SHEET_HINT[kind]));
  if (!name) throw new Error(`No sheet whose name contains "${SHEET_HINT[kind]}" was found.`);
  const grid = XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, defval: null });
  let years = null, title = null; const rows = {};
  grid.forEach((r, i) => {
    if (i === 0 && typeof r[0] === "string") title = r[0];
    if (r[0] === "Fiscal year") years = r.slice(1).filter((x) => x != null && x !== "").map(String);
    else if (years && typeof r[0] === "string") {
      const vals = r.slice(1, 1 + years.length);
      if (vals.some((x) => typeof x === "number")) rows[norm(r[0])] = vals.map((x) => (typeof x === "number" ? x : 0));
    }
  });
  if (!years) throw new Error(`Sheet "${name}" has no row starting with "Fiscal year".`);
  return { years, rows, title };
}

function parseWorkbook(wb, fileName) {
  const inc = readSheet(wb, "income"), bal = readSheet(wb, "balance"), cf = readSheet(wb, "cashflow");
  const years = inc.years, n = years.length, missing = [];
  const out = { years, income: {}, balance: {}, cashflow: {} };
  const src = { income: inc.rows, balance: bal.rows, cashflow: cf.rows };
  const fatal = [];
  for (const kind of Object.keys(SPEC)) {
    for (const [key, label, required] of SPEC[kind]) {
      const names = [].concat(label);
      const row = names.map((nm) => src[kind][norm(nm)]).find(Boolean);
      if (row) out[kind][key] = row.slice(0, n);
      else if (required) fatal.push(`"${names[0]}" (${kind} sheet)`);
      else { out[kind][key] = new Array(n).fill(0); missing.push(names[0]); }
    }
  }
  if (fatal.length) throw new Error("Required rows not found: " + fatal.join(", ") + ". Rename those rows in your Excel file to the names shown, or tell me the exact names used in your file.");
  const I = out.income, B = out.balance, C = out.cashflow;
  if (!src.income[norm("Gross profit")]) I.grossProfit = I.revenue.map((v, i) => v - I.costOfRevenue[i]);
  B.totalDebt = years.map((_, i) => B._d1[i] + B._d2[i] + B._d3[i] + B._d4[i]);
  B.netCash = years.map((_, i) => B.cash[i] + B.securities[i] - B.totalDebt[i]);
  ["_d1", "_d2", "_d3", "_d4"].forEach((k) => delete B[k]);
  // Outflows must be negative: accept sheets that show them as positive numbers
  I.interestExpense = I.interestExpense.map((v) => -Math.abs(v));
  ["capex", "buybacks", "dividends"].forEach((k) => (C[k] = C[k].map((v) => -Math.abs(v))));
  C.fcf = years.map((_, i) => C.ocf[i] + C.capex[i]);
  out.company = (inc.title || "").split(" - ")[0].trim() || fileName.replace(/\.[^.]+$/, "");
  const set = readSettings(wb);
  if (set.company) out.company = set.company;
  out.ticker = set.ticker || "";
  out.currency = currencyOf(set.currency);
  out.subtitle = set.subtitle || "";
  out.notes = set.notes;
  out.missing = missing;
  return out;
}

/* =====================================================================
   PART 3 - build the whole dashboard from a DATA object
   ===================================================================== */
let DATA = null, prices = [], priceChart = null, sparkN = 0;

function render(data) {
  DATA = data;
  CUR = (data.currency && data.currency.sym) || "$";
  const CODE = (data.currency && data.currency.code) || "USD";
  destroyCharts(); priceChart = null; sparkN = 0;
  $("welcome").hidden = true; $("dash").hidden = false;

  const Y = data.years, N = Y.length, L = N - 1;
  const I = data.income, B = data.balance, C = data.cashflow;
  const fy = Y[L], fyPrev = Y[L - 1];
  const map = (f) => Y.map((_, i) => f(i));
  const div = (a, b) => (b ? a / b : NaN);
  const yoy = (a, i = L) => div(a[i], a[i - 1]) - 1;

  document.title = data.company + " – Financial Dashboard";
  $("co-name").textContent = data.company;
  $("subtitle").textContent = data.subtitle || `${Y[0]}–${fy} · ${CODE} millions unless stated · annual data`;
  document.querySelectorAll(".cur").forEach((e) => (e.textContent = CUR));
  document.querySelectorAll(".fy").forEach((e) => (e.textContent = fy));

  /* ----- derived series ----- */
  const D = {
    gm: map((i) => div(I.grossProfit[i], I.revenue[i])), om: map((i) => div(I.operatingIncome[i], I.revenue[i])),
    nm: map((i) => div(I.netIncome[i], I.revenue[i])), fcfm: map((i) => div(C.fcf[i], I.revenue[i])),
    roe: map((i) => div(I.netIncome[i], B.equity[i])), roa: map((i) => div(I.netIncome[i], B.totalAssets[i])),
    rnd: map((i) => div(I.rnd[i], I.revenue[i])), cur: map((i) => div(B.currentAssets[i], B.currentLiabilities[i])),
    de: map((i) => div(B.totalDebt[i], B.equity[i])), cov: map((i) => div(I.operatingIncome[i], -I.interestExpense[i])),
    inv: map((i) => div(B.inventory[i], I.costOfRevenue[i]) * 365), dso: map((i) => div(B.receivables[i], I.revenue[i]) * 365),
    dpo: map((i) => div(B.payables[i], I.costOfRevenue[i]) * 365), ocfNi: map((i) => div(C.ocf[i], I.netIncome[i])),
    capex: map((i) => div(-C.capex[i], I.revenue[i])), ret: map((i) => div(-C.buybacks[i] - C.dividends[i], C.fcf[i])),
    growth: map((i) => (i === 0 ? null : div(I.revenue[i], I.revenue[i - 1]) - 1)),
  };
  const cagr = Math.pow(I.revenue[L] / I.revenue[0], 1 / (N - 1)) - 1;
  const B_ = (a) => a.map((v) => +bn(v));

  /* ----- KPI cards ----- */
  const dg = (v) => (isFinite(v) ? { up: v >= 0, text: Math.abs(v * 100).toFixed(1) + "%" } : null);
  const dpp = (a, b) => (isFinite(a - b) ? { up: a - b >= 0, text: Math.abs((a - b) * 100).toFixed(1) + " pp" } : null);
  function renderKpis(containerId, cards) {
    const built = cards.map(({ label, value, delta, sub, spark, id }) => {
      const sid = "sp" + sparkN++;
      const d = delta ? `<span class="delta ${delta.up ? "up" : "down"}">${arrow(delta.up)} ${delta.text}</span>` : "";
      return { sid, spark, html: `<div class="card"><div class="label">${label}</div>
        <div class="kpi-value" ${id ? `id="${id}-v"` : ""}>${value}</div>
        <div>${d} <span class="kpi-sub" ${id ? `id="${id}-s"` : ""}>${sub || ""}</span></div>
        ${spark ? `<div class="spark"><canvas id="${sid}"></canvas></div>` : ""}</div>` };
    });
    $(containerId).innerHTML = built.map((b) => b.html).join("");
    built.forEach((b) => b.spark && mk($(b.sid), { type: "line",
      data: { labels: Y, datasets: [{ data: b.spark, borderColor: col.s1, borderWidth: 2, pointRadius: 0, tension: 0.35 }] },
      options: { responsive: true, maintainAspectRatio: false, events: [], plugins: { legend: { display: false }, tooltip: { enabled: false } },
                 scales: { x: { display: false }, y: { display: false } } } }));
  }
  renderKpis("kpi-fin", [
    { label: `Revenue ${fy}`, value: money(I.revenue[L]), delta: dg(yoy(I.revenue)), sub: `${N - 1}-yr CAGR ${pct(cagr, 0)}`, spark: I.revenue },
    { label: `Net income ${fy}`, value: money(I.netIncome[L]), delta: dg(yoy(I.netIncome)), sub: `Net margin ${pct(D.nm[L])}`, spark: I.netIncome },
    { label: `Diluted EPS ${fy}`, value: CUR + I.epsDiluted[L].toFixed(2), delta: dg(yoy(I.epsDiluted)), sub: `DPS ${CUR}${I.dps[L].toFixed(3)}`, spark: I.epsDiluted },
    { label: `Free cash flow ${fy}`, value: money(C.fcf[L]), delta: dg(yoy(C.fcf)), sub: `FCF margin ${pct(D.fcfm[L])}`, spark: C.fcf },
    { label: `Operating margin ${fy}`, value: pct(D.om[L]), delta: dpp(D.om[L], D.om[L - 1]), sub: `Gross ${pct(D.gm[L])}`, spark: D.om },
    { label: `Net cash ${fy}`, value: money(B.netCash[L]), delta: dg(yoy(B.netCash)), sub: `Debt ${money(B.totalDebt[L])}`, spark: B.netCash },
  ]);
  renderKpis("kpi-mkt", [
    { label: "Share price", value: "—", id: "m1" }, { label: "1-year change", value: "—", id: "m2" },
    { label: "52-week range (close)", value: "—", id: "m3" }, { label: "Market cap (approx.)", value: "—", id: "m4" },
    { label: `P/E (${fy} EPS)`, value: "—", id: "m5" }, { label: `FCF yield (${fy})`, value: "—", id: "m6" },
  ]);

  /* ----- charts ----- */
  mk($("revChart"), { type: "bar", data: { labels: Y, datasets: [bar("Revenue", B_(I.revenue), col.s1), bar("Net income", B_(I.netIncome), col.s3)] },
    options: chartOpts({ fmt: (v) => CUR + num(v, 1) + "B" }) });
  mk($("marginChart"), { type: "line", data: { labels: Y, datasets: [line("Gross margin", D.gm.map((v) => v * 100), col.s1), line("Operating margin", D.om.map((v) => v * 100), col.s3), line("Net margin", D.nm.map((v) => v * 100), col.s5)] },
    options: chartOpts({ fmt: (v) => num(v, 1) + "%" }) });
  mk($("growthChart"), { type: "bar", data: { labels: Y.slice(1), datasets: [bar("Revenue growth", D.growth.slice(1).map((v) => v * 100), col.s1)] },
    options: chartOpts({ fmt: (v) => num(v, 1) + "%", legend: false }) });
  mk($("cashChart"), { type: "bar", data: { labels: Y, datasets: [bar("Operating cash flow", B_(C.ocf), col.s1), bar("Free cash flow", B_(C.fcf), col.s3), bar("Capex", B_(C.capex.map((v) => -v)), col.s2)] },
    options: chartOpts({ fmt: (v) => CUR + num(v, 1) + "B" }) });
  mk($("returnChart"), { type: "bar", data: { labels: Y, datasets: [bar("Buybacks", B_(C.buybacks.map((v) => -v)), col.s1), bar("Dividends", B_(C.dividends.map((v) => -v)), col.s5)] },
    options: chartOpts({ fmt: (v) => CUR + num(v, 1) + "B", stacked: true }) });
  mk($("opexChart"), { type: "bar", data: { labels: Y, datasets: [bar("R&D", B_(I.rnd), col.s1), bar("SG&A", B_(I.sga), col.s2)] },
    options: chartOpts({ fmt: (v) => CUR + num(v, 1) + "B", stacked: true }) });
  mk($("liqChart"), { type: "bar", data: { labels: Y, datasets: [bar("Cash + securities", B_(Y.map((_, i) => B.cash[i] + B.securities[i])), col.s3), bar("Total debt", B_(B.totalDebt), col.s2)] },
    options: chartOpts({ fmt: (v) => CUR + num(v, 1) + "B" }) });
  mk($("daysChart"), { type: "line", data: { labels: Y, datasets: [line("Inventory days", D.inv, col.s4), line("Receivable days (DSO)", D.dso, col.s1), line("Payable days (DPO)", D.dpo, col.s5)] },
    options: chartOpts({ fmt: (v) => num(v, 0) + " d" }) });
  mk($("epsChart"), { type: "bar", data: { labels: Y, datasets: [bar("Diluted EPS", I.epsDiluted, col.s1)] },
    options: chartOpts({ fmt: (v) => CUR + num(v, 2), legend: false }) });

  const gauge = (cid, tid, value, color) => {
    const v = isFinite(value) ? Math.max(0, Math.min(1, value)) : 0;
    $(tid).textContent = pct(value);
    mk($(cid), { type: "doughnut", data: { datasets: [{ data: [v * 100, 100 - v * 100], backgroundColor: [color, "#252b33"], borderWidth: 0, borderRadius: 5 }] },
      options: { responsive: true, maintainAspectRatio: false, rotation: -90, circumference: 180, cutout: "75%",
                 plugins: { legend: { display: false }, tooltip: { enabled: false } } } });
  };
  gauge("g1", "g1t", D.om[L], col.s3); gauge("g2", "g2t", D.fcfm[L], col.s1); gauge("g3", "g3t", D.roe[L], col.s5);

  const cashSec = B.cash[L] + B.securities[L], ppeGw = B.ppe[L] + B.goodwill[L];
  const other = Math.max(0, B.totalAssets[L] - cashSec - B.receivables[L] - B.inventory[L] - B.ltInvestments[L] - ppeGw);
  const parts = [["Cash & securities", cashSec, col.s1], ["Receivables", B.receivables[L], col.s3], ["Inventory", B.inventory[L], col.s4],
                 ["Long-term investments", B.ltInvestments[L], col.s5], ["PP&E + goodwill", ppeGw, col.s2], ["Other", other, col.other]];
  mk($("assetChart"), { type: "doughnut",
    data: { labels: parts.map((p) => p[0]), datasets: [{ data: parts.map((p) => +bn(p[1])), backgroundColor: parts.map((p) => p[2]), borderColor: "#161b22", borderWidth: 2 }] },
    options: { responsive: true, maintainAspectRatio: false, cutout: "58%",
               plugins: { legend: { position: "right", labels: { boxWidth: 8, font: { size: 10 } } },
                          tooltip: { callbacks: { label: (c) => ` ${c.label}: ${CUR}${num(c.parsed, 1)}B` } } } } });

  /* ----- tables ----- */
  function trendBadge(arr, kind, hib) {
    const a = arr[L], b = arr[L - 1];
    const diff = kind === "pct" ? (a - b) * 100 : (a / b - 1) * 100;
    if (!isFinite(diff) || Math.abs(diff) < 0.05) return `<span class="badge flat">–</span>`;
    const up = diff > 0, cls = hib === null ? "flat" : up === hib ? "up" : "down";
    return `<span class="badge ${cls}">${arrow(up)} ${Math.abs(diff).toFixed(1)}${kind === "pct" ? " pp" : "%"}</span>`;
  }
  const rows = [
    ["Profitability"], ["Gross margin", D.gm, "pct", true], ["Operating margin", D.om, "pct", true], ["Net margin", D.nm, "pct", true],
    ["Return on equity", D.roe, "pct", true], ["Return on assets", D.roa, "pct", true], ["R&D % of revenue", D.rnd, "pct", null],
    ["Liquidity & leverage"], ["Current ratio", D.cur, "x", true], ["Debt / equity", D.de, "x2", false], ["Interest coverage", D.cov, "x0", true],
    ["Efficiency"], ["Inventory days", D.inv, "d", false], ["Receivable days", D.dso, "d", false], ["Payable days", D.dpo, "d", null],
    ["Cash quality"], ["FCF margin", D.fcfm, "pct", true], ["Operating cash flow / net income", D.ocfNi, "pct", true],
    ["Capex % of revenue", D.capex, "pct", null], ["Buybacks + dividends % of FCF", D.ret, "pct", null],
  ];
  const fmtVal = { pct: (v) => pct(v), x: (v) => num(v, 1) + "x", x2: (v) => num(v, 2) + "x", x0: (v) => num(v, 0) + "x", d: (v) => num(v, 0) };
  $("ratio-table").innerHTML = `<thead><tr><th></th>${Y.map((y) => `<th>${y}</th>`).join("")}<th>vs ${fyPrev}</th></tr></thead><tbody>` +
    rows.map((r) => r.length === 1 ? `<tr class="group"><td colspan="${N + 2}">${r[0]}</td></tr>`
      : `<tr><td>${r[0]}</td>${r[1].map((v) => `<td>${fmtVal[r[2]](v)}</td>`).join("")}<td>${trendBadge(r[1], r[2] === "pct" ? "pct" : "rel", r[3])}</td></tr>`).join("") + "</tbody>";

  const inc = [["Revenue", I.revenue], ["Gross profit", I.grossProfit], ["R&D", I.rnd], ["SG&A", I.sga],
               ["Operating income", I.operatingIncome], ["Net income", I.netIncome], [`Diluted EPS (${CUR})`, I.epsDiluted, 2]];
  $("income-table").innerHTML = `<thead><tr><th></th>${Y.map((y) => `<th>${y}</th>`).join("")}<th>YoY</th></tr></thead><tbody>` +
    inc.map(([n, a, d]) => {
      const g = yoy(a), cost = n === "R&D" || n === "SG&A";
      const badge = isFinite(g) ? `<span class="badge ${cost ? "flat" : g >= 0 ? "up" : "down"}">${arrow(g >= 0)} ${Math.abs(g * 100).toFixed(0)}%</span>` : `<span class="badge flat">–</span>`;
      return `<tr><td>${n}</td>${a.map((v) => `<td>${num(v, d ?? 0)}</td>`).join("")}<td>${badge}</td></tr>`;
    }).join("") + "</tbody>";

  /* ----- footnotes (built from the data) ----- */
  const notes = [`<li>Source: your Excel workbook. Not financial advice.</li>`,
    `<li>Free cash flow = operating cash flow minus capital expenditures (acquisitions and licensing purchases are not deducted).</li>`,
    `<li>Market cap ≈ latest price × ${fy} diluted shares; P/E uses ${fy} diluted EPS.</li>`];
  const gainShare = I.netIncome[L] ? I.investmentGain[L] / I.netIncome[L] : 0;
  if (gainShare > 0.03) notes.push(`<li>${fy} net income includes a ${money(I.investmentGain[L])} gain on investments (~${(gainShare * 100).toFixed(0)}% of net income), so operating income is the cleaner measure.</li>`);
  (data.notes || []).forEach((n) => notes.push(`<li>${esc(n)}</li>`));
  if (data.missing.length) notes.push(`<li>Rows not found in the workbook and treated as 0: ${data.missing.join("; ")}.</li>`);
  $("notes").innerHTML = `<ul>${notes.join("")}</ul>`;

  loadPrices();
}

/* =====================================================================
   PART 4 - share prices (Twelve Data, key in config.js)
   ===================================================================== */
function setEmpty(msg) {
  ["price-empty", "perf-empty"].forEach((id) => { $(id).innerHTML = msg; $(id).hidden = false; });
  $("priceChart").parentElement.hidden = true; $("range-tabs").hidden = true; $("perf-table").hidden = true;
  ["m1", "m2", "m3", "m4", "m5", "m6"].forEach((k) => { $(k + "-v").textContent = "—"; $(k + "-s").textContent = "needs price data"; });
  $("top-price").innerHTML = "";
}
function resetStockUI() {
  ["price-empty", "perf-empty"].forEach((id) => ($(id).hidden = true));
  $("priceChart").parentElement.hidden = false; $("range-tabs").hidden = false; $("perf-table").hidden = false;
}
const onOrBefore = (d) => { for (let i = prices.length - 1; i >= 0; i--) if (prices[i].d <= d) return prices[i]; return null; };
function shift(dateStr, { y = 0, m = 0, d = 0 }) {
  const t = new Date(dateStr + "T00:00:00Z");
  t.setUTCFullYear(t.getUTCFullYear() - y); t.setUTCMonth(t.getUTCMonth() - m); t.setUTCDate(t.getUTCDate() - d);
  return t.toISOString().slice(0, 10);
}
const chgHtml = (v) => (v == null ? `<span class="muted">n/a</span>` : `<span class="badge ${v >= 0 ? "up" : "down"}">${arrow(v >= 0)} ${Math.abs(v * 100).toFixed(1)}%</span>`);

function drawPrice(n) {
  const pts = prices.slice(-n), ctx = $("priceChart").getContext("2d");
  const g = ctx.createLinearGradient(0, 0, 0, 300);
  g.addColorStop(0, "rgba(57,135,229,.30)"); g.addColorStop(1, "rgba(57,135,229,0)");
  if (priceChart) { priceChart.destroy(); charts.splice(charts.indexOf(priceChart), 1); }
  priceChart = mk(ctx, { type: "line",
    data: { labels: pts.map((p) => p.d), datasets: [{ label: "Close", data: pts.map((p) => p.c), borderColor: col.s1, backgroundColor: g, fill: true, borderWidth: 2, tension: 0.15, pointRadius: 0, pointHoverRadius: 4 }] },
    options: { responsive: true, maintainAspectRatio: false, interaction: { mode: "index", intersect: false },
      plugins: { legend: { display: false }, tooltip: { callbacks: { label: (c) => ` Close: ${CUR}${num(c.parsed.y, 2)}` } } },
      scales: { x: { grid: { display: false }, ticks: { maxTicksLimit: 7, maxRotation: 0 } }, y: { grid: { color: col.grid }, ticks: { callback: (v) => CUR + v } } } } });
}

function renderStock() {
  const I = DATA.income, B = DATA.balance, C = DATA.cashflow, L = DATA.years.length - 1, fy = DATA.years[L];
  const last = prices[prices.length - 1], prev = prices[prices.length - 2], dayChg = last.c / prev.c - 1;
  const ref = (o) => { const p = onOrBefore(shift(last.d, o)); return p && p.d >= prices[0].d ? last.c / p.c - 1 : null; };
  const ytdBase = onOrBefore(last.d.slice(0, 4) - 1 + "-12-31");
  const perf = [["1 day", dayChg], ["1 week", ref({ d: 7 })], ["1 month", ref({ m: 1 })], ["3 months", ref({ m: 3 })], ["6 months", ref({ m: 6 })],
                ["Year to date", ytdBase ? last.c / ytdBase.c - 1 : null], ["1 year", ref({ y: 1 })], ["5 years", ref({ y: 5 })]];
  $("perf-table").querySelector("tbody").innerHTML = perf.map(([n, v]) => `<tr><td>${n}</td><td>${chgHtml(v)}</td></tr>`).join("");

  const yr = prices.slice(-252), hi = Math.max(...yr.map((p) => p.c)), lo = Math.min(...yr.map((p) => p.c));
  const mcap = last.c * I.dilutedShares[L], y1 = ref({ y: 1 });
  $("top-price").innerHTML = `<div class="big">${CUR}${num(last.c, 2)}</div><div class="delta ${dayChg >= 0 ? "up" : "down"}">${arrow(dayChg >= 0)} ${Math.abs(dayChg * 100).toFixed(2)}% <span class="muted">close ${last.d}</span></div>`;
  $("m1-v").textContent = CUR + num(last.c, 2);
  $("m1-s").innerHTML = `<span class="delta ${dayChg >= 0 ? "up" : "down"}">${arrow(dayChg >= 0)} ${Math.abs(dayChg * 100).toFixed(2)}%</span> ${last.d}`;
  $("m2-v").textContent = y1 == null ? "n/a" : (y1 >= 0 ? "+" : "−") + Math.abs(y1 * 100).toFixed(1) + "%";
  $("m2-s").textContent = "vs. one year ago";
  $("m3-v").textContent = `${CUR}${num(lo, 0)} – ${CUR}${num(hi, 0)}`;
  $("m3-s").textContent = `now at ${((last.c - lo) / (hi - lo) * 100).toFixed(0)}% of range`;
  $("m4-v").textContent = money(mcap);
  $("m4-s").textContent = `P/S ${num(mcap / I.revenue[L], 1)}x · P/B ${num(mcap / B.equity[L], 1)}x`;
  $("m5-v").textContent = num(last.c / I.epsDiluted[L], 1) + "x";
  $("m5-s").textContent = `${fy} EPS ${CUR}${I.epsDiluted[L].toFixed(2)}`;
  $("m6-v").textContent = pct(C.fcf[L] / mcap, 2);
  $("m6-s").textContent = `Div. yield ${pct(I.dps[L] / last.c, 2)}`;

  [...$("range-tabs").children].forEach((b) => b.classList.toggle("on", b.dataset.r === "252"));
  drawPrice(252);
  $("range-tabs").onclick = (e) => {
    if (!e.target.dataset.r) return;
    [...$("range-tabs").children].forEach((b) => b.classList.toggle("on", b === e.target));
    drawPrice(+e.target.dataset.r);
  };
}

let priceRun = 0;

/* Best saved prices for a symbol: prices-data.js snapshot or browser cache, whichever is newer */
function savedPrices(symbol) {
  const cached = loadPriceCache(symbol);
  const lib = typeof PRICE_LIBRARY !== "undefined" && PRICE_LIBRARY[symbol];
  const cachedOk = cached && cached.prices && cached.prices.length > 1;
  const libOk = lib && lib.prices && lib.prices.length > 1;
  if (!cachedOk && !libOk) return null;
  const lastDay = (p) => p[p.length - 1].d;
  const useLib = libOk && (!cachedOk || lastDay(lib.prices) >= lastDay(cached.prices));
  return useLib
    ? { prices: lib.prices, note: `saved snapshot (prices-data.js, ${lib.updated})` }
    : { prices: cached.prices, note: `browser cache (${cacheAge(cached.timestamp)})` };
}

async function loadPrices() {
  if (!DATA) return;
  const run = ++priceRun, symbol = $("ticker").value.trim().toUpperCase();
  $("co-tick").textContent = symbol;
  resetStockUI();
  if (!symbol) return setEmpty(`Choose a company from the list at the top to load its share prices.`);

  const showSaved = (why) => {
    const s = savedPrices(symbol);
    if (!s) return false;
    prices = s.prices;
    renderStock();
    status(`${why} Showing ${s.note}.`);
    return true;
  };

  // No API key: use the saved data only
  if (!CONFIG.TWELVE_DATA_KEY) {
    if (showSaved("No Twelve Data key in config.js.")) return;
    return setEmpty(`No saved prices for ${symbol}.<br>Add it to <code>prices-data.js</code>, or paste a Twelve Data key into <code>config.js</code>.`);
  }

  // With a key: try live data, fall back to saved data
  try {
    const res = await fetch(`https://api.twelvedata.com/time_series?symbol=${encodeURIComponent(symbol)}&interval=1day&outputsize=1300&apikey=${CONFIG.TWELVE_DATA_KEY}`);
    const json = await res.json();
    if (run !== priceRun) return;
    if (json.status === "error" || !json.values) throw new Error(json.message || "No data returned");
    prices = json.values.map((v) => ({ d: v.datetime.slice(0, 10), c: parseFloat(v.close) })).reverse();
    savePriceCache(symbol, prices);
    renderStock();
  } catch (err) {
    if (run !== priceRun) return;
    if (showSaved(`Live prices unavailable (${err.message}).`)) return;
    setEmpty(`Could not load prices: ${err.message}`);
  }
}

/* =====================================================================
   PART 5 - loading the file (button, Live Server auto-load, last file)
   ===================================================================== */
function status(msg, isError) { const s = $("status"); s.textContent = msg || ""; s.className = "status" + (isError ? " err" : ""); }

function openData(data, fromName) {
  const prevCompany = store.get("dash-company");
  const name = data.company.toLowerCase();
  const match = TICKERS.find((t) => [t.name, ...(t.aka || [])].some((n) => name.includes(n.toLowerCase())));
  if (data.ticker) setTicker(data.ticker);                                    // ticker given in the Settings sheet
  else if (match) setTicker(match.symbol);                                    // company found in the list
  else if (prevCompany && prevCompany !== data.company) setTicker("");        // unknown new company: choose manually
  store.set("dash-company", data.company); store.set("dash-ticker", $("ticker").value);
  store.set("dash-data", JSON.stringify(data));
  status(`Loaded ${fromName}`);
  render(data);
}

async function loadFromBuffer(buf, name) {
  try { openData(parseWorkbook(XLSX.read(buf, { type: "array" }), name), name); }
  catch (e) { console.error(e); status(e.message, true); }
}

$("file").addEventListener("change", async (e) => {
  const f = e.target.files[0];
  if (f) await loadFromBuffer(await f.arrayBuffer(), f.name);
  e.target.value = "";
});
/* Companies stored in data.js: no Excel upload needed */
function loadFromLibrary(sym) {
  const d = typeof DATA_LIBRARY !== "undefined" && DATA_LIBRARY[sym];
  if (!d) return false;
  openData(JSON.parse(JSON.stringify(d)), "data.js (" + sym + ")");
  return true;
}

$("ticker").addEventListener("change", async () => {
  const sym = $("ticker").value;
  store.set("dash-ticker", sym);
  if (loadFromLibrary(sym)) return;

  // Find if this ticker has a pre-linked Excel file in config.js
  const match = TICKERS.find(t => t.symbol === sym);
  if (match && match.file && location.protocol.startsWith("http")) {
    try {
      const response = await fetch(match.file);
      if (response.ok) {
        const buffer = await response.arrayBuffer();
        loadFromBuffer(buffer, match.file);
        return; // Stop here once loaded
      }
    } catch (err) {
      console.log("Could not auto-load repo file, falling back to price update only.");
    }
  }

  // Otherwise, just update the live stock prices for the current data
  loadPrices();
});

(async function start() {
  fillTickers();
  setTicker(store.get("dash-ticker") || CONFIG.SYMBOL || "");
  if (loadFromLibrary($("ticker").value)) return;
  if (location.protocol.startsWith("http")) {                  // served by Live Server: try data.xlsx in this folder
    try {
      const r = await fetch("data.xlsx");
      if (r.ok) return loadFromBuffer(await r.arrayBuffer(), "data.xlsx");
    } catch (e) { /* no data.xlsx - fall through */ }
  }
  const saved = store.get("dash-data");                         // reopen the last file you loaded
  if (saved) { try { render(JSON.parse(saved)); status("Showing the last file you opened"); } catch (e) {} }
})();