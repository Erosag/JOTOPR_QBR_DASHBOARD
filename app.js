const fileInput = document.getElementById("fileInput");
const statusEl = document.getElementById("status");
const dashboardEl = document.getElementById("dashboard");
const shareBtn = document.getElementById("shareBtn");
const summaryBody = document.getElementById("summaryBody");
const placementsBody = document.getElementById("placementsBody");
const mentionsLegend = document.getElementById("mentionsLegend");
const audienceLegend = document.getElementById("audienceLegend");
const publicityLegend = document.getElementById("publicityLegend");
const tierFilter = document.getElementById("tierFilter");
const dateFromFilter = document.getElementById("dateFromFilter");
const dateToFilter = document.getElementById("dateToFilter");
const clearDateFiltersBtn = document.getElementById("clearDateFiltersBtn");
const exportPdfBtn = document.getElementById("exportPdfBtn");
const clientNameInput = document.getElementById("clientNameInput");

const charts = {
  mentions: null,
  audience: null,
  publicity: null,
  tier: null,
};

const TIER_ORDER = ["III", "IV", "V", "VI"];

const TIER_MULTIPLIERS = {
  "3": 1000,
  iii: 1000,
  "tier iii": 1000,
  "4": 1500,
  iv: 1500,
  "tier iv": 1500,
  "5": 3500,
  v: 3500,
  "tier v": 3500,
  "6": 7500,
  vi: 7500,
  "tier vi": 7500,
};

const columnAliases = {
  mediaType: ["media type", "mediatype", "media-type", "media_type", "media", "type", "type of media", "type pf media"],
  audience: [
    "audience/uvpm",
    "audience / uvpm",
    "audience/umpv",
    "audience / umpv",
    "audience/umvp",
    "audience / umvp",
    "audience uvpm",
    "audience umpv",
    "audience umvp",
    "audienceuvpm",
    "audienceumpv",
    "audienceumvp",
    "uvpm",
    "umpv",
    "umvp",
    "audience",
    "audience total",
    "total audience",
    "estimated audience",
    "audience estimate",
    "audience size",
    "audience reach",
    "uvpm audience",
    "audience uvpm total",
    "reach",
    "impressions",
  ],
  inches: ["inches", "inch", "size inches", "placement inches"],
  tier: ["tier", "media tier"],
  publicity: ["publicity value", "publicity", "pubicity value", "ave"],
  date: ["date", "publish date", "publication date"],
  publication: ["station/publication", "station / publication", "publication", "station", "outlet"],
  title: ["title", "headline"],
  link: ["link", "url"],
};

const centerTextPlugin = {
  id: "centerTextPlugin",
  afterDatasetsDraw(chart, args, pluginOptions) {
    const { ctx } = chart;
    const meta = chart.getDatasetMeta(0);
    if (!meta?.data?.length) return;

    const { x, y } = meta.data[0];
    const title = pluginOptions?.title || "";
    const value = pluginOptions?.value || "";

    const valueFontSize = Math.max(20, Math.min(44, 700 / Math.max(1, String(value).length)));
    const titleFontSize = Math.max(13, Math.min(20, 320 / Math.max(1, String(title).length)));

    ctx.save();
    ctx.textAlign = "center";
    ctx.textBaseline = "middle";
    ctx.fillStyle = "#101219";
    ctx.font = `700 ${valueFontSize}px Space Grotesk`;
    ctx.fillText(value, x, y - 8);
    ctx.font = `500 ${titleFontSize}px Space Grotesk`;
    ctx.fillText(title, x, y + 22);
    ctx.restore();
  },
};

Chart.register(centerTextPlugin);

let loadedRows = [];
let filteredRows = [];
let columnMap = null;
let currentFileLabel = "";

function setStatus(text, type = "") {
  statusEl.textContent = text;
  statusEl.className = `status ${type}`.trim();
}

function normalize(value) {
  return String(value || "")
    .trim()
    .toLowerCase();
}

function normalizeHeader(value) {
  return String(value || "")
    .replace(/^\uFEFF/, "")
    .replace(/\u00A0/g, " ")
    .replace(/["'`]/g, "")
    .replace(/[\r\n]+/g, " ")
    .replace(/[_-]+/g, " ")
    .replace(/\s*\/\s*/g, "/")
    .replace(/\s*&\s*/g, " and ")
    .replace(/[^\w/\s]+/g, " ")
    .replace(/\s+/g, " ")
    .trim()
    .toLowerCase();
}

function canonicalHeader(value) {
  return normalizeHeader(value).replace(/[^a-z0-9]/g, "");
}

function parseNumber(value) {
  if (typeof value === "number") return value;
  if (!value) return 0;

  let raw = String(value).trim();
  if (!raw) return 0;

  let multiplier = 1;
  const suffixMatch = raw.match(/([kmb])\s*$/i);
  if (suffixMatch) {
    const suffix = suffixMatch[1].toLowerCase();
    if (suffix === "k") multiplier = 1_000;
    if (suffix === "m") multiplier = 1_000_000;
    if (suffix === "b") multiplier = 1_000_000_000;
    raw = raw.replace(/([kmb])\s*$/i, "");
  }

  raw = raw.replace(/[^\d,.\-]/g, "");

  const commaCount = (raw.match(/,/g) || []).length;
  const dotCount = (raw.match(/\./g) || []).length;

  if (commaCount > 0 && dotCount > 0) {
    if (raw.lastIndexOf(",") > raw.lastIndexOf(".")) {
      raw = raw.replace(/\./g, "").replace(",", ".");
    } else {
      raw = raw.replace(/,/g, "");
    }
  } else if (commaCount > 0) {
    if (commaCount > 1) {
      raw = raw.replace(/,/g, "");
    } else {
      const [left, right] = raw.split(",");
      raw = right && right.length <= 2 ? `${left}.${right}` : `${left}${right || ""}`;
    }
  } else if (dotCount > 0) {
    if (dotCount > 1) {
      raw = raw.replace(/\./g, "");
    } else {
      const [left, right] = raw.split(".");
      raw = right && right.length === 3 ? `${left}${right}` : raw;
    }
  }

  const parsed = Number(raw);
  return Number.isFinite(parsed) ? parsed * multiplier : 0;
}

function findColumn(headers, aliasList) {
  for (const alias of aliasList) {
    const match = headers.find((h) => normalizeHeader(h) === normalizeHeader(alias));
    if (match) return match;
  }

  for (const alias of aliasList) {
    const aliasCanonical = canonicalHeader(alias);
    const match = headers.find((h) => canonicalHeader(h) === aliasCanonical);
    if (match) return match;
  }

  for (const alias of aliasList) {
    const match = headers.find((h) => normalizeHeader(h).includes(normalizeHeader(alias)));
    if (match) return match;
  }

  for (const alias of aliasList) {
    const aliasCanonical = canonicalHeader(alias);
    const match = headers.find((h) => canonicalHeader(h).includes(aliasCanonical));
    if (match) return match;
  }

  return null;
}

function headerMatchesAlias(header, alias) {
  const normalizedHeader = normalizeHeader(header);
  const normalizedAlias = normalizeHeader(alias);
  if (!normalizedHeader || !normalizedAlias) return false;
  if (normalizedHeader === normalizedAlias) return true;

  const canonicalH = canonicalHeader(header);
  const canonicalA = canonicalHeader(alias);
  if (canonicalH === canonicalA) return true;
  if (canonicalH.includes(canonicalA) || canonicalA.includes(canonicalH)) return true;
  if (normalizedHeader.includes(normalizedAlias) || normalizedAlias.includes(normalizedHeader)) return true;
  return false;
}

function findBestNumericColumn(rows, headers, aliasList) {
  const candidates = headers.filter((header) => aliasList.some((alias) => headerMatchesAlias(header, alias)));
  if (!candidates.length) return null;

  let best = null;
  for (const header of candidates) {
    let nonZeroCount = 0;
    let sum = 0;

    for (const row of rows) {
      const value = parseNumber(row[header]);
      if (value !== 0) nonZeroCount += 1;
      sum += Math.abs(value);
    }

    const score = { header, nonZeroCount, sum };
    if (
      !best ||
      score.nonZeroCount > best.nonZeroCount ||
      (score.nonZeroCount === best.nonZeroCount && score.sum > best.sum)
    ) {
      best = score;
    }
  }

  return best?.header || null;
}

function getMatchingColumns(headers, aliasList) {
  return headers.filter((header) => aliasList.some((alias) => headerMatchesAlias(header, alias)));
}

function looksLikeAudienceHeader(header) {
  const key = canonicalHeader(header);
  if (!key) return false;

  if (key.includes("audience")) return true;
  if (key.includes("uvpm") || key.includes("umpv") || key.includes("umvp")) return true;
  if (key.includes("reach")) return true;
  if (key.includes("impressions")) return true;
  return false;
}

function findBestNumericColumnFromCandidates(rows, candidates) {
  if (!Array.isArray(candidates) || candidates.length === 0) return null;

  let best = null;
  for (const header of candidates) {
    let nonZeroCount = 0;
    let sum = 0;

    for (const row of rows) {
      const value = parseNumber(row[header]);
      if (value !== 0) nonZeroCount += 1;
      sum += Math.abs(value);
    }

    const score = { header, nonZeroCount, sum };
    if (
      !best ||
      score.nonZeroCount > best.nonZeroCount ||
      (score.nonZeroCount === best.nonZeroCount && score.sum > best.sum)
    ) {
      best = score;
    }
  }

  return best?.header || null;
}

function normalizeTier(rawTier) {
  const key = normalize(rawTier).replace(/\s+/g, " ");
  if (["iii", "tier iii", "3"].includes(key)) return "III";
  if (["iv", "tier iv", "4"].includes(key)) return "IV";
  if (["v", "tier v", "5"].includes(key)) return "V";
  if (["vi", "tier vi", "6"].includes(key)) return "VI";
  return "";
}

function tierMultiplier(rawTier) {
  const key = normalize(rawTier).replace(/\s+/g, " ");
  return TIER_MULTIPLIERS[key] || 0;
}

function formatInt(num) {
  return new Intl.NumberFormat("en-US", { maximumFractionDigits: 0 }).format(num);
}

function formatCurrency(num) {
  return new Intl.NumberFormat("en-US", {
    style: "currency",
    currency: "USD",
    maximumFractionDigits: 0,
  }).format(num);
}

function formatMDY(date) {
  return new Intl.DateTimeFormat("en-US", {
    month: "numeric",
    day: "numeric",
    year: "numeric",
  }).format(date);
}

function getClientName() {
  const typed = String(clientNameInput?.value || "").trim();
  if (typed) return typed;

  const fromFile = String(currentFileLabel || "").replace(/\.[^/.]+$/, "").trim();
  return fromFile || "Client";
}

function detectColumns(rows) {
  if (!rows.length) throw new Error("The uploaded sheet has no rows.");

  const headers = Object.keys(rows[0]);
  const map = {
    mediaType: findColumn(headers, columnAliases.mediaType),
    audience: findColumn(headers, columnAliases.audience),
    publicity: findColumn(headers, columnAliases.publicity),
    inches: findColumn(headers, columnAliases.inches),
    tier: findColumn(headers, columnAliases.tier),
    date: findColumn(headers, columnAliases.date),
    publication: findColumn(headers, columnAliases.publication),
    title: findColumn(headers, columnAliases.title),
    link: findColumn(headers, columnAliases.link),
  };

  const aliasAudienceCandidates = getMatchingColumns(headers, columnAliases.audience);
  const keywordAudienceCandidates = headers.filter((header) => looksLikeAudienceHeader(header));
  map.audienceCandidates = [...new Set([...aliasAudienceCandidates, ...keywordAudienceCandidates])];

  const bestAudience = findBestNumericColumnFromCandidates(rows, map.audienceCandidates) || findBestNumericColumn(rows, headers, columnAliases.audience);
  if (bestAudience) map.audience = bestAudience;

  const missing = [];
  if (!map.mediaType) missing.push("Media Type");
  if (!map.audience) missing.push("Audience/UVPM (or Audience/UMPV)");
  if (!map.publicity && (!map.inches || !map.tier)) {
    missing.push("Publicity Value (or Inches + Tier)");
  }
  if (!map.publication) missing.push("Station/Publication");

  if (missing.length > 0) {
    const availableHeaders = headers.map((h) => String(h).trim()).filter(Boolean).join(", ");
    throw new Error(
      `Missing required column(s): ${missing.join(", ")}. Detected headers: ${availableHeaders || "(none)"}`
    );
  }

  return map;
}

function rowAudience(row, map) {
  const primary = parseNumber(row[map.audience]);
  if (primary !== 0) return primary;

  const candidates = Array.isArray(map.audienceCandidates) ? map.audienceCandidates : [];
  let fallback = 0;

  for (const candidate of candidates) {
    if (candidate === map.audience) continue;
    const value = parseNumber(row[candidate]);
    if (Math.abs(value) > Math.abs(fallback)) fallback = value;
  }

  if (fallback !== 0) return fallback;

  for (const [header, valueRaw] of Object.entries(row)) {
    if (!looksLikeAudienceHeader(header)) continue;
    const value = parseNumber(valueRaw);
    if (Math.abs(value) > Math.abs(fallback)) fallback = value;
  }

  return fallback;
}

function toDate(value) {
  if (value instanceof Date && !Number.isNaN(value.getTime())) return value;

  if (typeof value === "number" && Number.isFinite(value)) {
    if (value > 20000 && value < 100000) {
      const utcDays = Math.floor(value - 25569);
      return new Date(utcDays * 86400 * 1000);
    }
    const fromEpoch = new Date(value);
    if (!Number.isNaN(fromEpoch.getTime())) return fromEpoch;
  }

  const raw = String(value || "").trim();
  if (!raw) return null;

  const parsed = new Date(raw);
  if (!Number.isNaN(parsed.getTime())) return parsed;

  const mdy = raw.match(/^(\d{1,2})[/-](\d{1,2})[/-](\d{2,4})$/);
  if (mdy) {
    const mm = Number(mdy[1]);
    const dd = Number(mdy[2]);
    const yyyy = Number(mdy[3].length === 2 ? `20${mdy[3]}` : mdy[3]);
    const d = new Date(yyyy, mm - 1, dd);
    if (!Number.isNaN(d.getTime())) return d;
  }

  return null;
}

function rowPublicity(row, map) {
  if (map.publicity) return parseNumber(row[map.publicity]);
  const inches = parseNumber(row[map.inches]);
  const multiplier = tierMultiplier(row[map.tier]);
  return inches * multiplier;
}

function rowDateLabel(row, map) {
  if (!map.date) return "";
  const d = toDate(row[map.date]);
  if (!d) return String(row[map.date] || "");
  return formatMDY(d);
}

function mediaColor(label) {
  const key = normalize(label);
  if (key.includes("tv")) return "#2E86DE";
  if (key.includes("online") || key.includes("print")) return "#27AE60";
  if (key.includes("radio")) return "#F5B041";
  if (key.includes("podcast")) return "#8E44AD";
  if (key.includes("youtube")) return "#5D6D7E";
  return "#16A085";
}

function buildAggregation(rows, map) {
  const grouped = new Map();
  const tierCounts = { III: 0, IV: 0, V: 0, VI: 0 };

  for (const row of rows) {
    const mediaType = String(row[map.mediaType] || "Unknown").trim() || "Unknown";
    const audience = rowAudience(row, map);
    const publicity = rowPublicity(row, map);
    const tier = normalizeTier(row[map.tier]);

    if (!grouped.has(mediaType)) {
      grouped.set(mediaType, { mentions: 0, audience: 0, publicity: 0 });
    }

    const current = grouped.get(mediaType);
    current.mentions += 1;
    current.audience += audience;
    current.publicity += publicity;

    if (tierCounts[tier] !== undefined) tierCounts[tier] += 1;
  }

  const byType = Array.from(grouped.entries())
    .map(([mediaType, values]) => ({ mediaType, ...values }))
    .sort((a, b) => b.mentions - a.mentions);

  const byTier = TIER_ORDER.map((tier) => ({ tier, mentions: tierCounts[tier] || 0 }));

  const totals = byType.reduce(
    (acc, item) => {
      acc.mentions += item.mentions;
      acc.audience += item.audience;
      acc.publicity += item.publicity;
      return acc;
    },
    { mentions: 0, audience: 0, publicity: 0 },
  );

  return { byType, byTier, totals };
}

function destroyIfExists(chart) {
  if (chart) chart.destroy();
}

function makeDonutChart(canvasId, labels, data, centerTitle, centerValue, colors = null) {
  return new Chart(document.getElementById(canvasId), {
    type: "doughnut",
    data: {
      labels,
      datasets: [
        {
          data,
          backgroundColor: colors || labels.map((label) => mediaColor(label)),
          borderWidth: 2,
          borderColor: "#f7f9ff",
          hoverOffset: 8,
        },
      ],
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      cutout: "72%",
      plugins: {
        legend: {
          display: false,
        },
        tooltip: {
          bodyFont: { family: "Space Grotesk" },
          titleFont: { family: "Space Grotesk" },
        },
        centerTextPlugin: {
          title: centerTitle,
          value: centerValue,
        },
      },
    },
  });
}

function escapeHtml(value) {
  return String(value || "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}

function safeUrl(url) {
  try {
    const parsed = new URL(String(url || "").trim());
    if (parsed.protocol === "http:" || parsed.protocol === "https:") return parsed.toString();
  } catch {
    return "";
  }
  return "";
}

function renderSummaryTable(byType) {
  summaryBody.innerHTML = "";

  for (const row of byType) {
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${escapeHtml(row.mediaType)}</td>
      <td>${formatInt(row.mentions)}</td>
      <td>${formatInt(row.audience)}</td>
      <td>${formatCurrency(row.publicity)}</td>
    `;
    summaryBody.appendChild(tr);
  }
}

function renderPlacementTable(rows, map) {
  placementsBody.innerHTML = "";

  for (const row of rows) {
    const date = rowDateLabel(row, map);
    const publication = map.publication ? String(row[map.publication] || "") : "";
    const title = map.title ? String(row[map.title] || "") : "";
    const link = map.link ? safeUrl(row[map.link]) : "";
    const mediaType = String(row[map.mediaType] || "Unknown");
    const audience = rowAudience(row, map);
    const publicity = rowPublicity(row, map);

    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${escapeHtml(date)}</td>
      <td>${escapeHtml(publication)}</td>
      <td>${escapeHtml(title)}</td>
      <td>${link ? `<a href="${escapeHtml(link)}" target="_blank" rel="noopener noreferrer">${escapeHtml(link)}</a>` : ""}</td>
      <td>${escapeHtml(mediaType)}</td>
      <td>${formatInt(audience)}</td>
      <td>${formatCurrency(publicity)}</td>
    `;
    placementsBody.appendChild(tr);
  }
}

function renderMetricLegend(targetEl, byType, valueKey, formatValue) {
  if (!targetEl) return;

  const rows = [...byType].sort((a, b) => (b[valueKey] || 0) - (a[valueKey] || 0));
  targetEl.innerHTML = rows
    .map(
      (row) => `
        <div class="legend-item">
          <span class="legend-name">
            <span class="legend-dot" style="background:${mediaColor(row.mediaType)}"></span>
            <span>${escapeHtml(row.mediaType)}</span>
          </span>
          <span class="legend-value">${formatValue(row[valueKey] || 0)}</span>
        </div>
      `,
    )
    .join("");
}

function renderDashboard(report) {
  const mediaLabels = report.byType.map((x) => x.mediaType);

  destroyIfExists(charts.mentions);
  destroyIfExists(charts.audience);
  destroyIfExists(charts.publicity);
  destroyIfExists(charts.tier);

  charts.mentions = makeDonutChart(
    "mentionsChart",
    mediaLabels,
    report.byType.map((x) => x.mentions),
    "Mentions",
    formatInt(report.totals.mentions),
  );

  charts.audience = makeDonutChart(
    "audienceChart",
    mediaLabels,
    report.byType.map((x) => x.audience),
    "Audience",
    formatInt(report.totals.audience),
  );

  charts.publicity = makeDonutChart(
    "publicityChart",
    mediaLabels,
    report.byType.map((x) => x.publicity),
    "Publicity",
    formatCurrency(report.totals.publicity),
  );

  charts.tier = makeDonutChart(
    "tierChart",
    report.byTier.map((x) => x.tier),
    report.byTier.map((x) => x.mentions),
    "All Tiers",
    formatInt(report.byTier.reduce((sum, x) => sum + x.mentions, 0)),
    ["#6C9A8B", "#4C6EF5", "#F59F00", "#C2255C"],
  );

  renderSummaryTable(report.byType);
  renderMetricLegend(mentionsLegend, report.byType, "mentions", formatInt);
  renderMetricLegend(audienceLegend, report.byType, "audience", formatInt);
  renderMetricLegend(publicityLegend, report.byType, "publicity", formatCurrency);
  dashboardEl.classList.remove("hidden");
}

function applyFilters(rows, map, options = {}) {
  const selectedTier = options.ignoreTier ? "" : options.tierOverride ?? tierFilter.value;
  const fromRaw = dateFromFilter.value;
  const toRaw = dateToFilter.value;
  const fromDate = fromRaw ? new Date(`${fromRaw}T00:00:00`) : null;
  const toDateValue = toRaw ? new Date(`${toRaw}T23:59:59`) : null;

  if (fromDate && toDateValue && fromDate > toDateValue) {
    setStatus("Date From cannot be after Date To.", "error");
    return [];
  }

  return rows.filter((row) => {
    const tier = map.tier ? normalizeTier(row[map.tier]) : "";
    const tierOk = !selectedTier || tier === selectedTier;

    if (!tierOk) return false;

    if (!map.date || (!fromDate && !toDateValue)) return true;

    const d = toDate(row[map.date]);
    if (!d) return false;

    if (fromDate && d < fromDate) return false;
    if (toDateValue && d > toDateValue) return false;
    return true;
  });
}

function refreshDashboard() {
  if (!loadedRows.length || !columnMap) return;

  filteredRows = applyFilters(loadedRows, columnMap);
  const report = buildAggregation(filteredRows, columnMap);
  renderDashboard(report);
  renderPlacementTable(filteredRows, columnMap);

  if (columnMap.date) {
    setStatus(`Dashboard updated (${filteredRows.length} rows). Audience column: ${columnMap.audience || "not found"}.`, "ok");
  } else {
    setStatus(`Dashboard updated (${filteredRows.length} rows). Audience column: ${columnMap.audience || "not found"}. No Date column found, date filter disabled.`, "ok");
  }
}

function encodeState(payload) {
  const json = JSON.stringify(payload);
  const bytes = new TextEncoder().encode(json);
  let binary = "";
  for (const byte of bytes) binary += String.fromCharCode(byte);
  return btoa(binary);
}

function decodeState(encoded) {
  const binary = atob(encoded);
  const bytes = Uint8Array.from(binary, (char) => char.charCodeAt(0));
  const json = new TextDecoder().decode(bytes);
  return JSON.parse(json);
}

function enableShare() {
  shareBtn.disabled = false;
  shareBtn.onclick = async () => {
    const payload = {
      rows: loadedRows,
      map: columnMap,
      fileLabel: currentFileLabel,
      clientName: getClientName(),
      dateFrom: dateFromFilter.value,
      dateTo: dateToFilter.value,
      tier: tierFilter.value,
    };

    const url = new URL(window.location.href);
    url.searchParams.set("state", encodeState(payload));

    try {
      await navigator.clipboard.writeText(url.toString());
      setStatus("Share link copied. Anyone with the link can view this dashboard.", "ok");
    } catch {
      setStatus("Could not copy link. You can copy it manually from the browser address bar.", "error");
      window.history.replaceState({}, "", url.toString());
    }
  };
}

function parseWorkbook(fileBuffer) {
  const workbook = XLSX.read(fileBuffer, { type: "array" });
  const firstSheet = workbook.SheetNames[0];
  const worksheet = workbook.Sheets[firstSheet];
  return XLSX.utils.sheet_to_json(worksheet, { defval: "" });
}

function parseCsv(csvText) {
  const cleaned = String(csvText || "").replace(/^\uFEFF/, "").trim();
  if (!cleaned) throw new Error("CSV is empty.");

  const firstLine = cleaned.split(/\r\n|\n|\r/)[0] || "";
  const commaCount = (firstLine.match(/,/g) || []).length;
  const semicolonCount = (firstLine.match(/;/g) || []).length;
  const tabCount = (firstLine.match(/\t/g) || []).length;
  const delimiter = tabCount > commaCount && tabCount > semicolonCount ? "\t" : semicolonCount > commaCount ? ";" : ",";

  const rows = [];
  let row = [];
  let current = "";
  let inQuotes = false;

  for (let i = 0; i < cleaned.length; i += 1) {
    const char = cleaned[i];
    const next = cleaned[i + 1];

    if (char === '"') {
      if (inQuotes && next === '"') {
        current += '"';
        i += 1;
      } else {
        inQuotes = !inQuotes;
      }
    } else if (char === delimiter && !inQuotes) {
      row.push(current.trim());
      current = "";
    } else if ((char === "\n" || char === "\r") && !inQuotes) {
      if (char === "\r" && next === "\n") i += 1;
      row.push(current.trim());
      current = "";
      if (row.some((cell) => cell !== "")) rows.push(row);
      row = [];
    } else {
      current += char;
    }
  }

  if (current.length > 0 || row.length > 0) {
    row.push(current.trim());
    if (row.some((cell) => cell !== "")) rows.push(row);
  }

  if (rows.length < 2) {
    throw new Error("CSV must include a header row and at least one data row.");
  }

  const headers = rows[0];
  return rows.slice(1).map((values) => {
    const result = {};
    headers.forEach((header, idx) => {
      result[header] = values[idx] ?? "";
    });
    return result;
  });
}

async function parseFile(file) {
  const fileBuffer = await file.arrayBuffer();

  try {
    const rows = parseWorkbook(fileBuffer);
    if (rows.length > 0) return rows;
  } catch {
    // fallback to manual csv parse
  }

  const text = await file.text();
  return parseCsv(text);
}

function donutImage(labels, data, centerTitle, centerValue, colors = null) {
  const safeLabels = Array.isArray(labels) && labels.length ? labels : ["No data"];
  const safeData = Array.isArray(data) && data.some((v) => Number(v) > 0) ? data : [1];
  const safeColors =
    Array.isArray(colors) && colors.length
      ? colors
      : safeLabels[0] === "No data"
        ? ["#D1D5DB"]
        : safeLabels.map((label) => mediaColor(label));

  const canvas = document.createElement("canvas");
  canvas.width = 700;
  canvas.height = 460;
  canvas.style.position = "fixed";
  canvas.style.left = "-9999px";
  canvas.style.top = "-9999px";
  document.body.appendChild(canvas);

  const chart = new Chart(canvas.getContext("2d"), {
    type: "doughnut",
    data: {
      labels: safeLabels,
      datasets: [
        {
          data: safeData,
          backgroundColor: safeColors,
          borderWidth: 2,
          borderColor: "#f7f9ff",
          hoverOffset: 8,
        },
      ],
    },
    options: {
      responsive: false,
      maintainAspectRatio: false,
      animation: false,
      cutout: "72%",
      plugins: {
        legend: { display: false },
        centerTextPlugin: { title: centerTitle, value: centerValue },
      },
    },
  });

  const url = canvas.toDataURL("image/jpeg", 0.92);
  chart.destroy();
  canvas.remove();
  return url;
}

function drawPageHeader(doc, title, subtitleLines = []) {
  const pageW = doc.internal.pageSize.getWidth();
  doc.setFillColor(36, 39, 70);
  doc.rect(0, 0, pageW, 18, "F");

  doc.setTextColor(255, 255, 255);
  doc.setFontSize(14);
  doc.text(title, 10, 11.5);

  doc.setTextColor(55, 65, 81);
  doc.setFontSize(9);
  let y = 23;
  for (const line of subtitleLines) {
    doc.text(line, 10, y);
    y += 5;
  }
}

function drawThreeChartsRow(doc, chartItems, startY) {
  const pageW = doc.internal.pageSize.getWidth();
  const margin = 10;
  const gap = 5;
  const chartW = (pageW - margin * 2 - gap * 2) / 3;
  const chartH = 60;

  doc.setTextColor(17, 24, 39);
  doc.setFontSize(10);
  chartItems.forEach((item, idx) => {
    const x = margin + idx * (chartW + gap);
    doc.text(item.title, x, startY - 2);
    doc.addImage(item.image, "JPEG", x, startY, chartW, chartH);
  });
}

async function exportPdf() {
  if (!loadedRows.length || !columnMap) {
    setStatus("Load data first before exporting PDF.", "error");
    return;
  }

  const originalText = exportPdfBtn.textContent;
  exportPdfBtn.textContent = "Exporting...";
  exportPdfBtn.disabled = true;

  try {
    const { jsPDF } = window.jspdf;
    const doc = new jsPDF({ orientation: "landscape", unit: "mm", format: "a4" });

    const clientName = getClientName();
    const generatedOn = formatMDY(new Date());
    const dateRangeLabel = `${dateFromFilter.value || "-"} to ${dateToFilter.value || "-"}`;
    const selectedTier = tierFilter.value || "All tiers";

    const rowsAllTiers = applyFilters(loadedRows, columnMap, { ignoreTier: true });
    const rowsForTable = filteredRows.length ? filteredRows : rowsAllTiers;
    if (!rowsAllTiers.length) {
      setStatus("No rows available for export with current date filters.", "error");
      return;
    }

    const allReport = buildAggregation(rowsAllTiers, columnMap);
    const allMediaLabels = allReport.byType.map((x) => x.mediaType);

    drawPageHeader(doc, `${clientName} - QBR Dashboard`, [
      `File: ${currentFileLabel || "Uploaded file"} | Generated: ${generatedOn}`,
      `Date filter: ${dateRangeLabel} | Placement list tier filter: ${selectedTier}`,
      `All-tier rows: ${rowsAllTiers.length} | Placement list rows: ${rowsForTable.length}`,
    ]);

    const mentionsAll = donutImage(
      allMediaLabels,
      allReport.byType.map((x) => x.mentions),
      "Mentions",
      formatInt(allReport.totals.mentions),
    );
    const audienceAll = donutImage(
      allMediaLabels,
      allReport.byType.map((x) => x.audience),
      "Audience",
      formatInt(allReport.totals.audience),
    );
    const publicityAll = donutImage(
      allMediaLabels,
      allReport.byType.map((x) => x.publicity),
      "Publicity",
      formatCurrency(allReport.totals.publicity),
    );

    drawThreeChartsRow(
      doc,
      [
        { title: "All Tiers: Mentions by Media Type", image: mentionsAll },
        { title: "All Tiers: Audience by Media Type", image: audienceAll },
        { title: "All Tiers: Publicity by Media Type", image: publicityAll },
      ],
      36,
    );

    const tierMix = donutImage(
      allReport.byTier.map((x) => x.tier),
      allReport.byTier.map((x) => x.mentions),
      "All Tiers",
      formatInt(allReport.byTier.reduce((sum, x) => sum + x.mentions, 0)),
      ["#6C9A8B", "#4C6EF5", "#F59F00", "#C2255C"],
    );
    const pageW = doc.internal.pageSize.getWidth();
    doc.setFontSize(10);
    doc.text("All Tiers: Mentions Mix by Tier", 10, 102);
    doc.addImage(tierMix, "PNG", (pageW - 90) / 2, 104, 90, 60);

    for (const tier of TIER_ORDER) {
      const tierRows = rowsAllTiers.filter((row) => normalizeTier(row[columnMap.tier]) === tier);
      if (!tierRows.length) continue;

      const tierReport = buildAggregation(tierRows, columnMap);
      const tierLabels = tierReport.byType.map((x) => x.mediaType);

      doc.addPage();
      drawPageHeader(doc, `${clientName} - Tier ${tier} Charts`, [
        `Rows in tier ${tier}: ${tierRows.length}`,
        `Audience: ${formatInt(tierReport.totals.audience)} | Publicity: ${formatCurrency(tierReport.totals.publicity)}`,
      ]);

      drawThreeChartsRow(
        doc,
        [
          {
            title: `Tier ${tier}: Mentions by Media Type`,
            image: donutImage(tierLabels, tierReport.byType.map((x) => x.mentions), "Mentions", formatInt(tierReport.totals.mentions)),
          },
          {
            title: `Tier ${tier}: Audience by Media Type`,
            image: donutImage(tierLabels, tierReport.byType.map((x) => x.audience), "Audience", formatInt(tierReport.totals.audience)),
          },
          {
            title: `Tier ${tier}: Publicity by Media Type`,
            image: donutImage(tierLabels, tierReport.byType.map((x) => x.publicity), "Publicity", formatCurrency(tierReport.totals.publicity)),
          },
        ],
        40,
      );
    }

    doc.addPage();
    drawPageHeader(doc, `${clientName} - Placement List`, [
      `Rows: ${rowsForTable.length}`,
      `Same view as dashboard placement list (current filters applied).`,
    ]);

    const body = rowsForTable.map((row) => {
      const link = columnMap.link ? safeUrl(row[columnMap.link]) : "";
      return [
        rowDateLabel(row, columnMap),
        columnMap.publication ? String(row[columnMap.publication] || "") : "",
        columnMap.title ? String(row[columnMap.title] || "") : "",
        link,
        String(row[columnMap.mediaType] || "Unknown"),
        normalizeTier(row[columnMap.tier]),
        formatInt(rowAudience(row, columnMap)),
        formatCurrency(rowPublicity(row, columnMap)),
      ];
    });

    doc.autoTable({
      startY: 35,
      head: [["Date", "Station/Publication", "Title", "Link", "Media Type", "Tier", "Audience", "Publicity"]],
      body,
      styles: {
        fontSize: 8,
        cellPadding: 2,
        overflow: "linebreak",
        valign: "top",
      },
      headStyles: {
        fillColor: [36, 39, 70],
      },
      columnStyles: {
        0: { cellWidth: 20 },
        1: { cellWidth: 38 },
        2: { cellWidth: 58 },
        3: { cellWidth: 45, textColor: [30, 64, 175] },
        4: { cellWidth: 24 },
        5: { cellWidth: 14 },
        6: { cellWidth: 24 },
        7: { cellWidth: 24 },
      },
      didDrawCell(data) {
        if (data.section === "body" && data.column.index === 3) {
          const rowIndex = data?.row?.index;
          const rowValues = Number.isInteger(rowIndex) ? body[rowIndex] : null;
          const url = Array.isArray(rowValues) ? rowValues[3] : "";
          if (url) {
            doc.link(data.cell.x, data.cell.y, data.cell.width, data.cell.height, { url });
          }
        }
      },
      didDrawPage(data) {
        const pageCount = doc.internal.getNumberOfPages();
        doc.setFontSize(8);
        doc.text(`Page ${data.pageNumber} of ${pageCount}`, doc.internal.pageSize.getWidth() - 30, doc.internal.pageSize.getHeight() - 5);
      },
    });

    doc.save(`${clientName.toLowerCase().replace(/\s+/g, "-")}-qbr-dashboard.pdf`);
    setStatus("PDF exported successfully with all tiers and tier-by-tier charts.", "ok");
  } catch (error) {
    const detail = error && typeof error === "object" ? `${error.name || "Error"}: ${error.message || "Unknown export failure"}` : String(error || "Unknown export failure");
    setStatus(`Could not export PDF. ${detail}`, "error");
  } finally {
    exportPdfBtn.textContent = originalText;
    exportPdfBtn.disabled = false;
  }
}

fileInput.addEventListener("change", async (event) => {
  const file = event.target.files?.[0];
  if (!file) return;

  setStatus("Processing file...");

  try {
    const rows = await parseFile(file);
    const map = detectColumns(rows);

    loadedRows = rows;
    columnMap = map;
    currentFileLabel = file.name;

    // Reset filters on every new upload so totals come from the full file by default.
    tierFilter.value = "";
    dateFromFilter.value = "";
    dateToFilter.value = "";

    const hasDate = Boolean(columnMap.date);
    dateFromFilter.disabled = !hasDate;
    dateToFilter.disabled = !hasDate;

    refreshDashboard();
    enableShare();
  } catch (error) {
    setStatus(error.message || "Could not process the file.", "error");
  }
});

tierFilter.addEventListener("change", refreshDashboard);
dateFromFilter.addEventListener("change", refreshDashboard);
dateToFilter.addEventListener("change", refreshDashboard);

clearDateFiltersBtn.addEventListener("click", () => {
  tierFilter.value = "";
  dateFromFilter.value = "";
  dateToFilter.value = "";
  refreshDashboard();
});

exportPdfBtn.addEventListener("click", exportPdf);

(function initFromSharedLink() {
  const params = new URLSearchParams(window.location.search);
  const state = params.get("state");
  if (!state) return;

  try {
    const payload = decodeState(state);
    if (!payload.rows || !payload.map) {
      throw new Error("Invalid shared payload");
    }

    loadedRows = payload.rows;
    columnMap = payload.map;
    currentFileLabel = payload.fileLabel || "Shared report";
    if (clientNameInput) clientNameInput.value = payload.clientName || clientNameInput.value || "";

    dateFromFilter.value = payload.dateFrom || "";
    dateToFilter.value = payload.dateTo || "";
    tierFilter.value = payload.tier || "";

    const hasDate = Boolean(columnMap.date);
    dateFromFilter.disabled = !hasDate;
    dateToFilter.disabled = !hasDate;

    refreshDashboard();
    enableShare();
    setStatus("Loaded dashboard from share link.", "ok");
  } catch {
    setStatus("The share link data is invalid or incomplete.", "error");
  }
})();
