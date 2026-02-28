const API_URL = "https://script.google.com/macros/s/AKfycbxpudYsX6cBGA26wYkh05RlvlqVV96AI0-ce3or2SBnHSu_OwSeSLKme6pV6vbSzQ/exec";
const TV_LOGIN_KEY = "tvLoginId";
const THEME_KEY = "uiTheme";
const DEFAULT_THEME = "light";
const AUTO_REFRESH_MS = 5000;
const MAX_ROWS_PER_PANEL = 13;

let rawData = [];
let activeTvId = "";
let autoRefreshTimer = null;
let isRefreshInFlight = false;
let lastRowsFingerprint = "";

const appEl = document.getElementById("app");
const masterHeaderEl = document.getElementById("masterHeader");
const sectionHeaderEl = document.getElementById("sectionHeader");
const sectionListEl = document.getElementById("sectionList");
const logoutBtnEl = document.getElementById("logoutBtn");
const themeToggleBtnEl = document.getElementById("themeToggleBtn");
const fullscreenToggleBtnEl = document.getElementById("fullscreenToggleBtn");
const tvIdBadgeEl = document.getElementById("tvIdBadge");
const tvLoginOverlayEl = document.getElementById("tvLoginOverlay");
const tvLoginFormEl = document.getElementById("tvLoginForm");
const tvIdInputEl = document.getElementById("tvIdInput");
const tvLoginErrorEl = document.getElementById("tvLoginError");

function getCurrentTheme() {
  return document.documentElement.getAttribute("data-theme") === "light" ? "light" : "dark";
}

function updateThemeButtonUi() {
  if (!themeToggleBtnEl) return;
  const iconEl = themeToggleBtnEl.querySelector("i");
  const currentTheme = getCurrentTheme();
  const nextTheme = currentTheme === "dark" ? "light" : "dark";

  themeToggleBtnEl.setAttribute("aria-label", `Switch to ${nextTheme} theme`);
  themeToggleBtnEl.setAttribute("title", `Switch to ${nextTheme} theme`);
  if (iconEl) {
    iconEl.className = currentTheme === "dark" ? "bi bi-moon-stars-fill" : "bi bi-sun-fill";
  }
}

function applyTheme(theme, options = {}) {
  const { persist = false } = options;
  const normalizedTheme = theme === "light" ? "light" : "dark";
  document.documentElement.setAttribute("data-theme", normalizedTheme);
  updateThemeButtonUi();

  if (persist) {
    localStorage.setItem(THEME_KEY, normalizedTheme);
  }
}

function initTheme() {
  const savedTheme = localStorage.getItem(THEME_KEY);
  const initialTheme = savedTheme === "light" || savedTheme === "dark" ? savedTheme : DEFAULT_THEME;
  applyTheme(initialTheme);

  if (themeToggleBtnEl) {
    themeToggleBtnEl.addEventListener("click", () => {
      const nextTheme = getCurrentTheme() === "dark" ? "light" : "dark";
      applyTheme(nextTheme, { persist: true });
    });
  }
}

function updateFullscreenButtonUi() {
  if (!fullscreenToggleBtnEl) return;
  const iconEl = fullscreenToggleBtnEl.querySelector("i");
  const isFullscreen = Boolean(document.fullscreenElement);
  document.body.classList.toggle("fullscreen-mode", isFullscreen);
  fullscreenToggleBtnEl.setAttribute("aria-label", isFullscreen ? "Exit fullscreen" : "Enter fullscreen");
  fullscreenToggleBtnEl.setAttribute("title", isFullscreen ? "Exit fullscreen" : "Enter fullscreen");
  if (iconEl) {
    iconEl.className = isFullscreen ? "bi bi-fullscreen-exit" : "bi bi-arrows-fullscreen";
  }
}

async function toggleFullscreen() {
  if (!document.fullscreenEnabled) return;
  try {
    if (document.fullscreenElement) {
      await document.exitFullscreen();
    } else {
      await document.documentElement.requestFullscreen();
    }
  } catch (error) {
    // Keep view functional even if browser blocks fullscreen.
  } finally {
    updateFullscreenButtonUi();
  }
}

function normalizeTvId(value) {
  const text = String(value ?? "").trim();
  if (!text) return "";
  if (/^\d+$/.test(text)) return String(Number(text));
  return text;
}

function normalizeKey(value) {
  return String(value ?? "").toLowerCase().replace(/[^a-z0-9]/g, "");
}

function isTrueLike(value) {
  const v = String(value ?? "").trim().toLowerCase();
  return v === "true" || v === "1" || v === "yes" || v === "y";
}

function setLoginError(message) {
  if (!tvLoginErrorEl) return;
  const errorMessage = String(message ?? "").trim();
  if (!errorMessage) {
    tvLoginErrorEl.hidden = true;
    tvLoginErrorEl.textContent = "";
    return;
  }
  tvLoginErrorEl.hidden = false;
  tvLoginErrorEl.textContent = errorMessage;
}

function getRowValue(row, aliases) {
  if (!row || typeof row !== "object") return "";
  const normalizedAliases = aliases.map(normalizeKey);
  const directKey = Object.keys(row).find((key) => normalizedAliases.includes(normalizeKey(key)));
  return directKey ? row[directKey] : "";
}

function getRowTvId(row) {
  return normalizeTvId(getRowValue(row, ["tv", "tvid", "tv id", "tv_id"]));
}

function getRowMaster(row) {
  return String(getRowValue(row, ["master"])).trim();
}

function getRowSection(row) {
  return String(getRowValue(row, ["section", "group"])).trim();
}

function getRowProduct(row) {
  return String(getRowValue(row, ["product"])).trim();
}

function getRowPacked(row) {
  return getRowValue(row, ["packed"]);
}

function getRowTray(row) {
  return getRowValue(row, ["tray"]);
}

function getRowTotal(row) {
  return getRowValue(row, ["total"]);
}

function getRowStatus(row) {
  return String(getRowValue(row, ["status"])).trim();
}

function getRowPackedStatus(row) {
  return String(getRowValue(row, ["packed status", "packed_status", "packedstatus"])).trim();
}

function getRowTrayStatus(row) {
  return String(getRowValue(row, ["tray status", "tray_status", "traystatus"])).trim();
}

function isCompletedStatus(value) {
  const v = String(value ?? "").trim().toLowerCase();
  return v === "done" || v === "complete" || v === "completed" || v === "packed" || v === "tray";
}

function isPackedComplete(row) {
  const packedStatus = getRowPackedStatus(row);
  if (packedStatus) return isCompletedStatus(packedStatus);
  return isCompletedStatus(getRowStatus(row));
}

function isTrayComplete(row) {
  const trayStatus = getRowTrayStatus(row);
  if (trayStatus) return isCompletedStatus(trayStatus);
  return isCompletedStatus(getRowStatus(row));
}

function showLoginOverlay() {
  tvLoginOverlayEl.hidden = false;
  setLoginError("");
  window.setTimeout(() => tvIdInputEl.focus(), 0);
}

function hideLoginOverlay() {
  tvLoginOverlayEl.hidden = true;
}

function resetViewForLoggedOut() {
  rawData = [];
  lastRowsFingerprint = "";
  stopAutoRefresh();
  updateHeader([]);
  appEl.innerHTML = '<div class="loading-card single-card">Enter TV ID to continue.</div>';
  updateTvIdBadge("");
}

function updateTvIdBadge(tvId) {
  if (!tvIdBadgeEl) return;
  const normalized = normalizeTvId(tvId);
  tvIdBadgeEl.textContent = normalized ? `TV ID: ${normalized}` : "TV ID: -";
}

function stopAutoRefresh() {
  if (autoRefreshTimer) {
    window.clearInterval(autoRefreshTimer);
    autoRefreshTimer = null;
  }
}

function startAutoRefresh() {
  stopAutoRefresh();
  if (!activeTvId) return;
  autoRefreshTimer = window.setInterval(() => {
    refreshRowsInBackground();
  }, AUTO_REFRESH_MS);
}

function buildApiUrl(tvId, currentOnly) {
  const normalizedTvId = normalizeTvId(tvId);
  const queryParts = [];
  if (normalizedTvId) queryParts.push(`tvId=${encodeURIComponent(normalizedTvId)}`);
  if (currentOnly) queryParts.push("currentOnly=true");
  const separator = API_URL.includes("?") ? "&" : "?";
  return `${API_URL}${separator}${queryParts.join("&")}`;
}

async function fetchRowsForTvId(tvId, currentOnly) {
  const res = await fetch(buildApiUrl(tvId, currentOnly));
  const json = await res.json();
  return Array.isArray(json?.data) ? json.data : [];
}

function createRowsFingerprint(rows) {
  if (!Array.isArray(rows)) return "";
  return rows
    .map((row) => {
      if (!row || typeof row !== "object") return JSON.stringify(row);
      const sorted = {};
      Object.keys(row)
        .sort()
        .forEach((key) => {
          sorted[key] = row[key];
        });
      return JSON.stringify(sorted);
    })
    .sort()
    .join("|");
}

function applyRows(rows, tvId, options = {}) {
  const { forceRender = false, animate = false } = options;
  activeTvId = normalizeTvId(tvId);
  updateTvIdBadge(activeTvId);
  rawData = Array.isArray(rows) ? rows : [];

  const nextFingerprint = createRowsFingerprint(rawData);
  const changed = forceRender || nextFingerprint !== lastRowsFingerprint;
  if (changed) {
    lastRowsFingerprint = nextFingerprint;
    renderDataView({ animate });
  }

  return { matchCount: rawData.length, changed };
}

function escapeHtml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}

function formatNumberLike(value) {
  if (value === null || value === undefined) return "-";
  const text = String(value).trim();
  if (!text) return "-";
  if (text === "-") return "-";
  const num = Number(text);
  if (Number.isFinite(num)) return num.toFixed(2);
  return text;
}

function updateHeader(rows) {
  const masters = [...new Set(rows.map((row) => getRowMaster(row)).filter(Boolean))].sort();
  const sections = [...new Set(rows.map((row) => getRowSection(row)).filter(Boolean))].sort();

  if (!rows.length) {
    masterHeaderEl.textContent = "Master: -";
    sectionHeaderEl.textContent = "Section: -";
    sectionListEl.innerHTML = "";
    return;
  }

  if (masters.length === 1) {
    masterHeaderEl.textContent = "MASTER: " + masters[0].toUpperCase();
  } else if (masters.length > 1) {
    masterHeaderEl.textContent = "MASTER: " + masters.map((m) => m.toUpperCase()).join(" | ");
  } else {
    masterHeaderEl.textContent = "MASTER";
  }

  if (sections.length === 1) {
    sectionHeaderEl.textContent = sections[0].toUpperCase();
    sectionListEl.innerHTML = "";
    return;
  }

  sectionHeaderEl.textContent = sections.length ? `SECTIONS (${sections.length})` : "Section: -";
  sectionListEl.innerHTML = sections
    .map((section) => `<span class="section-pill">${escapeHtml(section.toUpperCase())}</span>`)
    .join("");
}

function buildSectionTableHtml(rows) {
  const sections = [...new Set(rows.map((row) => getRowSection(row)).filter(Boolean))];
  const sectionTitle = sections[0] || "SECTION";
  const masters = [...new Set(rows.map((row) => getRowMaster(row)).filter(Boolean))];
  const masterName = masters[0] || "MASTER";

  let summaryRow = rows.find((row) => {
    const product = getRowProduct(row);
    const master = getRowMaster(row);
    return product && master && product.toLowerCase() === master.toLowerCase();
  });

  // Fallback when explicit master-total row is not present in data.
  if (!summaryRow) {
    const totals = rows.reduce(
      (acc, row) => {
        const packed = Number(String(getRowPacked(row) ?? "").trim());
        const tray = Number(String(getRowTray(row) ?? "").trim());
        const total = Number(String(getRowTotal(row) ?? "").trim());
        if (Number.isFinite(packed)) acc.packed += packed;
        if (Number.isFinite(tray)) acc.tray += tray;
        if (Number.isFinite(total)) acc.total += total;
        return acc;
      },
      { packed: 0, tray: 0, total: 0 }
    );
    summaryRow = {
      product: masterName,
      packed: totals.packed.toFixed(2),
      tray: totals.tray.toFixed(2),
      total: totals.total.toFixed(2)
    };
  }

  const bodyRows = rows.filter((row) => {
    const product = getRowProduct(row);
    const master = getRowMaster(row);
    return !(product && master && product.toLowerCase() === master.toLowerCase());
  });
  const completedSums = bodyRows.reduce(
    (acc, row) => {
      const packedComplete = isPackedComplete(row);
      const trayComplete = isTrayComplete(row);
      const packed = Number(String(getRowPacked(row) ?? "").trim());
      const tray = Number(String(getRowTray(row) ?? "").trim());
      const total = Number(String(getRowTotal(row) ?? "").trim());
      if (packedComplete && Number.isFinite(packed)) acc.packed += packed;
      if (trayComplete && Number.isFinite(tray)) acc.tray += tray;
      if ((packedComplete || trayComplete) && Number.isFinite(total)) acc.total += total;
      return acc;
    },
    { packed: 0, tray: 0, total: 0 }
  );

  const chunks = [];
  const firstChunkSize = Math.max(1, MAX_ROWS_PER_PANEL - 1); // first panel includes summary row
  if (bodyRows.length <= firstChunkSize) {
    chunks.push(bodyRows);
  } else {
    chunks.push(bodyRows.slice(0, firstChunkSize));
    for (let i = firstChunkSize; i < bodyRows.length; i += MAX_ROWS_PER_PANEL) {
      chunks.push(bodyRows.slice(i, i + MAX_ROWS_PER_PANEL));
    }
  }
  if (!chunks.length) {
    chunks.push([]);
  }

  const renderBodyRows = (chunk) =>
    chunk
      .map((row) => {
        const product = getRowProduct(row) || "-";
        const packedComplete = isPackedComplete(row);
        const trayComplete = isTrayComplete(row);
        const completedRowClass = packedComplete && trayComplete ? "completed-product-row" : "";
        const packedCellClass = packedComplete ? "completed-cell" : "";
        const trayCellClass = trayComplete ? "completed-cell" : "";
        return `
          <tr class="${completedRowClass}">
            <td>${escapeHtml(product.toUpperCase())}</td>
            <td class="num ${packedCellClass}">${escapeHtml(formatNumberLike(getRowPacked(row)))}</td>
            <td class="num ${trayCellClass}">${escapeHtml(formatNumberLike(getRowTray(row)))}</td>
            <td class="num">${escapeHtml(formatNumberLike(getRowTotal(row)))}</td>
          </tr>
        `;
      })
      .join("");

  const tablesHtml = chunks
    .map((chunk, index) => {
      const summaryPackedNum = Number(String(summaryRow.packed ?? "").trim());
      const summaryTrayNum = Number(String(summaryRow.tray ?? "").trim());
      const summaryTotalNum = Number(String(summaryRow.total ?? "").trim());
      const balancePacked = Number.isFinite(summaryPackedNum) ? summaryPackedNum - completedSums.packed : null;
      const balanceTray = Number.isFinite(summaryTrayNum) ? summaryTrayNum - completedSums.tray : null;
      const balanceTotal = Number.isFinite(summaryTotalNum) ? summaryTotalNum - completedSums.total : null;
      const summaryHtml =
        index === 0
          ? `
            <tr class="master-summary-row">
              <td>${escapeHtml(String(summaryRow.product || masterName).toUpperCase())}</td>
              <td class="num">${escapeHtml(formatNumberLike(summaryRow.packed))}</td>
              <td class="num">${escapeHtml(formatNumberLike(summaryRow.tray))}</td>
              <td class="num">${escapeHtml(formatNumberLike(summaryRow.total))}</td>
            </tr>
          `
          : "";
      const completedSummaryHtml =
        index === chunks.length - 1
          ? `
            <tr class="completed-summary-row">
              <td>COMPLETED-BALANCE</td>
              <td class="num com-bal">${escapeHtml(
                `${formatNumberLike(completedSums.packed)}-${formatNumberLike(balancePacked)}`
              )}</td>
              <td class="num com-bal">${escapeHtml(
                `${formatNumberLike(completedSums.tray)}-${formatNumberLike(balanceTray)}`
              )}</td>
              <td class="num com-bal">${escapeHtml(
                `${formatNumberLike(completedSums.total)}-${formatNumberLike(balanceTotal)}`
              )}</td>
            </tr>
          `
          : "";
      return `
        <article class="section-table-card single-card">
          <table class="section-table">
            <thead>
              <tr>
                <th>${escapeHtml(sectionTitle.toUpperCase())}</th>
                <th>PACKED</th>
                <th>TRAY</th>
                <th>TOTAL</th>
              </tr>
            </thead>
            <tbody>
              ${summaryHtml}
              ${renderBodyRows(chunk)}
              ${completedSummaryHtml}
            </tbody>
          </table>
        </article>
      `;
    })
    .join("");

  const wrapperClass = chunks.length > 1 ? "split-view" : "split-view single";
  return `<div class="${wrapperClass}">${tablesHtml}</div>`;
}

function renderDataView(options = {}) {
  const { animate = false } = options;
  updateHeader(rawData);

  if (!rawData.length) {
    appEl.innerHTML = '<div class="empty-card single-card">No active section for this TV.</div>';
    return;
  }

  appEl.innerHTML = buildSectionTableHtml(rawData);
  if (animate) {
    appEl.classList.add("is-refreshing");
    window.requestAnimationFrame(() => {
      appEl.classList.remove("is-refreshing");
    });
  }
}

async function refreshRowsInBackground() {
  if (!activeTvId || isRefreshInFlight) return;
  isRefreshInFlight = true;
  try {
    const rows = await fetchRowsForTvId(activeTvId, true);
    applyRows(rows, activeTvId, { animate: true });
  } catch (error) {
    // Ignore transient network errors and keep the current TV view visible.
  } finally {
    isRefreshInFlight = false;
  }
}

if (fullscreenToggleBtnEl) {
  fullscreenToggleBtnEl.addEventListener("click", () => {
    toggleFullscreen();
  });
  updateFullscreenButtonUi();
}

document.addEventListener("fullscreenchange", () => {
  updateFullscreenButtonUi();
});

async function loadForActiveTv(tvId) {
  const allRows = await fetchRowsForTvId(tvId, false);
  if (!allRows.length) {
    return { hasTvRows: false, currentRows: [] };
  }
  const currentOnlyRows = await fetchRowsForTvId(tvId, true);
  const currentRows = currentOnlyRows.length
    ? currentOnlyRows
    : allRows.filter((row) => isTrueLike(getRowValue(row, ["screen"])));
  return { hasTvRows: true, currentRows };
}

async function init() {
  const savedTvId = normalizeTvId(localStorage.getItem(TV_LOGIN_KEY));
  if (savedTvId) {
    activeTvId = savedTvId;
    updateTvIdBadge(activeTvId);
    hideLoginOverlay();
  } else {
    showLoginOverlay();
    resetViewForLoggedOut();
  }

  tvLoginFormEl.addEventListener("submit", async (event) => {
    event.preventDefault();
    const enteredTvId = normalizeTvId(tvIdInputEl.value);
    if (!enteredTvId) {
      setLoginError("Please enter TV ID.");
      return;
    }

    appEl.innerHTML = '<div class="loading-card single-card">Loading...</div>';
    try {
      const result = await loadForActiveTv(enteredTvId);
      if (!result.hasTvRows) {
        setLoginError(`No data found for TV ID ${enteredTvId}.`);
        resetViewForLoggedOut();
        return;
      }

      applyRows(result.currentRows, enteredTvId, { forceRender: true });
      setLoginError("");
      localStorage.setItem(TV_LOGIN_KEY, enteredTvId);
      hideLoginOverlay();
      startAutoRefresh();
    } catch (error) {
      setLoginError("Unable to load data. Please try again.");
      resetViewForLoggedOut();
    }
  });

  logoutBtnEl.addEventListener("click", () => {
    localStorage.removeItem(TV_LOGIN_KEY);
    activeTvId = "";
    showLoginOverlay();
    resetViewForLoggedOut();
  });

  try {
    if (activeTvId) {
      const result = await loadForActiveTv(activeTvId);
      if (!result.hasTvRows) {
        localStorage.removeItem(TV_LOGIN_KEY);
        activeTvId = "";
        resetViewForLoggedOut();
        setLoginError("Saved TV ID has no matching data. Please login again.");
        showLoginOverlay();
      } else {
        applyRows(result.currentRows, activeTvId, { forceRender: true });
        startAutoRefresh();
      }
    }
  } catch (error) {
    appEl.innerHTML = '<div class="empty-card single-card">Unable to load data. Please try again.</div>';
  }
}

initTheme();
init();
