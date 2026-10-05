/* ============================================================
   VGRAT FMS — MONITORING
   ============================================================

   1. Watchlist + alerts
        Funds the viewer stars (saved in this browser), with
        1D / 1W / 1M moves, distance from the 52-week high and
        alert badges for the viewer's own limits.

   2. Unusual moves scan (all funds)
        - Unusual daily / weekly moves: bigger than N x the
          fund's usual move over the last 12 months
        - New 52-week highs / lows
        - Furthest below their 52-week high

   Data: data/bid_history.json (actual BID observations only)
         data/funds.json        (names, codes)

   Nothing is estimated: every figure compares actual BID
   observations. "Usual move" = standard deviation of the
   fund's observation-to-observation returns over the past
   12 months (the latest move itself excluded).
   ============================================================ */

import {
    getWatchlist,
    isWatched,
    addToWatchlist,
    removeFromWatchlist,
    onWatchlistChange
} from "./watchlist.js";


const STORAGE_KEYS = {
    rules: "vgrat-fms-alert-rules",
    sensitivity: "vgrat-fms-unusual-sensitivity"
};

const DEFAULT_RULES = {
    dayMove: { enabled: true, value: 2 },      // |1D| >= value %
    weekDrop: { enabled: true, value: 3 },     // 1W <= -value %
    monthDrop: { enabled: true, value: 5 },    // 1M <= -value %
    fromHigh: { enabled: true, value: 10 },    // below 52W high by >= value %
    newHigh: { enabled: true },
    newLow: { enabled: true }
};

const MIN_RETURNS_FOR_USUAL_MOVE = 60;

const LIST_LIMIT = 10;


const monitor = {
    initialized: false,
    funds: [],
    fundById: new Map(),
    metrics: new Map(),        // id -> metrics
    asOf: null,
    rules: structuredClone(DEFAULT_RULES),
    sensitivity: 2,
    openProfile: () => {}
};


/* ============================================================
   HELPERS
   ============================================================ */

function qs(selector, parent = document) {
    return parent.querySelector(selector);
}

function escapeHtml(value) {
    return String(value ?? "")
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;")
        .replace(/'/g, "&#039;");
}

function readStorage(key) {
    try {
        return localStorage.getItem(key);
    } catch {
        return null;
    }
}

function writeStorage(key, value) {
    try {
        localStorage.setItem(key, value);
    } catch {
        // ignore
    }
}

function fundId(fund) {
    return String(fund?.fundIdentifier ?? fund?.fundCode ?? fund?.fundName ?? "");
}

function isoMinusDays(iso, days) {
    const date = new Date(`${iso}T00:00:00Z`);

    date.setUTCDate(date.getUTCDate() - days);

    return date.toISOString().slice(0, 10);
}

function isoMinusMonths(iso, months) {
    const [year, month, day] = iso.split("-").map(Number);
    const target = new Date(Date.UTC(year, month - 1 - months, 1));
    const lastDay = new Date(Date.UTC(target.getUTCFullYear(), target.getUTCMonth() + 1, 0)).getUTCDate();

    target.setUTCDate(Math.min(day, lastDay));

    return target.toISOString().slice(0, 10);
}

/* Index of the latest observation on or before `date`, else -1. */
function indexOnOrBefore(observations, date) {
    let low = 0;
    let high = observations.length - 1;
    let result = -1;

    while (low <= high) {
        const middle = (low + high) >> 1;

        if (observations[middle].date <= date) {
            result = middle;
            low = middle + 1;
        } else {
            high = middle - 1;
        }
    }

    return result;
}

function pct(from, to) {
    return from > 0 ? ((to - from) / from) * 100 : null;
}

function formatPercent(value, digits = 2) {
    if (value === null || value === undefined || !Number.isFinite(value)) return "—";

    const sign = value > 0 ? "+" : "";

    return `${sign}${value.toFixed(digits)}%`;
}

function returnClass(value) {
    if (value === null || value === undefined || !Number.isFinite(value) || value === 0) return "";

    return value > 0 ? "return-positive" : "return-negative";
}

function formatDate(iso) {
    if (!iso) return "—";

    const [year, month, day] = iso.split("-").map(Number);

    return new Date(Date.UTC(year, month - 1, day)).toLocaleDateString("en-SG", {
        timeZone: "UTC",
        day: "2-digit",
        month: "short",
        year: "numeric"
    });
}

function pill(value, digits = 2) {
    return `
        <span class="return-cell ${returnClass(value)}">
            <span class="return-pill">${escapeHtml(formatPercent(value, digits))}</span>
        </span>
    `;
}


/* ============================================================
   METRICS (per fund, from actual BID observations)
   ============================================================ */

function computeMetrics(observations) {
    if (!Array.isArray(observations) || observations.length < 2) return null;

    const n = observations.length;
    const latest = observations[n - 1];
    const previous = observations[n - 2];

    const weekIndex = indexOnOrBefore(observations, isoMinusDays(latest.date, 7));
    const monthIndex = indexOnOrBefore(observations, isoMinusMonths(latest.date, 1));

    const d1 = pct(previous.bidPrice, latest.bidPrice);
    const w1 = weekIndex >= 0 && weekIndex < n - 1 ? pct(observations[weekIndex].bidPrice, latest.bidPrice) : null;
    const m1 = monthIndex >= 0 && monthIndex < n - 1 ? pct(observations[monthIndex].bidPrice, latest.bidPrice) : null;

    // 52-week yearObs (by calendar date, latest included)
    const yearStart = isoMinusDays(latest.date, 365);
    const firstInYear = observations.findIndex(item => item.date > yearStart);
    const yearObs = observations.slice(firstInYear < 0 ? n - 1 : firstInYear);

    const hasFullYear = observations[0].date <= yearStart;

    const high52 = Math.max(...yearObs.map(item => item.bidPrice));
    const low52 = Math.min(...yearObs.map(item => item.bidPrice));

    const priorWindow = yearObs.slice(0, -1);

    const priorHigh = priorWindow.length ? Math.max(...priorWindow.map(item => item.bidPrice)) : null;
    const priorLow = priorWindow.length ? Math.min(...priorWindow.map(item => item.bidPrice)) : null;

    const newHigh = hasFullYear && priorHigh !== null && latest.bidPrice > priorHigh;
    const newLow = hasFullYear && priorLow !== null && latest.bidPrice < priorLow;

    const fromHigh = pct(high52, latest.bidPrice);

    // Usual move: stdev of observation-to-observation returns in
    // the 52-week yearObs, EXCLUDING the latest move.
    const returns = [];

    for (let index = 1; index < yearObs.length - 1; index += 1) {
        const value = pct(yearObs[index - 1].bidPrice, yearObs[index].bidPrice);

        if (value !== null) returns.push(value);
    }

    let usual = null;

    if (returns.length >= MIN_RETURNS_FOR_USUAL_MOVE) {
        const mean = returns.reduce((sum, value) => sum + value, 0) / returns.length;
        const variance = returns.reduce((sum, value) => sum + (value - mean) ** 2, 0) / (returns.length - 1);

        usual = Math.sqrt(variance);
    }

    const weekSteps = weekIndex >= 0 ? n - 1 - weekIndex : null;
    const usualWeek = usual !== null && weekSteps ? usual * Math.sqrt(weekSteps) : null;

    return {
        latestDate: latest.date,
        latestBid: latest.bidPrice,
        previousDate: previous.date,
        d1,
        w1,
        m1,
        high52,
        low52,
        fromHigh,
        newHigh,
        newLow,
        hasFullYear,
        usual,
        usualWeek,
        zDay: usual ? d1 / usual : null,
        zWeek: usualWeek && w1 !== null ? w1 / usualWeek : null
    };
}


/* ============================================================
   ALERTS (watchlist)
   ============================================================ */

function alertsFor(metrics) {
    if (!metrics) return [];

    const rules = monitor.rules;
    const alerts = [];

    if (rules.dayMove.enabled && metrics.d1 !== null && Math.abs(metrics.d1) >= rules.dayMove.value) {
        alerts.push({
            tone: metrics.d1 < 0 ? "negative" : "positive",
            text: `1D ${formatPercent(metrics.d1)}`
        });
    }

    if (rules.weekDrop.enabled && metrics.w1 !== null && metrics.w1 <= -rules.weekDrop.value) {
        alerts.push({ tone: "negative", text: `1W ${formatPercent(metrics.w1)}` });
    }

    if (rules.monthDrop.enabled && metrics.m1 !== null && metrics.m1 <= -rules.monthDrop.value) {
        alerts.push({ tone: "negative", text: `1M ${formatPercent(metrics.m1)}` });
    }

    if (rules.fromHigh.enabled && metrics.fromHigh !== null && metrics.fromHigh <= -rules.fromHigh.value) {
        alerts.push({ tone: "negative", text: `${formatPercent(metrics.fromHigh, 1)} from 52W high` });
    }

    if (rules.newHigh.enabled && metrics.newHigh) {
        alerts.push({ tone: "positive", text: "New 52W high" });
    }

    if (rules.newLow.enabled && metrics.newLow) {
        alerts.push({ tone: "negative", text: "New 52W low" });
    }

    return alerts;
}


/* ============================================================
   RENDER: WATCHLIST
   ============================================================ */

function fundCell(id) {
    const fund = monitor.fundById.get(id);

    return `
        <div class="fund-name">${escapeHtml(fund?.fundName ?? id)}</div>
        <div class="fund-meta">${escapeHtml(fund?.fundCode ?? "")}</div>
    `;
}

function renderWatchlist() {
    const body = qs("#watchlist-body");
    const summary = qs("#watchlist-summary");

    if (!body) return;

    const ids = getWatchlist().filter(id => monitor.fundById.has(id));

    if (!ids.length) {
        if (summary) summary.textContent = "No funds watched yet";

        body.innerHTML = `
            <tr>
                <td colspan="8" class="empty-state monitor-empty">
                    Add funds above, or click <strong>☆ Watch</strong> in any Fund Explorer profile.
                </td>
            </tr>
        `;
        return;
    }

    let alertCount = 0;
    let fundsWithAlerts = 0;

    const rows = ids.map(id => {
        const metrics = monitor.metrics.get(id);
        const alerts = alertsFor(metrics);

        alertCount += alerts.length;
        if (alerts.length) fundsWithAlerts += 1;

        return { id, metrics, alerts };
    });

    // Funds with alerts first, then by name.
    rows.sort((a, b) =>
        (b.alerts.length > 0) - (a.alerts.length > 0) ||
        String(monitor.fundById.get(a.id)?.fundName).localeCompare(String(monitor.fundById.get(b.id)?.fundName))
    );

    if (summary) {
        summary.textContent = alertCount
            ? `${alertCount} alert${alertCount === 1 ? "" : "s"} on ${fundsWithAlerts} of ${ids.length} fund${ids.length === 1 ? "" : "s"}`
            : `${ids.length} fund${ids.length === 1 ? "" : "s"} watched · no alerts`;
    }

    body.innerHTML = rows
        .map(({ id, metrics, alerts }) => `
            <tr class="monitor-row${alerts.length ? " has-alerts" : ""}" data-monitor-open="${escapeHtml(id)}" tabindex="0">
                <td class="monitor-star-cell">
                    <button type="button" class="monitor-star is-on" data-unwatch="${escapeHtml(id)}" title="Stop watching" aria-label="Stop watching">★</button>
                </td>
                <td class="fund-cell">${fundCell(id)}</td>
                <td class="bid-cell">
                    ${metrics ? `${metrics.latestBid.toFixed(5)}<div class="fund-meta">${escapeHtml(formatDate(metrics.latestDate))}</div>` : "—"}
                </td>
                <td>${pill(metrics?.d1)}</td>
                <td>${pill(metrics?.w1)}</td>
                <td>${pill(metrics?.m1)}</td>
                <td>${pill(metrics?.fromHigh, 1)}</td>
                <td class="monitor-alerts-cell">
                    ${alerts.length
                        ? alerts.map(alert => `<span class="monitor-alert ${alert.tone}">${escapeHtml(alert.text)}</span>`).join("")
                        : `<span class="monitor-ok">No alerts</span>`}
                </td>
            </tr>
        `)
        .join("");
}


/* ============================================================
   RENDER: UNUSUAL MOVES
   ============================================================ */

function listRows(items, valueOf, detailOf, digits = 2) {
    if (!items.length) {
        return `<li class="monitor-list-empty">None today</li>`;
    }

    return items
        .map(({ id, metrics }) => `
            <li class="monitor-list-row" data-monitor-open="${escapeHtml(id)}" tabindex="0">
                <div class="monitor-list-fund">
                    ${fundCell(id)}
                </div>
                <div class="monitor-list-value">
                    ${pill(valueOf(metrics), digits)}
                    <div class="fund-meta">${detailOf(metrics)}</div>
                </div>
                ${isWatched(id) ? `<span class="monitor-list-star" title="On your watchlist">★</span>` : ""}
            </li>
        `)
        .join("");
}

function renderUnusual() {
    const container = qs("#unusual-lists");
    const asOf = qs("#unusual-asof");

    if (!container) return;

    if (asOf) asOf.textContent = monitor.asOf ? `Latest BID ${formatDate(monitor.asOf)}` : "";

    const all = [...monitor.metrics.entries()]
        .map(([id, metrics]) => ({ id, metrics }))
        .filter(item => item.metrics);

    const k = monitor.sensitivity;

    const unusualDay = all
        .filter(item => item.metrics.zDay !== null && Math.abs(item.metrics.zDay) >= k)
        .sort((a, b) => Math.abs(b.metrics.zDay) - Math.abs(a.metrics.zDay))
        .slice(0, LIST_LIMIT);

    const unusualWeek = all
        .filter(item => item.metrics.zWeek !== null && Math.abs(item.metrics.zWeek) >= k)
        .sort((a, b) => Math.abs(b.metrics.zWeek) - Math.abs(a.metrics.zWeek))
        .slice(0, LIST_LIMIT);

    const highs = all
        .filter(item => item.metrics.newHigh)
        .sort((a, b) => (b.metrics.d1 ?? 0) - (a.metrics.d1 ?? 0));

    const lows = all
        .filter(item => item.metrics.newLow)
        .sort((a, b) => (a.metrics.d1 ?? 0) - (b.metrics.d1 ?? 0));

    const furthest = all
        .filter(item => item.metrics.fromHigh !== null && item.metrics.fromHigh < 0)
        .sort((a, b) => a.metrics.fromHigh - b.metrics.fromHigh)
        .slice(0, LIST_LIMIT);

    const times = value => `${(Math.abs(value)).toFixed(1)}× usual`;

    const card = (title, hint, items) => `
        <section class="monitor-list-card">
            <header>
                <h3>${escapeHtml(title)}</h3>
                <span>${escapeHtml(hint)}</span>
            </header>
            <ol class="monitor-list">${items}</ol>
        </section>
    `;

    container.innerHTML =
        card(
            "Unusual daily moves",
            `Latest move ≥ ${k}× the fund's usual daily move`,
            listRows(unusualDay, m => m.d1, m => `${times(m.zDay)} (usual ±${m.usual.toFixed(2)}%)`)
        ) +
        card(
            "Unusual weekly moves",
            `1-week move ≥ ${k}× the fund's usual weekly move`,
            listRows(unusualWeek, m => m.w1, m => `${times(m.zWeek)} (usual ±${m.usualWeek.toFixed(2)}%)`)
        ) +
        card(
            "New 52-week highs",
            "Latest BID above every BID of the past year",
            listRows(highs, m => m.d1, m => `BID ${m.latestBid.toFixed(5)}`)
        ) +
        card(
            "New 52-week lows",
            "Latest BID below every BID of the past year",
            listRows(lows, m => m.d1, m => `BID ${m.latestBid.toFixed(5)}`)
        ) +
        card(
            "Furthest below 52-week high",
            "Current BID vs the highest BID of the past year",
            listRows(furthest, m => m.fromHigh, m => `52W high ${m.high52.toFixed(5)}`, 1)
        );
}


/* ============================================================
   ALERT RULES FORM
   ============================================================ */

function loadRules() {
    try {
        const saved = JSON.parse(readStorage(STORAGE_KEYS.rules) ?? "null");

        if (saved && typeof saved === "object") {
            for (const key of Object.keys(DEFAULT_RULES)) {
                if (saved[key]) monitor.rules[key] = { ...DEFAULT_RULES[key], ...saved[key] };
            }
        }
    } catch {
        // keep defaults
    }

    const sensitivity = Number(readStorage(STORAGE_KEYS.sensitivity));

    if ([2, 2.5, 3].includes(sensitivity)) monitor.sensitivity = sensitivity;
}

function syncRuleInputs() {
    for (const [key, rule] of Object.entries(monitor.rules)) {
        const toggle = qs(`[data-rule-enabled="${key}"]`);
        const input = qs(`[data-rule-value="${key}"]`);

        if (toggle) toggle.checked = Boolean(rule.enabled);
        if (input && rule.value !== undefined) {
            input.value = String(rule.value);
            input.disabled = !rule.enabled;
        }
    }

    const select = qs("#unusual-sensitivity");

    if (select) select.value = String(monitor.sensitivity);
}

function saveRules() {
    writeStorage(STORAGE_KEYS.rules, JSON.stringify(monitor.rules));
}


/* ============================================================
   PICKER (add to watchlist)
   ============================================================ */

function renderSuggestions() {
    const input = qs("#watchlist-search");
    const list = qs("#watchlist-suggestions");

    if (!input || !list) return;

    const query = input.value.trim().toLowerCase();

    const matches = monitor.funds
        .filter(fund => {
            const id = fundId(fund);

            if (isWatched(id) || !monitor.metrics.has(id)) return false;
            if (!query) return true;

            return (
                String(fund.fundName ?? "").toLowerCase().includes(query) ||
                String(fund.fundCode ?? "").toLowerCase().includes(query)
            );
        })
        .slice(0, 50);

    list.innerHTML = matches.length
        ? matches
            .map((fund, index) => `
                <li class="compare-suggestion${index === 0 ? " is-highlighted" : ""}" role="option" data-watch-add="${escapeHtml(fundId(fund))}">
                    <span class="compare-suggestion-name">${escapeHtml(fund.fundName ?? "")}</span>
                    <span class="compare-suggestion-code">${escapeHtml(fund.fundCode ?? "")}</span>
                </li>
            `)
            .join("")
        : `<li class="compare-suggestion-empty">No matching funds</li>`;

    list.hidden = false;
}

function closeSuggestions() {
    const list = qs("#watchlist-suggestions");

    if (list) list.hidden = true;
}


/* ============================================================
   EVENTS
   ============================================================ */

function openFromEvent(event) {
    if (event.target.closest("[data-unwatch]")) return;

    const row = event.target.closest("[data-monitor-open]");

    if (row) monitor.openProfile(row.dataset.monitorOpen);
}

function bindEvents() {
    const input = qs("#watchlist-search");
    const list = qs("#watchlist-suggestions");

    input?.addEventListener("focus", renderSuggestions);
    input?.addEventListener("input", renderSuggestions);

    input?.addEventListener("keydown", event => {
        if (event.key === "Enter") {
            event.preventDefault();

            const first = qs("#watchlist-suggestions [data-watch-add]");

            if (first) {
                addToWatchlist(first.dataset.watchAdd);
                input.value = "";
                renderSuggestions();
            }
        } else if (event.key === "Escape") {
            closeSuggestions();
            input.blur();
        }
    });

    list?.addEventListener("mousedown", event => {
        const item = event.target.closest("[data-watch-add]");

        if (!item) return;

        event.preventDefault();
        addToWatchlist(item.dataset.watchAdd);

        if (input) input.value = "";

        renderSuggestions();
    });

    document.addEventListener("click", event => {
        if (!event.target.closest(".watchlist-picker")) closeSuggestions();
    });

    const view = qs("#monitoring-view");

    view?.addEventListener("click", event => {
        const unwatch = event.target.closest("[data-unwatch]");

        if (unwatch) {
            event.stopPropagation();
            removeFromWatchlist(unwatch.dataset.unwatch);
            return;
        }

        openFromEvent(event);
    });

    view?.addEventListener("keydown", event => {
        if (event.key === "Enter") openFromEvent(event);
    });

    // Alert rules
    document.querySelectorAll("[data-rule-enabled]").forEach(toggle => {
        toggle.addEventListener("change", () => {
            const key = toggle.dataset.ruleEnabled;

            monitor.rules[key].enabled = toggle.checked;
            saveRules();
            syncRuleInputs();
            renderWatchlist();
        });
    });

    document.querySelectorAll("[data-rule-value]").forEach(field => {
        // Update as you type; ignore repeat events with no change
        // (so leaving the box doesn't redraw the table mid-click).
        const update = () => {
            const key = field.dataset.ruleValue;
            const value = Number(field.value);

            if (!Number.isFinite(value) || value <= 0) return;
            if (value === monitor.rules[key].value) return;

            monitor.rules[key].value = value;
            saveRules();
            renderWatchlist();
        };

        field.addEventListener("input", update);
        field.addEventListener("change", update);
    });

    qs("#alert-rules-reset")?.addEventListener("click", () => {
        monitor.rules = structuredClone(DEFAULT_RULES);
        saveRules();
        syncRuleInputs();
        renderWatchlist();
    });

    qs("#unusual-sensitivity")?.addEventListener("change", event => {
        monitor.sensitivity = Number(event.target.value) || 2;
        writeStorage(STORAGE_KEYS.sensitivity, String(monitor.sensitivity));
        renderUnusual();
    });

    onWatchlistChange(() => {
        renderWatchlist();
        renderUnusual();
    });
}


/* ============================================================
   PUBLIC API
   ============================================================ */

/**
 * @param {object}   options
 * @param {Array}    options.funds         funds.json records
 * @param {Map}      options.historyIndex  id -> sorted BID observations
 * @param {Function} options.openProfile   opens the Fund Explorer profile
 */
function initializeMonitoring({ funds, historyIndex, openProfile }) {
    monitor.funds = [...(funds ?? [])].sort((a, b) =>
        String(a.fundName ?? "").localeCompare(String(b.fundName ?? ""))
    );

    monitor.fundById = new Map(monitor.funds.map(fund => [fundId(fund), fund]));
    monitor.openProfile = openProfile ?? monitor.openProfile;

    monitor.metrics = new Map();

    for (const fund of monitor.funds) {
        const id = fundId(fund);
        const metrics = computeMetrics(historyIndex?.get(id));

        if (metrics) monitor.metrics.set(id, metrics);
    }

    monitor.asOf = [...monitor.metrics.values()]
        .map(item => item.latestDate)
        .sort()
        .pop() ?? null;

    if (!monitor.initialized) {
        loadRules();
        bindEvents();
        monitor.initialized = true;
    }

    syncRuleInputs();

    const loading = qs("#monitoring-loading");

    if (loading) loading.hidden = true;

    renderWatchlist();
    renderUnusual();
}


export {
    initializeMonitoring,
    computeMetrics
};
