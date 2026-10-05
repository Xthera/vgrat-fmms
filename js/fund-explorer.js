/* ============================================================
   VGRAT FMS — FUND EXPLORER
   ============================================================

   1. Fund list    searchable, sortable table of every fund
                   with filters (asset class, risk, geography,
                   sector, dividend) and a "holds company"
                   search across top-10 holdings.
   2. Profile      click a fund for a popup with its facts,
                   objective, Prudential's published returns,
                   top holdings, dividends and documents.

   Data sources:
     - data/funds.json         fund details, published returns,
                               holdings, dividends, documents
     - data/bid_history.json   ONLY for the profile's price
                               chart (shared with Market
                               Performance, so no extra
                               download)
   ============================================================ */

import {
    createChart,
    destroyChart,
    getFontScale
} from "./charts.js";

import {
    isWatched,
    toggleWatchlist
} from "./watchlist.js";

const PROFILE_CHART_ID = "fund-profile-chart";

/* Profile chart ranges (months back from the latest BID). */
const PROFILE_RANGES = [
    ["3M", 3],
    ["6M", 6],
    ["1Y", 12],
    ["3Y", 36],
    ["5Y", 60],
    ["10Y", 120],
    ["SI", null],
    ["CUSTOM", null]
];

const PRUDENTIAL_ORIGIN = "https://www.prudential.com.sg";

const RISK_ORDER = {
    "Lower Risk": 1,
    "Low to Medium Risk": 2,
    "Medium to High Risk": 3,
    "Higher Risk": 4
};

const explorer = {
    initialized: false,
    historyIndex: null,          // set when bid_history.json arrives
    chart: {
        fundId: null,
        range: "1Y",
        from: null,              // custom range (ISO dates)
        to: null
    },
    funds: [],
    rows: [],
    filters: {
        search: "",
        holding: "",
        assetClass: "ALL",
        risk: "ALL",
        geography: "ALL",
        sector: "ALL",
        dividend: "ALL"
    },
    sort: { key: "name", direction: "asc" },
    profile: { id: null }
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

function isBlank(value) {
    const text = String(value ?? "").trim();
    return !text || text === "-" || text === "—";
}

/* "4.14%" -> 4.14 ; "-" -> null */
function parsePercent(value) {
    if (isBlank(value)) return null;

    const number = Number(String(value).replace(/[%,\s]/g, ""));

    return Number.isFinite(number) ? number : null;
}

/* "$1.1079" -> 1.1079 */
function parsePrice(value) {
    if (isBlank(value)) return null;

    const number = Number(String(value).replace(/[^0-9.\-]/g, ""));

    return Number.isFinite(number) ? number : null;
}

function formatPercent(value) {
    if (value === null || value === undefined) return "—";

    const sign = value > 0 ? "+" : "";

    return `${sign}${value.toFixed(2)}%`;
}

function returnClass(value) {
    if (value === null || value === undefined) return "";

    return value >= 0 ? "return-positive" : "return-negative";
}

function absoluteUrl(path) {
    if (isBlank(path)) return "";

    const text = String(path).trim();

    if (/^https?:\/\//i.test(text)) return text;
    if (text.startsWith("/")) return PRUDENTIAL_ORIGIN + text;

    return "";
}

function fundId(fund) {
    return String(fund?.fundIdentifier ?? fund?.fundCode ?? fund?.fundName ?? "");
}

function formatIsoDate(isoDate) {
    if (!isoDate) return "—";

    const [year, month, day] = isoDate.split("-").map(Number);

    return new Date(Date.UTC(year, month - 1, day)).toLocaleDateString("en-SG", {
        timeZone: "UTC",
        day: "2-digit",
        month: "short",
        year: "numeric"
    });
}

function cleanGeo(value) {
    return isBlank(value) ? null : String(value).trim();
}


/* ============================================================
   ROW MODEL
   ============================================================ */

function buildRow(fund) {
    const details = fund.fund ?? {};
    const research = fund.research ?? {};

    const geographies = [research.geographic1, research.geographic2].map(cleanGeo).filter(Boolean);
    const sectors = [research.sector1, research.sector2].map(cleanGeo).filter(Boolean);

    const holdings = Array.isArray(fund.topHoldings?.holdings)
        ? fund.topHoldings.holdings
        : [];

    return {
        id: fundId(fund),
        fund,
        name: fund.fundName ?? "",
        code: fund.fundCode ?? "",
        assetClass: details.assetClass ?? "—",
        assetSubClass: details.assetSubClass ?? "",
        risk: details.riskClassification ?? "—",
        riskOrder: RISK_ORDER[details.riskClassification] ?? 99,
        bid: parsePrice(details.bidPrice),
        valuationDate: details.valuationDate ?? "",
        ytd: parsePercent(details.cumulativeYtd),
        m1: parsePercent(details.cumulative1m),
        y1: parsePercent(details.cumulative1y),
        y3: parsePercent(details.cumulative3y),
        hasDividend: Boolean(details.hasDividend),
        geographies,
        sectors,
        holdings,
        searchText: `${fund.fundName ?? ""} ${fund.fundCode ?? ""} ${fund.pruAccessName ?? ""}`.toLowerCase()
    };
}


/* ============================================================
   FILTER OPTIONS
   ============================================================ */

function fillSelect(selector, values, allLabel, labelFor = value => value) {
    const select = qs(selector);

    if (!select) return;

    const current = select.value || "ALL";

    select.innerHTML =
        `<option value="ALL">${escapeHtml(allLabel)}</option>` +
        values
            .map(value => `<option value="${escapeHtml(value)}">${escapeHtml(labelFor(value))}</option>`)
            .join("");

    select.value = values.includes(current) ? current : "ALL";
}

function countBy(rows, pick) {
    const counts = new Map();

    for (const row of rows) {
        for (const value of new Set(pick(row))) {
            counts.set(value, (counts.get(value) ?? 0) + 1);
        }
    }

    return counts;
}

function renderFilterOptions() {
    const rows = explorer.rows;

    const assetCounts = countBy(rows, row => [row.assetClass]);
    fillSelect(
        "#explorer-asset",
        [...assetCounts.keys()].sort(),
        "All asset classes",
        value => `${value} (${assetCounts.get(value)})`
    );

    const riskCounts = countBy(rows, row => [row.risk]);
    fillSelect(
        "#explorer-risk",
        [...riskCounts.keys()].sort((a, b) => (RISK_ORDER[a] ?? 99) - (RISK_ORDER[b] ?? 99)),
        "All risk levels",
        value => `${value} (${riskCounts.get(value)})`
    );

    const geoCounts = countBy(rows, row => row.geographies);
    fillSelect(
        "#explorer-geography",
        [...geoCounts.keys()].sort((a, b) => geoCounts.get(b) - geoCounts.get(a) || a.localeCompare(b)),
        "All geographies",
        value => `${value} (${geoCounts.get(value)})`
    );

    const sectorCounts = countBy(rows, row => row.sectors);
    fillSelect(
        "#explorer-sector",
        [...sectorCounts.keys()].sort((a, b) => sectorCounts.get(b) - sectorCounts.get(a) || a.localeCompare(b)),
        "All sectors",
        value => `${value} (${sectorCounts.get(value)})`
    );

    // Group spellings that differ only by case ("NVIDIA Corp" / "NVIDIA CORP")
    const holdingGroups = new Map();

    for (const row of rows) {
        const seen = new Set();

        for (const holding of row.holdings) {
            const name = String(holding?.name ?? "").trim();
            const key = name.toLowerCase();

            if (!name || seen.has(key)) continue;

            seen.add(key);

            const group = holdingGroups.get(key) ?? { count: 0, spellings: new Map() };

            group.count += 1;
            group.spellings.set(name, (group.spellings.get(name) ?? 0) + 1);
            holdingGroups.set(key, group);
        }
    }

    explorer.holdingNames = [...holdingGroups.values()]
        .map(group => ({
            name: [...group.spellings.entries()].sort((a, b) => b[1] - a[1] || a[0].localeCompare(b[0]))[0][0],
            count: group.count
        }))
        .sort((a, b) => a.name.localeCompare(b.name));
}


/* ============================================================
   HOLDINGS PICKER (scrollable suggestion list)
   ============================================================ */

function renderHoldingSuggestions() {
    const input = qs("#explorer-holding");
    const list = qs("#explorer-holding-options");

    if (!input || !list) return;

    const query = input.value.trim().toLowerCase();
    const names = explorer.holdingNames ?? [];

    const matches = query
        ? names
            .filter(item => item.name.toLowerCase().includes(query))
            .sort((a, b) =>
                Number(!a.name.toLowerCase().startsWith(query)) - Number(!b.name.toLowerCase().startsWith(query)) ||
                a.name.localeCompare(b.name)
            )
        : names;

    list.innerHTML = matches.length
        ? matches
            .map((item, index) => `
                <li class="compare-suggestion${index === 0 ? " is-highlighted" : ""}" role="option" data-holding-pick="${escapeHtml(item.name)}">
                    <span class="compare-suggestion-name">${escapeHtml(item.name)}</span>
                    <span class="compare-suggestion-code">${item.count} fund${item.count === 1 ? "" : "s"}</span>
                </li>
            `)
            .join("")
        : `<li class="compare-suggestion-empty">No matching holdings</li>`;

    list.hidden = false;
    list.scrollTop = 0;
    input.setAttribute("aria-expanded", "true");
}

function closeHoldingSuggestions() {
    const list = qs("#explorer-holding-options");

    if (list) list.hidden = true;

    qs("#explorer-holding")?.setAttribute("aria-expanded", "false");
}

function pickHolding(name) {
    const input = qs("#explorer-holding");

    if (!input) return;

    input.value = name;
    explorer.filters.holding = name;
    closeHoldingSuggestions();
    renderTable();
}

function moveHoldingHighlight(step) {
    const items = [...document.querySelectorAll("#explorer-holding-options [data-holding-pick]")];

    if (!items.length) return;

    const current = items.findIndex(item => item.classList.contains("is-highlighted"));
    const next = Math.min(Math.max(current + step, 0), items.length - 1);

    items.forEach((item, index) => item.classList.toggle("is-highlighted", index === next));
    items[next].scrollIntoView({ block: "nearest" });
}

function bindHoldingPicker() {
    const input = qs("#explorer-holding");
    const list = qs("#explorer-holding-options");

    if (!input || !list) return;

    input.addEventListener("focus", renderHoldingSuggestions);
    input.addEventListener("input", renderHoldingSuggestions);
    input.addEventListener("blur", () => setTimeout(closeHoldingSuggestions, 120));

    input.addEventListener("keydown", event => {
        if (event.key === "ArrowDown" || event.key === "ArrowUp") {
            event.preventDefault();

            if (list.hidden) renderHoldingSuggestions();

            moveHoldingHighlight(event.key === "ArrowDown" ? 1 : -1);
        } else if (event.key === "Enter") {
            const item = qs("#explorer-holding-options .is-highlighted[data-holding-pick]");

            if (!list.hidden && item) {
                event.preventDefault();
                pickHolding(item.dataset.holdingPick);
            }
        } else if (event.key === "Escape") {
            closeHoldingSuggestions();
        }
    });

    // mousedown keeps focus in the input; scrolling the list works normally
    list.addEventListener("mousedown", event => {
        event.preventDefault();

        const item = event.target.closest("[data-holding-pick]");

        if (item) pickHolding(item.dataset.holdingPick);
    });
}


/* ============================================================
   FILTER + SORT
   ============================================================ */

function holdingMatches(row, term) {
    if (!term) return [];

    return row.holdings.filter(holding =>
        String(holding?.name ?? "").toLowerCase().includes(term)
    );
}

function getVisibleRows() {
    const f = explorer.filters;
    const search = f.search.trim().toLowerCase();
    const holding = f.holding.trim().toLowerCase();

    let rows = explorer.rows.filter(row => {
        if (search && !row.searchText.includes(search)) return false;
        if (f.assetClass !== "ALL" && row.assetClass !== f.assetClass) return false;
        if (f.risk !== "ALL" && row.risk !== f.risk) return false;
        if (f.geography !== "ALL" && !row.geographies.includes(f.geography)) return false;
        if (f.sector !== "ALL" && !row.sectors.includes(f.sector)) return false;
        if (f.dividend === "YES" && !row.hasDividend) return false;
        if (f.dividend === "NO" && row.hasDividend) return false;
        if (holding && holdingMatches(row, holding).length === 0) return false;

        return true;
    });

    const { key, direction } = explorer.sort;
    const factor = direction === "asc" ? 1 : -1;

    const value = row => {
        switch (key) {
            case "code": return row.code;
            case "asset": return row.assetClass;
            case "risk": return row.riskOrder;
            case "bid": return row.bid;
            case "ytd": return row.ytd;
            case "m1": return row.m1;
            case "y1": return row.y1;
            case "y3": return row.y3;
            default: return row.name;
        }
    };

    rows = [...rows].sort((a, b) => {
        const va = value(a);
        const vb = value(b);

        // Missing values always last
        if (va === null || va === undefined) return 1;
        if (vb === null || vb === undefined) return -1;

        if (typeof va === "number" && typeof vb === "number") {
            return (va - vb) * factor || a.name.localeCompare(b.name);
        }

        return String(va).localeCompare(String(vb)) * factor || a.name.localeCompare(b.name);
    });

    return rows;
}


/* ============================================================
   TABLE
   ============================================================ */

function renderTable() {
    const body = qs("#explorer-table-body");
    const count = qs("#explorer-count");
    const clear = qs("#explorer-clear");

    if (!body) return;

    const rows = getVisibleRows();
    const holdingTerm = explorer.filters.holding.trim().toLowerCase();

    const active = Object.entries(explorer.filters).some(([key, value]) =>
        key === "search" || key === "holding" ? value.trim() !== "" : value !== "ALL"
    );

    if (count) {
        count.textContent = active
            ? `${rows.length} of ${explorer.rows.length} funds`
            : `${explorer.rows.length} funds`;
    }

    if (clear) clear.hidden = !active;

    document.querySelectorAll("[data-explorer-filter]").forEach(control => {
        const filtered = control.tagName === "SELECT"
            ? control.value !== "ALL"
            : control.value.trim() !== "";

        control.classList.toggle("is-filtered", filtered);
    });

    const sortSelect = qs("#explorer-sort-mobile");

    if (sortSelect) {
        const value = `${explorer.sort.key}:${explorer.sort.direction}`;

        if (![...sortSelect.options].some(option => option.value === value)) {
            sortSelect.querySelector("option[data-custom]")?.remove();
            sortSelect.insertAdjacentHTML("beforeend", `<option value="${value}" data-custom>Custom</option>`);
        }

        sortSelect.value = value;
    }

    document.querySelectorAll("[data-explorer-sort]").forEach(header => {
        const isActive = header.dataset.explorerSort === explorer.sort.key;

        header.classList.toggle("is-sorted", isActive);
        header.setAttribute(
            "aria-sort",
            isActive ? (explorer.sort.direction === "asc" ? "ascending" : "descending") : "none"
        );

        const arrow = qs(".sort-arrow", header);

        if (arrow) {
            arrow.textContent = isActive ? (explorer.sort.direction === "asc" ? "▲" : "▼") : "";
        }
    });

    if (!rows.length) {
        body.innerHTML = `
            <tr>
                <td colspan="8" class="empty-state">No funds match these filters.</td>
            </tr>
        `;
        return;
    }

    body.innerHTML = rows
        .map(row => {
            const matches = holdingMatches(row, holdingTerm);

            const holdingNote = matches.length
                ? `<div class="explorer-holding-match">Holds ${matches
                    .map(holding => `${escapeHtml(holding.name)} ${escapeHtml(holding.weightText ?? "")}`)
                    .join(", ")}</div>`
                : "";

            const cell = (value, label) => `<td class="return-cell ${returnClass(value)}" data-label="${label}">${formatPercent(value)}</td>`;

            return `
                <tr class="explorer-row" data-fund-open="${escapeHtml(row.id)}" tabindex="0" role="button">
                    <td class="fund-cell">
                        <div class="fund-name">${escapeHtml(row.name)}</div>
                        <div class="fund-meta">
                            ${escapeHtml(row.code)}${row.hasDividend ? ` · <span class="explorer-dividend-tag">Dividend</span>` : ""}
                        </div>
                        ${holdingNote}
                    </td>
                    <td class="explorer-asset-cell">${escapeHtml(row.assetClass)}</td>
                    <td class="explorer-risk-cell"><span class="risk-pill risk-${row.riskOrder}">${escapeHtml(row.risk)}</span></td>
                    <td class="bid-cell" data-label="BID price">${row.bid !== null ? row.bid.toFixed(4) : "—"}</td>
                    ${cell(row.ytd, "YTD")}
                    ${cell(row.m1, "1M")}
                    ${cell(row.y1, "1Y")}
                    ${cell(row.y3, "3Y")}
                </tr>
            `;
        })
        .join("");
}


/* ============================================================
   PROFILE POPUP
   ============================================================ */

function ensureDialog() {
    let dialog = document.getElementById("fund-dialog");

    if (!dialog) {
        dialog = document.createElement("dialog");
        dialog.id = "fund-dialog";
        dialog.className = "news-dialog fund-dialog";
        dialog.setAttribute("aria-labelledby", "fund-dialog-title");

        dialog.addEventListener("click", event => {
            if (event.target === dialog) dialog.close();
        });

        dialog.addEventListener("close", () => {
            document.body.classList.remove("has-news-dialog");
            destroyChart(PROFILE_CHART_ID);
        });

        document.body.appendChild(dialog);
    }

    return dialog;
}

function fact(label, value) {
    if (isBlank(value)) return "";

    return `
        <div class="fund-fact">
            <span class="fund-fact-label">${escapeHtml(label)}</span>
            <span class="fund-fact-value">${escapeHtml(value)}</span>
        </div>
    `;
}

function returnsTable(details) {
    /*
     * One row: short periods are cumulative, long periods
     * (3Y and longer) are annualised (per annum).
     */
    const returns = [
        ["YTD", details.cumulativeYtd, false],
        ["1M", details.cumulative1m, false],
        ["3M", details.cumulative3m, false],
        ["6M", details.cumulative6m, false],
        ["1Y", details.cumulative1y, false],
        ["3Y", details.annualised3y, true],
        ["5Y", details.annualised5y, true],
        ["10Y", details.annualised10y, true],
        ["Since launch", details.annualisedSinceLaunch, true]
    ];

    const cells = returns
        .map(([label, raw, annualised]) => {
            const value = parsePercent(raw);

            return `
                <div class="fund-return${annualised ? " is-annualised" : ""}">
                    <span class="fund-return-label">
                        ${escapeHtml(label)}${annualised ? ` <span class="fund-return-pa">p.a.</span>` : ""}
                    </span>
                    <span class="fund-return-value ${returnClass(value)}">${formatPercent(value)}</span>
                </div>
            `;
        })
        .join("");

    return `
        <section class="news-popup-section">
            <h3 class="news-popup-label">Returns</h3>
            <div class="fund-returns">${cells}</div>
        </section>

        <p class="fund-source-note">
            YTD to 1Y are cumulative; 3Y, 5Y, 10Y and Since launch are annualised (p.a.).
        </p>
    `;
}

function holdingsSection(fund, holdings) {
    const top = fund.topHoldings ?? {};

    if (!holdings.length) {
        return `
            <section class="news-popup-section">
                <h3 class="news-popup-label">Top holdings</h3>
                <p class="news-popup-muted">The factsheet does not publish a holdings list for this fund.</p>
            </section>
        `;
    }

    const max = Math.max(...holdings.map(holding => Number(holding.weightPercent) || 0), 1);
    const term = explorer.filters.holding.trim().toLowerCase();

    return `
        <section class="news-popup-section">
            <div class="fund-section-heading">
                <h3 class="news-popup-label">Top ${holdings.length} holdings</h3>
                ${top.factsheetDataAsAt ? `<span class="fund-heading-date">Data as at ${escapeHtml(top.factsheetDataAsAt)}</span>` : ""}
            </div>

            <ol class="fund-holdings">
                ${holdings
                    .map(holding => {
                        const weight = Number(holding.weightPercent) || 0;
                        const highlighted = term && String(holding.name ?? "").toLowerCase().includes(term);

                        return `
                            <li class="${highlighted ? "is-highlighted" : ""}">
                                <span class="fund-holding-rank">${escapeHtml(holding.rank ?? "")}</span>
                                <span class="fund-holding-name">${escapeHtml(holding.name ?? "")}</span>
                                <span class="fund-holding-bar" aria-hidden="true">
                                    <span style="width:${((weight / max) * 100).toFixed(1)}%"></span>
                                </span>
                                <span class="fund-holding-weight">${escapeHtml(holding.weightText ?? `${weight}%`)}</span>
                            </li>
                        `;
                    })
                    .join("")}
            </ol>

        </section>
    `;
}

function formatRate(rate, unit) {
    const number = Number(rate);
    const text = Number.isFinite(number) ? number.toFixed(2) : String(rate ?? "");

    // Unit as published in funds.json:
    //   "%", "% p.a.", "% per payout" -> "1.25%"
    //   "S$", "SGD", "US$"            -> "S$1.25"
    //   "cents per unit" etc.         -> "1.25 cents per unit"
    const clean = String(unit ?? "").trim();

    if (!clean) return text;
    if (clean.startsWith("%")) return `${text}%`;
    if (/^(S\$|US\$|\$|SGD|USD)$/i.test(clean)) return `${clean}${/^[A-Z]+$/i.test(clean) ? " " : ""}${text}`;

    return `${text} ${clean}`;
}

function dividendSection(details) {
    if (!details.hasDividend) return "";

    const history = Array.isArray(details.dividendHistory) ? details.dividendHistory : [];
    const unit = details.dividendUnit ?? "";

    if (!history.length) {
        return `
            <section class="news-popup-section">
                <h3 class="news-popup-label">Dividends</h3>
                <p>This fund pays dividends.</p>
            </section>
        `;
    }

    return `
        <section class="news-popup-section">
            <h3 class="news-popup-label">Dividend history${unit ? ` (${escapeHtml(unit)})` : ""}</h3>

            <div class="fund-dividends">
                ${history
                    .slice(0, 12)
                    .map(item => `
                        <div class="fund-dividend">
                            <span>${escapeHtml(formatIsoDate(item.exDate))}</span>
                            <strong>${escapeHtml(formatRate(item.rate, unit))}</strong>
                        </div>
                    `)
                    .join("")}
            </div>

            ${history.length > 12 ? `<p class="fund-source-note">Latest 12 of ${history.length} payouts shown.</p>` : ""}

            ${details.dividendUnitSource === "default" ? `
                <p class="fund-source-note">
                    Prudential did not show a unit for this fund's payouts; the default unit is displayed.
                </p>
            ` : ""}
        </section>
    `;
}

function documentsSection(fund) {
    const details = fund.fund ?? {};

    const links = [
        ["Fund page", fund.prudentialUrl],
        ["Factsheet", details.factsheetUrl],
        ["Product highlights", details.productHighlightSheetUrl],
        ["Prospectus", details.prospectusUrl],
        ["Annual report", details.annualReportUrl]
    ]
        .map(([label, url]) => [label, absoluteUrl(url)])
        .filter(([, url]) => url);

    if (!links.length) return "";

    return `
        <section class="news-popup-section">
            <h3 class="news-popup-label">Documents</h3>

            <div class="fund-documents">
                ${links
                    .map(([label, url]) => `
                        <a href="${escapeHtml(url)}" target="_blank" rel="noopener noreferrer">${escapeHtml(label)} ↗</a>
                    `)
                    .join("")}
            </div>
        </section>
    `;
}

/* ============================================================
   PROFILE PRICE CHART
   ============================================================
   Actual BID prices only. Gaps between dates are joined with
   a straight line; nothing is interpolated.
   ============================================================ */

function isoMinusMonths(isoDate, months) {
    const [year, month, day] = isoDate.split("-").map(Number);
    const target = new Date(Date.UTC(year, month - 1 - months, 1));
    const lastDay = new Date(Date.UTC(target.getUTCFullYear(), target.getUTCMonth() + 1, 0)).getUTCDate();

    target.setUTCDate(Math.min(day, lastDay));

    return target.toISOString().slice(0, 10);
}

function chartSectionHtml() {
    const buttons = PROFILE_RANGES
        .map(([key]) => `
            <button
                type="button"
                class="compare-range-tab"
                data-profile-range="${key}"
                role="tab"
                aria-selected="false"
            >${key === "CUSTOM" ? "Custom" : key}</button>
        `)
        .join("");

    return `
        <section class="news-popup-section fund-chart-section">

            <div class="fund-chart-header">
                <div class="fund-section-heading">
                    <h3 class="news-popup-label">BID price</h3>
                    <span id="fund-chart-summary" class="fund-heading-date"></span>
                </div>

                <div class="compare-range-selector" role="tablist" aria-label="Chart range">
                    ${buttons}
                </div>
            </div>

            <div id="fund-chart-custom" class="fund-chart-custom" hidden>
                <label>
                    <span>From</span>
                    <input id="fund-chart-from" type="date">
                </label>
                <label>
                    <span>To</span>
                    <input id="fund-chart-to" type="date">
                </label>
                <span id="fund-chart-custom-error" class="fund-chart-error" role="alert"></span>
            </div>

            <div id="fund-chart-wrap" class="fund-chart-wrap">
                <canvas
                    id="${PROFILE_CHART_ID}"
                    role="img"
                    aria-label="BID price history for this fund"
                ></canvas>
            </div>

            <p id="fund-chart-status" class="news-popup-muted fund-chart-status" hidden></p>

        </section>
    `;
}

function setChartStatus(message) {
    const status = qs("#fund-chart-status");
    const wrap = qs("#fund-chart-wrap");

    if (status) {
        status.textContent = message ?? "";
        status.hidden = !message;
    }

    if (wrap) {
        wrap.hidden = Boolean(message);
    }

    // No chart drawn -> no range summary either.
    if (message) {
        const summary = qs("#fund-chart-summary");

        if (summary) summary.textContent = "";
    }
}

function renderProfileChart() {
    const dialog = document.getElementById("fund-dialog");

    if (!dialog?.open || !qs("#fund-chart-wrap")) return;

    const { fundId, range } = explorer.chart;

    document.querySelectorAll("[data-profile-range]").forEach(button => {
        const active = button.dataset.profileRange === range;

        button.classList.toggle("active", active);
        button.setAttribute("aria-selected", String(active));
    });

    const custom = qs("#fund-chart-custom");

    if (custom) custom.hidden = range !== "CUSTOM";

    if (!explorer.historyIndex) {
        destroyChart(PROFILE_CHART_ID);
        setChartStatus("Loading price history…");
        return;
    }

    const observations = explorer.historyIndex.get(fundId) ?? [];

    if (!observations.length) {
        destroyChart(PROFILE_CHART_ID);
        setChartStatus("No BID price history is available for this fund.");
        return;
    }

    const first = observations[0].date;
    const last = observations[observations.length - 1].date;

    // Custom inputs: bounded to the available history.
    const fromInput = qs("#fund-chart-from");
    const toInput = qs("#fund-chart-to");
    const error = qs("#fund-chart-custom-error");

    for (const input of [fromInput, toInput]) {
        if (input) {
            input.min = first;
            input.max = last;
        }
    }

    let start = first;
    let end = last;

    if (range === "CUSTOM") {
        explorer.chart.from ??= isoMinusMonths(last, 12) < first ? first : isoMinusMonths(last, 12);
        explorer.chart.to ??= last;

        if (fromInput && !fromInput.value) fromInput.value = explorer.chart.from;
        if (toInput && !toInput.value) toInput.value = explorer.chart.to;

        start = explorer.chart.from;
        end = explorer.chart.to;

        if (start > end) {
            if (error) error.textContent = "“From” must be before “To”.";
            destroyChart(PROFILE_CHART_ID);
            setChartStatus("Choose a start date before the end date.");
            return;
        }

        if (error) error.textContent = "";
    } else {
        const months = PROFILE_RANGES.find(([key]) => key === range)?.[1];

        if (months) {
            const target = isoMinusMonths(last, months);

            // Start from the latest actual BID on or before the target.
            let base = observations[0];

            for (const observation of observations) {
                if (observation.date <= target) base = observation;
                else break;
            }

            start = base.date;
        }
    }

    const points = observations.filter(item => item.date >= start && item.date <= end);

    if (points.length < 2) {
        destroyChart(PROFILE_CHART_ID);
        setChartStatus("Not enough BID prices in this date range.");
        return;
    }

    setChartStatus("");

    const firstPoint = points[0];
    const lastPoint = points[points.length - 1];
    const change = ((lastPoint.bidPrice - firstPoint.bidPrice) / firstPoint.bidPrice) * 100;

    const summary = qs("#fund-chart-summary");

    if (summary) {
        summary.innerHTML = `
            ${escapeHtml(formatIsoDate(firstPoint.date))} – ${escapeHtml(formatIsoDate(lastPoint.date))}
            · <span class="${returnClass(change)}">${escapeHtml(formatPercent(change))}</span>
        `;
    }

    if (!window.Chart) {
        setChartStatus("Chart library unavailable.");
        return;
    }

    const style = getComputedStyle(document.documentElement);
    const css = (name, fallback) => style.getPropertyValue(name).trim() || fallback;

    const lineColor = css("--series-1", "#3987e5");
    const muted = css("--text-muted", "#8b98a5");
    const grid = css("--border-subtle", "rgba(255,255,255,0.08)");
    const surface = css("--bg-surface", "#101821");
    const text = css("--text-primary", "#edf2f6");
    const border = css("--border-medium", "#2a3947");
    const scale = getFontScale();

    const spanYears = (Date.parse(lastPoint.date) - Date.parse(firstPoint.date)) / (365.25 * 864e5);

    const crosshair = {
        id: "fundProfileCrosshair",
        afterDatasetsDraw(chart) {
            const active = chart.tooltip?.getActiveElements?.() ?? [];

            if (!active.length) return;

            const x = active[0].element.x;
            const { top, bottom } = chart.chartArea;

            chart.ctx.save();
            chart.ctx.beginPath();
            chart.ctx.moveTo(x, top);
            chart.ctx.lineTo(x, bottom);
            chart.ctx.lineWidth = 1;
            chart.ctx.strokeStyle = border;
            chart.ctx.stroke();
            chart.ctx.restore();
        }
    };

    createChart(PROFILE_CHART_ID, {
        type: "line",

        data: {
            labels: points.map(point => point.date),

            datasets: [
                {
                    label: "BID price",
                    data: points.map(point => point.bidPrice),
                    borderColor: lineColor,
                    backgroundColor: lineColor,
                    borderWidth: 2,
                    pointRadius: 0,
                    pointHoverRadius: 4,
                    pointHoverBorderColor: surface,
                    pointHoverBorderWidth: 2,
                    tension: 0
                }
            ]
        },

        plugins: [crosshair],

        options: {
            responsive: true,
            maintainAspectRatio: false,
            animation: false,

            interaction: { mode: "index", intersect: false },

            plugins: {
                legend: { display: false },

                tooltip: {
                    backgroundColor: surface,
                    titleColor: text,
                    bodyColor: text,
                    borderColor: border,
                    borderWidth: 1,
                    padding: 10,
                    displayColors: false,

                    callbacks: {
                        title: items => formatIsoDate(items[0]?.label),
                        label: item => {
                            const value = Number(item.raw);
                            const fromStart = ((value - firstPoint.bidPrice) / firstPoint.bidPrice) * 100;

                            return [
                                `BID ${value.toFixed(5)}`,
                                `${formatPercent(fromStart)} since ${formatIsoDate(firstPoint.date)}`
                            ];
                        }
                    }
                }
            },

            scales: {
                x: {
                    grid: { display: false },
                    border: { color: grid },

                    ticks: {
                        color: muted,
                        maxRotation: 0,
                        autoSkip: true,
                        maxTicksLimit: 6,
                        font: { size: Math.round(12 * scale) },

                        callback(value) {
                            const label = this.getLabelForValue(value);

                            if (!label) return "";

                            return spanYears > 2 ? label.slice(0, 4) : formatIsoDate(label);
                        }
                    }
                },

                y: {
                    grid: { color: grid },
                    border: { display: false },

                    ticks: {
                        color: muted,
                        maxTicksLimit: 5,
                        font: { size: Math.round(12 * scale) }
                    }
                }
            }
        }
    });
}

function bindChartControls(dialog) {
    dialog.querySelectorAll("[data-profile-range]").forEach(button => {
        button.addEventListener("click", () => {
            explorer.chart.range = button.dataset.profileRange;
            renderProfileChart();
        });
    });

    const fromInput = qs("#fund-chart-from", dialog);
    const toInput = qs("#fund-chart-to", dialog);

    fromInput?.addEventListener("change", () => {
        if (fromInput.value) explorer.chart.from = fromInput.value;
        renderProfileChart();
    });

    toInput?.addEventListener("change", () => {
        if (toInput.value) explorer.chart.to = toInput.value;
        renderProfileChart();
    });
}

function openProfile(id) {
    const row = explorer.rows.find(item => item.id === id);

    if (!row) return;

    explorer.profile.id = id;

    // New fund: reset the custom range, keep the chosen preset.
    explorer.chart.fundId = id;
    explorer.chart.from = null;
    explorer.chart.to = null;

    if (explorer.chart.range === "CUSTOM") explorer.chart.range = "1Y";

    const fund = row.fund;
    const details = fund.fund ?? {};
    const dialog = ensureDialog();

    dialog.innerHTML = `
        <div class="news-popup">

            <header class="news-popup-header">
                <div class="news-popup-meta">
                    <span class="news-source">${escapeHtml(row.code)}</span>
                    <span class="news-date">${escapeHtml(details.fundCurrency ?? "")}</span>
                </div>

                <div class="fund-popup-actions">
                    <button
                        type="button"
                        class="fund-watch-toggle${isWatched(id) ? " is-on" : ""}"
                        data-fund-watch="${escapeHtml(id)}"
                        aria-pressed="${isWatched(id)}"
                        title="Add to / remove from your Monitoring watchlist"
                    >${isWatched(id) ? "★ Watching" : "☆ Watch"}</button>

                    <button type="button" class="news-popup-close" aria-label="Close" title="Close" data-fund-close>×</button>
                </div>
            </header>

            <h2 id="fund-dialog-title" class="news-popup-title">${escapeHtml(row.name)}</h2>

            <div class="fund-popup-badge-row">

                <div class="news-card-badges news-popup-badges">
                    <span class="news-badge news-category">${escapeHtml(row.assetClass)}</span>
                    <span class="risk-pill risk-${row.riskOrder}">${escapeHtml(row.risk)}</span>
                    ${row.hasDividend ? `<span class="news-badge explorer-dividend-badge">Pays dividend</span>` : ""}
                </div>

                ${details.valuationDate ? `
                    <span class="fund-published-date">
                        Published by Prudential as at ${escapeHtml(details.valuationDate)}
                    </span>
                ` : ""}

            </div>

            <div class="news-popup-body">

                <div class="fund-facts">
                    ${fact("BID price", details.bidPrice)}
                    ${fact("Offer price", details.offerPrice)}
                    ${fact("Valuation date", details.valuationDate)}
                    ${fact("Inception", details.inceptionDate)}
                    ${fact("Sub-class", details.assetSubClass)}
                    ${fact("Geography", row.geographies.join(", "))}
                    ${fact("Sector", row.sectors.join(", "))}
                </div>

                ${!isBlank(details.fundObjective) ? `
                    <section class="news-popup-section">
                        <h3 class="news-popup-label">Objective</h3>
                        <p>${escapeHtml(details.fundObjective)}</p>
                    </section>
                ` : ""}

                ${!isBlank(details.investmentManager) ? `
                    <section class="news-popup-section">
                        <h3 class="news-popup-label">Investment manager</h3>
                        <p>${escapeHtml(details.investmentManager)}</p>
                    </section>
                ` : ""}

                ${returnsTable(details)}

                ${chartSectionHtml()}

                ${holdingsSection(fund, row.holdings)}

                ${dividendSection(details)}

                ${documentsSection(fund)}

            </div>

        </div>
    `;

    qs("[data-fund-close]", dialog)?.addEventListener("click", () => dialog.close());

    bindChartControls(dialog);

    qs("[data-fund-watch]", dialog)?.addEventListener("click", event => {
        const button = event.currentTarget;
        const on = toggleWatchlist(button.dataset.fundWatch);

        button.classList.toggle("is-on", on);
        button.setAttribute("aria-pressed", String(on));
        button.textContent = on ? "★ Watching" : "☆ Watch";
    });

    document.body.classList.add("has-news-dialog");

    if (!dialog.open) dialog.showModal();

    qs(".news-popup-body", dialog)?.scrollTo(0, 0);

    requestAnimationFrame(renderProfileChart);
}


/* ============================================================
   EVENTS
   ============================================================ */

function bindEvents() {
    bindHoldingPicker();

    const filterMap = {
        "#explorer-search": "search",
        "#explorer-holding": "holding",
        "#explorer-asset": "assetClass",
        "#explorer-risk": "risk",
        "#explorer-geography": "geography",
        "#explorer-sector": "sector",
        "#explorer-dividend": "dividend"
    };

    for (const [selector, key] of Object.entries(filterMap)) {
        const control = qs(selector);

        if (!control) continue;

        const update = () => {
            explorer.filters[key] = control.value ?? (key === "search" || key === "holding" ? "" : "ALL");
            renderTable();
        };

        control.addEventListener(control.tagName === "SELECT" ? "change" : "input", update);
    }

    qs("#explorer-clear")?.addEventListener("click", () => {
        for (const [selector, key] of Object.entries(filterMap)) {
            const control = qs(selector);
            const empty = key === "search" || key === "holding" ? "" : "ALL";

            if (control) control.value = empty;

            explorer.filters[key] = empty;
        }

        renderTable();
    });

    qs("#explorer-sort-mobile")?.addEventListener("change", event => {
        const [key, direction] = event.target.value.split(":");

        if (!key) return;

        explorer.sort.key = key;
        explorer.sort.direction = direction === "desc" ? "desc" : "asc";
        renderTable();
    });

    document.querySelectorAll("[data-explorer-sort]").forEach(header => {
        const activate = () => {
            const key = header.dataset.explorerSort;

            if (explorer.sort.key === key) {
                explorer.sort.direction = explorer.sort.direction === "asc" ? "desc" : "asc";
            } else {
                explorer.sort.key = key;
                // numbers: best first; text: A-Z
                explorer.sort.direction = ["name", "code", "asset", "risk"].includes(key) ? "asc" : "desc";
            }

            renderTable();
        };

        header.addEventListener("click", activate);

        header.addEventListener("keydown", event => {
            if (event.key === "Enter" || event.key === " ") {
                event.preventDefault();
                activate();
            }
        });
    });

    const body = qs("#explorer-table-body");

    body?.addEventListener("click", event => {
        const row = event.target.closest("[data-fund-open]");

        if (row) openProfile(row.dataset.fundOpen);
    });

    body?.addEventListener("keydown", event => {
        if (event.key !== "Enter" && event.key !== " ") return;

        const row = event.target.closest("[data-fund-open]");

        if (row) {
            event.preventDefault();
            openProfile(row.dataset.fundOpen);
        }
    });
}


/* ============================================================
   PUBLIC API
   ============================================================ */

/*
 * "Data updated" in the page header: when funds.json was
 * built (shown in Singapore time), plus Prudential's
 * valuation date for the prices.
 */
function renderUpdated(generatedAtUtc) {
    const updated = qs("#explorer-updated");
    const valuation = qs("#explorer-valuation");

    if (updated) {
        const date = generatedAtUtc ? new Date(generatedAtUtc) : null;

        updated.textContent = date && !Number.isNaN(date.getTime())
            ? date.toLocaleString("en-SG", {
                timeZone: "Asia/Singapore",
                day: "2-digit",
                month: "short",
                year: "numeric",
                hour: "2-digit",
                minute: "2-digit"
            }) + " SGT"
            : "—";
    }

    if (valuation) {
        const dates = [...new Set(
            explorer.rows
                .map(row => row.valuationDate)
                .filter(Boolean)
        )];

        valuation.textContent = dates.length === 1
            ? `Prices as at ${dates[0]}`
            : dates.length
                ? `Prices as at ${dates.sort().pop()} (latest)`
                : "";
    }
}

function initializeFundExplorer({ funds, generatedAtUtc = null }) {
    explorer.funds = Array.isArray(funds) ? funds : [];
    explorer.rows = explorer.funds.map(buildRow).filter(row => row.id);

    renderFilterOptions();

    renderUpdated(generatedAtUtc);

    if (!explorer.initialized) {
        bindEvents();
        explorer.initialized = true;
    }

    renderTable();
}

/**
 * Called once bid_history.json has loaded (shared with Market
 * Performance). Redraws the chart if a profile is open.
 */
function setFundExplorerHistory(historyIndex) {
    explorer.historyIndex = historyIndex ?? null;

    renderProfileChart();
}

/* Redraw the open profile chart (theme / text size change). */
function refreshFundExplorerChart() {
    renderProfileChart();
}

export {
    initializeFundExplorer,
    setFundExplorerHistory,
    refreshFundExplorerChart,
    openProfile as openFundProfile
};
