/* ============================================================
   VGRAT FMS — APPLICATION CONTROLLER
   ============================================================

   Main responsibilities:
   - Application startup
   - Data loading
   - Theme management
   - View/navigation management
   - Market Performance rendering
   - Market News rendering
   - Chart rendering
   - Application status/error handling

   Data sources:
   - data/funds.json
   - data/bid_history.json
   - data/market_news/analysis/current.json
   ============================================================ */

import {
    loadApplicationData
} from "./data-loader.js";

import {
    calculateAllPerformance,
    formatReturnPercent,
    formatBid,
    formatDisplayDate
} from "./performance.js";

import {
    createNewsState,
    getVisibleNews,
    getAvailableCategories,
    getAvailableYears,
    formatNewsDate,
    getNewsStatistics
} from "./market-news.js";

import {
    createNewsSentimentChart,
    createNewsCategoryChart,
    createNewsImportanceChart,
    createNewsAssetClassChart,
    createNewsGeographyChart,
    createNewsSectorChart,
    updateChartsForTheme,
    destroyAllCharts
} from "./charts.js";


/* ============================================================
   APPLICATION STATE
   ============================================================ */

const state = {

    initialized: false,

    loading: true,

    error: null,

    data: {

        funds: [],

        bidHistory: [],

        marketNews: {

            analyses: [],

            generatedAtSgt: null,

            timezone: null,

            timezoneLabel: null,

            windowDays: null,

            articleCount: 0

        }

    },

    performance: {

        DD: null,

        WW: null,

        MM: null,

        YY: null

    },

    selectedPeriod: "DD",

    currentView: "performance",

    news: {

        initialized: false,

        state: null,

        visibleItems: [],

        tab: "current",

        search: "",

        category: "ALL",

        importance: "ALL",

        sentiment: "ALL"

    }

};


/* ============================================================
   CONSTANTS
   ============================================================ */

const STORAGE_KEYS = {

    theme: "vgrat-fms-theme",

    view: "vgrat-fms-view",

    performancePeriod: "vgrat-fms-performance-period"

};


const VALID_VIEWS = [

    "performance",

    "news",

    "fund-explorer",

    "monitoring"

];


const VALID_PERIODS = [

    "DD",

    "WW",

    "MM",

    "YY"

];


/* ============================================================
   DOM HELPERS
   ============================================================ */

function qs(selector, parent = document) {

    return parent.querySelector(selector);

}


function qsa(selector, parent = document) {

    return Array.from(parent.querySelectorAll(selector));

}


function setText(selector, value) {

    const element = qs(selector);

    if (!element) {

        return;

    }

    element.textContent = value ?? "—";

}


function escapeHtml(value) {

    if (value === null || value === undefined) {

        return "";

    }

    return String(value)

        .replace(/&/g, "&amp;")

        .replace(/</g, "&lt;")

        .replace(/>/g, "&gt;")

        .replace(/"/g, "&quot;")

        .replace(/'/g, "&#039;");

}


function escapeAttribute(value) {

    return escapeHtml(value);

}


/* ============================================================
   THEME
   ============================================================ */

function getStoredTheme() {

    try {

        const stored = localStorage.getItem(
            STORAGE_KEYS.theme
        );

        if (stored === "light" || stored === "dark") {

            return stored;

        }

    } catch {

        // Ignore localStorage failures.

    }

    return "dark";

}


function applyTheme(theme) {

    const normalizedTheme =
        theme === "light"
            ? "light"
            : "dark";

    document.documentElement.dataset.theme =
        normalizedTheme;

    const button = qs("#theme-toggle");

    if (!button) {

        return;

    }

    const icon = qs(".theme-icon", button);

    const label = qs(".theme-label", button);

    const isLight =
        normalizedTheme === "light";

    button.setAttribute(
        "aria-pressed",
        String(isLight)
    );

    button.setAttribute(
        "aria-label",
        isLight
            ? "Switch to dark mode"
            : "Switch to light mode"
    );

    if (icon) {

        icon.textContent =
            isLight
                ? "☀"
                : "☾";

    }

    if (label) {

        label.textContent =
            isLight
                ? "Light"
                : "Dark";

    }

    try {

        localStorage.setItem(
            STORAGE_KEYS.theme,
            normalizedTheme
        );

    } catch {

        // Ignore localStorage failures.

    }

    /*
     * Give Chart.js a chance to update after CSS variables
     * have changed.
     */
    requestAnimationFrame(() => {

        try {

            updateChartsForTheme();

        } catch {

            // Chart layer is optional.

        }

    });

}


function initializeTheme() {

    applyTheme(
        getStoredTheme()
    );

}


function initializeThemeToggle() {

    const button = qs("#theme-toggle");

    if (!button) {

        return;

    }

    button.addEventListener(
        "click",
        () => {

            const current =
                document.documentElement.dataset.theme ===
                "light"
                    ? "light"
                    : "dark";

            applyTheme(
                current === "dark"
                    ? "light"
                    : "dark"
            );

        }
    );

}


/* ============================================================
   APPLICATION STATUS
   ============================================================ */

function setApplicationStatus(
    status,
    message
) {

    const container =
        qs("#app-status");

    const text =
        qs("#app-status-text");

    if (!container) {

        return;

    }

    container.classList.remove(

        "is-loading",

        "is-ready",

        "is-error"

    );

    container.classList.add(

        `is-${status}`

    );

    container.setAttribute(
        "aria-busy",
        status === "loading"
            ? "true"
            : "false"
    );

    if (text) {

        text.textContent =
            message;

    }

}


function setLoadingState() {

    setApplicationStatus(
        "loading",
        "Loading data…"
    );

}


function setReadyState() {

    setApplicationStatus(
        "ready",
        "Data loaded"
    );

}


function setErrorState(message) {

    setApplicationStatus(
        "error",
        "Data error"
    );

    const container =
        qs("#application-error");

    if (!container) {

        return;

    }

    container.innerHTML = `

        <div class="error-panel">

            <div class="error-panel-title">
                Unable to load VGrat FMS data
            </div>

            <div class="error-panel-message">
                ${escapeHtml(message)}
            </div>

            <button
                type="button"
                class="error-retry-button"
                data-action="retry"
            >
                Retry
            </button>

        </div>

    `;

    container.classList.add(
        "is-visible"
    );

    const retry =
        qs(
            '[data-action="retry"]',
            container
        );

    if (retry) {

        retry.addEventListener(
            "click",
            () => {

                window.location.reload();

            }
        );

    }

}


/* ============================================================
   LOCAL STORAGE HELPERS
   ============================================================ */

function getStoredValue(
    key,
    fallback
) {

    try {

        const value =
            localStorage.getItem(key);

        return value ?? fallback;

    } catch {

        return fallback;

    }

}


function setStoredValue(
    key,
    value
) {

    try {

        localStorage.setItem(
            key,
            value
        );

    } catch {

        // Ignore localStorage failures.

    }

}


/* ============================================================
   NAVIGATION
   ============================================================ */

function normalizeView(view) {

    return VALID_VIEWS.includes(view)
        ? view
        : "performance";

}


function setCurrentView(
    view,
    options = {}
) {

    const normalizedView =
        normalizeView(view);

    state.currentView =
        normalizedView;

    setStoredValue(
        STORAGE_KEYS.view,
        normalizedView
    );


    /* ----------------------------------------
       Navigation links
    ----------------------------------------- */

    qsa("[data-view-target]").forEach(
        link => {

            const target =
                link.dataset.viewTarget;

            const active =
                target === normalizedView;

            link.classList.toggle(
                "active",
                active
            );

            if (
                link.hasAttribute(
                    "aria-current"
                )
            ) {

                link.setAttribute(
                    "aria-current",
                    active
                        ? "page"
                        : "false"
                );

            }

        }
    );


    /* ----------------------------------------
       View panels
    ----------------------------------------- */

    qsa("[data-view-panel]").forEach(
        panel => {

            const target =
                panel.dataset.viewPanel;

            const active =
                target === normalizedView;

            panel.classList.toggle(
                "is-active",
                active
            );

            panel.hidden =
                !active;

        }
    );


    /* ----------------------------------------
       Legacy / explicit IDs
    ----------------------------------------- */

    const legacyMap = {

        performance:
            "#market-performance-view",

        news:
            "#market-news-view",

        "fund-explorer":
            "#fund-explorer-view",

        monitoring:
            "#monitoring-view"

    };


    Object.entries(
        legacyMap
    ).forEach(
        ([key, selector]) => {

            const element =
                qs(selector);

            if (!element) {

                return;

            }

            element.hidden =
                key !== normalizedView;

        }
    );


    /* ----------------------------------------
       News chart refresh
    ----------------------------------------- */

    if (
        normalizedView === "news" &&
        state.news.initialized
    ) {

        requestAnimationFrame(
            () => {

                renderNewsCharts();

            }
        );

    }


    /* ----------------------------------------
       URL hash
    ----------------------------------------- */

    if (
        options.updateHash !== false
    ) {

        const hash =
            normalizedView === "performance"
                ? "#market-performance"
                : normalizedView === "news"
                    ? "#market-news"
                    : normalizedView === "fund-explorer"
                        ? "#fund-explorer"
                        : "#monitoring";

        if (
            window.location.hash !== hash
        ) {

            history.replaceState(
                null,
                "",
                hash
            );

        }

    }

}


function getViewFromHash() {

    const hash =
        window.location.hash
            .replace("#", "")
            .trim()
            .toLowerCase();

    const map = {

        "market-performance":
            "performance",

        "performance":
            "performance",

        "market-news":
            "news",

        "news":
            "news",

        "fund-explorer":
            "fund-explorer",

        "monitoring":
            "monitoring"

    };

    return map[hash] ?? null;

}


function initializeNavigation() {

    qsa("[data-view-target]").forEach(
        element => {

            element.addEventListener(
                "click",
                event => {

                    event.preventDefault();

                    const target =
                        element.dataset.viewTarget;

                    setCurrentView(
                        target
                    );

                }
            );

        }
    );


    const hashView =
        getViewFromHash();


    const storedView =
        getStoredValue(
            STORAGE_KEYS.view,
            "performance"
        );


    setCurrentView(
        hashView ??
        storedView,
        {
            updateHash: true
        }
    );


    window.addEventListener(
        "hashchange",
        () => {

            const view =
                getViewFromHash();

            if (view) {

                setCurrentView(
                    view,
                    {
                        updateHash: false
                    }
                );

            }

        }
    );

}


/* ============================================================
   PERFORMANCE
   ============================================================ */

function normalizePeriod(period) {

    return VALID_PERIODS.includes(period)
        ? period
        : "DD";

}


function setSelectedPeriod(
    period
) {

    const normalizedPeriod =
        normalizePeriod(period);

    state.selectedPeriod =
        normalizedPeriod;

    setStoredValue(
        STORAGE_KEYS.performancePeriod,
        normalizedPeriod
    );


    qsa("[data-period]").forEach(
        button => {

            const active =
                button.dataset.period ===
                normalizedPeriod;

            button.classList.toggle(
                "active",
                active
            );

            button.setAttribute(
                "aria-selected",
                String(active)
            );

        }
    );


    renderSelectedPerformance();

}


function initializePerformanceTabs() {

    qsa("[data-period]").forEach(
        button => {

            button.addEventListener(
                "click",
                () => {

                    setSelectedPeriod(
                        button.dataset.period
                    );

                }
            );

        }
    );


    const stored =
        getStoredValue(
            STORAGE_KEYS.performancePeriod,
            "DD"
        );


    setSelectedPeriod(
        normalizePeriod(stored)
    );

}


function buildPerformanceResults() {

    if (
        !state.data.funds.length ||
        !state.data.bidHistory.length
    ) {

        state.performance = {

            DD: null,

            WW: null,

            MM: null,

            YY: null

        };

        return;

    }


    state.performance =
        calculateAllPerformance(

            state.data.funds,

            state.data.bidHistory,

            10

        );

}


function getSelectedPerformance() {

    return state.performance[
        state.selectedPeriod
    ] ?? null;

}


function renderPerformanceSummary(
    performance
) {

    if (!performance) {

        setText(
            "#selected-period-label",
            state.selectedPeriod
        );

        setText(
            "#selected-period-description",
            "No performance data available"
        );

        setText(
            "#performance-period-title",
            "Performance"
        );

        setText(
            "#performance-period-description",
            "No performance data available."
        );

        setText(
            "#performance-eligible-count",
            "—"
        );

        setText(
            "#eligible-fund-count",
            "—"
        );

        setText(
            "#performance-as-of",
            "Latest BID: —"
        );

        setText(
            "#performance-target-date",
            "Comparison target: —"
        );

        setText(
            "#top-winner-return",
            "—"
        );

        setText(
            "#top-winner-name",
            "—"
        );

        setText(
            "#top-loser-return",
            "—"
        );

        setText(
            "#top-loser-name",
            "—"
        );

        return;

    }


    setText(
        "#selected-period-label",
        performance.label
    );


    setText(
        "#selected-period-description",
        performance.description
    );


    setText(
        "#performance-period-title",
        performance.label
    );


    setText(
        "#performance-period-description",
        performance.description
    );


    setText(
        "#performance-eligible-count",
        `${performance.totalEligibleFunds} funds with valid comparison data`
    );


    setText(
        "#eligible-fund-count",
        performance.totalEligibleFunds
    );


    setText(
        "#performance-as-of",
        `Latest BID: ${formatDisplayDate(
            performance.asOfDate
        )}`
    );


    setText(
        "#performance-target-date",
        `Comparison target: ${formatDisplayDate(
            performance.targetDate
        )}`
    );


    const winner =
        performance.winners?.[0] ?? null;


    const loser =
        performance.losers?.[0] ?? null;


    if (winner) {

        setText(
            "#top-winner-return",
            formatReturnPercent(
                winner.returnPercent
            )
        );

        setText(
            "#top-winner-name",
            winner.fundName
        );

    } else {

        setText(
            "#top-winner-return",
            "—"
        );

        setText(
            "#top-winner-name",
            "No valid result"
        );

    }


    if (loser) {

        setText(
            "#top-loser-return",
            formatReturnPercent(
                loser.returnPercent
            )
        );

        setText(
            "#top-loser-name",
            loser.fundName
        );

    } else {

        setText(
            "#top-loser-return",
            "—"
        );

        setText(
            "#top-loser-name",
            "No valid result"
        );

    }

}


function renderPerformanceTable(
    selector,
    rows,
    emptyMessage
) {

    const tbody =
        qs(selector);

    if (!tbody) {

        return;

    }


    if (!rows || rows.length === 0) {

        tbody.innerHTML = `

            <tr>

                <td
                    colspan="5"
                    class="empty-state"
                >
                    ${escapeHtml(emptyMessage)}
                </td>

            </tr>

        `;

        return;

    }


    tbody.innerHTML =
        rows.map(
            (row, index) => {

                const returnValue =
                    Number(row.returnPercent);


                const returnClass =
                    returnValue >= 0
                        ? "return-positive"
                        : "return-negative";


                const sign =
                    returnValue > 0
                        ? "+"
                        : "";


                const fundName =
                    escapeHtml(
                        row.fundName
                    );


                const fundCode =
                    escapeHtml(
                        row.fundCode ??
                        "—"
                    );


                const bid =
                    formatBid(
                        row.currentBid
                    );


                return `

                    <tr
                        class="performance-row"
                        data-fund-identifier="${escapeAttribute(
                            row.fundIdentifier ??
                            ""
                        )}"
                    >

                        <td class="rank-cell">
                            ${index + 1}
                        </td>


                        <td class="fund-cell">

                            <div class="fund-name">
                                ${fundName}
                            </div>

                            <div class="fund-meta">
                                ${
                                    escapeHtml(
                                        row.currency ??
                                        ""
                                    )
                                }
                            </div>

                        </td>


                        <td class="fund-code-cell">
                            ${fundCode}
                        </td>


                        <td
                            class="return-cell ${returnClass}"
                        >
                            ${sign}${formatReturnPercent(
                                returnValue
                            )}
                        </td>


                        <td class="bid-cell">
                            ${escapeHtml(bid)}
                        </td>

                    </tr>

                `;

            }
        ).join("");

}


function renderPerformanceTables(
    performance
) {

    if (!performance) {

        renderPerformanceTable(
            "#winners-table-body",
            [],
            "No performance data available."
        );

        renderPerformanceTable(
            "#losers-table-body",
            [],
            "No performance data available."
        );

        return;

    }


    renderPerformanceTable(

        "#winners-table-body",

        performance.winners,

        "No valid winners for this period."

    );


    renderPerformanceTable(

        "#losers-table-body",

        performance.losers,

        "No valid losers for this period."

    );

}


function renderSelectedPerformance() {

    const performance =
        getSelectedPerformance();


    renderPerformanceSummary(
        performance
    );


    renderPerformanceTables(
        performance
    );

}


function initializePerformance() {

    buildPerformanceResults();

    renderSelectedPerformance();

}


/* ============================================================
   MARKET NEWS
   ============================================================ */

function  {

    const analyses =
        state.data.marketNews?.analyses ??
        [];


    state.news.state =
        createNewsState(
            analyses
        );


    state.news.initialized =
        true;


    state.news.search = "";

    state.news.category = "ALL";

    state.news.importance = "ALL";

    state.news.sentiment = "ALL";


    renderNewsFilterOptions();

    renderNewsSummary();

    renderMarketNews();

}


function renderNewsFilterOptions() {

    const select =
        qs("#market-news-category");

    if (!select) {

        return;

    }


    const analyses =
        state.news?.state?.analyses ??
        [];


    const categories =
        getAvailableCategories(
            analyses
        );


    const currentValue =
        state.news.category;


    select.innerHTML = `

        <option value="ALL">
            All categories
        </option>

        ${categories
            .filter(
                category =>
                    category !== "ALL"
            )
            .map(
                category => `

                    <option
                        value="${escapeAttribute(
                            category
                        )}"
                    >
                        ${escapeHtml(
                            category
                        )}
                    </option>

                `
            )
            .join("")}

    `;


    if (
        currentValue !== "ALL" &&
        categories.includes(
            currentValue
        )
    ) {

        select.value =
            currentValue;

    } else {

        select.value =
            "ALL";

        state.news.category =
            "ALL";

    }

}

function updateNewsFilters() {

    state.news.search =
        qs("#market-news-search")
            ?.value
            ?.trim()
            ?? "";


    state.news.category =
        qs("#market-news-category")
            ?.value
            ?? "ALL";


    state.news.importance =
        qs("#market-news-importance")
            ?.value
            ?? "ALL";


    state.news.sentiment =
        qs("#market-news-sentiment")
            ?.value
            ?? "ALL";


    renderNewsSummary();

    renderMarketNews();

}


function initializeNewsFilters() {

    const search =
        qs("#market-news-search");


    const category =
        qs("#market-news-category");


    const importance =
        qs("#market-news-importance");


    const sentiment =
        qs("#market-news-sentiment");


    if (search) {

        search.addEventListener(
            "input",
            updateNewsFilters
        );

    }


    if (category) {

        category.addEventListener(
            "change",
            updateNewsFilters
        );

    }


    if (importance) {

        importance.addEventListener(
            "change",
            updateNewsFilters
        );

    }


    if (sentiment) {

        sentiment.addEventListener(
            "change",
            updateNewsFilters
        );

    }

}


function getFilteredNews() {

    if (!state.news.state) {

        return [];

    }


    /*
     * Synchronize the application-level controls
     * with the Market News module state.
     */
    state.news.state.searchTerm =
        state.news.search ?? "";


    state.news.state.selectedCategory =
        state.news.category ?? "ALL";


    state.news.state.selectedImportance =
        state.news.importance ?? "ALL";


    state.news.state.selectedSentiment =
        state.news.sentiment ?? "ALL";


    return getVisibleNews(
        state.news.state
    );

}

function renderNewsSummary() {

    const visible =
        getFilteredNews();


    const total =
        state.data.marketNews?.analyses?.length ??
        0;


    const count =
        qs("#market-news-count");


    if (!count) {

        return;

    }


    if (
        state.news.search ||
        state.news.category !== "ALL" ||
        state.news.importance !== "ALL" ||
        state.news.sentiment !== "ALL"
    ) {

        count.textContent =
            `${visible.length} of ${total} articles`;

    } else {

        count.textContent =
            `${visible.length} analyzed articles`;

    }

}


function getImportanceClass(
    importance
) {

    const normalized =
        String(
            importance ?? ""
        ).toUpperCase();


    if (normalized === "HIGH") {

        return "news-importance-high";

    }


    if (normalized === "LOW") {

        return "news-importance-low";

    }


    return "news-importance-medium";

}


function getSentimentClass(
    sentiment
) {

    const normalized =
        String(
            sentiment ?? ""
        ).toUpperCase();


    if (
        normalized === "POSITIVE"
    ) {

        return "news-sentiment-positive";

    }


    if (
        normalized === "NEGATIVE"
    ) {

        return "news-sentiment-negative";

    }


    return "news-sentiment-neutral";

}


function renderTagList(
    values,
    className = ""
) {

    if (
        !Array.isArray(values) ||
        values.length === 0
    ) {

        return "";

    }


    return `

        <div class="news-tags ${className}">

            ${values.map(
                value => `

                    <span class="news-tag">
                        ${escapeHtml(value)}
                    </span>

                `
            ).join("")}

        </div>

    `;

}


function renderNewsCard(
    article
) {

    const title =
        escapeHtml(
            article.title ??
            "Untitled article"
        );


    const summary =
        escapeHtml(
            article.summary ??
            ""
        );


    const investorImpact =
        escapeHtml(
            article.investorImpact ??
            ""
        );


    const source =
        escapeHtml(
            article.source ??
            "Unknown source"
        );


    const category =
        escapeHtml(
            article.category ??
            "MARKET"
        );


    const importance =
        escapeHtml(
            article.importance ??
            "MEDIUM"
        );


    const sentiment =
        escapeHtml(
            article.sentiment ??
            "NEUTRAL"
        );


    const date =
        formatNewsDate(
            article.publishedAtSgt
        );


    const url =
        article.url ??
        "#";


    const assetClasses =
        Array.isArray(
            article.assetClasses
        )
            ? article.assetClasses
            : [];


    const geographies =
        Array.isArray(
            article.geographies
        )
            ? article.geographies
            : [];


    const sectors =
        Array.isArray(
            article.sectors
        )
            ? article.sectors
            : [];


    return `

        <article
            class="news-card"
            data-article-id="${escapeAttribute(
                article.articleId ??
                ""
            )}"
        >


            <div class="news-card-header">

                <div class="news-card-source">

                    <span class="news-source">
                        ${source}
                    </span>

                    <span class="news-date">
                        ${escapeHtml(date)}
                    </span>

                </div>


                <div class="news-card-badges">

                    <span
                        class="news-badge news-category"
                    >
                        ${category}
                    </span>

                    <span
                        class="news-badge ${getImportanceClass(
                            article.importance
                        )}"
                    >
                        ${importance}
                    </span>

                    <span
                        class="news-badge ${getSentimentClass(
                            article.sentiment
                        )}"
                    >
                        ${sentiment}
                    </span>

                </div>

            </div>


            <h3 class="news-card-title">

                <a
                    href="${escapeAttribute(url)}"
                    target="_blank"
                    rel="noopener noreferrer"
                >
                    ${title}
                </a>

            </h3>


            ${
                summary
                    ? `
                        <p class="news-card-summary">
                            ${summary}
                        </p>
                    `
                    : ""
            }


            ${
                investorImpact
                    ? `
                        <div class="news-impact">

                            <span class="news-impact-label">
                                Investor impact
                            </span>

                            <p>
                                ${investorImpact}
                            </p>

                        </div>
                    `
                    : ""
            }


            ${renderTagList(
                assetClasses,
                "news-asset-tags"
            )}


            ${renderTagList(
                geographies,
                "news-geography-tags"
            )}


            ${renderTagList(
                sectors,
                "news-sector-tags"
            )}

        </article>

    `;

}


function renderMarketNews() {

    /*
     * IMPORTANT:
     *
     * #market-news-current is the entire current-news panel.
     *
     * #market-news-articles is the actual article rendering
     * container.
     *
     * This is the selector correction made in this version.
     */
    const container =
        qs("#market-news-articles");


    if (!container) {

        return;

    }


    const articles =
        getFilteredNews();


    if (
        !articles ||
        articles.length === 0
    ) {

        container.innerHTML = `

            <div class="empty-news-state">

                <div class="empty-news-icon">
                    ◈
                </div>

                <h3>
                    No matching market news
                </h3>

                <p>
                    No articles match the current
                    search and filter criteria.
                </p>

            </div>

        `;

        renderNewsCharts();

        return;

    }


    container.innerHTML =
        articles.map(
            renderNewsCard
        ).join("");


    renderNewsCharts();

}


function renderNewsCharts() {

    if (
        !state.news.state
    ) {

        return;

    }


    const articles =
        getFilteredNews();


    try {

        createNewsSentimentChart(
            "news-sentiment-chart",
            articles
        );

        createNewsCategoryChart(
            "news-category-chart",
            articles
        );

        createNewsImportanceChart(
            "news-importance-chart",
            articles
        );

        createNewsAssetClassChart(
            "news-asset-class-chart",
            articles
        );

        createNewsGeographyChart(
            "news-geography-chart",
            articles
        );

        createNewsSectorChart(
            "news-sector-chart",
            articles
        );

    } catch (error) {

        /*
         * Charts are an enhancement.
         * A chart failure must not break the dashboard.
         */
        console.warn(
            "VGrat FMS: chart rendering failed.",
            error
        );

    }

}


function initializeNewsTabs() {

    qsa("[data-news-tab]").forEach(
        button => {

            button.addEventListener(
                "click",
                () => {

                    const tab =
                        button.dataset.newsTab;


                    if (
                        tab !== "current" &&
                        tab !== "history"
                    ) {

                        return;

                    }


                    state.news.tab =
                        tab;


                    qsa("[data-news-tab]")
                        .forEach(
                            other => {

                                const active =
                                    other.dataset.newsTab ===
                                    tab;


                                other.classList.toggle(
                                    "active",
                                    active
                                );


                                other.setAttribute(
                                    "aria-selected",
                                    String(active)
                                );

                            }
                        );


                    const currentPanel =
                        qs(
                            "#market-news-current"
                        );


                    const historyPanel =
                        qs(
                            "#market-news-history"
                        );


                    if (currentPanel) {

                        currentPanel.hidden =
                            tab !== "current";

                    }


                    if (historyPanel) {

                        historyPanel.hidden =
                            tab !== "history";

                    }


                    if (tab === "current") {

                        requestAnimationFrame(
                            () => {

                                renderNewsCharts();

                            }
                        );

                    }

                }
            );

        }
    );

}


function renderNewsMetadata() {

    const news =
        state.data.marketNews;


    if (!news) {

        return;

    }


    setText(
        "#market-news-generated",
        news.generatedAtSgt
            ? formatNewsDate(
                news.generatedAtSgt,
                {
                    includeTime: true
                }
            )
            : "—"
    );


    const windowLabel =
        news.windowDays
            ? `${news.windowDays}-day rolling window`
            : "Rolling window";


    setText(
        "#market-news-window",
        windowLabel
    );


    renderNewsSummary();

}


function initializeNewsState() {

    const marketNews =
        state.data.marketNews ?? {
            analyses: [],
            generatedAtSgt: null,
            timezone: "Asia/Singapore",
            timezoneLabel: "SGT",
            windowDays: 14,
            articleCount: 0
        };


    state.news.state =
        createNewsState(
            marketNews
        );


    state.news.initialized =
        true;


    state.news.search =
        "";

    state.news.category =
        "ALL";

    state.news.importance =
        "ALL";

    state.news.sentiment =
        "ALL";


    /*
     * Keep the module's internal filter state
     * synchronized with the application state.
     */
    state.news.state.searchTerm =
        "";

    state.news.state.selectedCategory =
        "ALL";

    state.news.state.selectedImportance =
        "ALL";

    state.news.state.selectedSentiment =
        "ALL";


    renderNewsFilterOptions();

    renderNewsSummary();

    renderMarketNews();

}


/* ============================================================
   APPLICATION DATA
   ============================================================ */

async function loadApplication() {

    state.loading = true;

    state.error = null;

    setLoadingState();


    try {

        const applicationData =
            await loadApplicationData();


        state.data.funds =
            applicationData.funds ??
            [];


        state.data.bidHistory =
            applicationData.bidHistory ??
            [];


        state.data.marketNews =
            applicationData.marketNews ?? {

                analyses: [],

                generatedAtSgt: null,

                timezone: null,

                timezoneLabel: null,

                windowDays: null,

                articleCount: 0

            };


        state.loading = false;

        state.initialized = true;


        const errorPanel =
            qs("#application-error");


        if (errorPanel) {

            errorPanel.classList.remove(
                "is-visible"
            );

            errorPanel.innerHTML = "";

        }


        setReadyState();


        initializePerformance();

        initializeMarketNews();


        /*
         * Allow the CSS/UI to display the actual number
         * of funds after loading.
         */
        updateFundUniverseStatus();


    } catch (error) {

        console.error(
            "VGrat FMS application initialization failed:",
            error
        );


        state.loading = false;

        state.initialized = false;

        state.error =
            error instanceof Error
                ? error
                : new Error(
                    String(error)
                );


        setErrorState(
            state.error.message
        );

    }

}


/* ============================================================
   FUND UNIVERSE STATUS
   ============================================================ */

function updateFundUniverseStatus() {

    const fundCount =
        state.data.funds.length;


    const element =
        qs("[data-fund-universe-count]");


    if (element) {

        element.textContent =
            String(fundCount);

    }


    /*
     * If the footer contains a dynamic fund count,
     * update it without requiring a dedicated element.
     */
    const footerSource =
        qs(".footer-source");


    if (
        footerSource &&
        fundCount > 0 &&
        !footerSource.dataset.initialized
    ) {

        footerSource.textContent =
            `Prudential Singapore fund data · ${fundCount} funds`;

        footerSource.dataset.initialized =
            "true";

    }

}


/* ============================================================
   WINDOW RESIZE
   ============================================================ */

function initializeResizeHandling() {

    let resizeTimer = null;


    window.addEventListener(
        "resize",
        () => {

            clearTimeout(
                resizeTimer
            );


            resizeTimer =
                setTimeout(
                    () => {

                        if (
                            state.currentView ===
                            "news"
                        ) {

                            try {

                                renderNewsCharts();

                            } catch {

                                // Ignore chart resize errors.

                            }

                        }

                    },
                    150
                );

        }
    );

}


/* ============================================================
   APPLICATION DIAGNOSTICS
   ============================================================ */

function exposeDiagnostics() {

    window.VGratFMS = {

        getState() {

            return state;

        },

        getPerformance(
            period = state.selectedPeriod
        ) {

            return state.performance[
                normalizePeriod(period)
            ];

        },

        getNews() {

            return {

                ...state.news,

                visibleItems:
                    getFilteredNews()

            };

        },

        reload() {

            return loadApplication();

        }

    };

}


/* ============================================================
   INITIALIZATION
   ============================================================ */

async function initializeApplication() {

    initializeTheme();

    initializeThemeToggle();

    initializeNavigation();

    initializePerformanceTabs();

    initializeResizeHandling();

    exposeDiagnostics();

    await loadApplication();

}


/* ============================================================
   DOM READY
   ============================================================ */

if (
    document.readyState ===
    "loading"
) {

    document.addEventListener(
        "DOMContentLoaded",
        () => {

            initializeApplication()
                .catch(
                    error => {

                        console.error(
                            "VGrat FMS startup error:",
                            error
                        );

                        setErrorState(
                            error instanceof Error
                                ? error.message
                                : String(error)
                        );

                    }
                );

        },
        {
            once: true
        }
    );

} else {

    initializeApplication()
        .catch(
            error => {

                console.error(
                    "VGrat FMS startup error:",
                    error
                );

                setErrorState(
                    error instanceof Error
                        ? error.message
                        : String(error)
                );

            }
        );

}


/* ============================================================
   END OF FILE
   ============================================================ */
