/* ============================================================
   VGRAT FMS
   APPLICATION CONTROLLER
   ============================================================

   Main responsibilities:
   - Load application data
   - Initialise theme
   - Control main navigation
   - Control performance period
   - Render performance rankings
   - Render Market News
   - Connect charts
   - Handle application errors

   Modules:
   - data-loader.js
   - performance.js
   - market-news.js
   - charts.js
============================================================ */

import {
    loadApplicationData
} from "./data-loader.js";

import {
    PERFORMANCE_PERIODS,
    calculateAllPerformance,
    formatReturnPercent,
    formatBid,
    formatDisplayDate
} from "./performance.js";

import {
    filterArticles,
    searchArticles,
    formatNewsDate
} from "./market-news.js";

import {
    createPerformanceBarChart,
    destroyAllCharts
} from "./charts.js";


/* ============================================================
   APPLICATION STATE
============================================================ */

const state = {
    initialized: false,
    loading: true,
    error: null,

    theme:
        localStorage.getItem("vgrat-fms-theme") || "dark",

    activeView:
        "market-performance",

    activePeriod:
        "DD",

    selectedFund:
        null,

    funds: [],
    bidHistory: [],

    marketNews: {
        analyses: [],
        generatedAtSgt: null,
        timezone: "Asia/Singapore",
        timezoneLabel: "SGT",
        windowDays: null,
        articleCount: 0
    },

    performance: null,

    marketNewsSearch: "",

    marketNewsFilters: {
        category: "",
        sentiment: "",
        importance: ""
    }
};


/* ============================================================
   DOM CACHE
============================================================ */

const dom = {};


function cacheDom() {

    dom.app =
        document.querySelector("#app");

    dom.dataStatus =
        document.querySelector("#data-status");

    dom.dataStatusText =
        document.querySelector("#data-status-text");

    dom.themeToggle =
        document.querySelector("#theme-toggle");

    dom.themeToggleLabel =
        document.querySelector("#theme-toggle-label");

    dom.pageTitle =
        document.querySelector("#page-title");

    dom.pageSubtitle =
        document.querySelector("#page-subtitle");

    dom.latestValuationDate =
        document.querySelector("#latest-valuation-date");

    dom.navLinks =
        document.querySelectorAll("[data-view]");

    dom.periodTabs =
        document.querySelectorAll("[data-period]");

    dom.performanceView =
        document.querySelector("#market-performance-view");

    dom.marketNewsView =
        document.querySelector("#market-news-view");

    dom.fundExplorerView =
        document.querySelector("#fund-explorer-view");

    dom.monitoringView =
        document.querySelector("#monitoring-view");

    dom.selectedPeriodLabel =
        document.querySelector("#selected-period-label");

    dom.selectedPeriodDescription =
        document.querySelector(
            "#selected-period-description"
        );

    dom.eligibleFundCount =
        document.querySelector("#eligible-fund-count");

    dom.winnerCount =
        document.querySelector("#winner-count");

    dom.loserCount =
        document.querySelector("#loser-count");

    dom.winnersTableBody =
        document.querySelector("#winners-table-body");

    dom.losersTableBody =
        document.querySelector("#losers-table-body");

    dom.winnersEmpty =
        document.querySelector("#winners-empty");

    dom.losersEmpty =
        document.querySelector("#losers-empty");

    dom.winnersCanvas =
        document.querySelector("#winners-chart");

    dom.losersCanvas =
        document.querySelector("#losers-chart");

    dom.marketNewsContainer =
        document.querySelector(
            "#market-news-container"
        );

    dom.marketNewsSearch =
        document.querySelector(
            "#market-news-search"
        );

    dom.marketNewsCategory =
        document.querySelector(
            "#market-news-category"
        );

    dom.marketNewsSentiment =
        document.querySelector(
            "#market-news-sentiment"
        );

    dom.marketNewsImportance =
        document.querySelector(
            "#market-news-importance"
        );

    dom.marketNewsCount =
        document.querySelector(
            "#market-news-count"
        );

    dom.marketNewsGenerated =
        document.querySelector(
            "#market-news-generated"
        );

    dom.marketNewsCurrentTab =
        document.querySelector(
            "[data-news-tab='current']"
        );

    dom.marketNewsHistoryTab =
        document.querySelector(
            "[data-news-tab='history']"
        );

    dom.marketNewsCurrentPanel =
        document.querySelector(
            "#market-news-current-panel"
        );

    dom.marketNewsHistoryPanel =
        document.querySelector(
            "#market-news-history-panel"
        );

    dom.globalError =
        document.querySelector("#global-error");
}


/* ============================================================
   THEME
============================================================ */

function applyTheme(theme) {

    const safeTheme =
        theme === "light"
            ? "light"
            : "dark";

    state.theme =
        safeTheme;

    document.documentElement.dataset.theme =
        safeTheme;

    localStorage.setItem(
        "vgrat-fms-theme",
        safeTheme
    );

    if (dom.themeToggle) {

        dom.themeToggle.setAttribute(
            "aria-pressed",
            safeTheme === "light"
                ? "true"
                : "false"
        );
    }

    if (dom.themeToggleLabel) {

        dom.themeToggleLabel.textContent =
            safeTheme === "light"
                ? "Light"
                : "Dark";
    }
}


function toggleTheme() {

    applyTheme(
        state.theme === "dark"
            ? "light"
            : "dark"
    );

    if (
        state.activeView ===
        "market-performance"
    ) {
        renderPerformance();
    }
}


/* ============================================================
   DATA STATUS
============================================================ */

function setDataStatus(
    status,
    message
) {

    if (dom.dataStatus) {

        dom.dataStatus.dataset.status =
            status;
    }

    if (dom.dataStatusText) {

        dom.dataStatusText.textContent =
            message;
    }
}


function setLoadingStatus() {

    setDataStatus(
        "loading",
        "Loading data"
    );
}


function setReadyStatus() {

    setDataStatus(
        "ready",
        "Data current"
    );
}


function setErrorStatus() {

    setDataStatus(
        "error",
        "Data error"
    );
}


/* ============================================================
   ERROR HANDLING
============================================================ */

function showGlobalError(error) {

    const message =
        error instanceof Error
            ? error.message
            : String(error);

    state.error =
        message;

    if (dom.globalError) {

        dom.globalError.hidden =
            false;

        dom.globalError.textContent =
            message;
    }

    setErrorStatus();
}


function clearGlobalError() {

    state.error =
        null;

    if (dom.globalError) {

        dom.globalError.hidden =
            true;

        dom.globalError.textContent =
            "";
    }
}


/* ============================================================
   MAIN VIEW NAVIGATION
============================================================ */

function setActiveView(view) {

    const validViews =
        new Set([
            "market-performance",
            "market-news",
            "fund-explorer",
            "monitoring"
        ]);

    if (!validViews.has(view)) {

        view =
            "market-performance";
    }

    state.activeView =
        view;


    /*
     * Navigation buttons.
     */
    for (
        const link of dom.navLinks || []
    ) {

        const isActive =
            link.dataset.view === view;

        link.classList.toggle(
            "active",
            isActive
        );

        if (
            link.hasAttribute(
                "aria-current"
            )
        ) {

            link.setAttribute(
                "aria-current",
                isActive
                    ? "page"
                    : "false"
            );
        }
    }


    /*
     * Page views.
     */
    const views = {
        "market-performance":
            dom.performanceView,

        "market-news":
            dom.marketNewsView,

        "fund-explorer":
            dom.fundExplorerView,

        "monitoring":
            dom.monitoringView
    };

    for (
        const [key, element]
        of Object.entries(views)
    ) {

        if (!element) {
            continue;
        }

        element.hidden =
            key !== view;
    }


    updatePageHeading();


    if (
        view ===
        "market-performance"
    ) {

        renderPerformance();
    }


    if (
        view ===
        "market-news"
    ) {

        renderMarketNews();
    }
}


/* ============================================================
   PAGE HEADING
============================================================ */

function updatePageHeading() {

    const headings = {

        "market-performance": {
            title:
                "Market Performance",

            subtitle:
                "Prudential Singapore fund performance across the full universe."
        },

        "market-news": {
            title:
                "Market News",

            subtitle:
                "AI-analyzed market developments relevant to investors and Prudential funds."
        },

        "fund-explorer": {
            title:
                "Fund Explorer",

            subtitle:
                "Explore Prudential Singapore funds, characteristics and holdings."
        },

        "monitoring": {
            title:
                "Monitoring",

            subtitle:
                "Monitor selected funds, market conditions and important changes."
        }
    };


    const current =
        headings[
            state.activeView
        ] ||
        headings["market-performance"];


    if (dom.pageTitle) {

        dom.pageTitle.textContent =
            current.title;
    }


    if (dom.pageSubtitle) {

        dom.pageSubtitle.textContent =
            current.subtitle;
    }
}


/* ============================================================
   PERFORMANCE PERIOD
============================================================ */

function setActivePeriod(
    periodKey
) {

    if (
        !Object.prototype.hasOwnProperty.call(
            PERFORMANCE_PERIODS,
            periodKey
        )
    ) {

        return;
    }

    state.activePeriod =
        periodKey;


    for (
        const tab of dom.periodTabs || []
    ) {

        const isActive =
            tab.dataset.period ===
            periodKey;

        tab.classList.toggle(
            "active",
            isActive
        );

        tab.setAttribute(
            "aria-selected",
            isActive
                ? "true"
                : "false"
        );
    }


    renderPerformance();
}


/* ============================================================
   PERFORMANCE RENDERING
============================================================ */

function renderPerformance() {

    if (
        !state.performance
    ) {

        renderPerformanceEmpty();

        return;
    }


    const period =
        state.performance[
            state.activePeriod
        ];


    if (!period) {

        renderPerformanceEmpty();

        return;
    }


    if (dom.selectedPeriodLabel) {

        dom.selectedPeriodLabel.textContent =
            period.label;
    }


    if (
        dom.selectedPeriodDescription
    ) {

        dom.selectedPeriodDescription.textContent =
            period.description;
    }


    if (dom.eligibleFundCount) {

        dom.eligibleFundCount.textContent =
            String(
                period.totalEligibleFunds
            );
    }


    if (dom.winnerCount) {

        dom.winnerCount.textContent =
            String(
                period.winners.length
            );
    }


    if (dom.loserCount) {

        dom.loserCount.textContent =
            String(
                period.losers.length
            );
    }


    renderPerformanceTable(
        dom.winnersTableBody,
        period.winners,
        "winners"
    );


    renderPerformanceTable(
        dom.losersTableBody,
        period.losers,
        "losers"
    );


    updateEmptyState(
        dom.winnersEmpty,
        period.winners.length === 0
    );


    updateEmptyState(
        dom.losersEmpty,
        period.losers.length === 0
    );


    renderPerformanceCharts(
        period
    );


    renderLatestValuationDate();
}


function renderPerformanceEmpty() {

    if (dom.winnersTableBody) {

        dom.winnersTableBody.innerHTML =
            "";
    }


    if (dom.losersTableBody) {

        dom.losersTableBody.innerHTML =
            "";
    }


    if (dom.eligibleFundCount) {

        dom.eligibleFundCount.textContent =
            "—";
    }


    if (dom.winnerCount) {

        dom.winnerCount.textContent =
            "0";
    }


    if (dom.loserCount) {

        dom.loserCount.textContent =
            "0";
    }


    updateEmptyState(
        dom.winnersEmpty,
        true
    );


    updateEmptyState(
        dom.losersEmpty,
        true
    );


    destroyAllCharts();
}


function updateEmptyState(
    element,
    visible
) {

    if (!element) {
        return;
    }

    element.hidden =
        !visible;
}


/* ============================================================
   PERFORMANCE TABLE
============================================================ */

function renderPerformanceTable(
    tableBody,
    results,
    type
) {

    if (!tableBody) {
        return;
    }


    tableBody.innerHTML =
        "";


    if (
        !Array.isArray(results) ||
        results.length === 0
    ) {

        return;
    }


    results.forEach(
        (
            result,
            index
        ) => {

            const row =
                document.createElement("tr");


            /*
             * Rank
             */
            const rankCell =
                document.createElement("td");

            rankCell.className =
                "rank-cell";

            rankCell.textContent =
                String(index + 1);


            /*
             * Fund
             */
            const fundCell =
                document.createElement("td");

            fundCell.className =
                "fund-cell";


            const fundButton =
                document.createElement("button");

            fundButton.type =
                "button";

            fundButton.className =
                "fund-link";

            fundButton.textContent =
                result.fundName ||
                "Unnamed fund";

            fundButton.dataset.fundIdentifier =
                result.fundIdentifier ||
                "";

            fundButton.dataset.fundCode =
                result.fundCode ||
                "";


            fundButton.addEventListener(
                "click",
                () => {
                    handleFundClick(
                        result
                    );
                }
            );


            fundCell.appendChild(
                fundButton
            );


            /*
             * Fund code
             */
            const codeCell =
                document.createElement("td");

            codeCell.className =
                "code-cell";

            codeCell.textContent =
                result.fundCode ||
                "—";


            /*
             * Return
             */
            const returnCell =
                document.createElement("td");

            returnCell.className =
                "return-cell";


            const returnValue =
                Number(
                    result.returnPercent
                );


            returnCell.textContent =
                formatReturnPercent(
                    returnValue
                );


            if (
                returnValue > 0
            ) {

                returnCell.classList.add(
                    "positive"
                );

            } else if (
                returnValue < 0
            ) {

                returnCell.classList.add(
                    "negative"
                );

            } else {

                returnCell.classList.add(
                    "neutral"
                );
            }


            /*
             * Current BID
             */
            const bidCell =
                document.createElement("td");

            bidCell.className =
                "bid-cell";

            bidCell.textContent =
                formatBid(
                    result.currentBid
                );


            /*
             * Comparison date
             */
            const comparisonCell =
                document.createElement("td");

            comparisonCell.className =
                "comparison-cell";

            comparisonCell.textContent =
                formatDisplayDate(
                    result.comparisonDate
                );


            row.dataset.comparisonDate =
                result.comparisonDate ||
                "";

            row.dataset.type =
                type;


            row.appendChild(
                rankCell
            );

            row.appendChild(
                fundCell
            );

            row.appendChild(
                codeCell
            );

            row.appendChild(
                returnCell
            );

            row.appendChild(
                bidCell
            );

            row.appendChild(
                comparisonCell
            );


            tableBody.appendChild(
                row
            );
        }
    );
}


/* ============================================================
   PERFORMANCE CHARTS
============================================================ */

function renderPerformanceCharts(
    period
) {

    if (
        !dom.winnersCanvas &&
        !dom.losersCanvas
    ) {

        return;
    }


    if (dom.winnersCanvas) {

        createPerformanceBarChart(
            dom.winnersCanvas,
            period.winners,
            {
                chartId:
                    "market-winners-chart",

                direction:
                    "winners",

                limit:
                    10
            }
        );
    }


    if (dom.losersCanvas) {

        createPerformanceBarChart(
            dom.losersCanvas,
            period.losers,
            {
                chartId:
                    "market-losers-chart",

                direction:
                    "losers",

                limit:
                    10
            }
        );
    }
}


/* ============================================================
   LATEST VALUATION DATE
============================================================ */

function renderLatestValuationDate() {

    if (
        !dom.latestValuationDate ||
        !Array.isArray(
            state.bidHistory
        )
    ) {

        return;
    }


    let latestDate =
        null;


    for (
        const record of state.bidHistory
    ) {

        const observations =
            record?.bidHistory?.observations;


        if (
            !Array.isArray(
                observations
            )
        ) {

            continue;
        }


        for (
            const observation of observations
        ) {

            const date =
                observation?.date;


            const bid =
                Number(
                    observation?.bidPrice
                );


            if (
                typeof date !== "string" ||
                !Number.isFinite(bid) ||
                bid <= 0
            ) {

                continue;
            }


            if (
                latestDate === null ||
                date > latestDate
            ) {

                latestDate =
                    date;
            }
        }
    }


    dom.latestValuationDate.textContent =
        latestDate
            ? formatDisplayDate(
                latestDate
            )
            : "—";
}


/* ============================================================
   FUND NAVIGATION
============================================================ */

function handleFundClick(
    result
) {

    const identifier =
        result?.fundIdentifier ||
        result?.fundCode ||
        "";


    if (!identifier) {
        return;
    }


    state.selectedFund =
        identifier;


    /*
     * Fund Explorer is not yet implemented.
     *
     * We still switch to it so that the eventual detail
     * implementation has a clean navigation path.
     */
    setActiveView(
        "fund-explorer"
    );
}


/* ============================================================
   MARKET NEWS
============================================================ */

function renderMarketNews() {

    if (
        !dom.marketNewsContainer
    ) {

        return;
    }


    const allArticles =
        Array.isArray(
            state.marketNews?.analyses
        )
            ? state.marketNews.analyses
            : [];


    let articles =
        filterArticles(
            allArticles,
            state.marketNewsFilters
        );


    articles =
        searchArticles(
            articles,
            state.marketNewsSearch
        );


    /*
     * Newest article first.
     */
    articles.sort(
        (a, b) => {

            const aTime =
                parseNewsTimestamp(
                    a?.publishedAtSgt
                );

            const bTime =
                parseNewsTimestamp(
                    b?.publishedAtSgt
                );

            return bTime - aTime;
        }
    );


    renderMarketNewsHeader(
        articles
    );


    if (
        articles.length === 0
    ) {

        dom.marketNewsContainer.innerHTML =
            `
            <div class="empty-state">

                <div class="empty-state-title">
                    No matching news
                </div>

                <div class="empty-state-text">
                    No analyzed articles match the
                    current filters.
                </div>

            </div>
            `;

        return;
    }


    dom.marketNewsContainer.innerHTML =
        articles
            .map(
                article =>
                    renderNewsArticle(
                        article
                    )
            )
            .join("");
}


function parseNewsTimestamp(
    value
) {

    if (
        typeof value !== "string"
    ) {

        return 0;
    }


    const timestamp =
        Date.parse(value);


    return Number.isFinite(
        timestamp
    )
        ? timestamp
        : 0;
}


/* ============================================================
   MARKET NEWS HEADER
============================================================ */

function renderMarketNewsHeader(
    articles
) {

    if (dom.marketNewsCount) {

        dom.marketNewsCount.textContent =
            String(
                articles.length
            );
    }


    if (dom.marketNewsGenerated) {

        dom.marketNewsGenerated.textContent =
            state.marketNews.generatedAtSgt
                ? formatNewsDate(
                    state.marketNews.generatedAtSgt
                )
                : "—";
    }
}


/* ============================================================
   MARKET NEWS ARTICLE
============================================================ */

function renderNewsArticle(
    article
) {

    const category =
        escapeHtml(
            article?.category ||
            "MARKET"
        );


    const sentiment =
        escapeHtml(
            article?.sentiment ||
            "NEUTRAL"
        );


    const importance =
        escapeHtml(
            article?.importance ||
            "MEDIUM"
        );


    const title =
        escapeHtml(
            article?.title ||
            "Untitled article"
        );


    const summary =
        escapeHtml(
            article?.summary ||
            ""
        );


    const investorImpact =
        escapeHtml(
            article?.investorImpact ||
            ""
        );


    const source =
        escapeHtml(
            article?.source ||
            "CNBC"
        );


    const published =
        escapeHtml(
            formatNewsDate(
                article?.publishedAtSgt
            )
        );


    const assetClasses =
        renderNewsTags(
            article?.assetClasses,
            "asset"
        );


    const geographies =
        renderNewsTags(
            article?.geographies,
            "geography"
        );


    const sectors =
        renderNewsTags(
            article?.sectors,
            "sector"
        );


    const sourceLink =
        article?.url
            ? `
                <a
                    class="news-source-link"
                    href="${escapeAttribute(
                        article.url
                    )}"
                    target="_blank"
                    rel="noopener noreferrer"
                >
                    Read source
                </a>
            `
            : "";


    const monitoringBadge =
        article?.fundMonitoringRelevant
            ? `
                <span class="news-badge news-badge-monitoring">
                    Fund Monitoring
                </span>
            `
            : "";


    return `
        <article
            class="news-article"
            data-article-id="${escapeAttribute(
                article?.articleId || ""
            )}"
        >

            <div class="news-article-header">

                <div class="news-article-meta">

                    <span class="news-badge">
                        ${category}
                    </span>

                    <span
                        class="news-badge news-sentiment-${escapeAttribute(
                            sentiment.toLowerCase()
                        )}"
                    >
                        ${sentiment}
                    </span>

                    <span class="news-badge">
                        ${importance}
                    </span>

                    ${monitoringBadge}

                </div>

                <time
                    class="news-published"
                    datetime="${escapeAttribute(
                        article?.publishedAtSgt || ""
                    )}"
                >
                    ${published}
                </time>

            </div>


            <h3 class="news-article-title">
                ${title}
            </h3>


            ${
                summary
                    ? `
                        <p class="news-article-summary">
                            ${summary}
                        </p>
                    `
                    : ""
            }


            ${
                investorImpact
                    ? `
                        <div class="news-investor-impact">

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


            ${
                assetClasses
                    ? `
                        <div class="news-tag-group">

                            <span class="news-tag-label">
                                Asset class
                            </span>

                            <div class="news-tags">
                                ${assetClasses}
                            </div>

                        </div>
                    `
                    : ""
            }


            ${
                geographies
                    ? `
                        <div class="news-tag-group">

                            <span class="news-tag-label">
                                Geography
                            </span>

                            <div class="news-tags">
                                ${geographies}
                            </div>

                        </div>
                    `
                    : ""
            }


            ${
                sectors
                    ? `
                        <div class="news-tag-group">

                            <span class="news-tag-label">
                                Sector
                            </span>

                            <div class="news-tags">
                                ${sectors}
                            </div>

                        </div>
                    `
                    : ""
            }


            <div class="news-article-footer">

                <span class="news-source">
                    ${source}
                </span>

                ${sourceLink}

            </div>

        </article>
    `;
}


function renderNewsTags(
    values,
    type
) {

    if (
        !Array.isArray(values) ||
        values.length === 0
    ) {

        return "";
    }


    return values
        .filter(
            value =>
                value !== null &&
                value !== undefined &&
                String(value).trim() !== ""
        )
        .map(
            value => `
                <span
                    class="news-tag news-tag-${escapeAttribute(
                        type
                    )}"
                >
                    ${escapeHtml(value)}
                </span>
            `
        )
        .join("");
}


/* ============================================================
   MARKET NEWS FILTER OPTIONS
============================================================ */

function buildNewsFilterOptions() {

    const articles =
        Array.isArray(
            state.marketNews?.analyses
        )
            ? state.marketNews.analyses
            : [];


    const categories =
        new Set();

    const sentiments =
        new Set();

    const importance =
        new Set();


    for (
        const article of articles
    ) {

        if (
            article?.category
        ) {

            categories.add(
                article.category
            );
        }


        if (
            article?.sentiment
        ) {

            sentiments.add(
                article.sentiment
            );
        }


        if (
            article?.importance
        ) {

            importance.add(
                article.importance
            );
        }
    }


    populateSelect(
        dom.marketNewsCategory,
        [...categories].sort(),
        "All categories"
    );


    populateSelect(
        dom.marketNewsSentiment,
        [...sentiments].sort(),
        "All sentiment"
    );


    populateSelect(
        dom.marketNewsImportance,
        [...importance].sort(),
        "All importance"
    );
}


function populateSelect(
    select,
    values,
    defaultLabel
) {

    if (!select) {
        return;
    }


    const currentValue =
        select.value;


    select.innerHTML =
        "";


    const defaultOption =
        document.createElement(
            "option"
        );


    defaultOption.value =
        "";

    defaultOption.textContent =
        defaultLabel;


    select.appendChild(
        defaultOption
    );


    for (
        const value of values || []
    ) {

        const option =
            document.createElement(
                "option"
            );


        option.value =
            value;

        option.textContent =
            value;


        select.appendChild(
            option
        );
    }


    if (
        values.includes(
            currentValue
        )
    ) {

        select.value =
            currentValue;
    }
}


/* ============================================================
   NEWS FILTER EVENTS
============================================================ */

function applyMarketNewsFilters() {

    state.marketNewsFilters = {

        category:
            dom.marketNewsCategory?.value ||
            "",

        sentiment:
            dom.marketNewsSentiment?.value ||
            "",

        importance:
            dom.marketNewsImportance?.value ||
            ""
    };


    renderMarketNews();
}


/* ============================================================
   NEWS TAB
============================================================ */

function setNewsTab(
    tab
) {

    const isCurrent =
        tab === "current";


    if (
        dom.marketNewsCurrentTab
    ) {

        dom.marketNewsCurrentTab.classList.toggle(
            "active",
            isCurrent
        );

        dom.marketNewsCurrentTab.setAttribute(
            "aria-selected",
            isCurrent
                ? "true"
                : "false"
        );
    }


    if (
        dom.marketNewsHistoryTab
    ) {

        dom.marketNewsHistoryTab.classList.toggle(
            "active",
            !isCurrent
        );

        dom.marketNewsHistoryTab.setAttribute(
            "aria-selected",
            !isCurrent
                ? "true"
                : "false"
        );
    }


    if (
        dom.marketNewsCurrentPanel
    ) {

        dom.marketNewsCurrentPanel.hidden =
            !isCurrent;
    }


    if (
        dom.marketNewsHistoryPanel
    ) {

        dom.marketNewsHistoryPanel.hidden =
            isCurrent;
    }
}


/* ============================================================
   EVENT BINDING
============================================================ */

function bindEvents() {

    /*
     * Theme.
     */
    dom.themeToggle?.addEventListener(
        "click",
        toggleTheme
    );


    /*
     * Main navigation.
     */
    for (
        const link of dom.navLinks || []
    ) {

        link.addEventListener(
            "click",
            event => {

                event.preventDefault();

                setActiveView(
                    link.dataset.view
                );
            }
        );
    }


    /*
     * Performance periods.
     */
    for (
        const tab of dom.periodTabs || []
    ) {

        tab.addEventListener(
            "click",
            () => {

                setActivePeriod(
                    tab.dataset.period
                );
            }
        );
    }


    /*
     * Market News search.
     */
    dom.marketNewsSearch?.addEventListener(
        "input",
        event => {

            state.marketNewsSearch =
                event.target.value || "";

            renderMarketNews();
        }
    );


    /*
     * Market News filters.
     */
    dom.marketNewsCategory?.addEventListener(
        "change",
        applyMarketNewsFilters
    );


    dom.marketNewsSentiment?.addEventListener(
        "change",
        applyMarketNewsFilters
    );


    dom.marketNewsImportance?.addEventListener(
        "change",
        applyMarketNewsFilters
    );


    /*
     * Market News tabs.
     */
    dom.marketNewsCurrentTab?.addEventListener(
        "click",
        () => {

            setNewsTab(
                "current"
            );
        }
    );


    dom.marketNewsHistoryTab?.addEventListener(
        "click",
        () => {

            setNewsTab(
                "history"
            );
        }
    );
}


/* ============================================================
   HTML ESCAPING
============================================================ */

function escapeHtml(value) {

    return String(value)
        .replaceAll("&", "&amp;")
        .replaceAll("<", "&lt;")
        .replaceAll(">", "&gt;")
        .replaceAll('"', "&quot;")
        .replaceAll("'", "&#039;");
}


function escapeAttribute(value) {

    return escapeHtml(
        value
    );
}


/* ============================================================
   APPLICATION INITIALISATION
============================================================ */

async function initialise() {

    cacheDom();

    applyTheme(
        state.theme
    );

    bindEvents();

    setLoadingStatus();

    clearGlobalError();


    try {

        /*
         * Load production data.
         */
        const data =
            await loadApplicationData();


        state.funds =
            Array.isArray(data.funds)
                ? data.funds
                : [];


        state.bidHistory =
            Array.isArray(data.bidHistory)
                ? data.bidHistory
                : [];


        state.marketNews = {

            analyses:
                Array.isArray(
                    data.marketNews?.analyses
                )
                    ? data.marketNews.analyses
                    : [],

            generatedAtSgt:
                data.marketNews
                    ?.generatedAtSgt ??
                null,

            timezone:
                data.marketNews
                    ?.timezone ??
                "Asia/Singapore",

            timezoneLabel:
                data.marketNews
                    ?.timezoneLabel ??
                "SGT",

            windowDays:
                data.marketNews
                    ?.windowDays ??
                null,

            articleCount:
                data.marketNews
                    ?.articleCount ??
                0
        };


        /*
         * Calculate all four performance periods.
         *
         * Performance is calculated from BID history,
         * not from cumulative performance fields.
         */
        state.performance =
            calculateAllPerformance(
                state.funds,
                state.bidHistory,
                10
            );


        /*
         * Build Market News filter values.
         */
        buildNewsFilterOptions();


        state.loading =
            false;

        state.initialized =
            true;


        setReadyStatus();

        clearGlobalError();


        /*
         * Initial UI.
         */
        renderLatestValuationDate();

        setActivePeriod(
            state.activePeriod
        );

        setActiveView(
            state.activeView
        );

        renderMarketNews();


    } catch (error) {

        state.loading =
            false;

        state.initialized =
            false;


        console.error(
            "VGrat FMS initialisation failed:",
            error
        );


        showGlobalError(
            error
        );


        renderPerformanceEmpty();
    }
}


/* ============================================================
   APPLICATION START
============================================================ */

if (
    document.readyState ===
    "loading"
) {

    document.addEventListener(
        "DOMContentLoaded",
        initialise,
        {
            once: true
        }
    );

} else {

    initialise();
}


/* ============================================================
   DEVELOPMENT / DEBUG ACCESS
============================================================ */

window.VGratFMS = {

    getState() {

        return {
            ...state,

            funds:
                [...state.funds],

            bidHistory:
                [...state.bidHistory]
        };
    },

    setPeriod:
        setActivePeriod,

    setView:
        setActiveView,

    setTheme:
        applyTheme,

    reload:
        initialise
};
