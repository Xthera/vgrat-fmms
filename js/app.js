/* ============================================================
   VGRAT FMS
   APPLICATION CONTROLLER
   ============================================================

   Main responsibilities:
   - Load application data
   - Initialise application modules
   - Control navigation
   - Control theme
   - Render Market Performance
   - Render Market News
   - Connect optional charts
   - Handle application-level errors

   Data calculation remains outside this file.
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
    loadCurrentMarketNews,
    createNewsState,
    getVisibleNews,
    getNewsStatistics,
    getTopItems,
    formatNewsDate,
    formatNewsTime
} from "./market-news.js";

import {
    isChartAvailable,
    createNewsSentimentChart,
    createNewsCategoryChart,
    createNewsImportanceChart,
    createNewsAssetClassChart,
    createNewsGeographyChart,
    createNewsSectorChart,
    updateChartTheme
} from "./charts.js";


/* ============================================================
   APPLICATION CONSTANTS
============================================================ */

const APP_VERSION =
    "1.0.0";


const STORAGE_KEYS = {

    theme:
        "vgrat-fms-theme",

    view:
        "vgrat-fms-view",

    period:
        "vgrat-fms-performance-period"

};


const DEFAULT_THEME =
    "dark";


const DEFAULT_VIEW =
    "performance";


const DEFAULT_PERIOD =
    "DD";


const PERFORMANCE_PERIODS = [
    "DD",
    "WW",
    "MM",
    "YY"
];


/* ============================================================
   APPLICATION STATE
============================================================ */

const state = {

    initialized:
        false,

    loading:
        true,

    error:
        null,

    data:
        null,

    performance:
        null,

    selectedPeriod:
        DEFAULT_PERIOD,

    currentView:
        DEFAULT_VIEW,

    news:
        null,

    newsState:
        null

};


/* ============================================================
   DOM HELPERS
============================================================ */

function qs(
    selector,
    root = document
) {

    return root.querySelector(
        selector
    );
}


function qsa(
    selector,
    root = document
) {

    return [
        ...root.querySelectorAll(
            selector
        )
    ];
}


function firstExisting(
    selectors
) {

    for (
        const selector
        of selectors
    ) {

        const element =
            qs(selector);


        if (element) {
            return element;
        }
    }


    return null;
}


function setText(
    selectors,
    value
) {

    const element =
        Array.isArray(selectors)
            ? firstExisting(
                selectors
            )
            : qs(
                selectors
            );


    if (!element) {
        return;
    }


    element.textContent =
        value ??
        "";
}


function setHTML(
    selectors,
    html
) {

    const element =
        Array.isArray(selectors)
            ? firstExisting(
                selectors
            )
            : qs(
                selectors
            );


    if (!element) {
        return;
    }


    element.innerHTML =
        html;
}


/* ============================================================
   INITIALISATION
============================================================ */

document.addEventListener(
    "DOMContentLoaded",
    init
);


async function init() {

    if (
        state.initialized
    ) {

        return;
    }


    state.initialized =
        true;


    initialiseTheme();

    initialiseNavigation();

    initialisePerformanceControls();

    initialiseNewsControls();

    restoreApplicationState();

    setApplicationLoadingState(
        true
    );


    try {

        const applicationData =
            await loadApplicationData();


        state.data =
            applicationData;


        state.performance =
            calculateAllPerformance(
                applicationData.funds,
                applicationData.bidHistory,
                10
            );


        /*
         * Market News is loaded through its dedicated
         * analysis module. This keeps the application
         * controller independent of the raw JSON schema.
         */
        try {

            state.news =
                await loadCurrentMarketNews();


            state.newsState =
                createNewsState(
                    state.news
                );

        } catch (
            newsError
        ) {

            /*
             * Market Performance should remain usable even
             * if Market News is temporarily unavailable.
             */
            console.error(
                "Market News failed to load:",
                newsError
            );


            state.news =
                null;


            state.newsState =
                null;
        }


        setApplicationLoadingState(
            false
        );


        updateDataStatus(
            true
        );


        renderPerformance();

        renderMarketNews();

        renderNewsCharts();

        renderApplicationMetadata();

        activateView(
            state.currentView
        );

        state.loading =
            false;

    } catch (
        error
    ) {

        console.error(
            "VGrat FMS initialisation failed:",
            error
        );


        state.error =
            error;


        state.loading =
            false;


        setApplicationLoadingState(
            false
        );


        updateDataStatus(
            false
        );


        renderApplicationError(
            error
        );
    }
}


/* ============================================================
   THEME
============================================================ */

function initialiseTheme() {

    let theme =
        DEFAULT_THEME;


    try {

        const saved =
            localStorage.getItem(
                STORAGE_KEYS.theme
            );


        if (
            saved === "light" ||
            saved === "dark"
        ) {

            theme =
                saved;
        }

    } catch (
        error
    ) {

        console.warn(
            "Unable to read saved theme:",
            error
        );
    }


    applyTheme(
        theme
    );


    const toggle =
        firstExisting([
            "#theme-toggle",
            "[data-theme-toggle]"
        ]);


    if (!toggle) {
        return;
    }


    toggle.addEventListener(
        "click",
        () => {

            const current =
                document.documentElement
                    .getAttribute(
                        "data-theme"
                    ) ||
                DEFAULT_THEME;


            const next =
                current === "dark"
                    ? "light"
                    : "dark";


            applyTheme(
                next
            );


            try {

                localStorage.setItem(
                    STORAGE_KEYS.theme,
                    next
                );

            } catch (
                error
            ) {

                console.warn(
                    "Unable to save theme:",
                    error
                );
            }


            updateChartTheme();
        }
    );
}


function applyTheme(
    theme
) {

    const normalized =
        theme === "light"
            ? "light"
            : "dark";


    document.documentElement
        .setAttribute(
            "data-theme",
            normalized
        );


    const toggle =
        firstExisting([
            "#theme-toggle",
            "[data-theme-toggle]"
        ]);


    if (!toggle) {
        return;
    }


    const isLight =
        normalized === "light";


    toggle.setAttribute(
        "aria-pressed",
        String(
            isLight
        )
    );


    toggle.setAttribute(
        "aria-label",
        isLight
            ? "Switch to dark mode"
            : "Switch to light mode"
    );


    const icon =
        qs(
            ".theme-icon",
            toggle
        );


    if (icon) {

        icon.textContent =
            isLight
                ? "☀"
                : "☾";
    }


    const label =
        qs(
            ".theme-label",
            toggle
        );


    if (label) {

        label.textContent =
            isLight
                ? "Light"
                : "Dark";
    }
}


/* ============================================================
   NAVIGATION
============================================================ */

function initialiseNavigation() {

    const links =
        qsa(
            "[data-view-target], [data-view-link]"
        );


    for (
        const link
        of links
    ) {

        link.addEventListener(
            "click",
            event => {

                const target =
                    link.dataset.viewTarget ||
                    link.dataset.viewLink;


                if (!target) {
                    return;
                }


                event.preventDefault();


                activateView(
                    target
                );


                try {

                    localStorage.setItem(
                        STORAGE_KEYS.view,
                        target
                    );

                } catch (
                    error
                ) {

                    console.warn(
                        "Unable to save current view:",
                        error
                    );
                }
            }
        );
    }
}


function activateView(
    requestedView
) {

    const normalized =
        normalizeViewName(
            requestedView
        );


    state.currentView =
        normalized;


    /*
     * Navigation items.
     */
    for (
        const link
        of qsa(
            "[data-view-target], [data-view-link]"
        )
    ) {

        const target =
            normalizeViewName(
                link.dataset.viewTarget ||
                link.dataset.viewLink
            );


        const active =
            target ===
            normalized;


        link.classList.toggle(
            "active",
            active
        );


        link.classList.toggle(
            "is-active",
            active
        );


        link.setAttribute(
            "aria-current",
            active
                ? "page"
                : "false"
        );
    }


    /*
     * View panels.
     *
     * Supports both:
     *
     * data-view-panel="performance"
     *
     * and IDs such as:
     *
     * market-performance-view
     */
    const panels =
        qsa(
            "[data-view-panel]"
        );


    if (
        panels.length > 0
    ) {

        for (
            const panel
            of panels
        ) {

            const panelView =
                normalizeViewName(
                    panel.dataset.viewPanel
                );


            const active =
                panelView ===
                normalized;


            panel.classList.toggle(
                "is-active",
                active
            );


            panel.hidden =
                !active;
        }

    } else {

        /*
         * Fallback for the current index.html structure.
         */
        const candidates =
            getViewPanelCandidates(
                normalized
            );


        for (
            const id
            of candidates
        ) {

            const panel =
                document.getElementById(
                    id
                );


            if (!panel) {
                continue;
            }


            panel.classList.add(
                "is-active"
            );


            panel.hidden =
                false;
        }


        /*
         * Hide known alternative views.
         */
        for (
            const view
            of [
                "performance",
                "market-news",
                "news",
                "fund-explorer",
                "monitoring"
            ]
        ) {

            if (
                normalizeViewName(
                    view
                ) ===
                normalized
            ) {

                continue;
            }


            for (
                const id
                of getViewPanelCandidates(
                    view
                )
            ) {

                const panel =
                    document.getElementById(
                        id
                    );


                if (!panel) {
                    continue;
                }


                panel.classList.remove(
                    "is-active"
                );


                panel.hidden =
                    true;
            }
        }
    }


    updateDocumentTitle(
        normalized
    );
}


function normalizeViewName(
    value
) {

    const normalized =
        String(
            value ??
            ""
        )
            .toLowerCase()
            .trim()
            .replace(
                /_/g,
                "-"
            );


    if (
        normalized ===
        "market-performance" ||
        normalized ===
        "market"
    ) {

        return "performance";
    }


    if (
        normalized ===
        "market-news" ||
        normalized ===
        "news"
    ) {

        return "news";
    }


    if (
        normalized ===
        "fund-explorer" ||
        normalized ===
        "funds"
    ) {

        return "fund-explorer";
    }


    if (
        normalized ===
        "monitoring" ||
        normalized ===
        "monitor"
    ) {

        return "monitoring";
    }


    if (
        normalized ===
        "performance"
    ) {

        return "performance";
    }


    return "performance";
}


function getViewPanelCandidates(
    view
) {

    switch (
        normalizeViewName(
            view
        )
    ) {

        case "performance":

            return [
                "market-performance-view",
                "view-market-performance",
                "performance-view"
            ];


        case "news":

            return [
                "market-news-view",
                "view-market-news",
                "news-view"
            ];


        case "fund-explorer":

            return [
                "fund-explorer-view",
                "view-fund-explorer"
            ];


        case "monitoring":

            return [
                "monitoring-view",
                "view-monitoring"
            ];


        default:

            return [];
    }
}


function updateDocumentTitle(
    view
) {

    const titles = {

        performance:
            "Market Performance",

        news:
            "Market News",

        "fund-explorer":
            "Fund Explorer",

        monitoring:
            "Monitoring"

    };


    const title =
        titles[
            normalizeViewName(
                view
            )
        ] ||
        titles.performance;


    document.title =
        `VGrat FMS — ${title}`;
}


/* ============================================================
   APPLICATION STATE RESTORATION
============================================================ */

function restoreApplicationState() {

    try {

        const savedView =
            localStorage.getItem(
                STORAGE_KEYS.view
            );


        if (savedView) {

            state.currentView =
                normalizeViewName(
                    savedView
                );
        }


        const savedPeriod =
            localStorage.getItem(
                STORAGE_KEYS.period
            );


        if (
            PERFORMANCE_PERIODS.includes(
                savedPeriod
            )
        ) {

            state.selectedPeriod =
                savedPeriod;
        }

    } catch (
        error
    ) {

        console.warn(
            "Unable to restore application state:",
            error
        );
    }
}


/* ============================================================
   PERFORMANCE CONTROLS
============================================================ */

function initialisePerformanceControls() {

    const tabs =
        qsa(
            "[data-period]"
        );


    for (
        const tab
        of tabs
    ) {

        tab.addEventListener(
            "click",
            () => {

                const period =
                    String(
                        tab.dataset.period
                    )
                        .toUpperCase();


                if (
                    !PERFORMANCE_PERIODS.includes(
                        period
                    )
                ) {

                    return;
                }


                state.selectedPeriod =
                    period;


                try {

                    localStorage.setItem(
                        STORAGE_KEYS.period,
                        period
                    );

                } catch (
                    error
                ) {

                    console.warn(
                        "Unable to save selected period:",
                        error
                    );
                }


                updatePerformanceTabs();

                renderPerformance();
            }
        );
    }
}


function updatePerformanceTabs() {

    for (
        const tab
        of qsa(
            "[data-period]"
        )
    ) {

        const period =
            String(
                tab.dataset.period
            )
                .toUpperCase();


        const active =
            period ===
            state.selectedPeriod;


        tab.classList.toggle(
            "active",
            active
        );


        tab.classList.toggle(
            "is-active",
            active
        );


        tab.setAttribute(
            "aria-selected",
            String(
                active
            )
        );
    }
}


/* ============================================================
   PERFORMANCE RENDERING
============================================================ */

function renderPerformance() {

    if (
        !state.performance
    ) {

        return;
    }


    updatePerformanceTabs();


    const periodData =
        state.performance[
            state.selectedPeriod
        ];


    if (!periodData) {

        renderPerformanceError(
            "Performance data is unavailable for this period."
        );


        return;
    }


    /*
     * Page summary.
     */
    setText(
        [
            "#selected-period-label",
            "#performance-period-title",
            "[data-performance-period-title]"
        ],
        periodData.label
    );


    setText(
        [
            "#selected-period-description",
            "#performance-period-description",
            "[data-performance-period-description]"
        ],
        getPeriodDescription(
            periodData
        )
    );


    setText(
        [
            "#performance-eligible-count",
            "#selected-period-count",
            "[data-performance-eligible-count]"
        ],
        `${periodData.totalEligibleFunds} eligible funds`
    );


    setText(
        [
            "#performance-as-of",
            "[data-performance-as-of]"
        ],
        periodData.asOfDate
            ? `Latest BID: ${formatDisplayDate(periodData.asOfDate)}`
            : "Latest BID: —"
    );


    setText(
        [
            "#performance-target-date",
            "[data-performance-target-date]"
        ],
        periodData.targetDate
            ? `Comparison target: ${formatDisplayDate(periodData.targetDate)}`
            : "Comparison target: —"
    );


    /*
     * Latest valuation date.
     */
    setText(
        [
            "#latest-valuation-date",
            "#market-latest-date",
            "[data-latest-valuation-date]"
        ],
        periodData.asOfDate
            ? formatDisplayDate(
                periodData.asOfDate
            )
            : "—"
    );


    /*
     * Ranking tables.
     */
    renderPerformanceTable(
        [
            "#winners-table-body",
            "#performance-winners-body",
            "[data-performance-winners]"
        ],
        periodData.winners,
        "winners"
    );


    renderPerformanceTable(
        [
            "#losers-table-body",
            "#performance-losers-body",
            "[data-performance-losers]"
        ],
        periodData.losers,
        "losers"
    );


    /*
     * Summary cards.
     */
    renderPerformanceSummary(
        periodData
    );
}


function getPeriodDescription(
    periodData
) {

    if (!periodData) {
        return "";
    }


    const descriptions = {

        DD:
            "Change versus the latest actual BID available on or before the previous calendar day.",

        WW:
            "Change versus the latest actual BID available on or before seven calendar days earlier.",

        MM:
            "Change versus the latest actual BID available on or before the same calendar day one month earlier.",

        YY:
            "Change versus the latest actual BID available on or before the same calendar day one year earlier."

    };


    return (
        descriptions[
            periodData.key
        ] ||
        periodData.description ||
        ""
    );
}


function renderPerformanceTable(
    selectors,
    rows,
    type
) {

    const tbody =
        Array.isArray(
            selectors
        )
            ? firstExisting(
                selectors
            )
            : qs(
                selectors
            );


    if (!tbody) {
        return;
    }


    const data =
        Array.isArray(
            rows
        )
            ? rows
            : [];


    if (
        data.length === 0
    ) {

        tbody.innerHTML = `

            <tr>

                <td
                    colspan="5"
                    class="empty-state"
                >
                    No eligible funds are available
                    for this period.
                </td>

            </tr>

        `;


        return;
    }


    tbody.innerHTML =
        data
            .map(
                row =>
                    renderPerformanceRow(
                        row,
                        type
                    )
            )
            .join("");
}


function renderPerformanceRow(
    row,
    type
) {

    const returnValue =
        Number(
            row.returnPercent
        );


    const returnClass =
        returnValue > 0
            ? "return-positive"
            : returnValue < 0
                ? "return-negative"
                : "return-neutral";


    const returnText =
        formatReturnPercent(
            returnValue
        );


    const fundName =
        escapeHTML(
            row.fundName ||
            "Unnamed fund"
        );


    const fundCode =
        escapeHTML(
            row.fundCode ||
            "—"
        );


    const identifier =
        escapeHTML(
            row.fundIdentifier ||
            ""
        );


    const bidText =
        formatBid(
            row.currentBid
        );


    const comparisonText =
        row.comparisonDate
            ? formatDisplayDate(
                row.comparisonDate
            )
            : "—";


    const currentDateText =
        row.currentDate
            ? formatDisplayDate(
                row.currentDate
            )
            : "—";


    /*
     * The fund name is intentionally rendered as a button-like
     * element for now. Fund detail navigation will be connected
     * once the Fund Explorer/detail view is built.
     */
    return `

        <tr
            data-fund-identifier="${identifier}"
            data-performance-type="${escapeHTML(type)}"
        >

            <td class="rank-cell">
                ${Number(row.rank) || "—"}
            </td>


            <td>

                <div class="fund-cell">

                    <div class="fund-name">
                        ${fundName}
                    </div>

                    <div class="fund-meta">
                        ${fundCode}
                    </div>

                </div>

            </td>


            <td class="fund-code-cell">
                ${fundCode}
            </td>


            <td class="${returnClass}">
                ${returnText}
            </td>


            <td>

                <div class="bid-cell">

                    <span class="bid-value">
                        ${escapeHTML(bidText)}
                    </span>

                    <span class="bid-date">
                        ${escapeHTML(currentDateText)}
                    </span>

                </div>

            </td>

        </tr>

    `;
}


function renderPerformanceSummary(
    periodData
) {

    const winners =
        Array.isArray(
            periodData.winners
        )
            ? periodData.winners
            : [];


    const losers =
        Array.isArray(
            periodData.losers
        )
            ? periodData.losers
            : [];


    const strongest =
        winners.length > 0
            ? winners[0]
            : null;


    const weakest =
        losers.length > 0
            ? losers[0]
            : null;


    setText(
        [
            "#top-winner-return",
            "[data-top-winner-return]"
        ],
        strongest
            ? formatReturnPercent(
                strongest.returnPercent
            )
            : "—"
    );


    setText(
        [
            "#top-winner-name",
            "[data-top-winner-name]"
        ],
        strongest?.fundName ||
        "—"
    );


    setText(
        [
            "#top-loser-return",
            "[data-top-loser-return]"
        ],
        weakest
            ? formatReturnPercent(
                weakest.returnPercent
            )
            : "—"
    );


    setText(
        [
            "#top-loser-name",
            "[data-top-loser-name]"
        ],
        weakest?.fundName ||
        "—"
    );


    setText(
        [
            "#eligible-fund-count",
            "[data-eligible-fund-count]"
        ],
        String(
            periodData.totalEligibleFunds ??
            0
        )
    );
}


function renderPerformanceError(
    message
) {

    const containers = [
        firstExisting([
            "#winners-table-body",
            "#performance-winners-body"
        ]),
        firstExisting([
            "#losers-table-body",
            "#performance-losers-body"
        ])
    ];


    for (
        const container
        of containers
    ) {

        if (!container) {
            continue;
        }


        container.innerHTML = `

            <tr>

                <td
                    colspan="5"
                    class="error-state"
                >
                    ${escapeHTML(message)}
                </td>

            </tr>

        `;
    }
}


/* ============================================================
   MARKET NEWS CONTROLS
============================================================ */

function initialiseNewsControls() {

    /*
     * Current / History tabs.
     */
    for (
        const tab
        of qsa(
            "[data-news-tab]"
        )
    ) {

        tab.addEventListener(
            "click",
            () => {

                const target =
                    tab.dataset.newsTab;


                switchNewsTab(
                    target
                );
            }
        );
    }


    /*
     * Search.
     */
    const search =
        firstExisting([
            "#market-news-search",
            "[data-market-news-search]"
        ]);


    if (search) {

        search.addEventListener(
            "input",
            () => {

                if (!state.newsState) {
                    return;
                }


                state.newsState.searchTerm =
                    search.value;


                renderMarketNews();
            }
        );
    }


    /*
     * Category filter.
     */
    const category =
        firstExisting([
            "#market-news-category",
            "[data-market-news-category]"
        ]);


    if (category) {

        category.addEventListener(
            "change",
            () => {

                if (!state.newsState) {
                    return;
                }


                state.newsState.selectedCategory =
                    category.value;


                renderMarketNews();
            }
        );
    }


    /*
     * Importance filter.
     */
    const importance =
        firstExisting([
            "#market-news-importance",
            "[data-market-news-importance]"
        ]);


    if (importance) {

        importance.addEventListener(
            "change",
            () => {

                if (!state.newsState) {
                    return;
                }


                state.newsState.selectedImportance =
                    importance.value;


                renderMarketNews();
            }
        );
    }


    /*
     * Sentiment filter.
     */
    const sentiment =
        firstExisting([
            "#market-news-sentiment",
            "[data-market-news-sentiment]"
        ]);


    if (sentiment) {

        sentiment.addEventListener(
            "change",
            () => {

                if (!state.newsState) {
                    return;
                }


                state.newsState.selectedSentiment =
                    sentiment.value;


                renderMarketNews();
            }
        );
    }
}


function switchNewsTab(
    target
) {

    const normalized =
        String(
            target ||
            "current"
        )
            .toLowerCase();


    for (
        const tab
        of qsa(
            "[data-news-tab]"
        )
    ) {

        const active =
            String(
                tab.dataset.newsTab
            )
                .toLowerCase() ===
            normalized;


        tab.classList.toggle(
            "active",
            active
        );


        tab.classList.toggle(
            "is-active",
            active
        );


        tab.setAttribute(
            "aria-selected",
            String(
                active
            )
        );
    }


    /*
     * Historical Market News UI will be connected to the
     * date archive once the history selector is added.
     */
    const current =
        firstExisting([
            "#market-news-current",
            "[data-market-news-current]"
        ]);


    const history =
        firstExisting([
            "#market-news-history",
            "[data-market-news-history]"
        ]);


    if (current) {

        current.hidden =
            normalized !==
            "current";
    }


    if (history) {

        history.hidden =
            normalized !==
            "history";
    }
}


/* ============================================================
   MARKET NEWS RENDERING
============================================================ */

function renderMarketNews() {

    const container =
        firstExisting([
            "#market-news-current",
            "[data-market-news-current]"
        ]);


    if (!container) {
        return;
    }


    if (
        !state.newsState
    ) {

        container.innerHTML = `

            <div class="empty-state">

                <strong>
                    Market News unavailable
                </strong>

                <span>
                    The latest AI analysis could not be loaded.
                </span>

            </div>

        `;


        return;
    }


    const visible =
        getVisibleNews(
            state.newsState
        );


    renderNewsMetadata();

    renderNewsFilters();


    if (
        visible.length === 0
    ) {

        container.innerHTML = `

            <div class="empty-state">

                <strong>
                    No matching articles
                </strong>

                <span>
                    Try changing the selected filters.
                </span>

            </div>

        `;


        return;
    }


    container.innerHTML =
        visible
            .map(
                renderNewsArticle
            )
            .join("");
}


function renderNewsMetadata() {

    if (
        !state.newsState
    ) {

        return;
    }


    setText(
        [
            "#market-news-count",
            "[data-market-news-count]"
        ],
        `${state.newsState.analyses.length} analyzed articles`
    );


    setText(
        [
            "#market-news-generated",
            "[data-market-news-generated]"
        ],
        state.newsState.generatedAtSgt
            ? `Updated ${formatNewsDateTimeSafe(
                state.newsState.generatedAtSgt
            )}`
            : "Update time unavailable"
    );


    setText(
        [
            "#market-news-window",
            "[data-market-news-window]"
        ],
        `${state.newsState.windowDays || 14}-day rolling window`
    );
}


function renderNewsFilters() {

    if (
        !state.newsState
    ) {

        return;
    }


    const category =
        firstExisting([
            "#market-news-category",
            "[data-market-news-category]"
        ]);


    if (
        category &&
        category.options.length === 0
    ) {

        const categories =
            [
                "ALL",
                ...new Set(
                    state.newsState.analyses
                        .map(
                            article =>
                                article.category
                        )
                        .filter(
                            Boolean
                        )
                )
            ];


        category.innerHTML =
            categories
                .map(
                    value => `

                        <option value="${escapeHTML(value)}">
                            ${escapeHTML(
                                value === "ALL"
                                    ? "All categories"
                                    : value
                            )}
                        </option>

                    `
                )
                .join("");
    }


    if (category) {

        category.value =
            state.newsState.selectedCategory;
    }


    const importance =
        firstExisting([
            "#market-news-importance",
            "[data-market-news-importance]"
        ]);


    if (importance) {

        importance.value =
            state.newsState.selectedImportance;
    }


    const sentiment =
        firstExisting([
            "#market-news-sentiment",
            "[data-market-news-sentiment]"
        ]);


    if (sentiment) {

        sentiment.value =
            state.newsState.selectedSentiment;
    }


    const search =
        firstExisting([
            "#market-news-search",
            "[data-market-news-search]"
        ]);


    if (
        search &&
        search.value !==
        state.newsState.searchTerm
    ) {

        search.value =
            state.newsState.searchTerm;
    }
}


function renderNewsArticle(
    article
) {

    const sentimentClass =
        getNewsStateClass(
            article.sentiment
        );


    const importanceClass =
        getNewsStateClass(
            article.importance
        );


    const categoryClass =
        getNewsStateClass(
            article.category
        );


    const tags = [
        ...article.assetClasses,
        ...article.geographies,
        ...article.sectors
    ]
        .filter(
            Boolean
        )
        .slice(
            0,
            8
        );


    const tagsHTML =
        tags.length > 0
            ? `

                <div class="news-tags">

                    ${tags
                        .map(
                            tag => `

                                <span class="news-tag">
                                    ${escapeHTML(tag)}
                                </span>

                            `
                        )
                        .join("")}

                </div>

            `
            : "";


    const sourceHTML =
        article.url
            ? `

                <a
                    class="news-source-link"
                    href="${escapeAttribute(article.url)}"
                    target="_blank"
                    rel="noopener noreferrer"
                >
                    ${escapeHTML(article.source || "Source")}
                </a>

            `
            : escapeHTML(
                article.source ||
                "Source"
            );


    const impactHTML =
        article.investorImpact
            ? `

                <div class="news-impact">

                    <div class="news-impact-label">
                        Investor impact
                    </div>

                    <div class="news-impact-text">
                        ${escapeHTML(
                            article.investorImpact
                        )}
                    </div>

                </div>

            `
            : "";


    const monitoringHTML =
        article.fundMonitoringRelevant
            ? `

                <span class="news-monitoring-badge">
                    Fund monitoring relevant
                </span>

            `
            : "";


    return `

        <article
            class="news-card"
            data-article-id="${escapeAttribute(
                article.articleId
            )}"
        >

            <div class="news-card-header">

                <div class="news-card-badges">

                    <span
                        class="news-badge news-category ${categoryClass}"
                    >
                        ${escapeHTML(
                            article.category
                        )}
                    </span>


                    <span
                        class="news-badge news-importance ${importanceClass}"
                    >
                        ${escapeHTML(
                            article.importance
                        )}
                    </span>


                    <span
                        class="news-badge news-sentiment ${sentimentClass}"
                    >
                        ${escapeHTML(
                            article.sentiment
                        )}
                    </span>

                    ${monitoringHTML}

                </div>


                <div class="news-card-date">

                    ${escapeHTML(
                        formatNewsDate(
                            article.publishedAtSgt
                        )
                    )}

                    <span>
                        ${escapeHTML(
                            formatNewsTime(
                                article.publishedAtSgt
                            )
                        )}
                    </span>

                </div>

            </div>


            <h3 class="news-title">

                ${
                    article.url
                        ? `

                            <a
                                href="${escapeAttribute(article.url)}"
                                target="_blank"
                                rel="noopener noreferrer"
                            >
                                ${escapeHTML(
                                    article.title
                                )}
                            </a>

                        `
                        : escapeHTML(
                            article.title
                        )
                }

            </h3>


            ${
                article.summary
                    ? `

                        <p class="news-summary">
                            ${escapeHTML(
                                article.summary
                            )}
                        </p>

                    `
                    : ""
            }


            ${impactHTML}


            ${tagsHTML}


            <div class="news-card-footer">

                <span class="news-source">
                    ${sourceHTML}
                </span>

            </div>

        </article>

    `;
}


/* ============================================================
   MARKET NEWS CHARTS
============================================================ */

function renderNewsCharts() {

    if (
        !state.newsState ||
        !isChartAvailable()
    ) {

        return;
    }


    const statistics =
        getNewsStatistics(
            state.newsState.analyses
        );


    /*
     * Only create charts when the corresponding canvas
     * actually exists in index.html.
     *
     * This keeps the module compatible with the initial
     * Market Performance-first layout.
     */

    const sentimentCanvas =
        firstExisting([
            "#news-sentiment-chart",
            "[data-news-sentiment-chart]"
        ]);


    if (sentimentCanvas) {

        createNewsSentimentChart(
            sentimentCanvas,
            statistics
        );
    }


    const categoryCanvas =
        firstExisting([
            "#news-category-chart",
            "[data-news-category-chart]"
        ]);


    if (categoryCanvas) {

        createNewsCategoryChart(
            categoryCanvas,
            statistics.categories
        );
    }


    const importanceCanvas =
        firstExisting([
            "#news-importance-chart",
            "[data-news-importance-chart]"
        ]);


    if (importanceCanvas) {

        createNewsImportanceChart(
            importanceCanvas,
            statistics
        );
    }


    const assetClassCanvas =
        firstExisting([
            "#news-asset-class-chart",
            "[data-news-asset-class-chart]"
        ]);


    if (assetClassCanvas) {

        createNewsAssetClassChart(
            assetClassCanvas,
            statistics
        );
    }


    const geographyCanvas =
        firstExisting([
            "#news-geography-chart",
            "[data-news-geography-chart]"
        ]);


    if (geographyCanvas) {

        createNewsGeographyChart(
            geographyCanvas,
            statistics
        );
    }


    const sectorCanvas =
        firstExisting([
            "#news-sector-chart",
            "[data-news-sector-chart]"
        ]);


    if (sectorCanvas) {

        createNewsSectorChart(
            sectorCanvas,
            statistics
        );
    }


    /*
     * Optional top-topic text summaries.
     */
    renderNewsTopTopics(
        statistics
    );
}


function renderNewsTopTopics(
    statistics
) {

    const mappings = [

        {
            selectors: [
                "#top-news-asset-classes",
                "[data-top-news-asset-classes]"
            ],
            values:
                getTopItems(
                    statistics.assetClasses,
                    5
                )
        },

        {
            selectors: [
                "#top-news-geographies",
                "[data-top-news-geographies]"
            ],
            values:
                getTopItems(
                    statistics.geographies,
                    5
                )
        },

        {
            selectors: [
                "#top-news-sectors",
                "[data-top-news-sectors]"
            ],
            values:
                getTopItems(
                    statistics.sectors,
                    5
                )
        }

    ];


    for (
        const mapping
        of mappings
    ) {

        const container =
            firstExisting(
                mapping.selectors
            );


        if (!container) {
            continue;
        }


        container.innerHTML =
            mapping.values.length > 0
                ? mapping.values
                    .map(
                        item => `

                            <span class="topic-item">

                                <span class="topic-name">
                                    ${escapeHTML(
                                        item.name
                                    )}
                                </span>

                                <span class="topic-count">
                                    ${item.count}
                                </span>

                            </span>

                        `
                    )
                    .join("")
                : `

                    <span class="empty-state">
                        No data
                    </span>

                `;
    }
}


/* ============================================================
   APPLICATION STATUS
============================================================ */

function setApplicationLoadingState(
    loading
) {

    const body =
        document.body;


    if (body) {

        body.classList.toggle(
            "is-loading",
            loading
        );
    }


    const status =
        firstExisting([
            "#app-status",
            "#data-status",
            "[data-app-status]"
        ]);


    if (!status) {
        return;
    }


    status.classList.toggle(
        "is-loading",
        loading
    );


    if (loading) {

        status.setAttribute(
            "aria-busy",
            "true"
        );

    } else {

        status.setAttribute(
            "aria-busy",
            "false"
        );
    }


    const text =
        firstExistingWithin(
            status,
            [
                "#app-status-text",
                ".status-text",
                "[data-status-text]"
            ]
        );


    if (text) {

        text.textContent =
            loading
                ? "Loading data…"
                : text.textContent;
    }
}


function updateDataStatus(
    success
) {

    const status =
        firstExisting([
            "#app-status",
            "#data-status",
            "[data-app-status]"
        ]);


    if (!status) {
        return;
    }


    status.classList.remove(
        "is-loading",
        "is-error",
        "is-success"
    );


    status.classList.add(
        success
            ? "is-success"
            : "is-error"
    );


    const text =
        firstExistingWithin(
            status,
            [
                "#app-status-text",
                ".status-text",
                "[data-status-text]"
            ]
        );


    if (text) {

        text.textContent =
            success
                ? "Data loaded"
                : "Data error";
    }


    status.setAttribute(
        "title",
        success
            ? "VGrat FMS data loaded successfully."
            : "VGrat FMS data failed to load."
    );
}


/* ============================================================
   APPLICATION METADATA
============================================================ */

function renderApplicationMetadata() {

    setText(
        [
            "#app-version",
            "[data-app-version]"
        ],
        `v${APP_VERSION}`
    );


    if (
        state.performance?.DD?.asOfDate
    ) {

        setText(
            [
                "#latest-data-date",
                "[data-latest-data-date]"
            ],
            formatDisplayDate(
                state.performance.DD.asOfDate
            )
        );
    }
}


/* ============================================================
   APPLICATION ERROR
============================================================ */

function renderApplicationError(
    error
) {

    const message =
        error?.message ||
        "Unable to load VGrat FMS data.";


    const containers = [

        firstExisting([
            "#application-error",
            "[data-application-error]"
        ]),

        firstExisting([
            "#winners-table-body",
            "#performance-winners-body"
        ]),

        firstExisting([
            "#losers-table-body",
            "#performance-losers-body"
        ])

    ].filter(
        Boolean
    );


    if (
        containers.length === 0
    ) {

        return;
    }


    const safeMessage =
        escapeHTML(
            message
        );


    const errorHTML = `

        <div class="error-state">

            <strong>
                Unable to load dashboard data
            </strong>

            <span>
                ${safeMessage}
            </span>

            <button
                type="button"
                class="secondary-button"
                data-retry
            >
                Retry
            </button>

        </div>

    `;


    /*
     * For an explicit application-error container,
     * use the complete block.
     */
    const primary =
        containers[0];


    if (
        primary &&
        (
            primary.id ===
            "application-error" ||
            primary.matches(
                "[data-application-error]"
            )
        )
    ) {

        primary.innerHTML =
            errorHTML;


        attachRetryButton(
            primary
        );


        return;
    }


    /*
     * Otherwise put a compact message into the
     * performance tables.
     */
    for (
        const container
        of containers
    ) {

        container.innerHTML = `

            <tr>

                <td
                    colspan="5"
                    class="error-state"
                >
                    ${safeMessage}
                </td>

            </tr>

        `;
    }
}


function attachRetryButton(
    container
) {

    const button =
        qs(
            "[data-retry]",
            container
        );


    if (!button) {
        return;
    }


    button.addEventListener(
        "click",
        () => {

            window.location.reload();
        }
    );
}


/* ============================================================
   DOM HELPER
============================================================ */

function firstExistingWithin(
    root,
    selectors
) {

    for (
        const selector
        of selectors
    ) {

        const element =
            root.querySelector(
                selector
            );


        if (element) {
            return element;
        }
    }


    return null;
}


/* ============================================================
   FORMATTING HELPERS
============================================================ */

function formatNewsDateTimeSafe(
    value
) {

    try {

        const date =
            new Date(
                value
            );


        if (
            Number.isNaN(
                date.getTime()
            )
        ) {

            return "—";
        }


        return date.toLocaleString(
            "en-SG",
            {
                day: "2-digit",
                month: "short",
                year: "numeric",
                hour: "2-digit",
                minute: "2-digit",
                timeZone: "Asia/Singapore"
            }
        );

    } catch (
        error
    ) {

        return "—";
    }
}


function getNewsStateClass(
    value
) {

    return String(
        value ??
        ""
    )
        .toLowerCase()
        .replace(
            /[^a-z0-9]+/g,
            "-"
        )
        .replace(
            /^-+|-+$/g,
            ""
        );
}


/* ============================================================
   HTML SAFETY
============================================================ */

function escapeHTML(
    value
) {

    return String(
        value ??
        ""
    )
        .replace(
            /&/g,
            "&amp;"
        )
        .replace(
            /</g,
            "&lt;"
        )
        .replace(
            />/g,
            "&gt;"
        )
        .replace(
            /"/g,
            "&quot;"
        )
        .replace(
            /'/g,
            "&#039;"
        );
}


function escapeAttribute(
    value
) {

    return escapeHTML(
        value
    );
}


/* ============================================================
   PUBLIC DEBUG API
============================================================ */

/*
 * Expose a very small read-only diagnostic interface.
 *
 * Useful during GitHub Pages testing:
 *
 * window.VGratFMS.getState()
 */
window.VGratFMS = {

    getState() {

        return {

            initialized:
                state.initialized,

            loading:
                state.loading,

            error:
                state.error
                    ? state.error.message
                    : null,

            selectedPeriod:
                state.selectedPeriod,

            currentView:
                state.currentView,

            fundCount:
                Array.isArray(
                    state.data?.funds
                )
                    ? state.data.funds.length
                    : 0,

            bidHistoryCount:
                Array.isArray(
                    state.data?.bidHistory
                )
                    ? state.data.bidHistory.length
                    : 0,

            newsCount:
                Array.isArray(
                    state.news?.analyses
                )
                    ? state.news.analyses.length
                    : 0

        };
    }

};
