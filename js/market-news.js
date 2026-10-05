/* ============================================================
   VGRAT FMS - MARKET NEWS
   ============================================================

   Visible Market News source:

       data/market_news/analysis/current.json

   Historical source:

       data/market_news/analysis/history/YYYY/MM/YYYY-MM-DD.json

   The frontend uses the AI-analysis output only.

   Raw collector files are intentionally NOT used here.
============================================================ */


/* ============================================================
   PATHS
============================================================ */

const MARKET_NEWS_PATHS = {
    current:
        "data/market_news/analysis/current.json",

    historyBase:
        "data/market_news/analysis/history"
};


/* ============================================================
   BASIC HELPERS
============================================================ */

function isObject(value) {
    return (
        value !== null &&
        typeof value === "object" &&
        !Array.isArray(value)
    );
}


function normaliseString(value) {
    if (value === null || value === undefined) {
        return "";
    }

    return String(value).trim();
}


function normaliseArray(value) {
    if (!Array.isArray(value)) {
        return [];
    }

    return value
        .map(item => normaliseString(item))
        .filter(Boolean);
}


/* ============================================================
   DATE HELPERS
============================================================ */

/**
 * Parse an SGT ISO timestamp.
 *
 * Example:
 * 2026-10-05T04:28:47+08:00
 */
function parseNewsDate(value) {
    if (!value) {
        return null;
    }

    const date = new Date(value);

    if (Number.isNaN(date.getTime())) {
        return null;
    }

    return date;
}


/**
 * Format a timestamp in Singapore time.
 */
function formatNewsDate(
    value,
    options = {}
) {
    const date = parseNewsDate(value);

    if (!date) {
        return "-";
    }

    const defaultOptions = {
        day: "2-digit",
        month: "short",
        year: "numeric",
        hour: "2-digit",
        minute: "2-digit",
        hour12: false,
        timeZone: "Asia/Singapore"
    };

    return new Intl.DateTimeFormat(
        "en-SG",
        {
            ...defaultOptions,
            ...options
        }
    ).format(date);
}


/**
 * Format a date as YYYY-MM-DD in Singapore time.
 */
function formatSingaporeDate(value) {
    const date = parseNewsDate(value);

    if (!date) {
        return null;
    }

    const parts = new Intl.DateTimeFormat(
        "en-CA",
        {
            year: "numeric",
            month: "2-digit",
            day: "2-digit",
            timeZone: "Asia/Singapore"
        }
    ).formatToParts(date);

    const map = {};

    for (const part of parts) {
        if (part.type !== "literal") {
            map[part.type] = part.value;
        }
    }

    if (!map.year || !map.month || !map.day) {
        return null;
    }

    return `${map.year}-${map.month}-${map.day}`;
}


/**
 * Extract YYYY/MM from a YYYY-MM-DD date.
 */
function getHistoryPathParts(dateString) {
    if (
        typeof dateString !== "string" ||
        !/^\d{4}-\d{2}-\d{2}$/.test(dateString)
    ) {
        return null;
    }

    const [year, month] = dateString.split("-");

    return {
        year,
        month
    };
}


/* ============================================================
   FETCH
============================================================ */

async function fetchJson(path) {
    const response = await fetch(
        path,
        {
            cache: "no-store"
        }
    );

    if (!response.ok) {
        throw new Error(
            `Unable to load ${path} (HTTP ${response.status})`
        );
    }

    try {
        return await response.json();
    } catch {
        throw new Error(
            `Invalid JSON returned from ${path}`
        );
    }
}


/* ============================================================
   ARTICLE NORMALISATION
============================================================ */

/**
 * Convert one raw analysis article into the stable structure
 * consumed by the UI.
 *
 * Hidden/internal fields such as reasoning and collector
 * metadata are retained internally but are not required by
 * the visible article renderer.
 */
function normaliseArticle(article) {
    if (!isObject(article)) {
        return null;
    }

    const articleId =
        normaliseString(article.articleId);

    const title =
        normaliseString(article.title);

    const summary =
        normaliseString(article.summary);

    /*
     * A visible article needs at minimum an ID and title.
     */
    if (!articleId || !title) {
        return null;
    }

    return {
        articleId,

        relevant:
            article.relevant !== false,

        category:
            normaliseString(article.category)
                .toUpperCase() || "MARKET",

        sentiment:
            normaliseString(article.sentiment)
                .toUpperCase() || "NEUTRAL",

        importance:
            normaliseString(article.importance)
                .toUpperCase() || "MEDIUM",

        summary,

        assetClasses:
            normaliseArray(article.assetClasses),

        geographies:
            normaliseArray(article.geographies),

        sectors:
            normaliseArray(article.sectors),

        investorImpact:
            normaliseString(article.investorImpact),

        fundMonitoringRelevant:
            article.fundMonitoringRelevant === true,

        source:
            normaliseString(article.source) || "CNBC",

        title,

        publishedAtSgt:
            normaliseString(article.publishedAtSgt),

        url:
            normaliseString(article.url),

        collectorCategory:
            normaliseString(article.collectorCategory),

        collectorRelevanceScore:
            Number.isFinite(
                Number(article.collectorRelevanceScore)
            )
                ? Number(article.collectorRelevanceScore)
                : null,

        collectorRelevanceReason:
            normaliseString(
                article.collectorRelevanceReason
            ),

        collectedAtSgt:
            normaliseString(article.collectedAtSgt),

        reasoning:
            normaliseString(article.reasoning)
    };
}


/* ============================================================
   ANALYSIS FILE NORMALISATION
============================================================ */

function normaliseAnalysisFile(data) {
    if (!isObject(data)) {
        throw new Error(
            "Market news analysis returned an invalid structure."
        );
    }

    if (!Array.isArray(data.analyses)) {
        throw new Error(
            "Market news analysis does not contain an analyses array."
        );
    }

    const analyses = data.analyses
        .map(normaliseArticle)
        .filter(Boolean);

    return {
        generatedAtSgt:
            normaliseString(data.generatedAtSgt) || null,

        timezone:
            normaliseString(data.timezone) || "Asia/Singapore",

        timezoneLabel:
            normaliseString(data.timezoneLabel) || "SGT",

        windowDays:
            Number.isFinite(Number(data.windowDays))
                ? Number(data.windowDays)
                : null,

        articleCount:
            Number.isFinite(Number(data.articleCount))
                ? Number(data.articleCount)
                : analyses.length,

        analyses
    };
}


/* ============================================================
   CURRENT NEWS
============================================================ */

/**
 * Load current analyzed Market News.
 */
async function loadCurrentMarketNews() {
    const raw =
        await fetchJson(
            MARKET_NEWS_PATHS.current
        );

    return normaliseAnalysisFile(raw);
}


/* ============================================================
   HISTORICAL NEWS
============================================================ */

/**
 * Build the path for a historical analysis file.
 *
 * Example:
 *
 * 2026-10-03
 *
 * becomes:
 *
 * data/market_news/analysis/history/2026/10/2026-10-03.json
 */
function buildHistoryPath(dateString) {
    const parts =
        getHistoryPathParts(dateString);

    if (!parts) {
        throw new Error(
            `Invalid history date: ${dateString}`
        );
    }

    return [
        MARKET_NEWS_PATHS.historyBase,
        parts.year,
        parts.month,
        `${dateString}.json`
    ].join("/");
}


/**
 * Load one historical Market News file.
 */
async function loadHistoricalMarketNews(
    dateString
) {
    const path =
        buildHistoryPath(dateString);

    const raw =
        await fetchJson(path);

    return {
        date: dateString,
        ...normaliseAnalysisFile(raw)
    };
}


/* ============================================================
   HISTORY AVAILABILITY
============================================================ */

/**
 * Test whether a historical date exists.
 *
 * Returns:
 *   true
 *   false
 */
async function historyExists(dateString) {
    try {
        await loadHistoricalMarketNews(
            dateString
        );

        return true;
    } catch {
        return false;
    }
}


/**
 * Generate dates between two YYYY-MM-DD values.
 *
 * Used when checking the rolling history window.
 */
function generateDateRange(
    startDateString,
    endDateString
) {
    const start =
        new Date(
            `${startDateString}T00:00:00Z`
        );

    const end =
        new Date(
            `${endDateString}T00:00:00Z`
        );

    if (
        Number.isNaN(start.getTime()) ||
        Number.isNaN(end.getTime()) ||
        start > end
    ) {
        return [];
    }

    const dates = [];

    const current =
        new Date(start.getTime());

    while (current <= end) {
        dates.push(
            current.toISOString().slice(0, 10)
        );

        current.setUTCDate(
            current.getUTCDate() + 1
        );
    }

    return dates;
}


/**
 * Find available historical dates inside a date range.
 *
 * This intentionally checks the filesystem structure through
 * HTTP requests rather than assuming every calendar date exists.
 */
async function findAvailableHistoryDates(
    startDateString,
    endDateString
) {
    const candidateDates =
        generateDateRange(
            startDateString,
            endDateString
        );

    if (candidateDates.length === 0) {
        return [];
    }

    const checks =
        await Promise.all(
            candidateDates.map(
                async dateString => ({
                    dateString,
                    exists:
                        await historyExists(
                            dateString
                        )
                })
            )
        );

    return checks
        .filter(item => item.exists)
        .map(item => item.dateString);
}


/* ============================================================
   HISTORY DATE EXTRACTION
============================================================ */

/**
 * Convert an article timestamp to its Singapore calendar date.
 */
function getArticleSingaporeDate(article) {
    return formatSingaporeDate(
        article?.publishedAtSgt
    );
}


/**
 * Group articles by Singapore publication date.
 */
function groupArticlesByDate(articles) {
    const groups = new Map();

    if (!Array.isArray(articles)) {
        return groups;
    }

    for (const article of articles) {
        const date =
            getArticleSingaporeDate(article);

        if (!date) {
            continue;
        }

        if (!groups.has(date)) {
            groups.set(date, []);
        }

        groups.get(date).push(article);
    }

    return groups;
}


/* ============================================================
   ARTICLE SORTING
============================================================ */

/**
 * Sort newest published articles first.
 */
function sortArticlesNewestFirst(
    articles
) {
    if (!Array.isArray(articles)) {
        return [];
    }

    return [...articles].sort(
        (a, b) => {
            const aTime =
                parseNewsDate(
                    a?.publishedAtSgt
                )?.getTime() ?? 0;

            const bTime =
                parseNewsDate(
                    b?.publishedAtSgt
                )?.getTime() ?? 0;

            return bTime - aTime;
        }
    );
}


/**
 * Sort oldest published articles first.
 */
function sortArticlesOldestFirst(
    articles
) {
    return sortArticlesNewestFirst(
        articles
    ).reverse();
}


/* ============================================================
   DEDUPLICATION
============================================================ */

/**
 * Remove duplicate articles by articleId.
 *
 * This protects the UI if the rolling files happen to overlap.
 */
function deduplicateArticles(
    articles
) {
    if (!Array.isArray(articles)) {
        return [];
    }

    const seen = new Set();
    const result = [];

    for (const article of articles) {
        const id =
            normaliseString(
                article?.articleId
            );

        if (!id) {
            continue;
        }

        if (seen.has(id)) {
            continue;
        }

        seen.add(id);
        result.push(article);
    }

    return result;
}


/* ============================================================
   FILTERING
============================================================ */

/**
 * Apply Market News filters.
 *
 * Supported filters:
 *
 * category
 * sentiment
 * importance
 * assetClass
 * geography
 * sector
 * fundMonitoringRelevant
 */
function filterArticles(
    articles,
    filters = {}
) {
    if (!Array.isArray(articles)) {
        return [];
    }

    const category =
        normaliseString(
            filters.category
        ).toUpperCase();

    const sentiment =
        normaliseString(
            filters.sentiment
        ).toUpperCase();

    const importance =
        normaliseString(
            filters.importance
        ).toUpperCase();

    const assetClass =
        normaliseString(
            filters.assetClass
        );

    const geography =
        normaliseString(
            filters.geography
        );

    const sector =
        normaliseString(
            filters.sector
        );

    const fundMonitoringOnly =
        filters.fundMonitoringRelevant === true;

    return articles.filter(
        article => {
            if (
                category &&
                article.category !== category
            ) {
                return false;
            }

            if (
                sentiment &&
                article.sentiment !== sentiment
            ) {
                return false;
            }

            if (
                importance &&
                article.importance !== importance
            ) {
                return false;
            }

            if (
                assetClass &&
                !article.assetClasses.includes(
                    assetClass
                )
            ) {
                return false;
            }

            if (
                geography &&
                !article.geographies.includes(
                    geography
                )
            ) {
                return false;
            }

            if (
                sector &&
                !article.sectors.includes(
                    sector
                )
            ) {
                return false;
            }

            if (
                fundMonitoringOnly &&
                article.fundMonitoringRelevant !== true
            ) {
                return false;
            }

            return true;
        }
    );
}


/* ============================================================
   SEARCH
============================================================ */

/**
 * Search visible article content.
 */
function searchArticles(
    articles,
    searchTerm
) {
    const query =
        normaliseString(
            searchTerm
        ).toLowerCase();

    if (!query) {
        return Array.isArray(articles)
            ? [...articles]
            : [];
    }

    if (!Array.isArray(articles)) {
        return [];
    }

    return articles.filter(
        article => {
            const searchableText = [
                article.title,
                article.summary,
                article.investorImpact,
                article.source,
                article.category,
                article.sentiment,
                article.importance,
                ...article.assetClasses,
                ...article.geographies,
                ...article.sectors
            ]
                .filter(Boolean)
                .join(" ")
                .toLowerCase();

            return searchableText.includes(
                query
            );
        }
    );
}


/* ============================================================
   FILTER OPTIONS
============================================================ */

/**
 * Build unique filter values from available articles.
 */
function getNewsFilterOptions(
    articles
) {
    const categories = new Set();
    const sentiments = new Set();
    const importance = new Set();
    const assetClasses = new Set();
    const geographies = new Set();
    const sectors = new Set();

    if (Array.isArray(articles)) {
        for (const article of articles) {
            if (article.category) {
                categories.add(
                    article.category
                );
            }

            if (article.sentiment) {
                sentiments.add(
                    article.sentiment
                );
            }

            if (article.importance) {
                importance.add(
                    article.importance
                );
            }

            for (
                const value of article.assetClasses
            ) {
                assetClasses.add(value);
            }

            for (
                const value of article.geographies
            ) {
                geographies.add(value);
            }

            for (
                const value of article.sectors
            ) {
                sectors.add(value);
            }
        }
    }

    return {
        categories:
            [...categories].sort(),

        sentiments:
            [...sentiments].sort(),

        importance:
            [...importance].sort(),

        assetClasses:
            [...assetClasses].sort(),

        geographies:
            [...geographies].sort(),

        sectors:
            [...sectors].sort()
    };
}


/* ============================================================
   SUMMARY STATISTICS
============================================================ */

/**
 * Calculate basic statistics for the Market News page.
 */
function getNewsStatistics(
    articles
) {
    const safeArticles =
        Array.isArray(articles)
            ? articles
            : [];

    const byCategory = {};
    const bySentiment = {};
    const byImportance = {};

    let fundMonitoringCount = 0;

    for (const article of safeArticles) {
        const category =
            article.category || "UNKNOWN";

        const sentiment =
            article.sentiment || "UNKNOWN";

        const importance =
            article.importance || "UNKNOWN";

        byCategory[category] =
            (byCategory[category] || 0) + 1;

        bySentiment[sentiment] =
            (bySentiment[sentiment] || 0) + 1;

        byImportance[importance] =
            (byImportance[importance] || 0) + 1;

        if (
            article.fundMonitoringRelevant
        ) {
            fundMonitoringCount += 1;
        }
    }

    return {
        total:
            safeArticles.length,

        byCategory,
        bySentiment,
        byImportance,

        fundMonitoringCount
    };
}


/* ============================================================
   CURRENT NEWS VIEW MODEL
============================================================ */

/**
 * Prepare current Market News data for app.js.
 */
async function getCurrentMarketNews() {
    const current =
        await loadCurrentMarketNews();

    const articles =
        sortArticlesNewestFirst(
            deduplicateArticles(
                current.analyses
            )
        );

    return {
        ...current,

        analyses: articles,

        filterOptions:
            getNewsFilterOptions(
                articles
            ),

        statistics:
            getNewsStatistics(
                articles
            )
    };
}


/* ============================================================
   HISTORICAL NEWS VIEW MODEL
============================================================ */

/**
 * Load one historical date and prepare it for display.
 */
async function getHistoricalMarketNews(
    dateString
) {
    const history =
        await loadHistoricalMarketNews(
            dateString
        );

    const articles =
        sortArticlesNewestFirst(
            deduplicateArticles(
                history.analyses
            )
        );

    return {
        ...history,

        analyses: articles,

        filterOptions:
            getNewsFilterOptions(
                articles
            ),

        statistics:
            getNewsStatistics(
                articles
            )
    };
}


/* ============================================================
   MONTH HELPERS
============================================================ */

/**
 * Build a list of months from YYYY-MM values.
 *
 * Example:
 *
 * [
 *   {
 *      key: "2026-10",
 *      year: 2026,
 *      month: 10,
 *      label: "October 2026"
 *   }
 * ]
 */
function buildMonthOptions(
    dates
) {
    const monthMap = new Map();

    if (!Array.isArray(dates)) {
        return [];
    }

    for (const dateString of dates) {
        if (
            typeof dateString !== "string" ||
            !/^\d{4}-\d{2}-\d{2}$/.test(
                dateString
            )
        ) {
            continue;
        }

        const [
            year,
            month
        ] = dateString.split("-");

        const key =
            `${year}-${month}`;

        if (!monthMap.has(key)) {
            const date =
                new Date(
                    Date.UTC(
                        Number(year),
                        Number(month) - 1,
                        1
                    )
                );

            monthMap.set(
                key,
                {
                    key,
                    year: Number(year),
                    month: Number(month),
                    label:
                        new Intl.DateTimeFormat(
                            "en-SG",
                            {
                                month: "long",
                                year: "numeric",
                                timeZone: "UTC"
                            }
                        ).format(date)
                }
            );
        }
    }

    return [...monthMap.values()]
        .sort(
            (a, b) =>
                b.key.localeCompare(a.key)
        );
}


/**
 * Filter a list of YYYY-MM-DD dates to a selected month.
 */
function getDatesForMonth(
    dates,
    year,
    month
) {
    const yearString =
        String(year);

    const monthString =
        String(month).padStart(2, "0");

    const prefix =
        `${yearString}-${monthString}-`;

    return (
        Array.isArray(dates)
            ? dates
                .filter(
                    date =>
                        typeof date === "string" &&
                        date.startsWith(prefix)
                )
                .sort()
                .reverse()
            : []
    );
}


/* ============================================================
   EXPORTS
============================================================ */

export {
    MARKET_NEWS_PATHS,

    parseNewsDate,
    formatNewsDate,
    formatSingaporeDate,

    getHistoryPathParts,
    buildHistoryPath,

    fetchJson,

    normaliseArticle,
    normaliseAnalysisFile,

    loadCurrentMarketNews,
    loadHistoricalMarketNews,

    historyExists,
    generateDateRange,
    findAvailableHistoryDates,

    getArticleSingaporeDate,
    groupArticlesByDate,

    sortArticlesNewestFirst,
    sortArticlesOldestFirst,

    deduplicateArticles,

    filterArticles,
    searchArticles,

    getNewsFilterOptions,
    getNewsStatistics,

    getCurrentMarketNews,
    getHistoricalMarketNews,

    buildMonthOptions,
    getDatesForMonth
};
