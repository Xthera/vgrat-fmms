/* ============================================================
   VGRAT FMS - PERFORMANCE ENGINE
   ============================================================

   Purpose:
   - Calculate D-D / W-W / M-M / Y-Y performance
   - Use bid_history.json as the authoritative price source
   - Use actual BID observations only
   - Never forward-fill missing observations
   - Rank all Prudential funds dynamically

   Methodology
   -----------
   Current BID:
       Latest valid BID observation available for that fund.

   D-D:
       Compare current BID against the latest actual BID
       observation on or before current date - 1 calendar day.

   W-W:
       Compare current BID against the latest actual BID
       observation on or before current date - 7 calendar days.

   M-M:
       Compare current BID against the latest actual BID
       observation on or before the same calendar day
       one calendar month earlier.

   Y-Y:
       Compare current BID against the latest actual BID
       observation on or before the same calendar day
       one calendar year earlier.

   Return:
       ((Current BID - Historical BID) / Historical BID) * 100

   Important:
       Missing dates are NOT filled forward.
============================================================ */


/* ============================================================
   PERIOD DEFINITIONS
============================================================ */

const PERFORMANCE_PERIODS = {
    DD: {
        key: "DD",
        label: "D-D",
        description: "Day over Day"
    },

    WW: {
        key: "WW",
        label: "W-W",
        description: "Week over Week"
    },

    MM: {
        key: "MM",
        label: "M-M",
        description: "Month over Month"
    },

    YY: {
        key: "YY",
        label: "Y-Y",
        description: "Year over Year"
    }
};


/* ============================================================
   BASIC DATE UTILITIES
============================================================ */

/**
 * Parse YYYY-MM-DD into a UTC Date.
 *
 * Using UTC avoids browser-local timezone shifts.
 */
function parseISODate(dateString) {
    if (typeof dateString !== "string") {
        return null;
    }

    const match = /^(\d{4})-(\d{2})-(\d{2})$/.exec(dateString);

    if (!match) {
        return null;
    }

    const year = Number(match[1]);
    const month = Number(match[2]);
    const day = Number(match[3]);

    const date = new Date(Date.UTC(year, month - 1, day));

    if (
        date.getUTCFullYear() !== year ||
        date.getUTCMonth() !== month - 1 ||
        date.getUTCDate() !== day
    ) {
        return null;
    }

    return date;
}


/**
 * Format a UTC Date as YYYY-MM-DD.
 */
function formatISODate(date) {
    if (!(date instanceof Date) || Number.isNaN(date.getTime())) {
        return null;
    }

    return [
        String(date.getUTCFullYear()).padStart(4, "0"),
        String(date.getUTCMonth() + 1).padStart(2, "0"),
        String(date.getUTCDate()).padStart(2, "0")
    ].join("-");
}


/**
 * Subtract a number of calendar days.
 */
function subtractDays(date, days) {
    const result = new Date(date.getTime());

    result.setUTCDate(result.getUTCDate() - days);

    return result;
}


/**
 * Subtract one calendar month while handling month-end dates.
 *
 * Example:
 * 2026-10-05 -> 2026-09-05
 *
 * For dates such as:
 * 2026-03-31 -> 2026-02-28
 */
function subtractOneMonth(date) {
    const originalDay = date.getUTCDate();

    const result = new Date(
        Date.UTC(
            date.getUTCFullYear(),
            date.getUTCMonth(),
            1
        )
    );

    result.setUTCMonth(result.getUTCMonth() - 1);

    const lastDayOfPreviousMonth = new Date(
        Date.UTC(
            result.getUTCFullYear(),
            result.getUTCMonth() + 1,
            0
        )
    ).getUTCDate();

    result.setUTCDate(
        Math.min(originalDay, lastDayOfPreviousMonth)
    );

    return result;
}


/**
 * Subtract one calendar year while handling leap day.
 *
 * Example:
 * 2025-10-05 -> 2024-10-05
 */
function subtractOneYear(date) {
    const originalMonth = date.getUTCMonth();
    const originalDay = date.getUTCDate();

    const result = new Date(
        Date.UTC(
            date.getUTCFullYear() - 1,
            originalMonth,
            1
        )
    );

    const lastDayOfTargetMonth = new Date(
        Date.UTC(
            result.getUTCFullYear(),
            originalMonth + 1,
            0
        )
    ).getUTCDate();

    result.setUTCDate(
        Math.min(originalDay, lastDayOfTargetMonth)
    );

    return result;
}


/* ============================================================
   BID VALIDATION
============================================================ */

/**
 * Convert a BID value into a valid positive number.
 *
 * Accepts:
 *   1.1079
 *   "1.1079"
 *   "$1.1079"
 *
 * Rejects:
 *   null
 *   "-"
 *   ""
 *   NaN
 *   zero
 *   negative values
 */
function parseBid(value) {
    if (typeof value === "number") {
        return Number.isFinite(value) && value > 0
            ? value
            : null;
    }

    if (typeof value !== "string") {
        return null;
    }

    const cleaned = value
        .replace(/[$,\s]/g, "")
        .trim();

    if (!cleaned || cleaned === "-") {
        return null;
    }

    const numericValue = Number(cleaned);

    if (!Number.isFinite(numericValue) || numericValue <= 0) {
        return null;
    }

    return numericValue;
}


/* ============================================================
   FUND MATCHING
============================================================ */

/**
 * Build a stable identity key for a fund record.
 *
 * fundIdentifier is the strongest identity.
 * fundCode and excelRow are fallbacks.
 */
function getFundKey(fund) {
    if (!fund || typeof fund !== "object") {
        return null;
    }

    if (fund.fundIdentifier != null && fund.fundIdentifier !== "") {
        return `identifier:${String(fund.fundIdentifier).trim()}`;
    }

    if (fund.fundCode != null && fund.fundCode !== "") {
        return `code:${String(fund.fundCode).trim()}`;
    }

    if (fund.excelRow != null && fund.excelRow !== "") {
        return `row:${String(fund.excelRow).trim()}`;
    }

    return null;
}


/**
 * Find the matching bid-history record for a fund.
 */
function findBidHistoryRecord(fund, bidHistoryRecords) {
    if (!Array.isArray(bidHistoryRecords)) {
        return null;
    }

    const fundIdentifier = fund?.fundIdentifier;
    const fundCode = fund?.fundCode;
    const excelRow = fund?.excelRow;

    if (fundIdentifier != null && fundIdentifier !== "") {
        const match = bidHistoryRecords.find(
            record =>
                String(record?.fundIdentifier ?? "").trim() ===
                String(fundIdentifier).trim()
        );

        if (match) {
            return match;
        }
    }

    if (fundCode != null && fundCode !== "") {
        const match = bidHistoryRecords.find(
            record =>
                String(record?.fundCode ?? "").trim() ===
                String(fundCode).trim()
        );

        if (match) {
            return match;
        }
    }

    if (excelRow != null && excelRow !== "") {
        const match = bidHistoryRecords.find(
            record =>
                String(record?.excelRow ?? "").trim() ===
                String(excelRow).trim()
        );

        if (match) {
            return match;
        }
    }

    return null;
}


/* ============================================================
   OBSERVATION NORMALISATION
============================================================ */

/**
 * Convert raw BID observations into a clean internal format.
 *
 * Output:
 * [
 *   {
 *     date: "2026-10-02",
 *     dateObject: Date,
 *     bid: 1.10792
 *   }
 * ]
 *
 * Invalid observations are excluded.
 */
function normaliseObservations(observations) {
    if (!Array.isArray(observations)) {
        return [];
    }

    const byDate = new Map();

    for (const observation of observations) {
        if (!observation || typeof observation !== "object") {
            continue;
        }

        const date = observation.date;

        const dateObject = parseISODate(date);

        if (!dateObject) {
            continue;
        }

        const bid = parseBid(observation.bidPrice);

        if (bid === null) {
            continue;
        }

        /*
         * If a duplicate date exists, retain the latest record
         * encountered in the source.
         *
         * The production pipeline is expected to contain
         * unique observation dates.
         */
        byDate.set(date, {
            date,
            dateObject,
            bid
        });
    }

    return Array.from(byDate.values()).sort(
        (a, b) => a.dateObject.getTime() - b.dateObject.getTime()
    );
}


/* ============================================================
   OBSERVATION LOOKUP
============================================================ */

/**
 * Find the latest actual observation on or before targetDate.
 *
 * IMPORTANT:
 * This function does NOT forward-fill.
 *
 * If the target date is a weekend or holiday, the function
 * returns the most recent actual observation before that date.
 */
function findObservationOnOrBefore(observations, targetDate) {
    if (
        !Array.isArray(observations) ||
        observations.length === 0 ||
        !targetDate
    ) {
        return null;
    }

    const targetTime = targetDate.getTime();

    let low = 0;
    let high = observations.length - 1;
    let result = null;

    while (low <= high) {
        const middle = Math.floor((low + high) / 2);

        const observation = observations[middle];

        if (observation.dateObject.getTime() <= targetTime) {
            result = observation;
            low = middle + 1;
        } else {
            high = middle - 1;
        }
    }

    return result;
}


/* ============================================================
   TARGET DATE CALCULATION
============================================================ */

/**
 * Calculate the target comparison date for a period.
 */
function getTargetDate(currentDate, periodKey) {
    switch (periodKey) {
        case PERFORMANCE_PERIODS.DD.key:
            return subtractDays(currentDate, 1);

        case PERFORMANCE_PERIODS.WW.key:
            return subtractDays(currentDate, 7);

        case PERFORMANCE_PERIODS.MM.key:
            return subtractOneMonth(currentDate);

        case PERFORMANCE_PERIODS.YY.key:
            return subtractOneYear(currentDate);

        default:
            return null;
    }
}


/* ============================================================
   RETURN CALCULATION
============================================================ */

/**
 * Calculate percentage return.
 */
function calculateReturn(currentBid, historicalBid) {
    if (
        !Number.isFinite(currentBid) ||
        !Number.isFinite(historicalBid) ||
        historicalBid <= 0
    ) {
        return null;
    }

    return ((currentBid - historicalBid) / historicalBid) * 100;
}


/* ============================================================
   SINGLE FUND PERFORMANCE
============================================================ */

/**
 * Calculate one fund's performance for one period.
 */
function calculateFundPerformance(
    fund,
    bidHistoryRecord,
    periodKey
) {
    if (!fund || !bidHistoryRecord) {
        return null;
    }

    const observations = normaliseObservations(
        bidHistoryRecord?.bidHistory?.observations
    );

    if (observations.length === 0) {
        return null;
    }

    /*
     * Current value is the fund's latest actual observation.
     */
    const currentObservation =
        observations[observations.length - 1];

    const currentDate = currentObservation.dateObject;

    const targetDate = getTargetDate(
        currentDate,
        periodKey
    );

    if (!targetDate) {
        return null;
    }

    /*
     * Historical value is the latest actual observation
     * on or before the target date.
     */
    const historicalObservation =
        findObservationOnOrBefore(
            observations,
            targetDate
        );

    if (!historicalObservation) {
        return null;
    }

    /*
     * Do not calculate a comparison against itself.
     *
     * This can occur only if the history is unusually short
     * or the target date resolves to the current observation.
     */
    if (
        historicalObservation.date === currentObservation.date
    ) {
        return null;
    }

    const returnPercent = calculateReturn(
        currentObservation.bid,
        historicalObservation.bid
    );

    if (returnPercent === null) {
        return null;
    }

    return {
        period: periodKey,

        fundIdentifier:
            fund.fundIdentifier ?? null,

        fundCode:
            fund.fundCode ?? null,

        fundName:
            fund.fundName ?? null,

        excelRow:
            fund.excelRow ?? null,

        currentDate:
            currentObservation.date,

        currentBid:
            currentObservation.bid,

        comparisonTargetDate:
            formatISODate(targetDate),

        comparisonDate:
            historicalObservation.date,

        comparisonBid:
            historicalObservation.bid,

        returnPercent
    };
}


/* ============================================================
   ALL FUND PERFORMANCE
============================================================ */

/**
 * Calculate performance for every fund for one period.
 *
 * Universe:
 *     funds.json
 *
 * Price source:
 *     bid_history.json
 *
 * Funds without valid history/comparison data are excluded.
 */
function calculatePerformance(
    funds,
    bidHistoryRecords,
    periodKey
) {
    if (!Array.isArray(funds)) {
        return [];
    }

    if (!Array.isArray(bidHistoryRecords)) {
        return [];
    }

    const results = [];

    for (const fund of funds) {
        const bidHistoryRecord =
            findBidHistoryRecord(
                fund,
                bidHistoryRecords
            );

        if (!bidHistoryRecord) {
            continue;
        }

        const performance =
            calculateFundPerformance(
                fund,
                bidHistoryRecord,
                periodKey
            );

        if (!performance) {
            continue;
        }

        results.push({
            ...performance,
            fund
        });
    }

    return results;
}


/* ============================================================
   RANKING
============================================================ */

/**
 * Sort performance results from highest to lowest return.
 */
function sortByReturnDescending(results) {
    return [...results].sort((a, b) => {
        const returnDifference =
            b.returnPercent - a.returnPercent;

        if (returnDifference !== 0) {
            return returnDifference;
        }

        /*
         * Stable deterministic tie-breaker.
         */
        return String(a.fundName ?? "").localeCompare(
            String(b.fundName ?? "")
        );
    });
}


/**
 * Sort performance results from lowest to highest return.
 */
function sortByReturnAscending(results) {
    return [...results].sort((a, b) => {
        const returnDifference =
            a.returnPercent - b.returnPercent;

        if (returnDifference !== 0) {
            return returnDifference;
        }

        /*
         * Stable deterministic tie-breaker.
         */
        return String(a.fundName ?? "").localeCompare(
            String(b.fundName ?? "")
        );
    });
}


/**
 * Create winners and losers rankings.
 *
 * Default:
 *     10 winners
 *     10 losers
 */
function rankPerformance(
    results,
    limit = 10
) {
    const validResults = Array.isArray(results)
        ? results.filter(
            item => Number.isFinite(item.returnPercent)
        )
        : [];

    const winners =
        sortByReturnDescending(validResults)
            .slice(0, limit);

    const losers =
        sortByReturnAscending(validResults)
            .slice(0, limit);

    return {
        totalEligibleFunds: validResults.length,
        winners,
        losers
    };
}


/* ============================================================
   ALL PERIODS
============================================================ */

/**
 * Calculate and rank all four performance periods.
 */
function calculateAllPerformance(
    funds,
    bidHistoryRecords,
    limit = 10
) {
    const periods = {};

    for (const period of Object.values(PERFORMANCE_PERIODS)) {
        const results =
            calculatePerformance(
                funds,
                bidHistoryRecords,
                period.key
            );

        periods[period.key] = {
            ...period,
            ...rankPerformance(results, limit),
            allResults: results
        };
    }

    return periods;
}


/* ============================================================
   MARKET LATEST DATE
============================================================ */

/**
 * Determine the latest BID observation date across the entire
 * Prudential universe.
 *
 * This is useful for the dashboard header.
 */
function getLatestMarketDate(bidHistoryRecords) {
    if (!Array.isArray(bidHistoryRecords)) {
        return null;
    }

    let latestDate = null;

    for (const record of bidHistoryRecords) {
        const observations =
            normaliseObservations(
                record?.bidHistory?.observations
            );

        if (observations.length === 0) {
            continue;
        }

        const fundLatest =
            observations[observations.length - 1].date;

        if (
            latestDate === null ||
            fundLatest > latestDate
        ) {
            latestDate = fundLatest;
        }
    }

    return latestDate;
}


/* ============================================================
   FORMATTING HELPERS
============================================================ */

/**
 * Format percentage for display.
 *
 * Example:
 *     4.123456 -> "+4.12%"
 *    -2.456789 -> "-2.46%"
 *     0 -> "0.00%"
 */
function formatReturnPercent(value) {
    if (!Number.isFinite(value)) {
        return "-";
    }

    const sign = value > 0
        ? "+"
        : "";

    return `${sign}${value.toFixed(2)}%`;
}


/**
 * Format BID for display.
 *
 * BID precision varies between funds, so preserve enough
 * precision for useful display without excessive decimals.
 */
function formatBid(value) {
    if (!Number.isFinite(value)) {
        return "-";
    }

    if (value >= 100) {
        return value.toFixed(2);
    }

    if (value >= 10) {
        return value.toFixed(3);
    }

    if (value >= 1) {
        return value.toFixed(4);
    }

    if (value >= 0.1) {
        return value.toFixed(5);
    }

    return value.toFixed(6);
}


/**
 * Format an ISO date for dashboard display.
 *
 * Example:
 *     2026-10-02 -> 02 Oct 2026
 */
function formatDisplayDate(dateString) {
    const date = parseISODate(dateString);

    if (!date) {
        return "-";
    }

    return new Intl.DateTimeFormat(
        "en-SG",
        {
            day: "2-digit",
            month: "short",
            year: "numeric",
            timeZone: "UTC"
        }
    ).format(date);
}


/* ============================================================
   EXPORTS
============================================================ */

export {
    PERFORMANCE_PERIODS,

    parseISODate,
    formatISODate,
    subtractDays,
    subtractOneMonth,
    subtractOneYear,

    parseBid,

    getFundKey,
    findBidHistoryRecord,

    normaliseObservations,
    findObservationOnOrBefore,

    getTargetDate,
    calculateReturn,

    calculateFundPerformance,
    calculatePerformance,

    sortByReturnDescending,
    sortByReturnAscending,

    rankPerformance,
    calculateAllPerformance,

    getLatestMarketDate,

    formatReturnPercent,
    formatBid,
    formatDisplayDate
};
