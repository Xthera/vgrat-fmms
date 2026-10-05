/* ============================================================
   VGRAT FMS - CHARTS
   ============================================================

   Chart.js is loaded globally by index.html.

   This module:
   - Creates and destroys charts safely
   - Handles dark/light theme colours
   - Provides reusable chart helpers
   - Keeps Chart.js logic outside app.js

   Current dashboard use:
   - Market Performance winners/losers
   - Future fund-detail charts
============================================================ */


/* ============================================================
   CHART REGISTRY
============================================================ */

/*
 * Every chart created by this module is registered here.
 *
 * This prevents duplicate Chart.js instances when the user
 * changes period, switches theme, or navigates between views.
 */
const chartRegistry = new Map();


/* ============================================================
   DEFAULT SETTINGS
============================================================ */

const CHART_DEFAULTS = {
    animationDuration: 250,

    fontFamily:
        "Inter, -apple-system, BlinkMacSystemFont, " +
        "\"Segoe UI\", Roboto, Helvetica, Arial, sans-serif"
};


/* ============================================================
   THEME
============================================================ */

/**
 * Read a CSS custom property from the current document.
 */
function getCssVariable(name) {
    return getComputedStyle(
        document.documentElement
    )
        .getPropertyValue(name)
        .trim();
}


/**
 * Return the current dashboard theme.
 */
function getChartTheme() {
    const isLight =
        document.documentElement.dataset.theme === "light";

    return {
        isLight,

        text:
            getCssVariable("--text-primary") ||
            (isLight ? "#111827" : "#F3F4F6"),

        muted:
            getCssVariable("--text-secondary") ||
            (isLight ? "#6B7280" : "#9CA3AF"),

        border:
            getCssVariable("--border") ||
            (isLight ? "#D1D5DB" : "#374151"),

        grid:
            getCssVariable("--border-subtle") ||
            (isLight ? "#E5E7EB" : "#1F2937"),

        surface:
            getCssVariable("--surface") ||
            (isLight ? "#FFFFFF" : "#111827"),

        positive:
            getCssVariable("--positive") ||
            "#16A34A",

        negative:
            getCssVariable("--negative") ||
            "#DC2626",

        accent:
            getCssVariable("--accent") ||
            "#2563EB"
    };
}


/* ============================================================
   CHART.JS VALIDATION
============================================================ */

/**
 * Ensure Chart.js is available.
 */
function ensureChartJs() {
    if (
        typeof window === "undefined" ||
        typeof window.Chart !== "function"
    ) {
        throw new Error(
            "Chart.js is not available. " +
            "Make sure Chart.js is loaded before charts.js."
        );
    }

    return window.Chart;
}


/* ============================================================
   DESTROY
============================================================ */

/**
 * Destroy a registered chart.
 */
function destroyChart(chartId) {
    const chart =
        chartRegistry.get(chartId);

    if (!chart) {
        return;
    }

    try {
        chart.destroy();
    } catch {
        /*
         * Chart.js may already have destroyed the instance.
         */
    }

    chartRegistry.delete(chartId);
}


/**
 * Destroy all charts.
 */
function destroyAllCharts() {
    for (
        const chartId of chartRegistry.keys()
    ) {
        destroyChart(chartId);
    }
}


/**
 * Register a chart.
 */
function registerChart(
    chartId,
    chart
) {
    destroyChart(chartId);

    chartRegistry.set(
        chartId,
        chart
    );

    return chart;
}


/**
 * Get a registered chart.
 */
function getChart(chartId) {
    return chartRegistry.get(
        chartId
    ) || null;
}


/* ============================================================
   COMMON OPTIONS
============================================================ */

/**
 * Build shared Chart.js options.
 */
function getCommonChartOptions(
    theme
) {
    return {
        responsive: true,

        maintainAspectRatio: false,

        animation: {
            duration:
                CHART_DEFAULTS.animationDuration
        },

        plugins: {
            legend: {
                display: false
            },

            tooltip: {
                enabled: true,

                backgroundColor:
                    theme.surface,

                titleColor:
                    theme.text,

                bodyColor:
                    theme.text,

                borderColor:
                    theme.border,

                borderWidth: 1,

                padding: 10,

                displayColors: false,

                titleFont: {
                    family:
                        CHART_DEFAULTS.fontFamily,

                    weight: "600"
                },

                bodyFont: {
                    family:
                        CHART_DEFAULTS.fontFamily
                }
            }
        },

        interaction: {
            mode: "index",

            intersect: false
        }
    };
}


/* ============================================================
   PERFORMANCE BAR CHART
============================================================ */

/**
 * Create a horizontal performance ranking chart.
 *
 * Intended for:
 *   - Top 10 winners
 *   - Top 10 losers
 *
 * Data format:
 *
 * [
 *   {
 *      fundName: "...",
 *      fundCode: "PAPB",
 *      returnPercent: 4.25
 *   }
 * ]
 */
function createPerformanceBarChart(
    canvas,
    results,
    options = {}
) {
    const Chart =
        ensureChartJs();

    if (!canvas) {
        return null;
    }

    const theme =
        getChartTheme();

    const chartId =
        options.chartId ||
        canvas.id ||
        `performance-${Date.now()}`;

    const limit =
        Number.isFinite(options.limit)
            ? options.limit
            : 10;

    const displayResults =
        Array.isArray(results)
            ? results.slice(0, limit)
            : [];

    /*
     * Horizontal charts are easier to read for long
     * Prudential fund names.
     */
    const labels =
        displayResults.map(
            item =>
                item.fundCode ||
                item.fundName ||
                "-"
        );

    const values =
        displayResults.map(
            item =>
                Number.isFinite(
                    item.returnPercent
                )
                    ? item.returnPercent
                    : 0
        );

    const isPositiveChart =
        options.direction !== "losers";

    const defaultBarColour =
        isPositiveChart
            ? theme.positive
            : theme.negative;

    const backgroundColours =
        values.map(
            value => {
                if (value > 0) {
                    return theme.positive;
                }

                if (value < 0) {
                    return theme.negative;
                }

                return theme.muted;
            }
        );

    const context =
        canvas.getContext("2d");

    if (!context) {
        return null;
    }

    const chart =
        new Chart(
            context,
            {
                type: "bar",

                data: {
                    labels,

                    datasets: [
                        {
                            label:
                                "Return",

                            data:
                                values,

                            backgroundColor:
                                backgroundColours.length
                                    ? backgroundColours
                                    : defaultBarColour,

                            borderWidth: 0,

                            borderRadius: 4,

                            barPercentage: 0.72,

                            categoryPercentage: 0.82
                        }
                    ]
                },

                options: {
                    ...getCommonChartOptions(
                        theme
                    ),

                    indexAxis: "y",

                    scales: {
                        x: {
                            beginAtZero: false,

                            border: {
                                display: false
                            },

                            grid: {
                                color:
                                    theme.grid,

                                drawTicks: false
                            },

                            ticks: {
                                color:
                                    theme.muted,

                                font: {
                                    family:
                                        CHART_DEFAULTS.fontFamily
                                },

                                callback:
                                    value =>
                                        `${Number(value).toFixed(1)}%`
                            }
                        },

                        y: {
                            border: {
                                display: false
                            },

                            grid: {
                                display: false
                            },

                            ticks: {
                                color:
                                    theme.text,

                                font: {
                                    family:
                                        CHART_DEFAULTS.fontFamily,

                                    size: 11
                                }
                            }
                        }
                    },

                    plugins: {
                        ...getCommonChartOptions(
                            theme
                        ).plugins,

                        tooltip: {
                            ...getCommonChartOptions(
                                theme
                            ).plugins.tooltip,

                            callbacks: {
                                title:
                                    tooltipItems => {
                                        const index =
                                            tooltipItems[0]?.dataIndex;

                                        return (
                                            displayResults[
                                                index
                                            ]?.fundName ||
                                            "-"
                                        );
                                    },

                                label:
                                    tooltipItem => {
                                        const value =
                                            tooltipItem.raw;

                                        return (
                                            Number(value)
                                                .toFixed(2) +
                                            "%"
                                        );
                                    }
                            }
                        }
                    }
                }
            }
        );

    return registerChart(
        chartId,
        chart
    );
}


/* ============================================================
   PERFORMANCE DISTRIBUTION
============================================================ */

/**
 * Create a compact performance distribution chart.
 *
 * This is useful later for showing the spread of all eligible
 * Prudential funds for a selected period.
 */
function createPerformanceDistributionChart(
    canvas,
    results,
    options = {}
) {
    const Chart =
        ensureChartJs();

    if (!canvas) {
        return null;
    }

    const theme =
        getChartTheme();

    const chartId =
        options.chartId ||
        canvas.id ||
        `distribution-${Date.now()}`;

    const values =
        Array.isArray(results)
            ? results
                .map(
                    item =>
                        Number(item?.returnPercent)
                )
                .filter(
                    value =>
                        Number.isFinite(value)
                )
            : [];

    if (values.length === 0) {
        destroyChart(chartId);
        return null;
    }

    /*
     * Build sensible buckets dynamically.
     */
    const min =
        Math.floor(
            Math.min(...values)
        );

    const max =
        Math.ceil(
            Math.max(...values)
        );

    const range =
        Math.max(
            max - min,
            1
        );

    const bucketCount =
        Math.min(
            12,
            Math.max(
                5,
                Math.ceil(
                    Math.sqrt(
                        values.length
                    )
                )
            )
        );

    const bucketSize =
        range / bucketCount;

    const buckets =
        [];

    for (
        let index = 0;
        index < bucketCount;
        index += 1
    ) {
        const start =
            min +
            index * bucketSize;

        const end =
            index === bucketCount - 1
                ? max
                : start + bucketSize;

        buckets.push({
            start,
            end,
            count: 0
        });
    }

    for (
        const value of values
    ) {
        let index =
            Math.floor(
                (value - min) /
                bucketSize
            );

        if (
            index < 0
        ) {
            index = 0;
        }

        if (
            index >= buckets.length
        ) {
            index =
                buckets.length - 1;
        }

        buckets[index].count += 1;
    }

    const labels =
        buckets.map(
            bucket =>
                `${bucket.start.toFixed(1)}%`
        );

    const counts =
        buckets.map(
            bucket =>
                bucket.count
        );

    const context =
        canvas.getContext("2d");

    if (!context) {
        return null;
    }

    const chart =
        new Chart(
            context,
            {
                type: "bar",

                data: {
                    labels,

                    datasets: [
                        {
                            label:
                                "Funds",

                            data:
                                counts,

                            backgroundColor:
                                theme.accent,

                            borderWidth: 0,

                            borderRadius: 3,

                            barPercentage: 0.9,

                            categoryPercentage: 0.95
                        }
                    ]
                },

                options: {
                    ...getCommonChartOptions(
                        theme
                    ),

                    scales: {
                        x: {
                            border: {
                                display: false
                            },

                            grid: {
                                display: false
                            },

                            ticks: {
                                color:
                                    theme.muted
                            }
                        },

                        y: {
                            beginAtZero: true,

                            border: {
                                display: false
                            },

                            grid: {
                                color:
                                    theme.grid
                            },

                            ticks: {
                                color:
                                    theme.muted,

                                precision: 0
                            }
                        }
                    },

                    plugins: {
                        ...getCommonChartOptions(
                            theme
                        ).plugins,

                        tooltip: {
                            ...getCommonChartOptions(
                                theme
                            ).plugins.tooltip,

                            callbacks: {
                                title:
                                    tooltipItems =>
                                        tooltipItems[0]
                                            ?.label ||
                                        "-",

                                label:
                                    tooltipItem =>
                                        `${tooltipItem.raw} funds`
                            }
                        }
                    }
                }
            }
        );

    return registerChart(
        chartId,
        chart
    );
}


/* ============================================================
   FUND PERFORMANCE LINE CHART
============================================================ */

/**
 * Create a line chart for historical fund BID data.
 *
 * Expected input:
 *
 * [
 *   {
 *      date: "2026-10-01",
 *      bid: 1.10893
 *   }
 * ]
 */
function createFundBidChart(
    canvas,
    observations,
    options = {}
) {
    const Chart =
        ensureChartJs();

    if (!canvas) {
        return null;
    }

    const theme =
        getChartTheme();

    const chartId =
        options.chartId ||
        canvas.id ||
        `fund-bid-${Date.now()}`;

    const validObservations =
        Array.isArray(observations)
            ? observations.filter(
                observation =>
                    observation &&
                    typeof observation.date ===
                        "string" &&
                    Number.isFinite(
                        Number(
                            observation.bid ??
                            observation.bidPrice
                        )
                    )
            )
            : [];

    if (
        validObservations.length === 0
    ) {
        destroyChart(chartId);
        return null;
    }

    const labels =
        validObservations.map(
            observation =>
                observation.date
        );

    const values =
        validObservations.map(
            observation =>
                Number(
                    observation.bid ??
                    observation.bidPrice
                )
        );

    const context =
        canvas.getContext("2d");

    if (!context) {
        return null;
    }

    const chart =
        new Chart(
            context,
            {
                type: "line",

                data: {
                    labels,

                    datasets: [
                        {
                            label:
                                options.label ||
                                "BID",

                            data:
                                values,

                            borderColor:
                                theme.accent,

                            backgroundColor:
                                "transparent",

                            borderWidth: 2,

                            pointRadius: 0,

                            pointHoverRadius: 4,

                            tension: 0.15,

                            fill: false
                        }
                    ]
                },

                options: {
                    ...getCommonChartOptions(
                        theme
                    ),

                    scales: {
                        x: {
                            border: {
                                display: false
                            },

                            grid: {
                                display: false
                            },

                            ticks: {
                                color:
                                    theme.muted,

                                maxTicksLimit:
                                    options.maxTicksLimit ||
                                    8,

                                maxRotation: 0
                            }
                        },

                        y: {
                            border: {
                                display: false
                            },

                            grid: {
                                color:
                                    theme.grid
                            },

                            ticks: {
                                color:
                                    theme.muted
                            }
                        }
                    },

                    plugins: {
                        ...getCommonChartOptions(
                            theme
                        ).plugins,

                        tooltip: {
                            ...getCommonChartOptions(
                                theme
                            ).plugins.tooltip,

                            callbacks: {
                                title:
                                    tooltipItems =>
                                        tooltipItems[0]
                                            ?.label ||
                                        "-",

                                label:
                                    tooltipItem =>
                                        `BID ${Number(
                                            tooltipItem.raw
                                        ).toFixed(5)}`
                            }
                        }
                    }
                }
            }
        );

    return registerChart(
        chartId,
        chart
    );
}


/* ============================================================
   CHART REFRESH
============================================================ */

/**
 * Rebuild a chart after theme changes.
 *
 * Chart.js does not automatically understand every CSS
 * variable change, so the simplest reliable approach is to
 * destroy and recreate the chart.
 */
function refreshChart(
    chartId,
    createFunction
) {
    const existing =
        getChart(chartId);

    if (!existing) {
        return null;
    }

    /*
     * The caller is expected to provide the original
     * creation function if a full rebuild is required.
     */
    if (typeof createFunction !== "function") {
        existing.update();
        return existing;
    }

    destroyChart(chartId);

    return createFunction();
}


/**
 * Refresh all registered charts.
 *
 * Chart.js can update dimensions/theme-related rendering
 * without rebuilding when necessary.
 */
function updateAllCharts() {
    for (
        const chart of chartRegistry.values()
    ) {
        try {
            chart.update();
        } catch {
            /*
             * Ignore charts that have already been destroyed.
             */
        }
    }
}


/* ============================================================
   EXPORTS
============================================================ */

export {
    CHART_DEFAULTS,

    chartRegistry,

    getCssVariable,
    getChartTheme,
    ensureChartJs,

    destroyChart,
    destroyAllCharts,
    registerChart,
    getChart,

    getCommonChartOptions,

    createPerformanceBarChart,
    createPerformanceDistributionChart,
    createFundBidChart,

    refreshChart,
    updateAllCharts
};
