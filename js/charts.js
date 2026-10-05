/* VGRAT FMS - CHARTS
 * ===================
 *
 * Optional Chart.js helpers.
 * The dashboard must remain functional even if Chart.js is unavailable.
 */

const chartRegistry = new Map();

function getChartConstructor() {
    return window.Chart ?? null;
}

function getCssVariable(name, fallback) {
    const value = getComputedStyle(document.documentElement)
        .getPropertyValue(name)
        .trim();

    return value || fallback;
}

function getThemeColors() {
    return {
        text: getCssVariable("--text-primary", "#f5f7fa"),
        muted: getCssVariable("--text-secondary", "#9aa4b2"),
        grid: getCssVariable("--border", "rgba(255,255,255,0.08)"),
        positive: getCssVariable("--positive", "#32c48d"),
        negative: getCssVariable("--negative", "#ef6b73"),
        accent: getCssVariable("--accent", "#4f8cff"),
        surface: getCssVariable("--surface", "#111827")
    };
}

function getCanvas(canvasId) {
    return document.getElementById(canvasId);
}

function destroyChart(canvasId) {
    const chart = chartRegistry.get(canvasId);

    if (!chart) {
        return;
    }

    try {
        chart.destroy();
    } catch {
        // Ignore cleanup errors.
    }

    chartRegistry.delete(canvasId);
}

function destroyAllCharts() {
    for (const canvasId of chartRegistry.keys()) {
        destroyChart(canvasId);
    }
}

function createChart(canvasId, config) {
    const ChartConstructor = getChartConstructor();

    if (!ChartConstructor) {
        return null;
    }

    const canvas = getCanvas(canvasId);

    if (!canvas) {
        return null;
    }

    destroyChart(canvasId);

    const chart = new ChartConstructor(canvas, config);

    chartRegistry.set(canvasId, chart);

    return chart;
}

function getBaseOptions() {
    const colors = getThemeColors();

    return {
        responsive: true,
        maintainAspectRatio: false,

        plugins: {
            legend: {
                labels: {
                    color: colors.text,
                    usePointStyle: true
                }
            },

            tooltip: {
                backgroundColor: colors.surface,
                titleColor: colors.text,
                bodyColor: colors.text,
                borderColor: colors.grid,
                borderWidth: 1
            }
        },

        scales: {
            x: {
                ticks: {
                    color: colors.muted
                },

                grid: {
                    color: colors.grid
                }
            },

            y: {
                ticks: {
                    color: colors.muted
                },

                grid: {
                    color: colors.grid
                }
            }
        }
    };
}

function createPerformanceChart(
    canvasId,
    labels,
    winners,
    losers
) {
    if (!labels?.length) {
        return null;
    }

    const colors = getThemeColors();

    return createChart(canvasId, {
        type: "bar",

        data: {
            labels,

            datasets: [
                {
                    label: "Winners",
                    data: winners,
                    backgroundColor: colors.positive,
                    borderRadius: 4
                },

                {
                    label: "Losers",
                    data: losers,
                    backgroundColor: colors.negative,
                    borderRadius: 4
                }
            ]
        },

        options: getBaseOptions()
    });
}

function createPerformanceLineChart(
    canvasId,
    labels,
    values
) {
    if (!labels?.length) {
        return null;
    }

    const colors = getThemeColors();

    return createChart(canvasId, {
        type: "line",

        data: {
            labels,

            datasets: [
                {
                    label: "Performance",
                    data: values,
                    borderColor: colors.accent,
                    backgroundColor: colors.accent,
                    tension: 0.25,
                    pointRadius: 3
                }
            ]
        },

        options: getBaseOptions()
    });
}

function createNewsSentimentChart(
    canvasId,
    statistics
) {
    if (!statistics) {
        return null;
    }

    const colors = getThemeColors();

    const labels = [
        "Positive",
        "Neutral",
        "Negative"
    ];

    const values = [
        statistics.sentiment?.POSITIVE ?? 0,
        statistics.sentiment?.NEUTRAL ?? 0,
        statistics.sentiment?.NEGATIVE ?? 0
    ];

    return createChart(canvasId, {
        type: "doughnut",

        data: {
            labels,

            datasets: [
                {
                    data: values,

                    backgroundColor: [
                        colors.positive,
                        colors.accent,
                        colors.negative
                    ],

                    borderWidth: 0
                }
            ]
        },

        options: {
            responsive: true,
            maintainAspectRatio: false,

            plugins: {
                legend: {
                    position: "bottom",

                    labels: {
                        color: colors.text,
                        usePointStyle: true
                    }
                }
            }
        }
    });
}

function createNewsCategoryChart(
    canvasId,
    statistics
) {
    return createNewsDistributionChart(
        canvasId,
        "Category",
        statistics?.categories
    );
}

function createNewsImportanceChart(
    canvasId,
    statistics
) {
    return createNewsDistributionChart(
        canvasId,
        "Importance",
        statistics?.importance
    );
}

function createNewsAssetClassChart(
    canvasId,
    statistics
) {
    return createNewsDistributionChart(
        canvasId,
        "Asset Class",
        statistics?.assetClasses
    );
}

function createNewsGeographyChart(
    canvasId,
    statistics
) {
    return createNewsDistributionChart(
        canvasId,
        "Geography",
        statistics?.geographies
    );
}

function createNewsSectorChart(
    canvasId,
    statistics
) {
    return createNewsDistributionChart(
        canvasId,
        "Sector",
        statistics?.sectors
    );
}

function createNewsDistributionChart(
    canvasId,
    label,
    source
) {
    if (!source) {
        destroyChart(canvasId);
        return null;
    }

    const entries = Array.isArray(source)
        ? source.map(item => [
            item.name ?? item.label ?? "Unknown",
            Number(item.value ?? item.count ?? 0)
        ])
        : Object.entries(source);

    if (!entries.length) {
        destroyChart(canvasId);
        return null;
    }

    const colors = getThemeColors();

    return createChart(canvasId, {
        type: "bar",

        data: {
            labels: entries.map(([key]) => key),

            datasets: [
                {
                    label,
                    data: entries.map(([, value]) => Number(value) || 0),
                    backgroundColor: colors.accent,
                    borderRadius: 4
                }
            ]
        },

        options: {
            ...getBaseOptions(),

            indexAxis: "y"
        }
    });
}

/*
 * Called when the dashboard theme changes.
 *
 * The existing chart instances are destroyed because their colours
 * are based on CSS variables. The application recreates the charts
 * after the theme change.
 */
function updateChartsForTheme() {
    destroyAllCharts();
}

function refreshChartsForTheme() {
    updateChartsForTheme();
}

export {
    chartRegistry,
    createChart,
    createPerformanceChart,
    createPerformanceLineChart,
    createNewsSentimentChart,
    createNewsCategoryChart,
    createNewsImportanceChart,
    createNewsAssetClassChart,
    createNewsGeographyChart,
    createNewsSectorChart,
    destroyChart,
    destroyAllCharts,
    updateChartsForTheme,
    refreshChartsForTheme
};
