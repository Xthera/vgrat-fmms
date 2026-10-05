/* VGRAT FMS - CHARTS
 * ===================
 *
 * Optional Chart.js helpers for the VGrat FMS dashboard.
 *
 * The application must continue to work if Chart.js is unavailable.
 */

const chartRegistry = new Map();

function getChartConstructor() {
    return window.Chart ?? null;
}

function getCssVariable(name, fallback = "") {
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

function destroyChart(canvasId) {
    const chart = chartRegistry.get(canvasId);

    if (!chart) {
        return;
    }

    try {
        chart.destroy();
    } catch {
        // Ignore Chart.js cleanup failures.
    }

    chartRegistry.delete(canvasId);
}

function destroyAllCharts() {
    for (const canvasId of chartRegistry.keys()) {
        destroyChart(canvasId);
    }
}

function getCanvas(canvasId) {
    const canvas = document.getElementById(canvasId);

    if (!canvas) {
        return null;
    }

    return canvas;
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

function baseChartOptions(extra = {}) {
    const colors = getThemeColors();

    return {
        responsive: true,
        maintainAspectRatio: false,

        interaction: {
            intersect: false,
            mode: "index"
        },

        plugins: {
            legend: {
                labels: {
                    color: colors.text,
                    usePointStyle: true,
                    padding: 16
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
        },

        ...extra
    };
}

function createPerformanceChart(canvasId, labels, winners, losers) {
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

        options: baseChartOptions()
    });
}

function createPerformanceLineChart(canvasId, labels, values) {
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

        options: baseChartOptions()
    });
}

function createNewsSentimentChart(canvasId, statistics) {
    if (!statistics) {
        return null;
    }

    const colors = getThemeColors();

    const labels = ["Positive", "Neutral", "Negative"];

    const data = [
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
                    data,

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
                        usePointStyle: true,
                        padding: 16
                    }
                }
            }
        }
    });
}

function createNewsCategoryChart(canvasId, statistics) {
    if (!statistics) {
        return null;
    }

    const entries = Object.entries(statistics.categories ?? {});

    return createNewsDistributionChart(
        canvasId,
        "Category",
        entries
    );
}

function createNewsImportanceChart(canvasId, statistics) {
    if (!statistics) {
        return null;
    }

    const entries = Object.entries(statistics.importance ?? {});

    return createNewsDistributionChart(
        canvasId,
        "Importance",
        entries
    );
}

function createNewsAssetClassChart(canvasId, statistics) {
    if (!statistics) {
        return null;
    }

    const entries = Object.entries(statistics.assetClasses ?? {});

    return createNewsDistributionChart(
        canvasId,
        "Asset Class",
        entries
    );
}

function createNewsGeographyChart(canvasId, statistics) {
    if (!statistics) {
        return null;
    }

    const entries = Object.entries(statistics.geographies ?? {});

    return createNewsDistributionChart(
        canvasId,
        "Geography",
        entries
    );
}

function createNewsSectorChart(canvasId, statistics) {
    if (!statistics) {
        return null;
    }

    const entries = Object.entries(statistics.sectors ?? {});

    return createNewsDistributionChart(
        canvasId,
        "Sector",
        entries
    );
}

function createNewsDistributionChart(canvasId, label, entries) {
    if (!entries.length) {
        destroyChart(canvasId);
        return null;
    }

    const colors = getThemeColors();

    const labels = entries.map(([key]) => key);
    const values = entries.map(([, value]) => Number(value) || 0);

    return createChart(canvasId, {
        type: "bar",

        data: {
            labels,

            datasets: [
                {
                    label,
                    data: values,
                    backgroundColor: colors.accent,
                    borderRadius: 4
                }
            ]
        },

        options: baseChartOptions({
            indexAxis: "y"
        })
    });
}

/*
 * Rebuild all currently registered charts using the current theme.
 *
 * Chart.js does not automatically rebuild every chart when CSS variables
 * change, so this function destroys the existing instances. The application
 * can recreate the charts from its current data after the theme changes.
 */
function updateChartsForTheme() {
    const existingCharts = Array.from(chartRegistry.entries());

    if (!existingCharts.length) {
        return;
    }

    /*
     * Capture the canvas IDs only. The actual chart data/configuration is
     * owned by the application and should be recreated by the caller.
     */
    const canvasIds = existingCharts.map(([canvasId]) => canvasId);

    for (const canvasId of canvasIds) {
        destroyChart(canvasId);
    }
}

/*
 * Compatibility alias.
 *
 * Some older application code may call refreshChartsForTheme().
 */
function refreshChartsForTheme() {
    updateChartsForTheme();
}

export {
    chartRegistry,
    destroyChart,
    destroyAllCharts,
    createChart,
    createPerformanceChart,
    createPerformanceLineChart,
    createNewsSentimentChart,
    createNewsCategoryChart,
    createNewsImportanceChart,
    createNewsAssetClassChart,
    createNewsGeographyChart,
    createNewsSectorChart,
    updateChartsForTheme,
    refreshChartsForTheme
};
