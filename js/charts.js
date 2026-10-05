 /* ============================================================
    VGRAT FMS
    CHARTS MODULE
    ============================================================

    Chart.js is loaded by index.html.

    This module:
    - Creates and destroys charts safely.
    - Uses the existing VGrat FMS theme variables.
    - Provides reusable chart helpers.
    - Does not calculate financial performance.
    - Does not fetch application data.

    Current dashboard charts:
    - Performance comparison
    - News sentiment
    - News category
    - Asset-class exposure
    - Geography
    - Sector distribution

    Charts are optional. If Chart.js is unavailable, the
    application continues to function without charts.
 ============================================================ */


/* ============================================================
   CHART REGISTRY
============================================================ */

const chartRegistry =
    new Map();


/* ============================================================
   CHART.JS AVAILABILITY
============================================================ */

function getChartLibrary() {

    if (
        typeof window === "undefined"
    ) {

        return null;
    }


    if (
        typeof window.Chart ===
        "undefined"
    ) {

        return null;
    }


    return window.Chart;
}


function isChartAvailable() {

    return (
        getChartLibrary() !== null
    );
}


/* ============================================================
   THEME HELPERS
============================================================ */

function isLightTheme() {

    return (
        document.documentElement
            .getAttribute(
                "data-theme"
            ) === "light"
    );
}


function getThemeColors() {

    const styles =
        getComputedStyle(
            document.documentElement
        );


    return {

        text:
            styles
                .getPropertyValue(
                    "--text-primary"
                )
                .trim() ||
            (
                isLightTheme()
                    ? "#17202a"
                    : "#f3f6f8"
            ),

        muted:
            styles
                .getPropertyValue(
                    "--text-secondary"
                )
                .trim() ||
            (
                isLightTheme()
                    ? "#68737d"
                    : "#9aa7b2"
            ),

        border:
            styles
                .getPropertyValue(
                    "--border"
                )
                .trim() ||
            (
                isLightTheme()
                    ? "#d9dee3"
                    : "#27323c"
            ),

        grid:
            styles
                .getPropertyValue(
                    "--chart-grid"
                )
                .trim() ||
            (
                isLightTheme()
                    ? "rgba(30, 42, 52, 0.08)"
                    : "rgba(255, 255, 255, 0.07)"
            ),

        positive:
            styles
                .getPropertyValue(
                    "--positive"
                )
                .trim() ||
            "#2bb673",

        negative:
            styles
                .getPropertyValue(
                    "--negative"
                )
                .trim() ||
            "#e45d68",

        accent:
            styles
                .getPropertyValue(
                    "--accent"
                )
                .trim() ||
            "#4c8dff",

        surface:
            styles
                .getPropertyValue(
                    "--surface"
                )
                .trim() ||
            (
                isLightTheme()
                    ? "#ffffff"
                    : "#111922"
            )

    };
}


/* ============================================================
   DEFAULT OPTIONS
============================================================ */

function getDefaultChartOptions(
    overrides = {}
) {

    const colors =
        getThemeColors();


    return {

        responsive:
            true,

        maintainAspectRatio:
            false,

        animation: {
            duration: 250
        },

        interaction: {
            mode: "index",
            intersect: false
        },

        plugins: {

            legend: {

                display:
                    true,

                labels: {

                    color:
                        colors.text,

                    usePointStyle:
                        true,

                    boxWidth:
                        8,

                    boxHeight:
                        8,

                    padding:
                        16
                }

            },

            tooltip: {

                backgroundColor:
                    colors.surface,

                titleColor:
                    colors.text,

                bodyColor:
                    colors.text,

                borderColor:
                    colors.border,

                borderWidth:
                    1,

                padding:
                    10,

                displayColors:
                    true
            }

        },

        scales: {

            x: {

                ticks: {

                    color:
                        colors.muted,

                    maxRotation:
                        0
                },

                grid: {

                    color:
                        colors.grid,

                    drawBorder:
                        false
                }

            },

            y: {

                ticks: {

                    color:
                        colors.muted
                },

                grid: {

                    color:
                        colors.grid,

                    drawBorder:
                        false
                }

            }

        },

        ...overrides

    };
}


/* ============================================================
   CHART REGISTRATION
============================================================ */

function destroyChart(
    chartId
) {

    const chart =
        chartRegistry.get(
            chartId
        );


    if (chart) {

        chart.destroy();

        chartRegistry.delete(
            chartId
        );
    }
}


function destroyAllCharts() {

    for (
        const chart
        of chartRegistry.values()
    ) {

        chart.destroy();
    }


    chartRegistry.clear();
}


function registerChart(
    chartId,
    chart
) {

    destroyChart(
        chartId
    );


    chartRegistry.set(
        chartId,
        chart
    );


    return chart;
}


function getChart(
    chartId
) {

    return (
        chartRegistry.get(
            chartId
        ) ??
        null
    );
}


/* ============================================================
   CANVAS RESOLUTION
============================================================ */

function resolveCanvas(
    target
) {

    if (
        typeof target ===
        "string"
    ) {

        return document.querySelector(
            target
        );
    }


    if (
        target instanceof
        HTMLCanvasElement
    ) {

        return target;
    }


    return null;
}


/* ============================================================
   PERFORMANCE BAR CHART
============================================================ */

function createPerformanceChart(
    target,
    performanceRows,
    options = {}
) {

    const Chart =
        getChartLibrary();


    if (!Chart) {

        return null;
    }


    const canvas =
        resolveCanvas(
            target
        );


    if (!canvas) {

        return null;
    }


    const rows =
        Array.isArray(
            performanceRows
        )
            ? performanceRows
            : [];


    const labels =
        rows.map(
            row =>
                row.fundName ??
                row.fundCode ??
                "Fund"
        );


    const values =
        rows.map(
            row =>
                Number(
                    row.returnPercent
                )
        );


    const colors =
        getThemeColors();


    const barColors =
        values.map(
            value =>
                value >= 0
                    ? colors.positive
                    : colors.negative
        );


    const chart =
        new Chart(
            canvas,
            {

                type:
                    "bar",

                data: {

                    labels,

                    datasets: [

                        {

                            label:
                                "Return",

                            data:
                                values,

                            backgroundColor:
                                barColors,

                            borderWidth:
                                0,

                            borderRadius:
                                3,

                            barPercentage:
                                0.72,

                            categoryPercentage:
                                0.78

                        }

                    ]

                },

                options:
                    getDefaultChartOptions({

                        indexAxis:
                            "y",

                        plugins: {

                            legend: {
                                display:
                                    false
                            },

                            tooltip: {

                                callbacks: {

                                    label:
                                        context =>
                                            formatPercent(
                                                context.raw
                                            )

                                }

                            }

                        },

                        scales: {

                            x: {

                                ticks: {

                                    callback:
                                        value =>
                                            `${value}%`

                                }

                            },

                            y: {

                                ticks: {

                                    color:
                                        colors.text

                                }

                            }

                        },

                        ...options

                    })

            }
        );


    return registerChart(
        getCanvasId(
            canvas,
            "performance-chart"
        ),
        chart
    );
}


/* ============================================================
   PERFORMANCE LINE CHART
============================================================ */

function createPerformanceLineChart(
    target,
    datasets,
    labels,
    options = {}
) {

    const Chart =
        getChartLibrary();


    if (!Chart) {

        return null;
    }


    const canvas =
        resolveCanvas(
            target
        );


    if (!canvas) {

        return null;
    }


    const colors =
        getThemeColors();


    const preparedDatasets =
        Array.isArray(
            datasets
        )
            ? datasets.map(
                (
                    dataset,
                    index
                ) => {

                    const defaultColor =
                        index === 0
                            ? colors.accent
                            : index % 2 === 0
                                ? colors.positive
                                : colors.negative;


                    return {

                        ...dataset,

                        borderColor:
                            dataset.borderColor ??
                            defaultColor,

                        backgroundColor:
                            dataset.backgroundColor ??
                            "transparent",

                        borderWidth:
                            dataset.borderWidth ??
                            2,

                        pointRadius:
                            dataset.pointRadius ??
                            0,

                        pointHoverRadius:
                            dataset.pointHoverRadius ??
                            4,

                        tension:
                            dataset.tension ??
                            0.2

                    };

                }
            )
            : [];


    const chart =
        new Chart(
            canvas,
            {

                type:
                    "line",

                data: {

                    labels:
                        labels || [],

                    datasets:
                        preparedDatasets

                },

                options:
                    getDefaultChartOptions({

                        plugins: {

                            tooltip: {

                                callbacks: {

                                    label:
                                        context =>
                                            `${context.dataset.label}: ${formatPercent(context.raw)}`

                                }

                            }

                        },

                        scales: {

                            y: {

                                ticks: {

                                    callback:
                                        value =>
                                            `${value}%`

                                }

                            }

                        },

                        ...options

                    })

            }
        );


    return registerChart(
        getCanvasId(
            canvas,
            "performance-line-chart"
        ),
        chart
    );
}


/* ============================================================
   NEWS SENTIMENT DOUGHNUT
============================================================ */

function createNewsSentimentChart(
    target,
    statistics,
    options = {}
) {

    const Chart =
        getChartLibrary();


    if (!Chart) {

        return null;
    }


    const canvas =
        resolveCanvas(
            target
        );


    if (!canvas) {

        return null;
    }


    const colors =
        getThemeColors();


    const values = [

        Number(
            statistics?.positive ??
            0
        ),

        Number(
            statistics?.neutral ??
            0
        ),

        Number(
            statistics?.negative ??
            0
        )

    ];


    const chart =
        new Chart(
            canvas,
            {

                type:
                    "doughnut",

                data: {

                    labels: [
                        "Positive",
                        "Neutral",
                        "Negative"
                    ],

                    datasets: [

                        {

                            data:
                                values,

                            backgroundColor: [

                                colors.positive,

                                colors.muted,

                                colors.negative

                            ],

                            borderWidth:
                                0,

                            hoverOffset:
                                4

                        }

                    ]

                },

                options: {

                    responsive:
                        true,

                    maintainAspectRatio:
                        false,

                    cutout:
                        "68%",

                    plugins: {

                        legend: {

                            position:
                                "bottom",

                            labels: {

                                color:
                                    colors.text,

                                usePointStyle:
                                    true,

                                padding:
                                    16

                            }

                        },

                        tooltip: {

                            callbacks: {

                                label:
                                    context => {

                                        const total =
                                            values.reduce(
                                                (
                                                    sum,
                                                    value
                                                ) =>
                                                    sum +
                                                    value,
                                                0
                                            );


                                        const value =
                                            Number(
                                                context.raw
                                            );


                                        const percentage =
                                            total > 0
                                                ? (
                                                    value /
                                                    total
                                                ) *
                                                100
                                                : 0;


                                        return (
                                            `${context.label}: ` +
                                            `${value} ` +
                                            `(${percentage.toFixed(1)}%)`
                                        );

                                    }

                            }

                        }

                    },

                    ...options

                }

            }
        );


    return registerChart(
        getCanvasId(
            canvas,
            "news-sentiment-chart"
        ),
        chart
    );
}


/* ============================================================
   NEWS CATEGORY CHART
============================================================ */

function createNewsCategoryChart(
    target,
    categoryCounts,
    options = {}
) {

    const Chart =
        getChartLibrary();


    if (!Chart) {

        return null;
    }


    const canvas =
        resolveCanvas(
            target
        );


    if (!canvas) {

        return null;
    }


    const entries =
        sortCounterEntries(
            categoryCounts
        );


    const labels =
        entries.map(
            entry =>
                entry[0]
        );


    const values =
        entries.map(
            entry =>
                entry[1]
        );


    const colors =
        getThemeColors();


    const chart =
        new Chart(
            canvas,
            {

                type:
                    "bar",

                data: {

                    labels,

                    datasets: [

                        {

                            label:
                                "Articles",

                            data:
                                values,

                            backgroundColor:
                                colors.accent,

                            borderRadius:
                                3,

                            borderWidth:
                                0

                        }

                    ]

                },

                options:
                    getDefaultChartOptions({

                        plugins: {

                            legend: {
                                display:
                                    false
                            }

                        },

                        scales: {

                            y: {

                                beginAtZero:
                                    true,

                                ticks: {

                                    precision:
                                        0

                                }

                            }

                        },

                        ...options

                    })

            }
        );


    return registerChart(
        getCanvasId(
            canvas,
            "news-category-chart"
        ),
        chart
    );
}


/* ============================================================
   GENERIC DISTRIBUTION BAR CHART
============================================================ */

function createDistributionChart(
    target,
    counter,
    options = {}
) {

    const Chart =
        getChartLibrary();


    if (!Chart) {

        return null;
    }


    const canvas =
        resolveCanvas(
            target
        );


    if (!canvas) {

        return null;
    }


    const entries =
        sortCounterEntries(
            counter
        );


    const labels =
        entries.map(
            entry =>
                entry[0]
        );


    const values =
        entries.map(
            entry =>
                entry[1]
        );


    const colors =
        getThemeColors();


    const chart =
        new Chart(
            canvas,
            {

                type:
                    "bar",

                data: {

                    labels,

                    datasets: [

                        {

                            label:
                                "Articles",

                            data:
                                values,

                            backgroundColor:
                                colors.accent,

                            borderWidth:
                                0,

                            borderRadius:
                                3,

                            barPercentage:
                                0.68,

                            categoryPercentage:
                                0.82

                        }

                    ]

                },

                options:
                    getDefaultChartOptions({

                        indexAxis:
                            "y",

                        plugins: {

                            legend: {
                                display:
                                    false
                            }

                        },

                        scales: {

                            x: {

                                beginAtZero:
                                    true,

                                ticks: {

                                    precision:
                                        0

                                }

                            }

                        },

                        ...options

                    })

            }
        );


    return registerChart(
        getCanvasId(
            canvas,
            "distribution-chart"
        ),
        chart
    );
}


/* ============================================================
   NEWS IMPORTANCE CHART
============================================================ */

function createNewsImportanceChart(
    target,
    statistics,
    options = {}
) {

    return createDistributionChart(
        target,
        {

            HIGH:
                Number(
                    statistics?.high ??
                    0
                ),

            MEDIUM:
                Number(
                    statistics?.medium ??
                    0
                ),

            LOW:
                Number(
                    statistics?.low ??
                    0
                )

        },
        options
    );
}


/* ============================================================
   NEWS ASSET CLASS CHART
============================================================ */

function createNewsAssetClassChart(
    target,
    statistics,
    options = {}
) {

    return createDistributionChart(
        target,
        statistics?.assetClasses ??
        {},
        options
    );
}


/* ============================================================
   NEWS GEOGRAPHY CHART
============================================================ */

function createNewsGeographyChart(
    target,
    statistics,
    options = {}
) {

    return createDistributionChart(
        target,
        statistics?.geographies ??
        {},
        options
    );
}


/* ============================================================
   NEWS SECTOR CHART
============================================================ */

function createNewsSectorChart(
    target,
    statistics,
    options = {}
) {

    return createDistributionChart(
        target,
        statistics?.sectors ??
        {},
        options
    );
}


/* ============================================================
   CHART UPDATE HELPERS
============================================================ */

function updateChartData(
    chartId,
    labels,
    datasets
) {

    const chart =
        getChart(
            chartId
        );


    if (!chart) {

        return false;
    }


    chart.data.labels =
        labels || [];


    chart.data.datasets =
        datasets || [];


    chart.update();


    return true;
}


function updateChartTheme() {

    if (
        chartRegistry.size === 0
    ) {

        return;
    }


    const colors =
        getThemeColors();


    for (
        const chart
        of chartRegistry.values()
    ) {

        if (
            chart.options?.plugins
                ?.legend
                ?.labels
        ) {

            chart.options
                .plugins
                .legend
                .labels
                .color =
                colors.text;
        }


        if (
            chart.options?.scales
        ) {

            for (
                const scale
                of Object.values(
                    chart.options.scales
                )
            ) {

                if (scale.ticks) {

                    scale.ticks.color =
                        colors.muted;
                }


                if (scale.grid) {

                    scale.grid.color =
                        colors.grid;
                }
            }
        }


        if (
            chart.options?.plugins
                ?.tooltip
        ) {

            chart.options
                .plugins
                .tooltip
                .backgroundColor =
                colors.surface;


            chart.options
                .plugins
                .tooltip
                .titleColor =
                colors.text;


            chart.options
                .plugins
                .tooltip
                .bodyColor =
                colors.text;


            chart.options
                .plugins
                .tooltip
                .borderColor =
                colors.border;
        }


        chart.update(
            "none"
        );
    }
}


/* ============================================================
   UTILITY HELPERS
============================================================ */

function sortCounterEntries(
    counter
) {

    return Object.entries(
        counter || {}
    )
        .filter(
            (
                [, value]
            ) =>
                Number.isFinite(
                    Number(value)
                )
        )
        .sort(
            (
                a,
                b
            ) => {

                const valueDifference =
                    Number(b[1]) -
                    Number(a[1]);


                if (
                    valueDifference !==
                    0
                ) {

                    return valueDifference;
                }


                return a[0]
                    .localeCompare(
                        b[0]
                    );
            }
        );
}


function formatPercent(
    value
) {

    const number =
        Number(value);


    if (
        !Number.isFinite(
            number
        )
    ) {

        return "—";
    }


    const sign =
        number > 0
            ? "+"
            : "";


    return (
        sign +
        number.toFixed(2) +
        "%"
    );
}


function getCanvasId(
    canvas,
    fallback
) {

    if (
        canvas.id
    ) {

        return canvas.id;
    }


    /*
     * Generate a stable internal identifier.
     * This is only used by the chart registry.
     */
    const generated =
        `${fallback}-${Math.random()
            .toString(36)
            .slice(2, 10)}`;


    canvas.dataset.chartId =
        generated;


    return generated;
}


/* ============================================================
   RESPONSIVE RESIZE
============================================================ */

function resizeAllCharts() {

    for (
        const chart
        of chartRegistry.values()
    ) {

        chart.resize();
    }
}


/* ============================================================
   CLEANUP
============================================================ */

function removeChart(
    chartId
) {

    destroyChart(
        chartId
    );
}


/* ============================================================
   PUBLIC API
============================================================ */

export {

    isChartAvailable,

    getThemeColors,

    getDefaultChartOptions,

    createPerformanceChart,

    createPerformanceLineChart,

    createNewsSentimentChart,

    createNewsCategoryChart,

    createNewsImportanceChart,

    createNewsAssetClassChart,

    createNewsGeographyChart,

    createNewsSectorChart,

    updateChartData,

    updateChartTheme,

    resizeAllCharts,

    destroyChart,

    destroyAllCharts,

    removeChart,

    getChart

};
