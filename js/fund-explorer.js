apen]");

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
