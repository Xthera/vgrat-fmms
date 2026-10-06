/* ============================================================
   VGRAT FMS — REPORT GENERATION (investment growth)
   ============================================================
   Builds a client-ready "Investment Growth Report" from the
   actual BID price history (data/bid_history.json):

   - a portfolio of up to 8 funds with a % split
   - an initial lump sum plus optional regular top-ups
     (monthly / quarterly / half-yearly / yearly)
   - units are bought at the BID price on the first price date
     on or after each investment date (no sales charge)
   - value on any date = units held x BID on or before that date
   - dividends paid out by distribution funds are NOT added back

   The report prints cleanly (Print / Save as PDF) on a white
   A4 page whatever theme the site is using.
   ============================================================ */

import { getFontScale } from "./charts.js";

const MAX_FUNDS = 8;

const FREQUENCIES = {
    none: { label: "No top-ups", months: 0 },
    monthly: { label: "Monthly", months: 1 },
    quarterly: { label: "Quarterly", months: 3 },
    "half-yearly": { label: "Half-yearly", months: 6 },
    yearly: { label: "Yearly", months: 12 }
};

const PRINT_COLORS = ["#1f5f99", "#c0661f", "#1a8a62", "#b38600", "#b0456f", "#2e7d32", "#6a5acd", "#c0392b"];

const report = {
    initialized: false,
    funds: [],
    fundById: new Map(),
    index: new Map(),
    selected: [],          // [{ id, weight }]
    chart: null,
    lastResult: null
};


/* ============================================================
   HELPERS
   ============================================================ */

const qs = (selector, root = document) => root.querySelector(selector);

function escapeHtml(value) {
    return String(value ?? "")
        .replaceAll("&", "&amp;")
        .replaceAll("<", "&lt;")
        .replaceAll(">", "&gt;")
        .replaceAll('"', "&quot;")
        .replaceAll("'", "&#039;");
}

function fundId(fund) {
    return String(fund?.fundIdentifier ?? fund?.fundCode ?? fund?.excelRow ?? "");
}

function observationsFor(id) {
    const list = report.index.get(id);

    return Array.isArray(list) ? list : [];
}

function currencyOf(id) {
    const details = report.fundById.get(id)?.fund ?? {};

    return String(details.unitCurrency ?? details.fundCurrency ?? "SGD").toUpperCase();
}

function money(value, currency) {
    if (value === null || value === undefined || !Number.isFinite(value)) return "—";

    const prefix = currency === "USD" ? "US$" : currency === "SGD" ? "S$" : `${currency} `;
    const sign = value < 0 ? "−" : "";

    return `${sign}${prefix}${Math.abs(value).toLocaleString("en-SG", { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`;
}

function percent(value, digits = 2, signed = true) {
    if (value === null || value === undefined || !Number.isFinite(value)) return "—";

    const sign = signed && value > 0 ? "+" : value < 0 ? "−" : "";

    return `${sign}${Math.abs(value).toFixed(digits)}%`;
}

function formatDate(iso) {
    if (!iso) return "—";

    const [year, month, day] = String(iso).split("-").map(Number);
    const months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];

    return `${String(day).padStart(2, "0")} ${months[month - 1]} ${year}`;
}

function addMonths(iso, months) {
    const [year, month, day] = iso.split("-").map(Number);
    const target = new Date(Date.UTC(year, month - 1 + months, 1));
    const lastDay = new Date(Date.UTC(target.getUTCFullYear(), target.getUTCMonth() + 1, 0)).getUTCDate();

    target.setUTCDate(Math.min(day, lastDay));

    return target.toISOString().slice(0, 10);
}

function yearsBetween(a, b) {
    return (Date.parse(b) - Date.parse(a)) / (365.25 * 86400000);
}

/* first observation on or after the date (buying) */
function priceOnOrAfter(observations, iso) {
    let low = 0;
    let high = observations.length - 1;
    let found = -1;

    while (low <= high) {
        const mid = (low + high) >> 1;

        if (observations[mid].date >= iso) {
            found = mid;
            high = mid - 1;
        } else {
            low = mid + 1;
        }
    }

    return found >= 0 ? observations[found] : null;
}

/* last observation on or before the date (valuing) */
function priceOnOrBefore(observations, iso) {
    let low = 0;
    let high = observations.length - 1;
    let found = -1;

    while (low <= high) {
        const mid = (low + high) >> 1;

        if (observations[mid].date <= iso) {
            found = mid;
            low = mid + 1;
        } else {
            high = mid - 1;
        }
    }

    return found >= 0 ? observations[found] : null;
}

/* Money-weighted annual return (XIRR) of dated cash flows */
function xirr(flows) {
    if (flows.length < 2) return null;

    const t0 = Date.parse(flows[0].date);
    const value = rate => flows.reduce(
        (sum, flow) => sum + flow.amount / Math.pow(1 + rate, (Date.parse(flow.date) - t0) / (365.25 * 86400000)),
        0
    );

    let low = -0.9999;
    let high = 10;
    let fLow = value(low);
    const fHigh = value(high);

    if (!Number.isFinite(fLow) || !Number.isFinite(fHigh) || fLow * fHigh > 0) return null;

    for (let step = 0; step < 200; step += 1) {
        const mid = (low + high) / 2;
        const fMid = value(mid);

        if (Math.abs(fMid) < 1e-7) return mid * 100;

        if (fLow * fMid < 0) {
            high = mid;
        } else {
            low = mid;
            fLow = fMid;
        }
    }

    return ((low + high) / 2) * 100;
}


/* ============================================================
   DATE LIMITS FOR THE CHOSEN FUNDS
   ============================================================ */

function dateLimits() {
    const ids = report.selected.map(item => item.id);

    if (!ids.length) return null;

    let earliest = null;
    let latest = null;

    for (const id of ids) {
        const obs = observationsFor(id);

        if (!obs.length) return null;

        const first = obs[0].date;
        const last = obs[obs.length - 1].date;

        if (!earliest || first > earliest) earliest = first;
        if (!latest || last < latest) latest = last;
    }

    return earliest <= latest ? { earliest, latest } : null;
}


/* ============================================================
   CALCULATION
   ============================================================ */

function calculate(input) {
    const { lumpSum, topUp, frequency, start, end } = input;
    const months = FREQUENCIES[frequency]?.months ?? 0;

    // Investment dates: the start date, then every period after it
    const investDates = [start];

    if (topUp > 0 && months > 0) {
        let next = addMonths(start, months);

        while (next <= end) {
            investDates.push(next);
            next = addMonths(next, months);
        }
    }

    const skipped = [];

    const funds = report.selected.map((item, slot) => {
        const obs = observationsFor(item.id);
        const weight = item.weight / 100;
        const purchases = [];

        investDates.forEach((date, index) => {
            const amount = (index === 0 ? lumpSum + topUp : topUp) * weight;

            if (amount <= 0) return;

            const price = priceOnOrAfter(obs, date);

            if (!price || price.date > end) {
                // No BID price yet for this date (after the latest price)
                skipped.push({ date, amount });
                return;
            }

            purchases.push({ date: price.date, amount, units: amount / price.bidPrice, bid: price.bidPrice });
        });

        return { ...item, slot, obs, purchases, fund: report.fundById.get(item.id) };
    });

    // All price dates in the period, across the chosen funds
    const dateSet = new Set();

    for (const fund of funds) {
        for (const observation of fund.obs) {
            if (observation.date >= start && observation.date <= end) dateSet.add(observation.date);
        }
    }

    const dates = [...dateSet].sort();

    const series = dates.map(date => {
        let value = 0;
        let invested = 0;

        for (const fund of funds) {
            const units = fund.purchases.filter(p => p.date <= date).reduce((sum, p) => sum + p.units, 0);
            const spent = fund.purchases.filter(p => p.date <= date).reduce((sum, p) => sum + p.amount, 0);
            const price = priceOnOrBefore(fund.obs, date);

            invested += spent;
            value += price ? units * price.bidPrice : 0;
        }

        return { date, value, invested };
    }).filter(point => point.invested > 0);

    const endDate = series.length ? series[series.length - 1].date : end;

    const perFund = funds.map(fund => {
        const units = fund.purchases.reduce((sum, p) => sum + p.units, 0);
        const invested = fund.purchases.reduce((sum, p) => sum + p.amount, 0);
        const endPrice = priceOnOrBefore(fund.obs, endDate);
        const value = endPrice ? units * endPrice.bidPrice : 0;

        return {
            ...fund,
            units,
            invested,
            startBid: fund.purchases[0]?.bid ?? null,
            startDate: fund.purchases[0]?.date ?? null,
            endBid: endPrice?.bidPrice ?? null,
            endDate: endPrice?.date ?? null,
            averageCost: units ? invested / units : null,
            value,
            gain: value - invested,
            gainPct: invested ? ((value - invested) / invested) * 100 : null
        };
    });

    const totalInvested = perFund.reduce((sum, fund) => sum + fund.invested, 0);
    const totalValue = perFund.reduce((sum, fund) => sum + fund.value, 0);

    // Cash flows for the annualised (money-weighted) return
    const flowMap = new Map();

    for (const fund of perFund) {
        for (const purchase of fund.purchases) {
            flowMap.set(purchase.date, (flowMap.get(purchase.date) ?? 0) - purchase.amount);
        }
    }

    const flows = [...flowMap.entries()]
        .sort((a, b) => a[0].localeCompare(b[0]))
        .map(([date, amount]) => ({ date, amount }));

    flows.push({ date: endDate, amount: totalValue });

    const firstDate = series[0]?.date ?? start;
    const annualised = yearsBetween(firstDate, endDate) >= 1 ? xirr(flows) : null;

    // Year-by-year: each calendar year end in the period + the end date
    const yearRows = [];
    const firstYear = Number(firstDate.slice(0, 4));
    const lastYear = Number(endDate.slice(0, 4));

    const pointOnOrBefore = iso => {
        let found = null;

        for (const point of series) {
            if (point.date <= iso) found = point;
            else break;
        }

        return found;
    };

    for (let year = firstYear; year <= lastYear; year += 1) {
        const target = year === lastYear ? endDate : `${year}-12-31`;
        const point = pointOnOrBefore(target);

        if (!point) continue;
        if (yearRows.length && yearRows[yearRows.length - 1].date === point.date) continue;

        yearRows.push({
            label: year === lastYear ? `${formatDate(point.date)} (latest)` : `End ${year}`,
            date: point.date,
            invested: point.invested,
            value: point.value,
            gain: point.value - point.invested,
            gainPct: point.invested ? ((point.value - point.invested) / point.invested) * 100 : null
        });
    }

    return {
        input,
        firstDate,
        endDate,
        series,
        perFund,
        totalInvested,
        totalValue,
        gain: totalValue - totalInvested,
        gainPct: totalInvested ? ((totalValue - totalInvested) / totalInvested) * 100 : null,
        annualised,
        contributions: flows.length - 1,
        yearRows,
        requestedEnd: end,
        priceDateBeforeEnd: endDate < end,
        skippedAmount: skipped.reduce((sum, item) => sum + item.amount, 0),
        skippedCount: new Set(skipped.map(item => item.date)).size
    };
}


/* ============================================================
   SETUP FORM
   ============================================================ */

function renderSelected() {
    const body = qs("#report-funds-body");
    const total = qs("#report-allocation-total");

    if (!body) return;

    if (!report.selected.length) {
        body.innerHTML = `<tr><td colspan="4" class="report-empty">Add funds above (up to ${MAX_FUNDS}).</td></tr>`;
    } else {
        body.innerHTML = report.selected
            .map((item, index) => {
                const fund = report.fundById.get(item.id);

                return `
                    <tr>
                        <td>
                            <div class="fund-name">${escapeHtml(fund?.fundName ?? item.id)}</div>
                            <div class="fund-meta">${escapeHtml(fund?.fundCode ?? "")} · ${escapeHtml(currencyOf(item.id))}</div>
                        </td>
                        <td class="report-weight-cell">
                            <input type="number" min="0" max="100" step="1" value="${item.weight}" data-report-weight="${index}" aria-label="Allocation % for ${escapeHtml(fund?.fundName ?? item.id)}">
                            <span>%</span>
                        </td>
                        <td class="report-remove-cell">
                            <button type="button" class="compare-remove" data-report-remove="${index}" aria-label="Remove ${escapeHtml(fund?.fundName ?? item.id)}" title="Remove">×</button>
                        </td>
                    </tr>
                `;
            })
            .join("");
    }

    const sum = report.selected.reduce((acc, item) => acc + (Number(item.weight) || 0), 0);

    if (total) {
        total.textContent = `Total ${Math.round(sum * 100) / 100}%`;
        total.classList.toggle("is-ok", Math.abs(sum - 100) < 0.01);
        total.classList.toggle("is-off", report.selected.length > 0 && Math.abs(sum - 100) >= 0.01);
    }

    syncDateLimits();
}

function syncDateLimits() {
    const startInput = qs("#report-start");
    const endInput = qs("#report-end");
    const note = qs("#report-date-note");
    const limits = dateLimits();

    if (!startInput || !endInput) return;

    if (!limits) {
        startInput.removeAttribute("min");
        startInput.removeAttribute("max");
        endInput.removeAttribute("min");
        endInput.removeAttribute("max");
        if (note) note.textContent = "";
        return;
    }

    const today = new Date().toISOString().slice(0, 10);

    // Start: must fall within the price history.
    startInput.min = limits.earliest;
    startInput.max = limits.latest;

    // End: any date from the first price onwards (not limited to the
    // latest BID date). The report states the price date actually used.
    endInput.min = limits.earliest;
    endInput.max = today;     // no future dates

    if (!endInput.value || endInput.value < limits.earliest || endInput.value > today) {
        endInput.value = today;
    }

    if (!startInput.value) {
        const threeYears = addMonths(limits.latest, -36);

        startInput.value = threeYears > limits.earliest ? threeYears : limits.earliest;
    }

    if (startInput.value < limits.earliest) startInput.value = limits.earliest;
    if (startInput.value > limits.latest) startInput.value = limits.latest;

    if (note) {
        note.textContent =
            `Price history for the chosen funds: ${formatDate(limits.earliest)} to ${formatDate(limits.latest)}. ` +
            `The end date can be up to today; if it is after ${formatDate(limits.latest)}, the report values the portfolio at that latest BID price and says so.`;
    }
}

function splitEqually() {
    const n = report.selected.length;

    if (!n) return;

    const base = Math.floor(100 / n);
    let remainder = 100 - base * n;

    report.selected = report.selected.map(item => {
        const extra = remainder > 0 ? 1 : 0;

        remainder -= extra;

        return { ...item, weight: base + extra };
    });

    renderSelected();
}

function addFund(id) {
    if (!id || report.selected.some(item => item.id === id) || report.selected.length >= MAX_FUNDS) return;

    report.selected.push({ id, weight: 0 });
    splitEqually();
}

function renderSuggestions() {
    const input = qs("#report-fund-search");
    const list = qs("#report-fund-suggestions");

    if (!input || !list) return;

    const query = input.value.trim().toLowerCase();
    const chosen = new Set(report.selected.map(item => item.id));

    const matches = report.funds
        .filter(fund => {
            const id = fundId(fund);

            if (chosen.has(id) || !observationsFor(id).length) return false;
            if (!query) return true;

            return (
                String(fund.fundName ?? "").toLowerCase().includes(query) ||
                String(fund.fundCode ?? "").toLowerCase().includes(query)
            );
        });

    list.innerHTML = matches.length
        ? matches
            .map(fund => `
                <li class="compare-suggestion" role="option" data-report-add="${escapeHtml(fundId(fund))}">
                    <span class="compare-suggestion-name">${escapeHtml(fund.fundName ?? "")}</span>
                    <span class="compare-suggestion-code">${escapeHtml(fund.fundCode ?? "")} · ${escapeHtml(currencyOf(fundId(fund)))}</span>
                </li>
            `)
            .join("")
        : `<li class="compare-suggestion-empty">${report.selected.length >= MAX_FUNDS ? `Maximum ${MAX_FUNDS} funds` : "No matching funds"}</li>`;

    list.hidden = false;
}

function closeSuggestions() {
    const list = qs("#report-fund-suggestions");

    if (list) list.hidden = true;
}

function readInput() {
    const number = selector => {
        const value = Number(String(qs(selector)?.value ?? "").replace(/[^0-9.]/g, ""));

        return Number.isFinite(value) ? value : 0;
    };

    return {
        clientName: qs("#report-client")?.value.trim() ?? "",
        adviserName: qs("#report-adviser")?.value.trim() ?? "",
        adviserContact: qs("#report-contact")?.value.trim() ?? "",
        lumpSum: number("#report-lump"),
        topUp: number("#report-topup"),
        frequency: qs("#report-frequency")?.value ?? "monthly",
        start: qs("#report-start")?.value ?? "",
        end: qs("#report-end")?.value ?? ""
    };
}

function validate(input) {
    if (!report.selected.length) return "Add at least one fund.";

    const sum = report.selected.reduce((acc, item) => acc + (Number(item.weight) || 0), 0);

    if (Math.abs(sum - 100) >= 0.01) return `The allocation adds up to ${Math.round(sum * 100) / 100}%. It needs to total 100%.`;

    if (report.selected.some(item => !(Number(item.weight) > 0))) return "Every fund needs an allocation above 0%. Remove funds you don't want.";

    const currencies = new Set(report.selected.map(item => currencyOf(item.id)));

    if (currencies.size > 1) return `The chosen funds are in different currencies (${[...currencies].join(" and ")}). Pick funds in one currency for a report.`;

    if (input.lumpSum <= 0 && input.topUp <= 0) return "Enter an initial investment, a regular top-up, or both.";

    if (input.topUp > 0 && input.frequency === "none") return "Choose how often the regular top-up is made.";

    const limits = dateLimits();

    if (!limits) return "No overlapping price history for these funds.";

    if (!input.start || !input.end) return "Choose a start and end date.";

    if (input.start >= input.end) return "The start date must be before the end date.";

    const today = new Date().toISOString().slice(0, 10);

    if (input.end > today) return `The end date can't be in the future. Choose ${formatDate(today)} or earlier.`;

    if (input.start < limits.earliest || input.start > limits.latest) {
        return `Choose a start date between ${formatDate(limits.earliest)} and ${formatDate(limits.latest)}.`;
    }

    return null;
}


/* ============================================================
   REPORT OUTPUT
   ============================================================ */

function renderReport(result) {
    const sheet = qs("#report-sheet");
    const wrap = qs("#report-output");

    if (!sheet || !wrap) return;

    const { input, perFund } = result;
    const currency = currencyOf(perFund[0].id);
    const m = value => money(value, currency);
    const tone = value => (value > 0 ? "is-up" : value < 0 ? "is-down" : "");
    const frequency = FREQUENCIES[input.topUp > 0 ? input.frequency : "none"];
    const today = formatDate(new Date().toISOString().slice(0, 10));

    const planText = [
        input.lumpSum > 0 ? `${m(input.lumpSum)} invested on ${formatDate(result.firstDate)}` : "",
        input.topUp > 0 ? `${m(input.topUp)} ${frequency.label.toLowerCase()} top-ups` : ""
    ].filter(Boolean).join(" plus ");

    sheet.innerHTML = `
        <header class="rpt-header">
            <img src="assets/logo.png" alt="" class="rpt-logo" width="56" height="56">
            <div class="rpt-title-block">
                <div class="rpt-eyebrow">VGrat FMS · Funds Monitoring System</div>
                <h2 class="rpt-title">Investment Growth Report</h2>
                <div class="rpt-subtitle">${escapeHtml(formatDate(result.firstDate))} to ${escapeHtml(formatDate(result.requestedEnd))}</div>
            </div>
            <div class="rpt-meta">
                ${input.clientName ? `<div><span>Prepared for</span><strong>${escapeHtml(input.clientName)}</strong></div>` : ""}
                ${input.adviserName ? `<div><span>Prepared by</span><strong>${escapeHtml(input.adviserName)}</strong>${input.adviserContact ? `<em>${escapeHtml(input.adviserContact)}</em>` : ""}</div>` : ""}
                <div><span>Report date</span><strong>${escapeHtml(today)}</strong></div>
            </div>
        </header>

        ${result.priceDateBeforeEnd ? `
            <section class="rpt-note">
                <strong>Price date:</strong> the latest BID price available is for
                ${escapeHtml(formatDate(result.endDate))}, so the portfolio is valued at that date
                (end date requested: ${escapeHtml(formatDate(result.requestedEnd))}).
                ${result.skippedCount ? `${result.skippedCount} top-up${result.skippedCount === 1 ? "" : "s"} totalling ${m(result.skippedAmount)} scheduled after ${escapeHtml(formatDate(result.endDate))} ${result.skippedCount === 1 ? "is" : "are"} not included, as no price is available yet.` : ""}
            </section>
        ` : ""}

        <section class="rpt-plan">
            <strong>Investment plan:</strong> ${escapeHtml(planText)}, across ${perFund.length} fund${perFund.length === 1 ? "" : "s"}
            (${perFund.map(fund => `${escapeHtml(fund.fund?.fundCode ?? "")} ${fund.weight}%`).join(", ")}).
        </section>

        <section class="rpt-tiles">
            <div class="rpt-tile">
                <span>Total invested</span>
                <strong>${m(result.totalInvested)}</strong>
                <em>${result.contributions} investment${result.contributions === 1 ? "" : "s"}</em>
            </div>
            <div class="rpt-tile">
                <span>Value on ${escapeHtml(formatDate(result.endDate))}</span>
                <strong>${m(result.totalValue)}</strong>
                <em>${result.priceDateBeforeEnd ? "latest available BID price" : "at BID price"}</em>
            </div>
            <div class="rpt-tile ${tone(result.gain)}">
                <span>Gain / loss</span>
                <strong>${m(result.gain)}</strong>
                <em>${percent(result.gainPct)} on amount invested</em>
            </div>
            <div class="rpt-tile ${tone(result.annualised ?? 0)}">
                <span>Annualised return</span>
                <strong>${result.annualised === null ? "—" : percent(result.annualised)}</strong>
                <em>${result.annualised === null ? "shown for periods of 1 year or more" : "per year, money-weighted"}</em>
            </div>
        </section>

        <section class="rpt-section">
            <h3>Growth of the investment</h3>
            <div class="rpt-chart-wrap">
                <canvas id="report-chart" aria-label="Portfolio value compared with the amount invested"></canvas>
            </div>
            <div class="rpt-legend">
                <span><i style="background:#1f5f99"></i>Portfolio value</span>
                <span><i class="rpt-legend-dash"></i>Total invested</span>
            </div>
        </section>

        <section class="rpt-section">
            <h3>Portfolio breakdown</h3>
            <table class="rpt-table">
                <thead>
                    <tr>
                        <th>Fund</th>
                        <th>Allocation</th>
                        <th>Invested</th>
                        <th>Units held</th>
                        <th>Average cost</th>
                        <th>BID on ${escapeHtml(formatDate(result.endDate))}</th>
                        <th>Value</th>
                        <th>Gain / loss</th>
                    </tr>
                </thead>
                <tbody>
                    ${perFund.map(fund => `
                        <tr>
                            <td>
                                <i class="rpt-swatch" style="background:${PRINT_COLORS[fund.slot % PRINT_COLORS.length]}"></i>
                                <strong>${escapeHtml(fund.fund?.fundName ?? fund.id)}</strong>
                                <div class="rpt-muted">${escapeHtml(fund.fund?.fundCode ?? "")}</div>
                            </td>
                            <td>${fund.weight}%</td>
                            <td>${m(fund.invested)}</td>
                            <td>${fund.units.toLocaleString("en-SG", { minimumFractionDigits: 4, maximumFractionDigits: 4 })}</td>
                            <td>${fund.averageCost === null ? "—" : fund.averageCost.toFixed(5)}</td>
                            <td>${fund.endBid === null ? "—" : fund.endBid.toFixed(5)}</td>
                            <td>${m(fund.value)}</td>
                            <td class="${tone(fund.gain)}">${m(fund.gain)}<div class="rpt-muted">${percent(fund.gainPct)}</div></td>
                        </tr>
                    `).join("")}
                </tbody>
                <tfoot>
                    <tr>
                        <td>Total</td>
                        <td>100%</td>
                        <td>${m(result.totalInvested)}</td>
                        <td></td>
                        <td></td>
                        <td></td>
                        <td>${m(result.totalValue)}</td>
                        <td class="${tone(result.gain)}">${m(result.gain)}<div class="rpt-muted">${percent(result.gainPct)}</div></td>
                    </tr>
                </tfoot>
            </table>
        </section>

        <section class="rpt-section">
            <h3>Year by year</h3>
            <table class="rpt-table rpt-table-years">
                <thead>
                    <tr>
                        <th>As at</th>
                        <th>Total invested</th>
                        <th>Portfolio value</th>
                        <th>Gain / loss</th>
                        <th>Return on invested</th>
                    </tr>
                </thead>
                <tbody>
                    ${result.yearRows.map(row => `
                        <tr>
                            <td>${escapeHtml(row.label)}</td>
                            <td>${m(row.invested)}</td>
                            <td>${m(row.value)}</td>
                            <td class="${tone(row.gain)}">${m(row.gain)}</td>
                            <td class="${tone(row.gain)}">${percent(row.gainPct)}</td>
                        </tr>
                    `).join("")}
                </tbody>
            </table>
        </section>

        <section class="rpt-section rpt-funds">
            <h3>About the funds</h3>
            ${perFund.map(fund => {
                const details = fund.fund?.fund ?? {};

                return `
                    <div class="rpt-fund">
                        <div class="rpt-fund-head">
                            <i class="rpt-swatch" style="background:${PRINT_COLORS[fund.slot % PRINT_COLORS.length]}"></i>
                            <strong>${escapeHtml(fund.fund?.fundName ?? fund.id)}</strong>
                            <span class="rpt-muted">${escapeHtml(fund.fund?.fundCode ?? "")}</span>
                        </div>
                        <div class="rpt-fund-facts">
                            <span><b>Asset class:</b> ${escapeHtml(details.assetClass ?? "—")}${details.assetSubClass ? ` (${escapeHtml(details.assetSubClass)})` : ""}</span>
                            <span><b>Risk:</b> ${escapeHtml(details.riskClassification ?? "—")}</span>
                            <span><b>Currency:</b> ${escapeHtml(currencyOf(fund.id))}</span>
                            <span><b>First purchase:</b> ${escapeHtml(formatDate(fund.startDate))} at BID ${fund.startBid === null ? "—" : fund.startBid.toFixed(5)}</span>
                        </div>
                        ${details.fundObjective ? `<p>${escapeHtml(String(details.fundObjective).replace(/\s+/g, " ").trim())}</p>` : ""}
                    </div>
                `;
            }).join("")}
        </section>

        <footer class="rpt-disclaimer">
            <strong>Important notes.</strong>
            This is a historical illustration based on actual published BID prices of the funds between the dates shown.
            Units are assumed to be bought at the BID price on the first pricing day on or after each investment date, with no
            sales charge, fees or taxes deducted. Dividends paid out by distribution funds are not included, so the
            figures for those funds may understate their total return. Past performance is not necessarily indicative of
            future performance. The value of units and the income from them may fall as well as rise. This report is for
            illustration only and does not constitute financial advice or an offer to buy or sell any investment.
            Please refer to the fund's prospectus and product highlights sheet before investing.
            <div class="rpt-source">Source: VGrat FMS, Prudential Singapore fund prices. Generated ${escapeHtml(today)}.</div>
        </footer>
    `;

    wrap.hidden = false;

    renderChart(result);

    requestAnimationFrame(() => wrap.scrollIntoView({ behavior: "smooth", block: "start" }));
}

function renderChart(result) {
    const canvas = qs("#report-chart");
    const ChartConstructor = window.Chart;

    report.chart?.destroy?.();
    report.chart = null;

    if (!canvas || !ChartConstructor) return;

    const currency = currencyOf(result.perFund[0].id);
    const font = Math.round(11 * getFontScale());

    report.chart = new ChartConstructor(canvas, {
        type: "line",
        data: {
            labels: result.series.map(point => point.date),
            datasets: [
                {
                    label: "Portfolio value",
                    data: result.series.map(point => point.value),
                    borderColor: "#1f5f99",
                    backgroundColor: "rgba(31, 95, 153, 0.10)",
                    fill: true,
                    borderWidth: 2,
                    pointRadius: 0,
                    tension: 0
                },
                {
                    label: "Total invested",
                    data: result.series.map(point => point.invested),
                    borderColor: "#7a8590",
                    borderDash: [5, 4],
                    borderWidth: 1.5,
                    pointRadius: 0,
                    stepped: true,
                    fill: false
                }
            ]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            animation: false,
            interaction: { mode: "index", intersect: false },
            plugins: {
                legend: { display: false },
                tooltip: {
                    backgroundColor: "#ffffff",
                    titleColor: "#1b2530",
                    bodyColor: "#1b2530",
                    borderColor: "#cfd6dd",
                    borderWidth: 1,
                    callbacks: {
                        title: items => formatDate(items[0]?.label),
                        label: item => `${item.dataset.label}: ${money(item.parsed.y, currency)}`
                    }
                }
            },
            scales: {
                x: {
                    type: "category",
                    ticks: {
                        color: "#5b6670",
                        font: { size: font },
                        maxTicksLimit: 8,
                        autoSkip: true,
                        callback(value) {
                            const label = this.getLabelForValue(value);

                            return formatDate(label).slice(3);
                        }
                    },
                    grid: { color: "#eef1f4" }
                },
                y: {
                    ticks: {
                        color: "#5b6670",
                        font: { size: font },
                        callback: value => money(value, currency).replace(/\.00$/, "")
                    },
                    grid: { color: "#eef1f4" }
                }
            }
        }
    });
}


/* ============================================================
   EVENTS
   ============================================================ */

function bindEvents() {
    const search = qs("#report-fund-search");
    const list = qs("#report-fund-suggestions");

    search?.addEventListener("focus", renderSuggestions);
    search?.addEventListener("input", renderSuggestions);
    search?.addEventListener("blur", () => setTimeout(closeSuggestions, 150));

    search?.addEventListener("keydown", event => {
        if (event.key === "Enter") {
            event.preventDefault();

            if (!search.value.trim()) return;

            const first = qs("#report-fund-suggestions [data-report-add]");

            if (first) {
                addFund(first.dataset.reportAdd);
                search.value = "";
                renderSuggestions();
            }
        } else if (event.key === "Escape") {
            closeSuggestions();
            search.blur();
        }
    });

    list?.addEventListener("mousedown", event => {
        event.preventDefault();

        const item = event.target.closest("[data-report-add]");

        if (!item) return;

        addFund(item.dataset.reportAdd);
        search.value = "";
        renderSuggestions();
    });

    const body = qs("#report-funds-body");

    body?.addEventListener("input", event => {
        const field = event.target.closest("[data-report-weight]");

        if (!field) return;

        report.selected[Number(field.dataset.reportWeight)].weight = Math.max(0, Number(field.value) || 0);

        const sum = report.selected.reduce((acc, item) => acc + (Number(item.weight) || 0), 0);
        const total = qs("#report-allocation-total");

        if (total) {
            total.textContent = `Total ${Math.round(sum * 100) / 100}%`;
            total.classList.toggle("is-ok", Math.abs(sum - 100) < 0.01);
            total.classList.toggle("is-off", Math.abs(sum - 100) >= 0.01);
        }
    });

    body?.addEventListener("click", event => {
        const remove = event.target.closest("[data-report-remove]");

        if (!remove) return;

        report.selected.splice(Number(remove.dataset.reportRemove), 1);
        renderSelected();
    });

    qs("#report-split")?.addEventListener("click", splitEqually);

    qs("#report-topup")?.addEventListener("input", () => {
        const frequency = qs("#report-frequency");
        const topUp = Number(qs("#report-topup").value) || 0;

        if (frequency && topUp > 0 && frequency.value === "none") frequency.value = "monthly";
    });

    qs("#report-form")?.addEventListener("submit", event => {
        event.preventDefault();

        const error = qs("#report-error");
        const input = readInput();
        const problem = validate(input);

        if (error) {
            error.textContent = problem ?? "";
            error.hidden = !problem;
        }

        if (problem) return;

        const result = calculate(input);

        if (!result.series.length) {
            if (error) {
                error.textContent = "No prices were found in this period for the chosen funds.";
                error.hidden = false;
            }

            return;
        }

        report.lastResult = result;
        renderReport(result);
    });

    qs("#report-print")?.addEventListener("click", () => {
        document.body.classList.add("is-printing-report");

        // Let Chart.js resize for the page before printing
        requestAnimationFrame(() => {
            window.print();
        });
    });

    window.addEventListener("afterprint", () => {
        document.body.classList.remove("is-printing-report");
        report.chart?.resize?.();
    });
}


/* ============================================================
   PUBLIC API
   ============================================================ */

function initializeReports({ funds, historyIndex }) {
    report.funds = [...(funds ?? [])].sort((a, b) =>
        String(a.fundName ?? "").localeCompare(String(b.fundName ?? ""))
    );

    report.fundById = new Map(report.funds.map(fund => [fundId(fund), fund]));
    report.index = historyIndex ?? new Map();

    if (!report.initialized) {
        bindEvents();
        report.initialized = true;
    }

    const loading = qs("#report-loading");

    if (loading) loading.hidden = true;

    const form = qs("#report-form");

    if (form) form.hidden = false;

    renderSelected();
}


export {
    initializeReports,
    calculate as calculateInvestmentGrowth
};
