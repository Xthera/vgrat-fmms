
#!/usr/bin/env python3

"""
PruAccess historical BID price extractor.

Test fund:
    Funds Links.xlsm
        Column A = Prudential fund URL
        Column B = PruAccess fund name

Workflow:
    Excel
       ↓
    Prudential fund page/API
       ↓
    Get official inception date
       ↓
    PruAccess
       ↓
    Exact fund-name match
       ↓
    Table view
       ↓
    Start = inception date
    End = today
       ↓
    Direct pagination:
        page.page=1
        page.page=2
        page.page=3
        ...
       ↓
    Extract BID prices only

Important:
- No synthetic data.
- No estimated data.
- No interpolation.
- No fuzzy fund matching.
- Fund Price Type is NOT changed.
- Historical BID prices come directly from PruAccess.
- Offer prices are not stored in bid_history.json.
"""

from __future__ import annotations

import json
import re
from datetime import datetime
from pathlib import Path
from urllib.parse import urljoin

from openpyxl import load_workbook
from playwright.sync_api import sync_playwright


# ============================================================
# CONFIG
# ============================================================

EXCEL_FILE = Path("Funds Links.xlsm")

OUTPUT_DIR = Path(
    "output_pruaccess"
)

PRUACCESS_URL = (
    "https://pruaccess.prudential.com.sg/"
    "prulinkfund/viewFundPerformance.do"
)

PAGE_SIZE = 20

MAX_PAGES = 1000

PAGE_WAIT_MS = 500


# ============================================================
# HELPERS
# ============================================================

def clean_text(value) -> str:

    if value is None:
        return ""

    return re.sub(
        r"\s+",
        " ",
        str(value),
    ).strip()


def normalize_name(value: str) -> str:

    value = clean_text(
        value
    ).lower()

    value = value.replace(
        "–",
        "-"
    )

    value = value.replace(
        "—",
        "-"
    )

    value = re.sub(
        r"\s+",
        " ",
        value,
    )

    return value.strip()


def save_json(
    path: Path,
    data,
) -> None:

    path.write_text(
        json.dumps(
            data,
            indent=2,
            ensure_ascii=False,
        ),
        encoding="utf-8",
    )


def parse_prudential_date(
    value: str,
) -> datetime:

    return datetime.strptime(
        value.strip(),
        "%d/%m/%Y",
    )


def pruaccess_date(
    value: datetime,
) -> str:

    return value.strftime(
        "%d-%b-%Y"
    )


# ============================================================
# EXCEL
# ============================================================

def read_excel_fund():

    print(
        "\n=== Reading Funds Links.xlsm ==="
    )

    if not EXCEL_FILE.exists():

        raise FileNotFoundError(
            f"Excel file not found: "
            f"{EXCEL_FILE}"
        )

    workbook = load_workbook(
        EXCEL_FILE,
        read_only=True,
        keep_vba=True,
        data_only=True,
    )

    worksheet = workbook.active

    prudential_url = clean_text(
        worksheet["A2"].value
    )

    pruaccess_name = clean_text(
        worksheet["B2"].value
    )

    workbook.close()

    if not prudential_url:

        raise RuntimeError(
            "Funds Links.xlsm A2 is empty."
        )

    if not pruaccess_name:

        raise RuntimeError(
            "Funds Links.xlsm B2 is empty."
        )

    print(
        f"Excel A2: {prudential_url}"
    )

    print(
        f"Excel B2: {pruaccess_name}"
    )

    return (
        prudential_url,
        pruaccess_name,
    )


# ============================================================
# PRUDENTIAL API
# ============================================================

def get_prudential_fund_data(
    page,
    prudential_url: str,
) -> dict:

    print(
        "\n=== Opening Prudential fund page ==="
    )

    print(
        prudential_url
    )

    captured_urls = []

    def handle_response(
        response,
    ):

        if (
            "ilpfunds.json"
            in response.url
        ):

            captured_urls.append(
                response.url
            )

    page.on(
        "response",
        handle_response,
    )

    page.goto(
        prudential_url,
        wait_until="domcontentloaded",
        timeout=120000,
    )

    page.wait_for_timeout(
        3000
    )

    if not captured_urls:

        raise RuntimeError(
            "Could not capture "
            "Prudential ilpfunds.json."
        )

    api_url = captured_urls[0]

    print(
        "\nPrudential API:"
    )

    print(
        api_url
    )

    response = page.request.get(
        api_url,
        timeout=120000,
    )

    if not response.ok:

        raise RuntimeError(
            "Prudential API returned "
            f"HTTP {response.status}"
        )

    data = response.json()

    if (
        not isinstance(data, list)
        or not data
    ):

        raise RuntimeError(
            "Unexpected Prudential "
            "API response."
        )

    fund = data[0]

    result = {
        "apiUrl": api_url,

        "fundIdentifier":
            fund.get(
                "fundIdentifier"
            ),

        "fundName":
            fund.get(
                "fundName"
            ),

        "fundCode":
            fund.get(
                "fundCode"
            ),

        "fundCurrency":
            fund.get(
                "fundCurrency"
            ),

        "unitCurrency":
            fund.get(
                "unitCurrency"
            ),

        "assetClass":
            fund.get(
                "assetClass"
            ),

        "assetSubClass":
            fund.get(
                "assetSubClass"
            ),

        "riskClassification":
            fund.get(
                "riskClassification"
            ),

        "bidPrice":
            fund.get(
                "bidPrice"
            ),

        "offerPrice":
            fund.get(
                "offerPrice"
            ),

        "valuationDate":
            fund.get(
                "valuationDate"
            ),

        "inceptionDate":
            fund.get(
                "inceptionDate"
            ),

        "cumulativeYtd":
            fund.get(
                "cumulativeYtd"
            ),

        "cumulative1m":
            fund.get(
                "cumulative1m"
            ),

        "cumulative3m":
            fund.get(
                "cumulative3m"
            ),

        "cumulative6m":
            fund.get(
                "cumulative6m"
            ),

        "cumulative1y":
            fund.get(
                "cumulative1y"
            ),

        "cumulative3y":
            fund.get(
                "cumulative3y"
            ),

        "cumulative5y":
            fund.get(
                "cumulative5y"
            ),

        "annualised3y":
            fund.get(
                "annualised3y"
            ),

        "annualised5y":
            fund.get(
                "annualised5y"
            ),

        "annualised10y":
            fund.get(
                "annualised10y"
            ),

        "annualisedSinceLaunch":
            fund.get(
                "annualisedSinceLaunch"
            ),

        "factsheetUrl":
            fund.get(
                "factsheetUrl"
            ),

        "fundObjective":
            fund.get(
                "fundObjective"
            ),

        "investmentManager":
            fund.get(
                "investmentManager"
            ),

        "hasDividend":
            fund.get(
                "hasDividend"
            ),

        "dividendRate":
            fund.get(
                "dividendRate"
            ),

        "raw": fund,
    }

    return result


# ============================================================
# PRUACCESS FUND OPTIONS
# ============================================================

def get_fund_options(
    page,
) -> list[dict]:

    options = page.locator(
        "#fundName option"
    )

    count = options.count()

    print(
        f"\nPruAccess fund options found: "
        f"{count}"
    )

    result = []

    for index in range(count):

        option = options.nth(
            index
        )

        result.append(
            {
                "text": clean_text(
                    option.inner_text()
                ),
                "value":
                    option.get_attribute(
                        "value"
                    ),
            }
        )

    return result


# ============================================================
# READONLY DATE INPUT
# ============================================================

def set_readonly_input_value(
    page,
    selector: str,
    value: str,
) -> None:

    locator = page.locator(
        selector
    )

    if locator.count() == 0:

        raise RuntimeError(
            f"Could not find input: "
            f"{selector}"
        )

    locator.evaluate(
        """
        (element, value) => {
            element.value = value;

            element.dispatchEvent(
                new Event("input", {
                    bubbles: true
                })
            );

            element.dispatchEvent(
                new Event("change", {
                    bubbles: true
                })
            );
        }
        """,
        value,
    )

    actual = locator.input_value()

    print(
        f"{selector}: {actual}"
    )

    if actual != value:

        raise RuntimeError(
            f"Could not set {selector}. "
            f"Expected {value}, "
            f"got {actual}."
        )


# ============================================================
# PRICE TABLE EXTRACTION
# ============================================================

def extract_price_rows(
    page,
) -> list[dict]:

    """
    Extract rows from the current PruAccess
    price table.

    Expected:

        Date
        Bid Price
        Offer Price

    We intentionally store only:

        Date
        Bid Price
    """

    tables = page.locator(
        "table"
    )

    table_count = tables.count()

    all_rows = []

    for table_index in range(
        table_count
    ):

        table = tables.nth(
            table_index
        )

        rows = table.locator(
            "tr"
        )

        row_count = rows.count()

        for row_index in range(
            row_count
        ):

            row = rows.nth(
                row_index
            )

            cells = row.locator(
                "th, td"
            )

            cell_count = cells.count()

            if cell_count < 2:
                continue

            values = []

            for cell_index in range(
                cell_count
            ):

                values.append(
                    clean_text(
                        cells.nth(
                            cell_index
                        ).inner_text()
                    )
                )

            if len(values) < 2:
                continue

            date_value = values[0]

            if not re.match(
                r"^\d{1,2}-[A-Za-z]{3}-\d{4}$",
                date_value,
            ):
                continue

            bid_value = values[1]

            if not re.match(
                r"^-?\d+(?:\.\d+)?$",
                bid_value,
            ):
                continue

            all_rows.append(
                {
                    "date": date_value,
                    "bidPrice": bid_value,
                }
            )

    # Remove duplicates while preserving order.
    seen = set()

    result = []

    for row in all_rows:

        key = (
            row["date"],
            row["bidPrice"],
        )

        if key in seen:
            continue

        seen.add(key)

        result.append(
            row
        )

    return result


# ============================================================
# PAGINATION URL
# ============================================================

def build_page_url(
    base_url: str,
    fund_id: str,
    view_type: str,
    start_date: str,
    end_date: str,
    page_number: int,
) -> str:

    return (
        f"{base_url}"
        f"?fundId={fund_id}"
        f"&viewType={view_type}"
        f"&startDate={start_date}"
        f"&endDate={end_date}"
        f"&page.page={page_number}"
        f"&page.size={PAGE_SIZE}"
    )


# ============================================================
# FULL PAGINATION EXTRACTION
# ============================================================

def extract_all_pages(
    page,
    fund_id: str,
    start_date: str,
    end_date: str,
) -> tuple[list[dict], list[dict]]:

    """
    Directly request page 1, page 2, page 3, etc.

    We do NOT click pagination controls.

    PruAccess has already shown that its pagination
    URLs use:

        page.page=N
        page.size=20
    """

    print(
        "\n============================================================"
    )

    print(
        "Starting direct PruAccess pagination"
    )

    print(
        "============================================================"
    )

    all_rows = {}

    page_diagnostics = []

    for page_number in range(
        1,
        MAX_PAGES + 1,
    ):

        url = build_page_url(
            PRUACCESS_URL,
            fund_id,
            "TBL",
            start_date,
            end_date,
            page_number,
        )

        print(
            f"\nFetching page "
            f"{page_number}:"
        )

        print(
            url
        )

        response = page.goto(
            url,
            wait_until="domcontentloaded",
            timeout=120000,
        )

        page.wait_for_timeout(
            PAGE_WAIT_MS
        )

        current_rows = extract_price_rows(
            page
        )

        print(
            f"Rows found: "
            f"{len(current_rows)}"
        )

        # ----------------------------------------------------
        # If there are no rows, pagination is finished.
        # ----------------------------------------------------

        if not current_rows:

            print(
                "No rows found. "
                "End of pagination."
            )

            page_diagnostics.append(
                {
                    "page": page_number,
                    "url": url,
                    "rowCount": 0,
                    "status": "empty",
                }
            )

            break

        # ----------------------------------------------------
        # Record rows.
        # ----------------------------------------------------

        for row in current_rows:

            key = (
                row["date"],
                row["bidPrice"],
            )

            all_rows[key] = row

        page_diagnostics.append(
            {
                "page": page_number,
                "url": url,
                "rowCount": len(
                    current_rows
                ),
                "firstDate":
                    current_rows[0][
                        "date"
                    ],
                "lastDate":
                    current_rows[-1][
                        "date"
                    ],
                "status": "success",
            }
        )

        print(
            f"Date range on page: "
            f"{current_rows[0]['date']} "
            f"→ "
            f"{current_rows[-1]['date']}"
        )

        # ----------------------------------------------------
        # If fewer than PAGE_SIZE rows are returned,
        # this is the final page.
        # ----------------------------------------------------

        if len(current_rows) < PAGE_SIZE:

            print(
                "\nFinal page detected."
            )

            break

    # --------------------------------------------------------
    # Sort newest → oldest.
    # --------------------------------------------------------

    def sort_date(row):

        try:

            return datetime.strptime(
                row["date"],
                "%d-%b-%Y",
            )

        except Exception:

            return datetime.min

    final_rows = sorted(
        all_rows.values(),
        key=sort_date,
        reverse=True,
    )

    return (
        final_rows,
        page_diagnostics,
    )


# ============================================================
# MAIN
# ============================================================

def main():

    OUTPUT_DIR.mkdir(
        parents=True,
        exist_ok=True,
    )

    print(
        "============================================================"
    )

    print(
        "PruAccess Historical BID Price Extraction"
    )

    print(
        "============================================================"
    )

    # ========================================================
    # READ EXCEL
    # ========================================================

    (
        prudential_url,
        excel_pruaccess_name,
    ) = read_excel_fund()

    # ========================================================
    # PLAYWRIGHT
    # ========================================================

    with sync_playwright() as playwright:

        browser = playwright.chromium.launch(
            headless=True
        )

        context = browser.new_context(
            viewport={
                "width": 1440,
                "height": 1000,
            }
        )

        page = context.new_page()

        # ====================================================
        # PRUDENTIAL
        # ====================================================

        prudential = (
            get_prudential_fund_data(
                page,
                prudential_url,
            )
        )

        save_json(
            OUTPUT_DIR
            / "prudential_fund.json",
            prudential,
        )

        print(
            "\n=== Prudential fund ==="
        )

        print(
            prudential.get(
                "fundName"
            )
        )

        print(
            f"Citicode / identifier: "
            f"{prudential.get('fundIdentifier')}"
        )

        print(
            f"Fund code: "
            f"{prudential.get('fundCode')}"
        )

        print(
            f"Current bid: "
            f"{prudential.get('bidPrice')}"
        )

        print(
            f"Current offer: "
            f"{prudential.get('offerPrice')}"
        )

        print(
            f"Inception: "
            f"{prudential.get('inceptionDate')}"
        )

        # ====================================================
        # DATE RANGE
        # ====================================================

        inception_raw = prudential.get(
            "inceptionDate"
        )

        if not inception_raw:

            raise RuntimeError(
                "Prudential inception date "
                "is missing."
            )

        inception_date = (
            parse_prudential_date(
                inception_raw
            )
        )

        start_date = pruaccess_date(
            inception_date
        )

        end_date = datetime.now().strftime(
            "%d-%b-%Y"
        )

        print(
            "\nPruAccess start date:"
        )

        print(
            start_date
        )

        print(
            "\nPruAccess end date:"
        )

        print(
            end_date
        )

        # ====================================================
        # OPEN PRUACCESS
        # ====================================================

        print(
            "\n=== Opening PruAccess ==="
        )

        page.goto(
            PRUACCESS_URL,
            wait_until="domcontentloaded",
            timeout=120000,
        )

        page.wait_for_timeout(
            1500
        )

        (
            OUTPUT_DIR
            / "pre_submit_visible_text.txt"
        ).write_text(
            page.locator(
                "body"
            ).inner_text(),
            encoding="utf-8",
        )

        # ====================================================
        # FUND OPTIONS
        # ====================================================

        options = get_fund_options(
            page
        )

        save_json(
            OUTPUT_DIR
            / "fund_options.json",
            options,
        )

        target_name = normalize_name(
            excel_pruaccess_name
        )

        matches = []

        for option in options:

            if normalize_name(
                option["text"]
            ) == target_name:

                matches.append(
                    option
                )

        if not matches:

            raise RuntimeError(
                "Exact PruAccess fund-name "
                "match failed."
            )

        if len(matches) > 1:

            raise RuntimeError(
                "Multiple exact PruAccess "
                "fund-name matches found."
            )

        selected = matches[0]

        print(
            "\nMatched PruAccess fund:"
        )

        print(
            selected
        )

        fund_id = selected[
            "value"
        ]

        # ====================================================
        # SELECT FUND
        # ====================================================

        page.locator(
            "#fundName"
        ).select_option(
            fund_id
        )

        # ====================================================
        # TABLE VIEW
        # ====================================================

        page.locator(
            "#viewType"
        ).select_option(
            "TBL"
        )

        # ====================================================
        # FUND PRICE TYPE
        #
        # DO NOT MODIFY.
        # ====================================================

        fund_price_type = page.locator(
            "#fundPriceType"
        )

        if fund_price_type.count():

            fund_price_type_value = (
                fund_price_type.input_value()
            )

        else:

            fund_price_type_value = None

        print(
            "\nFund Price Type:"
        )

        print(
            f"{fund_price_type_value} "
            "(left untouched)"
        )

        # ====================================================
        # SET DATES
        # ====================================================

        print(
            "\n=== Setting PruAccess dates ==="
        )

        set_readonly_input_value(
            page,
            'input[name="startDate"]',
            start_date,
        )

        set_readonly_input_value(
            page,
            'input[name="endDate"]',
            end_date,
        )

        # ====================================================
        # CSRF
        # ====================================================

        csrf = page.locator(
            'input[name="_csrf"]'
        )

        if csrf.count() == 0:

            raise RuntimeError(
                "PruAccess CSRF input not found."
            )

        csrf_value = csrf.input_value()

        if not csrf_value:

            raise RuntimeError(
                "PruAccess CSRF token is empty."
            )

        # ====================================================
        # VERIFY
        # ====================================================

        actual_start = page.locator(
            'input[name="startDate"]'
        ).input_value()

        actual_end = page.locator(
            'input[name="endDate"]'
        ).input_value()

        print(
            "\nVerified dates:"
        )

        print(
            f"Start: {actual_start}"
        )

        print(
            f"End:   {actual_end}"
        )

        if actual_start != start_date:

            raise RuntimeError(
                "Start date verification failed."
            )

        if actual_end != end_date:

            raise RuntimeError(
                "End date verification failed."
            )

        # ====================================================
        # PRE-SUBMIT
        # ====================================================

        pre_submit = {
            "prudentialUrl":
                prudential_url,

            "excelPruAccessName":
                excel_pruaccess_name,

            "matchedPruAccessName":
                selected["text"],

            "matchedPruAccessValue":
                selected["value"],

            "startDate":
                start_date,

            "endDate":
                end_date,

            "viewType":
                "TBL",

            "fundPriceType":
                fund_price_type_value,

            "csrfPresent":
                bool(csrf_value),
        }

        save_json(
            OUTPUT_DIR
            / "pre_submit.json",
            pre_submit,
        )

        # ====================================================
        # SUBMIT
        # ====================================================

        print(
            "\n=== Submitting PruAccess form ==="
        )

        page.locator(
            "#fundForm"
        ).evaluate(
            """
            form => {
                form.submit();
            }
            """
        )

        page.wait_for_load_state(
            "domcontentloaded",
            timeout=120000,
        )

        page.wait_for_timeout(
            2500
        )

        # ====================================================
        # SAVE FIRST RESULT
        # ====================================================

        (
            OUTPUT_DIR
            / "result.html"
        ).write_text(
            page.content(),
            encoding="utf-8",
        )

        (
            OUTPUT_DIR
            / "result_visible_text.txt"
        ).write_text(
            page.locator(
                "body"
            ).inner_text(),
            encoding="utf-8",
        )

        print(
            "\nReturned URL:"
        )

        print(
            page.url
        )

        # ====================================================
        # EXTRACT ALL PAGES
        # ====================================================

        (
            historical_rows,
            page_diagnostics,
        ) = extract_all_pages(
            page,
            fund_id,
            start_date,
            end_date,
        )

        # ====================================================
        # SAVE BID HISTORY
        # ====================================================

        bid_history = {
            "source":
                "PruAccess",

            "fundName":
                prudential.get(
                    "fundName"
                ),

            "pruAccessFundName":
                selected["text"],

            "fundId":
                fund_id,

            "fundIdentifier":
                prudential.get(
                    "fundIdentifier"
                ),

            "fundCode":
                prudential.get(
                    "fundCode"
                ),

            "currency":
                prudential.get(
                    "fundCurrency"
                ),

            "startDate":
                start_date,

            "endDate":
                end_date,

            "priceType":
                "BID",

            "pageSize":
                PAGE_SIZE,

            "observationCount":
                len(
                    historical_rows
                ),

            "observations":
                historical_rows,
        }

        save_json(
            OUTPUT_DIR
            / "bid_history.json",
            bid_history,
        )

        # ====================================================
        # SAVE PAGINATION
        # ====================================================

        save_json(
            OUTPUT_DIR
            / "pagination.json",
            page_diagnostics,
        )

        # ====================================================
        # SUMMARY
        # ====================================================

        summary = {
            "status":
                "success",

            "prudentialFundName":
                prudential.get(
                    "fundName"
                ),

            "excelPruAccessName":
                excel_pruaccess_name,

            "matchedPruAccessName":
                selected["text"],

            "pruAccessFundId":
                fund_id,

            "fundIdentifier":
                prudential.get(
                    "fundIdentifier"
                ),

            "fundCode":
                prudential.get(
                    "fundCode"
                ),

            "inceptionDate":
                inception_raw,

            "startDate":
                start_date,

            "endDate":
                end_date,

            "currentBidPrice":
                prudential.get(
                    "bidPrice"
                ),

            "currentOfferPrice":
                prudential.get(
                    "offerPrice"
                ),

            "fundPriceType":
                fund_price_type_value,

            "historicalPriceType":
                "BID",

            "pageSize":
                PAGE_SIZE,

            "pagesExtracted":
                len(
                    page_diagnostics
                ),

            "historicalObservationCount":
                len(
                    historical_rows
                ),

            "newestObservation":
                (
                    historical_rows[0]
                    if historical_rows
                    else None
                ),

            "oldestObservation":
                (
                    historical_rows[-1]
                    if historical_rows
                    else None
                ),
        }

        save_json(
            OUTPUT_DIR
            / "summary.json",
            summary,
        )

        # ====================================================
        # CONSOLE SUMMARY
        # ====================================================

        print(
            "\n============================================================"
        )

        print(
            "EXTRACTION COMPLETE"
        )

        print(
            "============================================================"
        )

        print(
            f"Pages extracted: "
            f"{len(page_diagnostics)}"
        )

        print(
            f"Historical BID observations: "
            f"{len(historical_rows)}"
        )

        if historical_rows:

            print(
                "\nNewest BID:"
            )

            print(
                historical_rows[0]
            )

            print(
                "\nOldest BID:"
            )

            print(
                historical_rows[-1]
            )

        browser.close()

    # ========================================================
    # FILES
    # ========================================================

    print(
        "\nFiles created:"
    )

    for file in sorted(
        OUTPUT_DIR.iterdir()
    ):

        print(
            f" - {file}"
        )

    print(
        "\nDone."
    )


if __name__ == "__main__":
    main()

