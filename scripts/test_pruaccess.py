#!/usr/bin/env python3

"""
PruAccess historical BID price extractor.

MASTER SOURCE:
    Funds Links.xlsm

Excel:
    Column A = Prudential fund URL
    Column B = PruAccess fund name

Workflow:

    Funds Links.xlsm
          ↓
    Read every populated Excel row
          ↓
    Prudential fund page/API
          ↓
    Get official fund information
          ↓
    Get official inception date
          ↓
    PruAccess
          ↓
    Exact fund-name match
          ↓
    Table view
          ↓
    Start = Prudential inception date
    End   = request date
          ↓
    Direct pagination
          ↓
    Extract BID prices only
          ↓
    Repeat for every Excel fund
          ↓
    Consolidated JSON

Important:
- Funds Links.xlsm determines the fund universe.
- Column A determines which funds exist.
- Column B provides the PruAccess fund name.
- No hardcoded fund count.
- No synthetic data.
- No estimated data.
- No interpolation.
- No fuzzy fund matching.
- Fund Price Type is NOT changed.
- Historical BID prices come directly from PruAccess.
- Offer prices are not stored as historical BID prices.
- If a fund fails, the failure is recorded.
- A failed fund is NOT replaced with fake or previous data.
"""

from __future__ import annotations

import json
import re
from datetime import datetime
from pathlib import Path

from openpyxl import load_workbook
from playwright.sync_api import sync_playwright


# ============================================================
# CONFIG
# ============================================================

EXCEL_FILE = Path("Funds Links.xlsm")

OUTPUT_DIR = Path(
    "output_pruaccess"
)

FUNDS_OUTPUT_DIR = (
    OUTPUT_DIR / "funds"
)

PRUACCESS_URL = (
    "https://pruaccess.prudential.com.sg/"
    "prulinkfund/viewFundPerformance.do"
)

PAGE_SIZE = 20

MAX_PAGES = 1000

PAGE_WAIT_MS = 500

PRUDENTIAL_WAIT_MS = 3000

PRUACCESS_INITIAL_WAIT_MS = 1500

PRUACCESS_RESULT_WAIT_MS = 2500


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

    path.parent.mkdir(
        parents=True,
        exist_ok=True,
    )

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


def safe_filename(
    value: str,
) -> str:

    value = clean_text(
        value
    )

    value = re.sub(
        r"[^A-Za-z0-9._-]+",
        "_",
        value,
    )

    value = value.strip(
        "._"
    )

    if not value:

        value = "fund"

    return value[:120]


# ============================================================
# EXCEL
# ============================================================

def read_excel_funds() -> list[dict]:

    print(
        "\n============================================================"
    )

    print(
        "READING FUNDS LINKS.XLSM"
    )

    print(
        "============================================================"
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

    funds = []

    for row_number in range(
        2,
        worksheet.max_row + 1,
    ):

        prudential_url = clean_text(
            worksheet.cell(
                row=row_number,
                column=1,
            ).value
        )

        pruaccess_name = clean_text(
            worksheet.cell(
                row=row_number,
                column=2,
            ).value
        )

        # ----------------------------------------------------
        # Column A is the master universe.
        # ----------------------------------------------------

        if not prudential_url:

            continue

        funds.append(
            {
                "excelRow":
                    row_number,

                "prudentialUrl":
                    prudential_url,

                "pruAccessName":
                    pruaccess_name,
            }
        )

    workbook.close()

    if not funds:

        raise RuntimeError(
            "No populated Prudential fund URLs "
            "were found in Column A."
        )

    print(
        f"\nPopulated fund URLs found: "
        f"{len(funds)}"
    )

    print(
        "\nExcel universe:"
    )

    for fund in funds:

        print(
            f"Row {fund['excelRow']}: "
            f"{fund['pruAccessName'] or '[NO PRUACCESS NAME]'}"
        )

        print(
            f"    {fund['prudentialUrl']}"
        )

    return funds


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

    try:

        page.goto(
            prudential_url,
            wait_until="domcontentloaded",
            timeout=120000,
        )

        page.wait_for_timeout(
            PRUDENTIAL_WAIT_MS
        )

    finally:

        page.remove_listener(
            "response",
            handle_response,
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

    fund_identifier = clean_text(
        fund.get(
            "fundIdentifier"
        )
    )

    fund_name = clean_text(
        fund.get(
            "fundName"
        )
    )

    inception_date = clean_text(
        fund.get(
            "inceptionDate"
        )
    )

    if not fund_identifier:

        raise RuntimeError(
            "Prudential fundIdentifier "
            "is missing."
        )

    if not fund_name:

        raise RuntimeError(
            "Prudential fundName "
            "is missing."
        )

    if not inception_date:

        raise RuntimeError(
            "Prudential inceptionDate "
            "is missing."
        )

    result = {
        "apiUrl":
            api_url,

        "fundIdentifier":
            fund_identifier,

        "fundName":
            fund_name,

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
            inception_date,

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

        "raw":
            fund,
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

    for index in range(
        count
    ):

        option = options.nth(
            index
        )

        result.append(
            {
                "text":
                    clean_text(
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
# PRUACCESS EXACT MATCH
# ============================================================

def find_exact_pruaccess_match(
    options: list[dict],
    target_name: str,
) -> dict:

    target_normalized = normalize_name(
        target_name
    )

    if not target_normalized:

        raise RuntimeError(
            "Excel PruAccess fund name "
            "is empty."
        )

    matches = []

    for option in options:

        option_name = normalize_name(
            option["text"]
        )

        if option_name == target_normalized:

            matches.append(
                option
            )

    if not matches:

        raise RuntimeError(
            "Exact PruAccess fund-name "
            "match failed for: "
            f"{target_name}"
        )

    if len(matches) > 1:

        raise RuntimeError(
            "Multiple exact PruAccess "
            "fund-name matches found for: "
            f"{target_name}"
        )

    return matches[0]


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
                    "date":
                        date_value,

                    "bidPrice":
                        bid_value,
                }
            )

    # --------------------------------------------------------
    # Remove duplicate date/price records.
    # --------------------------------------------------------

    seen = set()

    result = []

    for row in all_rows:

        key = (
            row["date"],
            row["bidPrice"],
        )

        if key in seen:

            continue

        seen.add(
            key
        )

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

        page.goto(
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

        if not current_rows:

            print(
                "No rows found. "
                "End of pagination."
            )

            page_diagnostics.append(
                {
                    "page":
                        page_number,

                    "url":
                        url,

                    "rowCount":
                        0,

                    "status":
                        "empty",
                }
            )

            break

        for row in current_rows:

            key = (
                row["date"],
                row["bidPrice"],
            )

            all_rows[key] = row

        page_diagnostics.append(
            {
                "page":
                    page_number,

                "url":
                    url,

                "rowCount":
                    len(
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

                "status":
                    "success",
            }
        )

        print(
            f"Date range on page: "
            f"{current_rows[0]['date']} "
            f"→ "
            f"{current_rows[-1]['date']}"
        )

        if len(
            current_rows
        ) < PAGE_SIZE:

            print(
                "\nFinal page detected."
            )

            break

    else:

        raise RuntimeError(
            f"Pagination exceeded "
            f"MAX_PAGES={MAX_PAGES}."
        )

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
# SINGLE FUND EXTRACTION
# ============================================================

def extract_single_fund(
    page,
    excel_fund: dict,
    pruaccess_options: list[dict],
) -> dict:

    excel_row = excel_fund[
        "excelRow"
    ]

    prudential_url = excel_fund[
        "prudentialUrl"
    ]

    excel_pruaccess_name = excel_fund[
        "pruAccessName"
    ]

    print(
        "\n\n"
        "################################################################"
    )

    print(
        f"PROCESSING EXCEL ROW {excel_row}"
    )

    print(
        "################################################################"
    )

    print(
        f"\nPrudential URL:"
    )

    print(
        prudential_url
    )

    print(
        f"\nExcel PruAccess name:"
    )

    print(
        excel_pruaccess_name
    )

    # ========================================================
    # PRUDENTIAL
    # ========================================================

    prudential = (
        get_prudential_fund_data(
            page,
            prudential_url,
        )
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

    # ========================================================
    # DATE RANGE
    # ========================================================

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

    # ========================================================
    # EXACT PRUACCESS MATCH
    # ========================================================

    selected = (
        find_exact_pruaccess_match(
            pruaccess_options,
            excel_pruaccess_name,
        )
    )

    print(
        "\nMatched PruAccess fund:"
    )

    print(
        selected
    )

    fund_id = selected[
        "value"
    ]

    if not fund_id:

        raise RuntimeError(
            "Matched PruAccess fund "
            "has no fund ID."
        )

    # ========================================================
    # OPEN PRUACCESS
    # ========================================================

    print(
        "\n=== Opening PruAccess ==="
    )

    page.goto(
        PRUACCESS_URL,
        wait_until="domcontentloaded",
        timeout=120000,
    )

    page.wait_for_timeout(
        PRUACCESS_INITIAL_WAIT_MS
    )

    # ========================================================
    # SELECT FUND
    # ========================================================

    fund_selector = page.locator(
        "#fundName"
    )

    if fund_selector.count() == 0:

        raise RuntimeError(
            "PruAccess #fundName "
            "selector not found."
        )

    fund_selector.select_option(
        fund_id
    )

    # ========================================================
    # TABLE VIEW
    # ========================================================

    view_selector = page.locator(
        "#viewType"
    )

    if view_selector.count() == 0:

        raise RuntimeError(
            "PruAccess #viewType "
            "selector not found."
        )

    view_selector.select_option(
        "TBL"
    )

    # ========================================================
    # FUND PRICE TYPE
    #
    # DO NOT MODIFY.
    # ========================================================

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

    # ========================================================
    # SET DATES
    # ========================================================

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

    # ========================================================
    # CSRF
    # ========================================================

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

    # ========================================================
    # VERIFY DATES
    # ========================================================

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

    # ========================================================
    # PRE-SUBMIT DATA
    # ========================================================

    pre_submit = {
        "excelRow":
            excel_row,

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
            bool(
                csrf_value
            ),
    }

    # ========================================================
    # SUBMIT
    # ========================================================

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
        PRUACCESS_RESULT_WAIT_MS
    )

    # ========================================================
    # EXTRACT ALL PAGES
    # ========================================================

    (
        historical_rows,
        page_diagnostics,
    ) = extract_all_pages(
        page,
        fund_id,
        start_date,
        end_date,
    )

    if not historical_rows:

        raise RuntimeError(
            "No historical BID observations "
            "were extracted."
        )

    # ========================================================
    # CREATE FUND RESULT
    # ========================================================

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

    # ========================================================
    # SUMMARY
    # ========================================================

    summary = {
        "status":
            "success",

        "excelRow":
            excel_row,

        "prudentialUrl":
            prudential_url,

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
            historical_rows[0],

        "oldestObservation":
            historical_rows[-1],
    }

    # ========================================================
    # CONSOLIDATED FUND RECORD
    # ========================================================

    fund_record = {
        "status":
            "success",

        "excelRow":
            excel_row,

        "prudentialUrl":
            prudential_url,

        "excelPruAccessName":
            excel_pruaccess_name,

        "prudential":
            prudential,

        "pruAccess": {
            "fundName":
                selected["text"],

            "fundId":
                fund_id,

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
        },

        "summary":
            summary,

        "preSubmit":
            pre_submit,

        "pagination":
            page_diagnostics,

        "bidHistory":
            bid_history,
    }

    return fund_record


# ============================================================
# MAIN
# ============================================================

def main():

    OUTPUT_DIR.mkdir(
        parents=True,
        exist_ok=True,
    )

    FUNDS_OUTPUT_DIR.mkdir(
        parents=True,
        exist_ok=True,
    )

    print(
        "============================================================"
    )

    print(
        "PRUACCESS HISTORICAL BID PRICE EXTRACTION"
    )

    print(
        "ALL EXCEL FUNDS"
    )

    print(
        "============================================================"
    )

    # ========================================================
    # READ EXCEL UNIVERSE
    # ========================================================

    funds = read_excel_funds()

    # ========================================================
    # RUN METADATA
    # ========================================================

    run_started = datetime.utcnow()

    successful_funds = []

    failed_funds = []

    all_bid_history = {}

    all_fund_records = []

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
        # LOAD PRUACCESS OPTIONS ONCE
        # ====================================================

        print(
            "\n============================================================"
        )

        print(
            "LOADING PRUACCESS FUND OPTIONS"
        )

        print(
            "============================================================"
        )

        page.goto(
            PRUACCESS_URL,
            wait_until="domcontentloaded",
            timeout=120000,
        )

        page.wait_for_timeout(
            PRUACCESS_INITIAL_WAIT_MS
        )

        pruaccess_options = (
            get_fund_options(
                page
            )
        )

        save_json(
            OUTPUT_DIR
            / "fund_options.json",
            pruaccess_options,
        )

        # ====================================================
        # PROCESS EVERY EXCEL FUND
        # ====================================================

        total_funds = len(
            funds
        )

        for fund_index, excel_fund in enumerate(
            funds,
            start=1,
        ):

            print(
                "\n\n"
                "############################################################"
            )

            print(
                f"FUND {fund_index} / {total_funds}"
            )

            print(
                f"EXCEL ROW {excel_fund['excelRow']}"
            )

            print(
                "############################################################"
            )

            try:

                result = extract_single_fund(
                    page,
                    excel_fund,
                    pruaccess_options,
                )

                all_fund_records.append(
                    result
                )

                successful_funds.append(
                    result
                )

                # ------------------------------------------------
                # Individual fund directory.
                # ------------------------------------------------

                prudential = result[
                    "prudential"
                ]

                fund_identifier = (
                    clean_text(
                        prudential.get(
                            "fundIdentifier"
                        )
                    )
                )

                fund_code = (
                    clean_text(
                        prudential.get(
                            "fundCode"
                        )
                    )
                )

                fund_name = (
                    clean_text(
                        prudential.get(
                            "fundName"
                        )
                    )
                )

                directory_name = (
                    f"{excel_fund['excelRow']}_"
                    f"{safe_filename(fund_identifier or fund_code or fund_name)}"
                )

                fund_output_dir = (
                    FUNDS_OUTPUT_DIR
                    / directory_name
                )

                fund_output_dir.mkdir(
                    parents=True,
                    exist_ok=True,
                )

                save_json(
                    fund_output_dir
                    / "prudential_fund.json",
                    prudential,
                )

                save_json(
                    fund_output_dir
                    / "pre_submit.json",
                    result[
                        "preSubmit"
                    ],
                )

                save_json(
                    fund_output_dir
                    / "pagination.json",
                    result[
                        "pagination"
                    ],
                )

                save_json(
                    fund_output_dir
                    / "bid_history.json",
                    result[
                        "bidHistory"
                    ],
                )

                save_json(
                    fund_output_dir
                    / "summary.json",
                    result[
                        "summary"
                    ],
                )

                # ------------------------------------------------
                # Store consolidated BID history keyed by
                # Prudential fund identifier.
                # ------------------------------------------------

                history_key = (
                    fund_identifier
                    or fund_code
                    or str(
                        excel_fund[
                            "excelRow"
                        ]
                    )
                )

                all_bid_history[
                    history_key
                ] = result[
                    "bidHistory"
                ]

                print(
                    "\n============================================================"
                )

                print(
                    "FUND SUCCESS"
                )

                print(
                    "============================================================"
                )

                print(
                    f"Fund: {fund_name}"
                )

                print(
                    f"Identifier: {fund_identifier}"
                )

                print(
                    f"PruAccess ID: "
                    f"{result['pruAccess']['fundId']}"
                )

                print(
                    f"Pages: "
                    f"{result['pruAccess']['observationCount']}"
                    " observations"
                )

                print(
                    f"Historical BID observations: "
                    f"{result['pruAccess']['observationCount']}"
                )

            except Exception as error:

                error_text = clean_text(
                    str(error)
                )

                print(
                    "\n============================================================"
                )

                print(
                    "FUND FAILED"
                )

                print(
                    "============================================================"
                )

                print(
                    f"Excel row: "
                    f"{excel_fund['excelRow']}"
                )

                print(
                    f"Prudential URL: "
                    f"{excel_fund['prudentialUrl']}"
                )

                print(
                    f"PruAccess name: "
                    f"{excel_fund['pruAccessName']}"
                )

                print(
                    f"Error: "
                    f"{error_text}"
                )

                failure = {
                    "status":
                        "failed",

                    "excelRow":
                        excel_fund[
                            "excelRow"
                        ],

                    "prudentialUrl":
                        excel_fund[
                            "prudentialUrl"
                        ],

                    "excelPruAccessName":
                        excel_fund[
                            "pruAccessName"
                        ],

                    "error":
                        error_text,
                }

                failed_funds.append(
                    failure
                )

                # ------------------------------------------------
                # Save individual failure record.
                # ------------------------------------------------

                failure_dir_name = (
                    f"{excel_fund['excelRow']}_"
                    f"failed"
                )

                failure_output_dir = (
                    FUNDS_OUTPUT_DIR
                    / failure_dir_name
                )

                failure_output_dir.mkdir(
                    parents=True,
                    exist_ok=True,
                )

                save_json(
                    failure_output_dir
                    / "failure.json",
                    failure,
                )

                # ------------------------------------------------
                # Continue with the next fund.
                # ------------------------------------------------

                continue

        browser.close()

    # ========================================================
    # CONSOLIDATED FUND DATA
    # ========================================================

    run_finished = datetime.utcnow()

    all_funds = {
        "source":
            "Funds Links.xlsm",

        "generatedAtUtc":
            run_finished.isoformat()
            + "Z",

        "fundUniverseCount":
            len(
                funds
            ),

        "successfulFundCount":
            len(
                successful_funds
            ),

        "failedFundCount":
            len(
                failed_funds
            ),

        "funds":
            all_fund_records,

        "failed":
            failed_funds,
    }

    save_json(
        OUTPUT_DIR
        / "all_funds.json",
        all_funds,
    )

    # ========================================================
    # CONSOLIDATED BID HISTORY
    # ========================================================

    consolidated_bid_history = {
        "source":
            "PruAccess",

        "priceType":
            "BID",

        "generatedAtUtc":
            run_finished.isoformat()
            + "Z",

        "fundCount":
            len(
                all_bid_history
            ),

        "funds":
            all_bid_history,
    }

    save_json(
        OUTPUT_DIR
        / "all_bid_history.json",
        consolidated_bid_history,
    )

    # ========================================================
    # RUN SUMMARY
    # ========================================================

    total_observations = 0

    for fund_record in successful_funds:

        total_observations += (
            fund_record[
                "pruAccess"
            ][
                "observationCount"
            ]
        )

    run_summary = {
        "status":
            (
                "success"
                if not failed_funds
                else "partial"
            ),

        "startedAtUtc":
            run_started.isoformat()
            + "Z",

        "completedAtUtc":
            run_finished.isoformat()
            + "Z",

        "excelFile":
            str(
                EXCEL_FILE
            ),

        "fundUniverseCount":
            len(
                funds
            ),

        "successfulFundCount":
            len(
                successful_funds
            ),

        "failedFundCount":
            len(
                failed_funds
            ),

        "totalHistoricalBidObservations":
            total_observations,

        "historicalPriceType":
            "BID",

        "pageSize":
            PAGE_SIZE,

        "successfulFunds":
            [
                {
                    "excelRow":
                        fund[
                            "excelRow"
                        ],

                    "fundName":
                        fund[
                            "prudential"
                        ].get(
                            "fundName"
                        ),

                    "fundIdentifier":
                        fund[
                            "prudential"
                        ].get(
                            "fundIdentifier"
                        ),

                    "fundCode":
                        fund[
                            "prudential"
                        ].get(
                            "fundCode"
                        ),

                    "pruAccessFundName":
                        fund[
                            "pruAccess"
                        ].get(
                            "fundName"
                        ),

                    "pruAccessFundId":
                        fund[
                            "pruAccess"
                        ].get(
                            "fundId"
                        ),

                    "observationCount":
                        fund[
                            "pruAccess"
                        ].get(
                            "observationCount"
                        ),

                    "newestObservation":
                        fund[
                            "summary"
                        ].get(
                            "newestObservation"
                        ),

                    "oldestObservation":
                        fund[
                            "summary"
                        ].get(
                            "oldestObservation"
                        ),
                }

                for fund in successful_funds
            ],

        "failedFunds":
            failed_funds,
    }

    save_json(
        OUTPUT_DIR
        / "run_summary.json",
        run_summary,
    )

    # ========================================================
    # CONSOLE SUMMARY
    # ========================================================

    print(
        "\n\n"
        "============================================================"
    )

    print(
        "ALL-FUND EXTRACTION COMPLETE"
    )

    print(
        "============================================================"
    )

    print(
        f"Excel fund universe: "
        f"{len(funds)}"
    )

    print(
        f"Successful funds: "
        f"{len(successful_funds)}"
    )

    print(
        f"Failed funds: "
        f"{len(failed_funds)}"
    )

    print(
        f"Total historical BID observations: "
        f"{total_observations}"
    )

    print(
        "\nSuccessful funds:"
    )

    for fund in successful_funds:

        prudential = fund[
            "prudential"
        ]

        print(
            f" - Row "
            f"{fund['excelRow']}: "
            f"{prudential.get('fundName')} "
            f"| "
            f"{prudential.get('fundIdentifier')} "
            f"| "
            f"{fund['pruAccess']['observationCount']} observations"
        )

    if failed_funds:

        print(
            "\nFailed funds:"
        )

        for failure in failed_funds:

            print(
                f" - Row "
                f"{failure['excelRow']}: "
                f"{failure['excelPruAccessName']} "
                f"| "
                f"{failure['error']}"
            )

    print(
        "\nFiles created:"
    )

    for file in sorted(
        OUTPUT_DIR.iterdir()
    ):

        if file.is_file():

            print(
                f" - {file}"
            )

    print(
        "\nIndividual fund results:"
    )

    print(
        f" - {FUNDS_OUTPUT_DIR}"
    )

    print(
        "\nDone."
    )


if __name__ == "__main__":
    main()
