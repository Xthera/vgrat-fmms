#!/usr/bin/env python3

"""
PruAccess historical BID price extractor.

MASTER SOURCE:
    Funds Links.xlsm

Excel:
    Column A = Prudential fund URL
    Column B = PruAccess fund name

WORKFLOW
========

    Funds Links.xlsm
          |
          v
    Read every populated Excel row
          |
          v
    Prudential fund page/API
          |
          v
    Official Prudential fund information
          |
          v
    Official Prudential inception date
          |
          v
    PruAccess
          |
          v
    Exact fund-name match
          |
          v
    Table view
          |
          v
    Start = Prudential inception date
    End   = current Singapore date
          |
          v
    Submit PruAccess search
          |
          v
    Establish authenticated/session cookies
          |
          v
    Direct HTTP pagination
          |
          v
    Extract BID prices only
          |
          v
    Validate complete pagination
          |
          v
    Repeat for every Excel fund
          |
          v
    Consolidated JSON


IMPORTANT RULES
===============

1. Funds Links.xlsm determines the fund universe.

2. Column A determines which funds exist.

3. Column B provides the exact PruAccess fund name.

4. No hardcoded fund count.

5. No fuzzy matching.

6. No synthetic data.

7. No estimated data.

8. No interpolation.

9. No fabricated observations.

10. Historical prices come directly from PruAccess.

11. Historical price type is BID only.

12. Offer prices are never stored as historical BID prices.

13. Fund Price Type is never changed.

14. Start date is the official Prudential inception date.

15. End date is the current Singapore date.

16. All pagination pages must be successfully retrieved.

17. A permanently failed page fails the entire fund.

18. A fund is never marked successful with partial history.

19. Duplicate observations are not silently removed by the extractor.

20. Validation detects duplicate observations.

21. Current Prudential BID is stored separately from historical PruAccess BID.

22. Prudential fund name is authoritative.

23. Citicode/fundIdentifier is technical identity metadata.

24. No PruAccess fund is added to the universe unless it exists
    in Funds Links.xlsm.

25. If the extractor cannot confidently obtain data, it fails loudly.

26. Code is intentionally API/HTTP-first for pagination.

27. Playwright is used only where needed to establish the PruAccess
    session and submit the initial search.

28. Direct HTTP pagination uses the same PruAccess session cookies.

29. Full pagination is validated before a fund becomes successful.

30. No partial-success fund is published as successful.
"""

from __future__ import annotations

import json
import re
import time
from datetime import datetime
from pathlib import Path
from zoneinfo import ZoneInfo

from bs4 import BeautifulSoup
from openpyxl import load_workbook
from playwright.sync_api import sync_playwright


# ============================================================
# CONFIG
# ============================================================

EXCEL_FILE = Path("Funds Links.xlsm")

OUTPUT_DIR = Path("output_pruaccess")

FUNDS_OUTPUT_DIR = (
    OUTPUT_DIR / "funds"
)

PRUACCESS_URL = (
    "https://pruaccess.prudential.com.sg/"
    "prulinkfund/viewFundPerformance.do"
)

PAGE_SIZE = 20

MAX_PAGES = 1000

# Number of retries for an individual pagination request.
PAGE_RETRY_COUNT = 4

# Seconds between retries.
PAGE_RETRY_DELAY_SECONDS = 5

# Direct HTTP request timeout.
PAGE_TIMEOUT_MS = 120000

# Initial Playwright navigation timeout.
INITIAL_PAGE_TIMEOUT_MS = 120000

# Wait after opening Prudential.
PRUDENTIAL_WAIT_MS = 3000

# Wait after opening PruAccess.
PRUACCESS_INITIAL_WAIT_MS = 1500

# Wait after submitting PruAccess form.
PRUACCESS_RESULT_WAIT_MS = 2500

# Headless browser.
BROWSER_HEADLESS = True

# Singapore timezone.
SINGAPORE_TZ = ZoneInfo(
    "Asia/Singapore"
)


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


def singapore_today() -> str:

    return datetime.now(
        SINGAPORE_TZ
    ).strftime(
        "%d-%b-%Y"
    )


def utc_now_iso() -> str:

    return (
        datetime.now(
            ZoneInfo("UTC")
        ).isoformat()
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

        # Column A is the master universe.
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

            if response.url not in captured_urls:

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
            timeout=INITIAL_PAGE_TIMEOUT_MS,
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
        timeout=INITIAL_PAGE_TIMEOUT_MS,
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

    if not isinstance(
        fund,
        dict,
    ):

        raise RuntimeError(
            "Unexpected Prudential "
            "fund record."
        )

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

    return {
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

        "prospectusUrl":
            fund.get(
                "prospectusUrl"
            ),

        "productHighlightSheetUrl":
            fund.get(
                "productHighlightSheetUrl"
            ),

        "annualReportUrl":
            fund.get(
                "annualReportUrl"
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
            option.get(
                "text",
                "",
            )
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

    if actual != value:

        raise RuntimeError(
            f"Could not set {selector}. "
            f"Expected {value}, "
            f"got {actual}."
        )


# ============================================================
# HTML PRICE TABLE EXTRACTION
# ============================================================

def extract_price_rows_from_html(
    html: str,
) -> list[dict]:

    soup = BeautifulSoup(
        html,
        "html.parser",
    )

    rows = []

    for table in soup.find_all(
        "table"
    ):

        for tr in table.find_all(
            "tr"
        ):

            cells = tr.find_all(
                ["th", "td"]
            )

            if len(cells) < 2:
                continue

            values = [
                clean_text(
                    cell.get_text(
                        " ",
                        strip=True,
                    )
                )
                for cell in cells
            ]

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
                r"^\d+(?:\.\d+)?$",
                bid_value,
            ):
                continue

            rows.append(
                {
                    "date":
                        date_value,

                    "bidPrice":
                        bid_value,
                }
            )

    return rows


# ============================================================
# PAGINATION PAGE COUNT DISCOVERY
# ============================================================

def discover_max_page(
    html: str,
) -> int | None:

    soup = BeautifulSoup(
        html,
        "html.parser",
    )

    page_numbers = []

    for link in soup.find_all(
        "a"
    ):

        href = link.get(
            "href",
            "",
        )

        if not href:
            continue

        matches = re.findall(
            r"page\.page=(\d+)",
            href,
        )

        for match in matches:

            try:

                page_numbers.append(
                    int(match)
                )

            except ValueError:

                continue

    # Some versions of the site may expose
    # page numbers in plain text rather than links.
    text = soup.get_text(
        " ",
        strip=True,
    )

    text_matches = re.findall(
        r"[Pp]age\s+(\d+)\s+(?:of|/)\s+(\d+)",
        text,
    )

    for _, total in text_matches:

        try:

            page_numbers.append(
                int(total)
            )

        except ValueError:

            pass

    if not page_numbers:

        return None

    maximum = max(
        page_numbers
    )

    if maximum < 1:

        return None

    return maximum


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
# DIRECT HTTP PAGINATION REQUEST
# ============================================================

def request_page_html(
    request_context,
    url: str,
    page_number: int,
) -> tuple[str, int]:

    last_error = None

    for attempt in range(
        1,
        PAGE_RETRY_COUNT + 1,
    ):

        print(
            f"  HTTP page {page_number} "
            f"attempt {attempt}/{PAGE_RETRY_COUNT}"
        )

        try:

            response = request_context.get(
                url,
                timeout=PAGE_TIMEOUT_MS,
                fail_on_status_code=False,
            )

            status = response.status

            if status != 200:

                raise RuntimeError(
                    f"HTTP {status}"
                )

            html = response.text()

            if not html:

                raise RuntimeError(
                    "Empty HTTP response"
                )

            # Verify that this is actually a
            # PruAccess fund-performance page.
            if (
                "PruLink Fund Price Table"
                not in html
                and
                "Fund Performance"
                not in html
            ):

                raise RuntimeError(
                    "Response does not appear "
                    "to be a PruAccess fund "
                    "performance page."
                )

            return (
                html,
                status,
            )

        except Exception as error:

            last_error = error

            print(
                f"  Page {page_number} failed: "
                f"{clean_text(str(error))}"
            )

            if attempt < PAGE_RETRY_COUNT:

                print(
                    f"  Waiting "
                    f"{PAGE_RETRY_DELAY_SECONDS}s "
                    f"before retry..."
                )

                time.sleep(
                    PAGE_RETRY_DELAY_SECONDS
                )

    raise RuntimeError(
        f"PruAccess page {page_number} "
        f"failed after "
        f"{PAGE_RETRY_COUNT} attempts: "
        f"{clean_text(str(last_error))}"
    )


# ============================================================
# FULL PAGINATION EXTRACTION
# ============================================================

def extract_all_pages(
    request_context,
    fund_id: str,
    start_date: str,
    end_date: str,
) -> tuple[list[dict], list[dict]]:

    print(
        "\n============================================================"
    )

    print(
        "STARTING DIRECT HTTP PRUACCESS PAGINATION"
    )

    print(
        "============================================================"
    )

    first_url = build_page_url(
        PRUACCESS_URL,
        fund_id,
        "TBL",
        start_date,
        end_date,
        1,
    )

    first_html, first_status = (
        request_page_html(
            request_context,
            first_url,
            1,
        )
    )

    first_rows = (
        extract_price_rows_from_html(
            first_html
        )
    )

    if not first_rows:

        raise RuntimeError(
            "PruAccess page 1 returned "
            "no historical BID rows."
        )

    discovered_max_page = (
        discover_max_page(
            first_html
        )
    )

    print(
        f"Page 1 rows: "
        f"{len(first_rows)}"
    )

    print(
        f"Discovered maximum page: "
        f"{discovered_max_page}"
    )

    page_diagnostics = []

    all_rows = []

    # --------------------------------------------------------
    # Store page 1.
    # --------------------------------------------------------

    page_diagnostics.append(
        {
            "page":
                1,

            "url":
                first_url,

            "httpStatus":
                first_status,

            "rowCount":
                len(first_rows),

            "firstDate":
                first_rows[0]["date"],

            "lastDate":
                first_rows[-1]["date"],

            "status":
                "success",
        }
    )

    all_rows.extend(
        first_rows
    )

    # --------------------------------------------------------
    # If page 1 is already a partial page and no explicit
    # pagination says otherwise, page 1 is the final page.
    # --------------------------------------------------------

    if (
        discovered_max_page is None
        and len(first_rows) < PAGE_SIZE
    ):

        max_page = 1

    elif discovered_max_page is not None:

        max_page = min(
            discovered_max_page,
            MAX_PAGES,
        )

    else:

        max_page = MAX_PAGES

    # --------------------------------------------------------
    # Fetch remaining pages.
    # --------------------------------------------------------

    for page_number in range(
        2,
        max_page + 1,
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
            f"{page_number}/{max_page}"
        )

        html, status = request_page_html(
            request_context,
            url,
            page_number,
        )

        rows = extract_price_rows_from_html(
            html
        )

        print(
            f"  Rows found: "
            f"{len(rows)}"
        )

        # ----------------------------------------------------
        # Empty pages are NOT automatically accepted.
        # request_page_html already retried the request.
        # ----------------------------------------------------

        if not rows:

            # If we already know the page count,
            # an empty expected page is a failure.
            if (
                discovered_max_page is not None
                and page_number <= discovered_max_page
            ):

                raise RuntimeError(
                    f"Expected PruAccess page "
                    f"{page_number} but it returned "
                    f"zero BID observations."
                )

            # Otherwise this marks the end.
            page_diagnostics.append(
                {
                    "page":
                        page_number,

                    "url":
                        url,

                    "httpStatus":
                        status,

                    "rowCount":
                        0,

                    "status":
                        "empty",
                }
            )

            break

        page_diagnostics.append(
            {
                "page":
                    page_number,

                "url":
                    url,

                "httpStatus":
                    status,

                "rowCount":
                    len(rows),

                "firstDate":
                    rows[0]["date"],

                "lastDate":
                    rows[-1]["date"],

                "status":
                    "success",
            }
        )

        all_rows.extend(
            rows
        )

        print(
            f"  Date range: "
            f"{rows[0]['date']} "
            f"-> "
            f"{rows[-1]['date']}"
        )

        # ----------------------------------------------------
        # If no explicit page count exists, a partial page
        # ends pagination.
        # ----------------------------------------------------

        if (
            discovered_max_page is None
            and len(rows) < PAGE_SIZE
        ):

            print(
                "\nFinal partial page detected."
            )

            break

    else:

        if (
            discovered_max_page is None
            and max_page == MAX_PAGES
        ):

            raise RuntimeError(
                "Pagination reached MAX_PAGES "
                f"({MAX_PAGES}) without determining "
                "the final page."
            )

    # --------------------------------------------------------
    # If pagination explicitly told us N pages exist,
    # make absolutely sure every page was fetched.
    # --------------------------------------------------------

    if discovered_max_page is not None:

        successful_pages = {
            item["page"]
            for item in page_diagnostics
            if item.get("status") == "success"
        }

        expected_pages = set(
            range(
                1,
                max_page + 1,
            )
        )

        missing_pages = sorted(
            expected_pages
            - successful_pages
        )

        if missing_pages:

            raise RuntimeError(
                "Pagination incomplete. "
                f"Missing pages: {missing_pages}"
            )

    # --------------------------------------------------------
    # IMPORTANT:
    #
    # Do NOT silently deduplicate here.
    #
    # The raw observations are preserved so the validator can
    # detect duplicates.
    # --------------------------------------------------------

    def sort_date(row):

        return datetime.strptime(
            row["date"],
            "%d-%b-%Y",
        )

    final_rows = sorted(
        all_rows,
        key=sort_date,
        reverse=True,
    )

    return (
        final_rows,
        page_diagnostics,
    )


# ============================================================
# SUBMIT PRUACCESS FORM
# ============================================================

def submit_pruaccess_search(
    page,
    fund_id: str,
    start_date: str,
    end_date: str,
) -> dict:

    print(
        "\n=== Opening PruAccess ==="
    )

    page.goto(
        PRUACCESS_URL,
        wait_until="domcontentloaded",
        timeout=INITIAL_PAGE_TIMEOUT_MS,
    )

    page.wait_for_timeout(
        PRUACCESS_INITIAL_WAIT_MS
    )

    # --------------------------------------------------------
    # Fund
    # --------------------------------------------------------

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

    # --------------------------------------------------------
    # Table view
    # --------------------------------------------------------

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

    # --------------------------------------------------------
    # Fund Price Type
    #
    # NEVER MODIFY.
    # --------------------------------------------------------

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

    # --------------------------------------------------------
    # Dates
    # --------------------------------------------------------

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

    # --------------------------------------------------------
    # CSRF
    # --------------------------------------------------------

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

    # --------------------------------------------------------
    # Verify
    # --------------------------------------------------------

    actual_start = page.locator(
        'input[name="startDate"]'
    ).input_value()

    actual_end = page.locator(
        'input[name="endDate"]'
    ).input_value()

    if actual_start != start_date:

        raise RuntimeError(
            "Start date verification failed."
        )

    if actual_end != end_date:

        raise RuntimeError(
            "End date verification failed."
        )

    # --------------------------------------------------------
    # Submit.
    #
    # We intentionally use the form's normal submit.
    # This establishes the PruAccess session/cookies.
    # --------------------------------------------------------

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
        timeout=INITIAL_PAGE_TIMEOUT_MS,
    )

    page.wait_for_timeout(
        PRUACCESS_RESULT_WAIT_MS
    )

    # --------------------------------------------------------
    # Verify that the resulting page contains the table.
    # --------------------------------------------------------

    result_text = clean_text(
        page.locator(
            "body"
        ).inner_text()
    )

    if (
        "PruLink Fund Price Table"
        not in result_text
    ):

        raise RuntimeError(
            "PruAccess form submission did not "
            "produce the expected fund price table."
        )

    return {
        "fundPriceType":
            fund_price_type_value,

        "csrfPresent":
            bool(csrf_value),

        "startDate":
            start_date,

        "endDate":
            end_date,

        "viewType":
            "TBL",

        "fundId":
            fund_id,
    }


# ============================================================
# SINGLE FUND EXTRACTION
# ============================================================

def extract_single_fund(
    context,
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
        "\nPrudential fund:"
    )

    print(
        prudential["fundName"]
    )

    print(
        f"Identifier: "
        f"{prudential['fundIdentifier']}"
    )

    print(
        f"Fund code: "
        f"{prudential.get('fundCode')}"
    )

    print(
        f"Current BID: "
        f"{prudential.get('bidPrice')}"
    )

    print(
        f"Inception: "
        f"{prudential['inceptionDate']}"
    )

    # ========================================================
    # DATE RANGE
    # ========================================================

    inception_date = parse_prudential_date(
        prudential["inceptionDate"]
    )

    start_date = pruaccess_date(
        inception_date
    )

    end_date = singapore_today()

    print(
        f"\nPruAccess date range: "
        f"{start_date} -> {end_date}"
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

    fund_id = clean_text(
        selected.get(
            "value"
        )
    )

    if not fund_id:

        raise RuntimeError(
            "Matched PruAccess fund "
            "has no fund ID."
        )

    print(
        "\nMatched PruAccess fund:"
    )

    print(
        selected
    )

    # ========================================================
    # SUBMIT SEARCH
    # ========================================================

    pre_submit = submit_pruaccess_search(
        page,
        fund_id,
        start_date,
        end_date,
    )

    pre_submit.update(
        {
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
        }
    )

    # ========================================================
    # DIRECT HTTP PAGINATION
    #
    # Browser is no longer used for every historical page.
    # The BrowserContext request context shares the browser
    # session cookies.
    # ========================================================

    historical_rows, page_diagnostics = (
        extract_all_pages(
            context.request,
            fund_id,
            start_date,
            end_date,
        )
    )

    if not historical_rows:

        raise RuntimeError(
            "No historical BID observations "
            "were extracted."
        )

    # ========================================================
    # BASIC EXTRACTION INTEGRITY
    # ========================================================

    if (
        page_diagnostics
        and page_diagnostics[0]["page"] != 1
    ):

        raise RuntimeError(
            "Pagination does not start at page 1."
        )

    expected_pages = list(
        range(
            1,
            len(page_diagnostics) + 1,
        )
    )

    actual_pages = [
        item["page"]
        for item in page_diagnostics
        if item.get("status") == "success"
    ]

    if actual_pages != expected_pages:

        raise RuntimeError(
            "Pagination page sequence is incomplete "
            "or out of order."
        )

    # ========================================================
    # BID HISTORY
    # ========================================================

    bid_history = {
        "source":
            "PruAccess",

        "fundName":
            prudential["fundName"],

        "pruAccessFundName":
            selected["text"],

        "fundId":
            fund_id,

        "fundIdentifier":
            prudential["fundIdentifier"],

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
            prudential["fundName"],

        "excelPruAccessName":
            excel_pruaccess_name,

        "matchedPruAccessName":
            selected["text"],

        "pruAccessFundId":
            fund_id,

        "fundIdentifier":
            prudential["fundIdentifier"],

        "fundCode":
            prudential.get(
                "fundCode"
            ),

        "inceptionDate":
            prudential["inceptionDate"],

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
            pre_submit.get(
                "fundPriceType"
            ),

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
    # FUND RECORD
    # ========================================================

    return {
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


# ============================================================
# SAVE SUCCESSFUL FUND
# ============================================================

def save_successful_fund(
    result: dict,
) -> None:

    prudential = result[
        "prudential"
    ]

    excel_row = result[
        "excelRow"
    ]

    fund_identifier = clean_text(
        prudential.get(
            "fundIdentifier"
        )
    )

    fund_code = clean_text(
        prudential.get(
            "fundCode"
        )
    )

    fund_name = clean_text(
        prudential.get(
            "fundName"
        )
    )

    # Calculate the safe filename separately.
    # This avoids nested quotation marks inside the f-string.
    safe_fund_name = safe_filename(
        fund_identifier
        or fund_code
        or fund_name
    )

    directory_name = (
        f"{excel_row}_"
        f"{safe_fund_name}"
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
        result["preSubmit"],
    )

    save_json(
        fund_output_dir
        / "pagination.json",
        result["pagination"],
    )

    save_json(
        fund_output_dir
        / "bid_history.json",
        result["bidHistory"],
    )

    save_json(
        fund_output_dir
        / "summary.json",
        result["summary"],
    )


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
        "DIRECT HTTP PAGINATION VERSION"
    )

    print(
        "============================================================"
    )

    # ========================================================
    # READ MASTER EXCEL
    # ========================================================

    funds = read_excel_funds()

    # ========================================================
    # RUN METADATA
    # ========================================================

    run_started = utc_now_iso()

    successful_funds = []

    failed_funds = []

    all_bid_history = {}

    # ========================================================
    # PLAYWRIGHT
    # ========================================================

    with sync_playwright() as playwright:

        browser = playwright.chromium.launch(
            headless=BROWSER_HEADLESS
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
            timeout=INITIAL_PAGE_TIMEOUT_MS,
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
            {
                "source":
                    PRUACCESS_URL,

                "retrievedAtUtc":
                    utc_now_iso(),

                "count":
                    len(
                        pruaccess_options
                    ),

                "options":
                    pruaccess_options,
            },
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
                    context,
                    page,
                    excel_fund,
                    pruaccess_options,
                )

                save_successful_fund(
                    result
                )

                successful_funds.append(
                    result
                )

                fund_identifier = clean_text(
                    result["prudential"].get(
                        "fundIdentifier"
                    )
                )

                fund_code = clean_text(
                    result["prudential"].get(
                        "fundCode"
                    )
                )

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
                    f"Fund: "
                    f"{result['prudential']['fundName']}"
                )

                print(
                    f"Identifier: "
                    f"{fund_identifier}"
                )

                print(
                    f"PruAccess ID: "
                    f"{result['pruAccess']['fundId']}"
                )

                print(
                    f"Pages: "
                    f"{result['summary']['pagesExtracted']}"
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

                failure_dir_name = (
                    f"{excel_fund['excelRow']}_failed"
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

                continue

        browser.close()

    # ========================================================
    # CONSOLIDATED DATA
    # ========================================================

    run_finished = utc_now_iso()

    total_observations = sum(
        fund["pruAccess"][
            "observationCount"
        ]
        for fund in successful_funds
    )

    # ========================================================
    # ALL FUNDS
    # ========================================================

    all_funds = {
        "source":
            "Funds Links.xlsm",

        "generatedAtUtc":
            run_finished,

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
            successful_funds,

        "failed":
            failed_funds,

        "failedFunds":
            failed_funds,
    }

    save_json(
        OUTPUT_DIR
        / "all_funds.json",
        all_funds,
    )

    # ========================================================
    # ALL BID HISTORY
    # ========================================================

    consolidated_bid_history = {
        "source":
            "PruAccess",

        "priceType":
            "BID",

        "generatedAtUtc":
            run_finished,

        "fundCount":
            len(
                all_bid_history
            ),

        "totalHistoricalBidObservations":
            total_observations,

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

    run_status = (
        "success"
        if not failed_funds
        else "partial"
    )

    run_summary = {
        "status":
            run_status,

        "startedAtUtc":
            run_started,

        "completedAtUtc":
            run_finished,

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

        "pageRetryCount":
            PAGE_RETRY_COUNT,

        "pageRetryDelaySeconds":
            PAGE_RETRY_DELAY_SECONDS,

        "pageTimeoutMs":
            PAGE_TIMEOUT_MS,

        "successfulFunds":
            [
                {
                    "excelRow":
                        fund["excelRow"],

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

                    "pagesExtracted":
                        fund[
                            "summary"
                        ].get(
                            "pagesExtracted"
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

                for fund
                in successful_funds
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
