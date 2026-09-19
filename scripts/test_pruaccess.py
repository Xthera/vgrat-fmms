#!/usr/bin/env python3

"""
PruAccess historical BID price extractor.

MASTER SOURCE:
    Funds Links.xlsm

Excel:
    Column A = Prudential fund URL
    Column B = exact PruAccess fund name

================================================================
WORKFLOW
================================================================

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
    PruAccess exact fund-name match
          |
          v
    Automatically split history into <=10-year windows
          |
          v
    Submit each PruAccess window
          |
          v
    Establish authenticated/session cookies
          |
          v
    Discover complete pagination
          |
          v
    Parallel HTTP page retrieval
          |
          v
    Retry failed pages
          |
          v
    Validate every expected page
          |
          v
    Extract BID observations only
          |
          v
    Reconstruct complete chronology
          |
          v
    Validate inception coverage
          |
          v
    Validate duplicates
          |
          v
    Save fund
          |
          v
    Consolidated JSON

================================================================
IMPORTANT RULES
================================================================

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

14. The historical extraction begins from the official
    Prudential inception date.

15. PruAccess date windows must never exceed 10 years.

16. The historical end date is the current Singapore date.

17. Every expected pagination page must be retrieved.

18. A permanently failed page fails the entire window.

19. A failed window fails the entire fund.

20. A fund is never marked successful with partial history.

21. Duplicate observations are NOT silently removed.

22. Duplicate observations are detected and cause validation
    failure.

23. Current Prudential BID is stored separately from historical
    PruAccess BID.

24. Prudential fund name is authoritative.

25. Citicode/fundIdentifier is technical identity metadata.

26. No PruAccess fund is added to the universe unless it exists
    in Funds Links.xlsm.

27. If the extractor cannot confidently obtain data, it fails
    loudly.

28. Code is HTTP/API-first for pagination.

29. Playwright is used where required to establish the PruAccess
    session and submit each initial search window.

30. Direct HTTP pagination uses the same Playwright browser
    context session cookies.

31. Pagination pages are retrieved in parallel.

32. Every page is individually validated.

33. Full pagination is validated before a window becomes
    successful.

34. Every window must be successful before a fund becomes
    successful.

35. All windows are reconstructed chronologically.

36. Historical date gaps are allowed.

37. Missing calendar dates are NOT filled.

38. Weekends, public holidays, non-valuation days and other
    legitimate PruAccess gaps are allowed.

39. The oldest historical BID observation must be either:
        - the Prudential inception date, OR
        - no more than 1 calendar day after inception.

40. The validator does NOT require continuous daily observations.

41. No observation is fabricated to satisfy inception coverage.

42. If PruAccess genuinely provides no observation on the
    inception date but provides one on the following calendar
    day, that is valid.

43. Any history beginning more than 1 calendar day after the
    official inception date fails validation.

44. The extractor does not silently deduplicate overlapping or
    repeated observations.

45. If duplicate date observations are found, validation fails.

46. Every code change should result in the COMPLETE updated
    script being supplied, not a patch.
"""

from __future__ import annotations

import asyncio
import json
import re
import time
from concurrent.futures import ThreadPoolExecutor
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Any
from zoneinfo import ZoneInfo

from bs4 import BeautifulSoup
from openpyxl import load_workbook
from playwright.async_api import (
    Browser,
    BrowserContext,
    Page,
    async_playwright,
)


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

# ------------------------------------------------------------
# PruAccess pagination
# ------------------------------------------------------------

PAGE_SIZE = 20

MAX_PAGES = 1000

# Number of attempts for each HTTP page.
PAGE_RETRY_COUNT = 4

# Seconds between retries.
PAGE_RETRY_DELAY_SECONDS = 5

# HTTP timeout in milliseconds.
PAGE_TIMEOUT_MS = 120000

# ------------------------------------------------------------
# Parallelism
# ------------------------------------------------------------

# Number of historical pagination pages retrieved concurrently.
#
# This does NOT change the number of pages retrieved.
# It only controls concurrency.
PARALLEL_PAGE_WORKERS = 6

# ------------------------------------------------------------
# Browser timing
# ------------------------------------------------------------

INITIAL_PAGE_TIMEOUT_MS = 120000

PRUDENTIAL_WAIT_MS = 3000

PRUACCESS_INITIAL_WAIT_MS = 1500

PRUACCESS_RESULT_WAIT_MS = 2500

BROWSER_HEADLESS = True

# ------------------------------------------------------------
# PruAccess date limitation
# ------------------------------------------------------------

# PruAccess allows a maximum historical date range of 10 years.
#
# To remain safely inside that limit, each window is constructed
# as:
#
#     start -> start + 10 calendar years - 1 day
#
# The next window starts the following calendar day.
#
# Therefore no window exceeds exactly 10 calendar years.
MAX_WINDOW_YEARS = 10

# ------------------------------------------------------------
# Singapore timezone
# ------------------------------------------------------------

SINGAPORE_TZ = ZoneInfo(
    "Asia/Singapore"
)


# ============================================================
# GENERAL HELPERS
# ============================================================

def clean_text(value: Any) -> str:

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
    data: Any,
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


def parse_pruaccess_date(
    value: str,
) -> date:

    return datetime.strptime(
        value.strip(),
        "%d-%b-%Y",
    ).date()


def format_pruaccess_date(
    value: date,
) -> str:

    return value.strftime(
        "%d-%b-%Y"
    )


def singapore_today() -> date:

    return datetime.now(
        SINGAPORE_TZ
    ).date()


def singapore_today_pruaccess() -> str:

    return format_pruaccess_date(
        singapore_today()
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
# DATE WINDOW GENERATION
# ============================================================

def add_calendar_years(
    value: date,
    years: int,
) -> date:

    """
    Add calendar years while handling February 29.

    Example:

        29-Feb-2024 + 10 years
        -> 28-Feb-2034
    """

    try:

        return value.replace(
            year=value.year + years
        )

    except ValueError:

        # February 29 -> February 28
        return value.replace(
            year=value.year + years,
            day=28,
        )


def build_date_windows(
    inception_date: date,
    end_date: date,
) -> list[dict]:

    """
    Split inception -> end into sequential windows where each
    window is safely <= 10 calendar years.

    Example:

        01-Jan-2000 -> 31-Dec-2010

    becomes approximately:

        01-Jan-2000 -> 31-Dec-2009
        01-Jan-2010 -> 31-Dec-2010

    No dates are skipped.

    Windows are adjacent, not overlapping, so legitimate
    duplicate observations are not created merely by the
    extraction design.
    """

    if end_date < inception_date:

        raise RuntimeError(
            "PruAccess end date is earlier "
            "than Prudential inception date."
        )

    windows = []

    current_start = inception_date

    window_number = 1

    while current_start <= end_date:

        ten_year_boundary = (
            add_calendar_years(
                current_start,
                MAX_WINDOW_YEARS,
            )
            - timedelta(days=1)
        )

        current_end = min(
            ten_year_boundary,
            end_date,
        )

        if current_end < current_start:

            raise RuntimeError(
                "Generated invalid PruAccess "
                "date window."
            )

        windows.append(
            {
                "windowNumber":
                    window_number,

                "startDate":
                    current_start,

                "endDate":
                    current_end,

                "startDateText":
                    format_pruaccess_date(
                        current_start
                    ),

                "endDateText":
                    format_pruaccess_date(
                        current_end
                    ),

                "calendarDays":
                    (
                        current_end
                        - current_start
                    ).days + 1,
            }
        )

        next_start = (
            current_end
            + timedelta(days=1)
        )

        if next_start <= current_start:

            raise RuntimeError(
                "Date-window generator did not "
                "advance chronologically."
            )

        current_start = next_start

        window_number += 1

    # --------------------------------------------------------
    # Final coverage validation.
    # --------------------------------------------------------

    if not windows:

        raise RuntimeError(
            "No PruAccess date windows generated."
        )

    if (
        windows[0]["startDate"]
        != inception_date
    ):

        raise RuntimeError(
            "First PruAccess window does not "
            "start at Prudential inception date."
        )

    if (
        windows[-1]["endDate"]
        != end_date
    ):

        raise RuntimeError(
            "Last PruAccess window does not "
            "reach the current Singapore date."
        )

    # --------------------------------------------------------
    # Ensure no gaps or overlaps between windows.
    # --------------------------------------------------------

    for index in range(
        1,
        len(windows),
    ):

        previous = windows[
            index - 1
        ]

        current = windows[
            index
        ]

        expected_start = (
            previous["endDate"]
            + timedelta(days=1)
        )

        if (
            current["startDate"]
            != expected_start
        ):

            raise RuntimeError(
                "PruAccess windows contain a "
                "gap or overlap."
            )

    return windows


# ============================================================
# EXCEL MASTER UNIVERSE
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

async def get_prudential_fund_data(
    page: Page,
    prudential_url: str,
) -> dict:

    print(
        "\n=== Opening Prudential fund page ==="
    )

    print(
        prudential_url
    )

    captured_urls = []

    async def handle_response(
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

        await page.goto(
            prudential_url,
            wait_until="domcontentloaded",
            timeout=INITIAL_PAGE_TIMEOUT_MS,
        )

        await page.wait_for_timeout(
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

    response = await page.request.get(
        api_url,
        timeout=INITIAL_PAGE_TIMEOUT_MS,
    )

    if not response.ok:

        raise RuntimeError(
            "Prudential API returned "
            f"HTTP {response.status}"
        )

    data = await response.json()

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

async def get_fund_options(
    page: Page,
) -> list[dict]:

    options = page.locator(
        "#fundName option"
    )

    count = await options.count()

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
                        await option.inner_text()
                    ),

                "value":
                    await option.get_attribute(
                        "value"
                    ),
            }
        )

    return result


# ============================================================
# EXACT PRUACCESS MATCH
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

async def set_readonly_input_value(
    page: Page,
    selector: str,
    value: str,
) -> None:

    locator = page.locator(
        selector
    )

    if await locator.count() == 0:

        raise RuntimeError(
            f"Could not find input: "
            f"{selector}"
        )

    await locator.evaluate(
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

    actual = await locator.input_value()

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

            # Verify the date can actually be parsed.
            try:

                parse_pruaccess_date(
                    date_value
                )

            except ValueError:

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
# RESPONSE VALIDATION
# ============================================================

def validate_pagination_html(
    html: str,
    page_number: int,
) -> None:

    if not html:

        raise RuntimeError(
            f"Page {page_number} returned "
            "an empty HTTP response."
        )

    if (
        "PruLink Fund Price Table"
        not in html
        and
        "Fund Performance"
        not in html
    ):

        raise RuntimeError(
            f"Page {page_number} response does not "
            "appear to be a PruAccess fund-performance "
            "page."
        )


# ============================================================
# ASYNC HTTP PAGE RETRIEVAL
# ============================================================

async def request_page_html(
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
            f"    HTTP page {page_number} "
            f"attempt {attempt}/{PAGE_RETRY_COUNT}"
        )

        try:

            response = await request_context.get(
                url,
                timeout=PAGE_TIMEOUT_MS,
                fail_on_status_code=False,
            )

            status = response.status

            if status != 200:

                raise RuntimeError(
                    f"HTTP {status}"
                )

            html = await response.text()

            validate_pagination_html(
                html,
                page_number,
            )

            return (
                html,
                status,
            )

        except Exception as error:

            last_error = error

            print(
                f"    Page {page_number} failed: "
                f"{clean_text(str(error))}"
            )

            if attempt < PAGE_RETRY_COUNT:

                print(
                    f"    Waiting "
                    f"{PAGE_RETRY_DELAY_SECONDS}s "
                    f"before retry..."
                )

                await asyncio.sleep(
                    PAGE_RETRY_DELAY_SECONDS
                )

    raise RuntimeError(
        f"PruAccess page {page_number} "
        f"failed after "
        f"{PAGE_RETRY_COUNT} attempts: "
        f"{clean_text(str(last_error))}"
    )


# ============================================================
# PARALLEL PAGE RETRIEVAL
# ============================================================

async def fetch_pages_parallel(
    request_context,
    page_jobs: list[tuple[int, str]],
) -> dict[int, dict]:

    """
    Retrieve pagination pages concurrently.

    A permanently failed page raises an exception.

    The caller therefore cannot accidentally publish a
    partially retrieved window.
    """

    semaphore = asyncio.Semaphore(
        PARALLEL_PAGE_WORKERS
    )

    async def fetch_one(
        page_number: int,
        url: str,
    ):

        async with semaphore:

            html, status = (
                await request_page_html(
                    request_context,
                    url,
                    page_number,
                )
            )

            rows = extract_price_rows_from_html(
                html
            )

            if not rows:

                raise RuntimeError(
                    f"Expected PruAccess page "
                    f"{page_number} but it returned "
                    "zero BID observations."
                )

            return (
                page_number,
                {
                    "html":
                        html,

                    "status":
                        status,

                    "rows":
                        rows,

                    "url":
                        url,
                },
            )

    tasks = [
        asyncio.create_task(
            fetch_one(
                page_number,
                url,
            )
        )
        for page_number, url
        in page_jobs
    ]

    try:

        results = await asyncio.gather(
            *tasks
        )

    except Exception:

        for task in tasks:

            if not task.done():

                task.cancel()

        await asyncio.gather(
            *tasks,
            return_exceptions=True,
        )

        raise

    return {
        page_number:
            data

        for page_number, data
        in results
    }


# ============================================================
# FIRST PAGE + COMPLETE PAGINATION
# ============================================================

async def extract_all_pages(
    request_context,
    fund_id: str,
    start_date: str,
    end_date: str,
) -> tuple[list[dict], list[dict]]:

    print(
        "\n============================================================"
    )

    print(
        "STARTING PRUACCESS PAGINATION WINDOW"
    )

    print(
        f"Fund ID: {fund_id}"
    )

    print(
        f"Date range: {start_date} -> {end_date}"
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

    # --------------------------------------------------------
    # Page 1 is requested first because it tells us the
    # expected pagination count.
    # --------------------------------------------------------

    first_html, first_status = (
        await request_page_html(
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

    # --------------------------------------------------------
    # Determine final page.
    # --------------------------------------------------------

    if discovered_max_page is not None:

        if discovered_max_page > MAX_PAGES:

            raise RuntimeError(
                f"PruAccess reports "
                f"{discovered_max_page} pages, "
                f"which exceeds MAX_PAGES="
                f"{MAX_PAGES}."
            )

        max_page = discovered_max_page

    else:

        if len(first_rows) < PAGE_SIZE:

            max_page = 1

        else:

            # We do not accept an unknown pagination state
            # that could potentially exceed MAX_PAGES.
            #
            # We probe sequentially until a partial page is
            # found, but every page must still be successful.
            max_page = None

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

    # ========================================================
    # CASE A:
    # Explicit page count
    # ========================================================

    if max_page is not None:

        remaining_jobs = []

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

            remaining_jobs.append(
                (
                    page_number,
                    url,
                )
            )

        if remaining_jobs:

            print(
                "\nParallel page retrieval:"
            )

            print(
                f"  Pages: "
                f"2-{max_page}"
            )

            print(
                f"  Workers: "
                f"{PARALLEL_PAGE_WORKERS}"
            )

            page_results = (
                await fetch_pages_parallel(
                    request_context,
                    remaining_jobs,
                )
            )

            # ------------------------------------------------
            # Validate EVERY expected page.
            # ------------------------------------------------

            expected_pages = set(
                range(
                    1,
                    max_page + 1,
                )
            )

            actual_pages = set(
                page_results.keys()
            )

            actual_pages.add(
                1
            )

            missing_pages = sorted(
                expected_pages
                - actual_pages
            )

            if missing_pages:

                raise RuntimeError(
                    "Pagination incomplete. "
                    f"Missing pages: "
                    f"{missing_pages}"
                )

            # ------------------------------------------------
            # Preserve chronological page ordering as
            # returned by PruAccess.
            # ------------------------------------------------

            for page_number in range(
                2,
                max_page + 1,
            ):

                result = page_results[
                    page_number
                ]

                rows = result[
                    "rows"
                ]

                url = result[
                    "url"
                ]

                status = result[
                    "status"
                ]

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

        successful_pages = {
            item["page"]
            for item in page_diagnostics
            if item.get("status")
            == "success"
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
                "Complete pagination validation "
                "failed. Missing pages: "
                f"{missing_pages}"
            )

    # ========================================================
    # CASE B:
    # No explicit page count.
    #
    # We must continue until a short page is found.
    # These pages cannot be parallelized safely because the
    # termination condition is not known until each page is
    # inspected.
    # ========================================================

    else:

        current_page = 2

        while True:

            if current_page > MAX_PAGES:

                raise RuntimeError(
                    "Pagination reached MAX_PAGES "
                    f"({MAX_PAGES}) without determining "
                    "the final page."
                )

            url = build_page_url(
                PRUACCESS_URL,
                fund_id,
                "TBL",
                start_date,
                end_date,
                current_page,
            )

            print(
                f"\nFetching page "
                f"{current_page}"
            )

            html, status = (
                await request_page_html(
                    request_context,
                    url,
                    current_page,
                )
            )

            rows = extract_price_rows_from_html(
                html
            )

            if not rows:

                raise RuntimeError(
                    f"PruAccess page "
                    f"{current_page} returned "
                    "zero BID observations."
                )

            page_diagnostics.append(
                {
                    "page":
                        current_page,

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

            if len(rows) < PAGE_SIZE:

                print(
                    "\nFinal partial page detected."
                )

                break

            current_page += 1

    # ========================================================
    # FINAL PAGE SEQUENCE VALIDATION
    # ========================================================

    page_diagnostics.sort(
        key=lambda item:
            item["page"]
    )

    actual_pages = [
        item["page"]
        for item in page_diagnostics
        if item.get("status")
        == "success"
    ]

    expected_pages = list(
        range(
            1,
            len(page_diagnostics) + 1,
        )
    )

    if actual_pages != expected_pages:

        raise RuntimeError(
            "Pagination page sequence is "
            "incomplete or out of order."
        )

    # ========================================================
    # IMPORTANT:
    #
    # DO NOT DEDUPLICATE.
    #
    # Duplicates are intentionally preserved so validation
    # can detect them.
    # ========================================================

    def sort_date(row):

        return parse_pruaccess_date(
            row["date"]
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
# SUBMIT PRUACCESS SEARCH
# ============================================================

async def submit_pruaccess_search(
    page: Page,
    fund_id: str,
    start_date: str,
    end_date: str,
) -> dict:

    print(
        "\n=== Opening PruAccess ==="
    )

    await page.goto(
        PRUACCESS_URL,
        wait_until="domcontentloaded",
        timeout=INITIAL_PAGE_TIMEOUT_MS,
    )

    await page.wait_for_timeout(
        PRUACCESS_INITIAL_WAIT_MS
    )

    # --------------------------------------------------------
    # Fund
    # --------------------------------------------------------

    fund_selector = page.locator(
        "#fundName"
    )

    if await fund_selector.count() == 0:

        raise RuntimeError(
            "PruAccess #fundName "
            "selector not found."
        )

    await fund_selector.select_option(
        fund_id
    )

    # --------------------------------------------------------
    # Table view
    # --------------------------------------------------------

    view_selector = page.locator(
        "#viewType"
    )

    if await view_selector.count() == 0:

        raise RuntimeError(
            "PruAccess #viewType "
            "selector not found."
        )

    await view_selector.select_option(
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

    if await fund_price_type.count():

        fund_price_type_value = (
            await fund_price_type.input_value()
        )

        fund_price_type_label = clean_text(
            await fund_price_type.locator(
                "option:checked"
            ).inner_text()
        )

    else:

        fund_price_type_value = None

        fund_price_type_label = None

    print(
        "\nFund Price Type:"
    )

    print(
        f"{fund_price_type_value} "
        f"| {fund_price_type_label} "
        "(left untouched)"
    )

    # --------------------------------------------------------
    # Dates
    # --------------------------------------------------------

    await set_readonly_input_value(
        page,
        'input[name="startDate"]',
        start_date,
    )

    await set_readonly_input_value(
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

    if await csrf.count() == 0:

        raise RuntimeError(
            "PruAccess CSRF input not found."
        )

    csrf_value = await csrf.input_value()

    if not csrf_value:

        raise RuntimeError(
            "PruAccess CSRF token is empty."
        )

    # --------------------------------------------------------
    # Verify dates before submit.
    # --------------------------------------------------------

    actual_start = await page.locator(
        'input[name="startDate"]'
    ).input_value()

    actual_end = await page.locator(
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
    # Verify window length.
    # --------------------------------------------------------

    start_parsed = parse_pruaccess_date(
        start_date
    )

    end_parsed = parse_pruaccess_date(
        end_date
    )

    if (
        end_parsed
        < start_parsed
    ):

        raise RuntimeError(
            "PruAccess window has an end date "
            "before its start date."
        )

    max_allowed_end = (
        add_calendar_years(
            start_parsed,
            MAX_WINDOW_YEARS,
        )
        - timedelta(days=1)
    )

    if end_parsed > max_allowed_end:

        raise RuntimeError(
            "PruAccess date window exceeds "
            "the configured 10-year limit: "
            f"{start_date} -> {end_date}"
        )

    # --------------------------------------------------------
    # Submit.
    #
    # Normal form submission establishes the PruAccess
    # session/cookies.
    # --------------------------------------------------------

    print(
        "\n=== Submitting PruAccess form ==="
    )

    await page.locator(
        "#fundForm"
    ).evaluate(
        """
        form => {
            form.submit();
        }
        """
    )

    await page.wait_for_load_state(
        "domcontentloaded",
        timeout=INITIAL_PAGE_TIMEOUT_MS,
    )

    await page.wait_for_timeout(
        PRUACCESS_RESULT_WAIT_MS
    )

    # --------------------------------------------------------
    # Verify result.
    # --------------------------------------------------------

    result_text = clean_text(
        await page.locator(
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

        "fundPriceTypeLabel":
            fund_price_type_label,

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
# DUPLICATE VALIDATION
# ============================================================

def find_duplicate_dates(
    observations: list[dict],
) -> list[dict]:

    by_date = {}

    for observation in observations:

        observation_date = clean_text(
            observation.get(
                "date"
            )
        )

        by_date.setdefault(
            observation_date,
            [],
        ).append(
            observation
        )

    duplicates = []

    for observation_date, rows in by_date.items():

        if len(rows) > 1:

            duplicates.append(
                {
                    "date":
                        observation_date,

                    "count":
                        len(rows),

                    "observations":
                        rows,
                }
            )

    duplicates.sort(
        key=lambda item:
            parse_pruaccess_date(
                item["date"]
            )
    )

    return duplicates


# ============================================================
# CHRONOLOGICAL RECONSTRUCTION
# ============================================================

def reconstruct_history(
    window_results: list[dict],
) -> tuple[list[dict], list[dict]]:

    """
    Combine all successfully retrieved windows.

    No observations are removed.

    Returns:
        chronological_history
        duplicate_dates
    """

    all_rows = []

    for window_result in window_results:

        all_rows.extend(
            window_result[
                "observations"
            ]
        )

    if not all_rows:

        raise RuntimeError(
            "No historical observations exist "
            "after window reconstruction."
        )

    duplicates = find_duplicate_dates(
        all_rows
    )

    # --------------------------------------------------------
    # Preserve every raw observation.
    #
    # Sort chronologically.
    # --------------------------------------------------------

    chronological = sorted(
        all_rows,
        key=lambda row:
            parse_pruaccess_date(
                row["date"]
            )
    )

    return (
        chronological,
        duplicates,
    )


# ============================================================
# FULL-HISTORY VALIDATION
# ============================================================

def validate_full_history_to_inception(
    historical_rows: list[dict],
    inception_date: date,
    end_date: date,
    duplicate_dates: list[dict],
) -> dict:

    if not historical_rows:

        raise RuntimeError(
            "Historical BID history is empty."
        )

    # --------------------------------------------------------
    # Duplicate validation.
    # --------------------------------------------------------

    if duplicate_dates:

        duplicate_dates_text = ", ".join(
            item["date"]
            for item in duplicate_dates
        )

        raise RuntimeError(
            "Duplicate historical BID observations "
            "detected for date(s): "
            f"{duplicate_dates_text}"
        )

    # --------------------------------------------------------
    # Parse all dates.
    # --------------------------------------------------------

    parsed_dates = []

    for row in historical_rows:

        try:

            parsed_dates.append(
                parse_pruaccess_date(
                    row["date"]
                )
            )

        except Exception as error:

            raise RuntimeError(
                "Invalid historical BID date: "
                f"{row.get('date')}"
            ) from error

    oldest_date = min(
        parsed_dates
    )

    newest_date = max(
        parsed_dates
    )

    # --------------------------------------------------------
    # The oldest observation must be:
    #
    # inception date
    # OR
    # inception date + 1 calendar day
    #
    # Date gaps are otherwise allowed.
    # --------------------------------------------------------

    latest_allowed_oldest_date = (
        inception_date
        + timedelta(days=1)
    )

    if oldest_date < inception_date:

        raise RuntimeError(
            "Historical PruAccess data contains "
            "an observation earlier than the "
            "official Prudential inception date. "
            f"Inception={format_pruaccess_date(inception_date)}, "
            f"oldest={format_pruaccess_date(oldest_date)}"
        )

    if oldest_date > latest_allowed_oldest_date:

        raise RuntimeError(
            "Historical PruAccess data does not "
            "reach the Prudential inception date "
            "within the permitted +1 calendar day. "
            f"Inception="
            f"{format_pruaccess_date(inception_date)}, "
            f"oldest="
            f"{format_pruaccess_date(oldest_date)}"
        )

    # --------------------------------------------------------
    # Newest historical observation must not be after the
    # requested Singapore end date.
    # --------------------------------------------------------

    if newest_date > end_date:

        raise RuntimeError(
            "Historical PruAccess data contains "
            "an observation later than the requested "
            f"end date {format_pruaccess_date(end_date)}."
        )

    # --------------------------------------------------------
    # Chronological ordering.
    # --------------------------------------------------------

    for index in range(
        1,
        len(parsed_dates),
    ):

        if (
            parsed_dates[index]
            < parsed_dates[index - 1]
        ):

            raise RuntimeError(
                "Historical BID reconstruction "
                "is not chronological."
            )

    # --------------------------------------------------------
    # Date gaps are deliberately NOT treated as errors.
    # --------------------------------------------------------

    date_gaps = []

    for index in range(
        1,
        len(parsed_dates),
    ):

        previous_date = parsed_dates[
            index - 1
        ]

        current_date = parsed_dates[
            index
        ]

        gap_days = (
            current_date
            - previous_date
        ).days

        if gap_days > 1:

            date_gaps.append(
                {
                    "fromDate":
                        format_pruaccess_date(
                            previous_date
                        ),

                    "toDate":
                        format_pruaccess_date(
                            current_date
                        ),

                    "missingCalendarDays":
                        gap_days - 1,
                }
            )

    return {
        "status":
            "passed",

        "inceptionDate":
            format_pruaccess_date(
                inception_date
            ),

        "oldestHistoricalDate":
            format_pruaccess_date(
                oldest_date
            ),

        "newestHistoricalDate":
            format_pruaccess_date(
                newest_date
            ),

        "allowedLatestOldestDate":
            format_pruaccess_date(
                latest_allowed_oldest_date
            ),

        "inceptionCoverage":
            (
                "exact"
                if oldest_date == inception_date
                else "inception+1-day"
            ),

        "observationCount":
            len(
                historical_rows
            ),

        "duplicateDateCount":
            len(
                duplicate_dates
            ),

        "dateGapCount":
            len(
                date_gaps
            ),

        "dateGaps":
            date_gaps,
    }


# ============================================================
# WINDOW VALIDATION
# ============================================================

def validate_window_coverage(
    window: dict,
    observations: list[dict],
) -> dict:

    if not observations:

        raise RuntimeError(
            f"Window {window['windowNumber']} "
            "contains no observations."
        )

    window_start = window[
        "startDate"
    ]

    window_end = window[
        "endDate"
    ]

    parsed_dates = []

    for observation in observations:

        observation_date = (
            parse_pruaccess_date(
                observation["date"]
            )
        )

        if observation_date < window_start:

            raise RuntimeError(
                f"Window {window['windowNumber']} "
                "contains an observation before "
                "its requested start date: "
                f"{observation['date']}"
            )

        if observation_date > window_end:

            raise RuntimeError(
                f"Window {window['windowNumber']} "
                "contains an observation after "
                "its requested end date: "
                f"{observation['date']}"
            )

        parsed_dates.append(
            observation_date
        )

    # --------------------------------------------------------
    # Duplicate detection within the window.
    # --------------------------------------------------------

    duplicate_dates = (
        find_duplicate_dates(
            observations
        )
    )

    if duplicate_dates:

        dates = ", ".join(
            item["date"]
            for item in duplicate_dates
        )

        raise RuntimeError(
            f"Window {window['windowNumber']} "
            "contains duplicate historical "
            f"BID dates: {dates}"
        )

    return {
        "status":
            "passed",

        "windowNumber":
            window["windowNumber"],

        "requestedStartDate":
            window["startDateText"],

        "requestedEndDate":
            window["endDateText"],

        "oldestObservation":
            min(
                parsed_dates
            ).strftime(
                "%d-%b-%Y"
            ),

        "newestObservation":
            max(
                parsed_dates
            ).strftime(
                "%d-%b-%Y"
            ),

        "observationCount":
            len(
                observations
            ),

        "duplicateDateCount":
            0,
    }


# ============================================================
# SINGLE FUND EXTRACTION
# ============================================================

async def extract_single_fund(
    context: BrowserContext,
    page: Page,
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
        await get_prudential_fund_data(
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
    # DATES
    # ========================================================

    inception_datetime = (
        parse_prudential_date(
            prudential["inceptionDate"]
        )
    )

    inception_date = (
        inception_datetime.date()
    )

    end_date = singapore_today()

    print(
        "\nHistorical extraction:"
    )

    print(
        f"Inception: "
        f"{format_pruaccess_date(inception_date)}"
    )

    print(
        f"End: "
        f"{format_pruaccess_date(end_date)}"
    )

    # ========================================================
    # AUTOMATIC DATE WINDOWS
    # ========================================================

    windows = build_date_windows(
        inception_date,
        end_date,
    )

    print(
        f"\nPruAccess windows required: "
        f"{len(windows)}"
    )

    for window in windows:

        print(
            f"  Window {window['windowNumber']}: "
            f"{window['startDateText']} -> "
            f"{window['endDateText']} "
            f"({window['calendarDays']} calendar days)"
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
    # WINDOW EXTRACTION
    # ========================================================

    window_results = []

    window_diagnostics = []

    request_context = context.request

    for window in windows:

        window_number = window[
            "windowNumber"
        ]

        start_date = window[
            "startDateText"
        ]

        end_date_text = window[
            "endDateText"
        ]

        print(
            "\n\n"
            "------------------------------------------------------------"
        )

        print(
            f"WINDOW {window_number} / "
            f"{len(windows)}"
        )

        print(
            f"{start_date} -> {end_date_text}"
        )

        print(
            "------------------------------------------------------------"
        )

        # ----------------------------------------------------
        # Establish PruAccess session for this window.
        # ----------------------------------------------------

        pre_submit = (
            await submit_pruaccess_search(
                page,
                fund_id,
                start_date,
                end_date_text,
            )
        )

        # ----------------------------------------------------
        # Direct HTTP pagination using the same browser
        # context session cookies.
        # ----------------------------------------------------

        observations, page_diagnostics = (
            await extract_all_pages(
                request_context,
                fund_id,
                start_date,
                end_date_text,
            )
        )

        # ----------------------------------------------------
        # Window validation.
        # ----------------------------------------------------

        window_validation = (
            validate_window_coverage(
                window,
                observations,
            )
        )

        window_result = {
            "windowNumber":
                window_number,

            "startDate":
                start_date,

            "endDate":
                end_date_text,

            "fundId":
                fund_id,

            "fundPriceType":
                pre_submit.get(
                    "fundPriceType"
                ),

            "fundPriceTypeLabel":
                pre_submit.get(
                    "fundPriceTypeLabel"
                ),

            "historicalPriceType":
                "BID",

            "pageSize":
                PAGE_SIZE,

            "pagesExtracted":
                len(
                    page_diagnostics
                ),

            "observationCount":
                len(
                    observations
                ),

            "observations":
                observations,

            "pagination":
                page_diagnostics,

            "validation":
                window_validation,
        }

        window_results.append(
            window_result
        )

        window_diagnostics.append(
            {
                "windowNumber":
                    window_number,

                "startDate":
                    start_date,

                "endDate":
                    end_date_text,

                "pagesExtracted":
                    len(
                        page_diagnostics
                    ),

                "observationCount":
                    len(
                        observations
                    ),

                "oldestObservation":
                    window_validation[
                        "oldestObservation"
                    ],

                "newestObservation":
                    window_validation[
                        "newestObservation"
                    ],

                "status":
                    "success",
            }
        )

        print(
            f"\nWINDOW {window_number} SUCCESS"
        )

        print(
            f"Pages: "
            f"{len(page_diagnostics)}"
        )

        print(
            f"Observations: "
            f"{len(observations)}"
        )

    # ========================================================
    # CHRONOLOGICAL RECONSTRUCTION
    # ========================================================

    (
        chronological_history,
        duplicate_dates,
    ) = reconstruct_history(
        window_results
    )

    print(
        "\n============================================================"
    )

    print(
        "CHRONOLOGICAL RECONSTRUCTION"
    )

    print(
        "============================================================"
    )

    print(
        f"Total observations: "
        f"{len(chronological_history)}"
    )

    print(
        f"Duplicate dates: "
        f"{len(duplicate_dates)}"
    )

    # ========================================================
    # FULL HISTORY VALIDATION
    # ========================================================

    history_validation = (
        validate_full_history_to_inception(
            chronological_history,
            inception_date,
            end_date,
            duplicate_dates,
        )
    )

    print(
        "\n============================================================"
    )

    print(
        "FULL-HISTORY VALIDATION PASSED"
    )

    print(
        "============================================================"
    )

    print(
        f"Inception: "
        f"{history_validation['inceptionDate']}"
    )

    print(
        f"Oldest BID: "
        f"{history_validation['oldestHistoricalDate']}"
    )

    print(
        f"Newest BID: "
        f"{history_validation['newestHistoricalDate']}"
    )

    print(
        f"Inception coverage: "
        f"{history_validation['inceptionCoverage']}"
    )

    print(
        f"Date gaps: "
        f"{history_validation['dateGapCount']}"
    )

    # ========================================================
    # FINAL BID HISTORY
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
            format_pruaccess_date(
                inception_date
            ),

        "endDate":
            format_pruaccess_date(
                end_date
            ),

        "priceType":
            "BID",

        "pageSize":
            PAGE_SIZE,

        "windowCount":
            len(
                windows
            ),

        "observationCount":
            len(
                chronological_history
            ),

        "observations":
            chronological_history,
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
            format_pruaccess_date(
                inception_date
            ),

        "endDate":
            format_pruaccess_date(
                end_date
            ),

        "currentBidPrice":
            prudential.get(
                "bidPrice"
            ),

        "currentOfferPrice":
            prudential.get(
                "offerPrice"
            ),

        "historicalPriceType":
            "BID",

        "pageSize":
            PAGE_SIZE,

        "windowCount":
            len(
                windows
            ),

        "pagesExtracted":
            sum(
                item["pagesExtracted"]
                for item
                in window_diagnostics
            ),

        "historicalObservationCount":
            len(
                chronological_history
            ),

        "newestObservation":
            chronological_history[-1],

        "oldestObservation":
            chronological_history[0],

        "inceptionValidation":
            history_validation,

        "windows":
            window_diagnostics,
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
                format_pruaccess_date(
                    inception_date
                ),

            "endDate":
                format_pruaccess_date(
                    end_date
                ),

            "priceType":
                "BID",

            "pageSize":
                PAGE_SIZE,

            "windowCount":
                len(
                    windows
                ),

            "observationCount":
                len(
                    chronological_history
                ),
        },

        "dateWindows":
            windows,

        "windowResults":
            window_results,

        "windowDiagnostics":
            window_diagnostics,

        "inceptionValidation":
            history_validation,

        "summary":
            summary,

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
        / "date_windows.json",
        result["dateWindows"],
    )

    save_json(
        fund_output_dir
        / "window_diagnostics.json",
        result["windowDiagnostics"],
    )

    save_json(
        fund_output_dir
        / "inception_validation.json",
        result["inceptionValidation"],
    )

    save_json(
        fund_output_dir
        / "window_results.json",
        result["windowResults"],
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

async def main():

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
        "10-YEAR WINDOW + PARALLEL PAGINATION VERSION"
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

    async with async_playwright() as playwright:

        browser: Browser = (
            await playwright.chromium.launch(
                headless=BROWSER_HEADLESS
            )
        )

        context: BrowserContext = (
            await browser.new_context(
                viewport={
                    "width": 1440,
                    "height": 1000,
                }
            )
        )

        page = await context.new_page()

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

        await page.goto(
            PRUACCESS_URL,
            wait_until="domcontentloaded",
            timeout=INITIAL_PAGE_TIMEOUT_MS,
        )

        await page.wait_for_timeout(
            PRUACCESS_INITIAL_WAIT_MS
        )

        pruaccess_options = (
            await get_fund_options(
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

                result = (
                    await extract_single_fund(
                        context,
                        page,
                        excel_fund,
                        pruaccess_options,
                    )
                )

                # ------------------------------------------------
                # DO NOT save the fund until every window and
                # full-history validation has passed.
                # ------------------------------------------------

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
                    f"Windows: "
                    f"{result['pruAccess']['windowCount']}"
                )

                print(
                    f"Pages: "
                    f"{result['summary']['pagesExtracted']}"
                )

                print(
                    f"Historical BID observations: "
                    f"{result['pruAccess']['observationCount']}"
                )

                print(
                    f"Oldest BID: "
                    f"{result['summary']['oldestObservation']}"
                )

                print(
                    f"Newest BID: "
                    f"{result['summary']['newestObservation']}"
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

        await browser.close()

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

    total_windows = sum(
        fund["pruAccess"][
            "windowCount"
        ]
        for fund in successful_funds
    )

    total_pages = sum(
        fund["summary"][
            "pagesExtracted"
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

        "totalWindows":
            total_windows,

        "totalPages":
            total_pages,

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
    #
    # IMPORTANT:
    #
    # Any failed fund means the complete run is NOT a
    # successful production extraction.
    # ========================================================

    run_status = (
        "success"
        if (
            not failed_funds
            and len(successful_funds)
            == len(funds)
        )
        else "failed"
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

        "totalWindows":
            total_windows,

        "totalPages":
            total_pages,

        "totalHistoricalBidObservations":
            total_observations,

        "historicalPriceType":
            "BID",

        "maximumWindowYears":
            MAX_WINDOW_YEARS,

        "pageSize":
            PAGE_SIZE,

        "parallelPageWorkers":
            PARALLEL_PAGE_WORKERS,

        "pageRetryCount":
            PAGE_RETRY_COUNT,

        "pageRetryDelaySeconds":
            PAGE_RETRY_DELAY_SECONDS,

        "pageTimeoutMs":
            PAGE_TIMEOUT_MS,

        "inceptionCoverageRule":
            {
                "oldestObservationMustBeOnOrAfterInception":
                    True,

                "maximumAllowedDaysAfterInception":
                    1,

                "dateGapsAllowed":
                    True,

                "duplicateDatesAllowed":
                    False,

                "syntheticDataAllowed":
                    False,

                "interpolationAllowed":
                    False,

                "estimationAllowed":
                    False,
            },

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

                    "windowCount":
                        fund[
                            "pruAccess"
                        ].get(
                            "windowCount"
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

                    "inceptionValidation":
                        fund.get(
                            "inceptionValidation"
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
    # VALIDATION SUMMARY
    # ========================================================

    validation_summary = {
        "status":
            run_status,

        "generatedAtUtc":
            run_finished,

        "excelFundUniverse":
            len(
                funds
            ),

        "successfulFunds":
            len(
                successful_funds
            ),

        "failedFunds":
            len(
                failed_funds
            ),

        "totalHistoricalBidObservations":
            total_observations,

        "totalWindows":
            total_windows,

        "totalPages":
            total_pages,

        "rules": {
            "automaticTenYearWindows":
                True,

            "parallelPageRetrieval":
                True,

            "retryHandling":
                True,

            "completePaginationValidation":
                True,

            "noSyntheticData":
                True,

            "chronologicalReconstruction":
                True,

            "fullHistoryToInceptionValidation":
                True,

            "firstObservationMayBeOneDayAfterInception":
                True,

            "dateGapsAllowed":
                True,

            "duplicateDatesCauseFailure":
                True,

            "exactPruAccessNameMatch":
                True,

            "excelMasterUniverse":
                True,

            "hardcodedFundCount":
                False,
        },

        "failedFunds":
            failed_funds,
    }

    save_json(
        OUTPUT_DIR
        / "validation.json",
        validation_summary,
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
        f"Total windows: "
        f"{total_windows}"
    )

    print(
        f"Total pages: "
        f"{total_pages}"
    )

    print(
        f"Total historical BID observations: "
        f"{total_observations}"
    )

    print(
        "\nInception rule:"
    )

    print(
        " - Exact inception date: PASS"
    )

    print(
        " - Inception + 1 calendar day: PASS"
    )

    print(
        " - Date gaps: ALLOWED"
    )

    print(
        " - Synthetic data: FORBIDDEN"
    )

    print(
        " - Interpolation: FORBIDDEN"
    )

    print(
        " - Duplicate dates: FAIL"
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

    # ========================================================
    # PRODUCTION GATE
    # ========================================================

    if failed_funds:

        print(
            "\n============================================================"
        )

        print(
            "VALIDATION FAILED"
        )

        print(
            "============================================================"
        )

        print(
            "One or more Excel-master funds failed."
        )

        print(
            "The run must NOT be treated as a complete "
            "production dataset."
        )

        raise SystemExit(1)

    if len(successful_funds) != len(funds):

        print(
            "\n============================================================"
        )

        print(
            "VALIDATION FAILED"
        )

        print(
            "============================================================"
        )

        print(
            "Successful fund count does not equal "
            "the Excel master-universe count."
        )

        raise SystemExit(1)

    print(
        "\n============================================================"
    )

    print(
        "VALIDATION PASSED"
    )

    print(
        "============================================================"
    )

    print(
        f"{len(successful_funds)} / "
        f"{len(funds)} Excel-master funds passed."
    )

    print(
        "All historical BID windows passed."
    )

    print(
        "All pagination passed."
    )

    print(
        "All inception-coverage checks passed."
    )

    print(
        "No synthetic data was generated."
    )

    print(
        "Date gaps were allowed without filling."
    )

    print(
        "\nDone."
    )


# ============================================================
# ENTRY POINT
# ============================================================

if __name__ == "__main__":

    asyncio.run(
        main()
    )
