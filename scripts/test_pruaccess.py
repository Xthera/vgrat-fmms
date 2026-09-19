#!/usr/bin/env python3

"""
PruAccess historical BID price extractor.

MASTER SOURCE
=============

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
    Prudential fund page
          |
          v
    Official Prudential ilpfunds.json
          |
          v
    Get fund identity/current data/inception
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
    Submit search
          |
          v
    Discover pagination
          |
          v
    Fetch every page
          |
          v
    Extract BID prices only
          |
          v
    Validate completeness
          |
          v
    Save individual + consolidated JSON

IMPORTANT RULES
===============

1. Funds Links.xlsm determines the fund universe.
2. Column A determines which funds exist.
3. Column B provides the exact PruAccess fund name.
4. No hardcoded fund count.
5. No fuzzy fund matching.
6. No synthetic data.
7. No estimated data.
8. No interpolation.
9. No fabricated data.
10. No carry-forward values.
11. Historical prices come directly from PruAccess.
12. Historical price type is BID only.
13. Offer prices are not historical BID prices.
14. Fund Price Type is never changed.
15. Start date is Prudential inception date.
16. End date is current Singapore date.
17. Every pagination page must successfully extract.
18. Temporary page failures are retried.
19. A permanently failed page fails the entire fund.
20. Partial historical data is never marked successful.
21. Extraction preserves returned rows; validation detects duplicates.
22. Prudential fund identity is authoritative.
23. PruAccess is not allowed to define the fund universe.
"""

from __future__ import annotations

import json
import re
import time
from datetime import datetime, timezone
from pathlib import Path
from zoneinfo import ZoneInfo
from urllib.parse import urlencode

from openpyxl import load_workbook
from playwright.sync_api import sync_playwright


# ============================================================
# CONFIG
# ============================================================

EXCEL_FILE = Path("Funds Links.xlsm")

OUTPUT_DIR = Path("output_pruaccess")

FUNDS_OUTPUT_DIR = OUTPUT_DIR / "funds"

PRUACCESS_URL = (
    "https://pruaccess.prudential.com.sg/"
    "prulinkfund/viewFundPerformance.do"
)

PAGE_SIZE = 20

MAX_PAGES = 1000

PAGE_RETRY_COUNT = 4

PAGE_RETRY_DELAY_SECONDS = 5

PAGE_TIMEOUT_MS = 120000

INITIAL_PAGE_TIMEOUT_MS = 120000

PRUDENTIAL_WAIT_MS = 3000

PRUACCESS_INITIAL_WAIT_MS = 1500

PRUACCESS_RESULT_WAIT_MS = 2500

SINGAPORE_TIMEZONE = ZoneInfo("Asia/Singapore")


# ============================================================
# HELPERS
# ============================================================

def clean_text(value) -> str:
    """Normalize whitespace while preserving meaningful text."""

    if value is None:
        return ""

    return re.sub(
        r"\s+",
        " ",
        str(value),
    ).strip()


def normalize_name(value: str) -> str:
    """
    Normalize names for exact comparison.

    This is NOT fuzzy matching.

    Only harmless formatting differences are normalized:
        - whitespace
        - case
        - en dash
        - em dash
    """

    value = clean_text(value).lower()

    value = value.replace("–", "-")
    value = value.replace("—", "-")

    value = re.sub(
        r"\s+",
        " ",
        value,
    )

    return value.strip()


def utc_now() -> datetime:
    """Return timezone-aware UTC datetime."""

    return datetime.now(timezone.utc)


def utc_iso() -> str:
    """Return timezone-aware UTC ISO timestamp."""

    return utc_now().isoformat().replace(
        "+00:00",
        "Z",
    )


def singapore_today_string() -> str:
    """
    Return the current Singapore calendar date in
    PruAccess format.
    """

    now_singapore = datetime.now(
        SINGAPORE_TIMEZONE
    )

    return now_singapore.strftime(
        "%d-%b-%Y"
    )


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
        clean_text(value),
        "%d/%m/%Y",
    )


def pruaccess_date(
    value: datetime,
) -> str:

    return value.strftime(
        "%d-%b-%Y"
    )


def parse_pruaccess_date(
    value: str,
) -> datetime:

    return datetime.strptime(
        clean_text(value),
        "%d-%b-%Y",
    )


def safe_filename(
    value: str,
) -> str:

    value = clean_text(value)

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


def is_valid_price(
    value: str,
) -> bool:

    return bool(
        re.fullmatch(
            r"\d+(?:\.\d+)?",
            clean_text(value),
        )
    )


def price_as_float(
    value: str,
) -> float:

    return float(
        clean_text(value)
    )


# ============================================================
# EXCEL
# ============================================================

def read_excel_funds() -> list[dict]:
    """
    Read every populated Column A URL.

    Column A is the master universe.

    Column B is the PruAccess exact lookup name.
    """

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
            f"Excel file not found: {EXCEL_FILE}"
        )

    workbook = load_workbook(
        EXCEL_FILE,
        read_only=True,
        keep_vba=True,
        data_only=True,
    )

    worksheet = workbook.active

    funds: list[dict] = []

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
        # Column A determines the universe.
        # ----------------------------------------------------

        if not prudential_url:
            continue

        funds.append(
            {
                "excelRow": row_number,
                "prudentialUrl": prudential_url,
                "pruAccessName": pruaccess_name,
            }
        )

    workbook.close()

    if not funds:

        raise RuntimeError(
            "No populated Prudential fund URLs "
            "were found in Column A."
        )

    print(
        f"\nPopulated fund URLs found: {len(funds)}"
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
    """
    Open the official Prudential fund page and capture
    the official ilpfunds.json request.

    The structured Prudential API is used rather than
    depending on HTML layout.
    """

    print(
        "\n=== Opening Prudential fund page ==="
    )

    print(
        prudential_url
    )

    captured_urls: list[str] = []

    def handle_response(response):

        response_url = response.url

        if (
            "ilpfunds.json"
            in response_url
        ):
            captured_urls.append(
                response_url
            )

    page.on(
        "response",
        handle_response,
    )

    try:

        page.goto(
            prudential_url,
            wait_until="domcontentloaded",
            timeout=PAGE_TIMEOUT_MS,
        )

        page.wait_for_timeout(
            PRUDENTIAL_WAIT_MS
        )

    finally:

        page.remove_listener(
            "response",
            handle_response,
        )

    # --------------------------------------------------------
    # Remove duplicates while preserving order.
    # --------------------------------------------------------

    unique_urls = []

    seen_urls = set()

    for url in captured_urls:

        if url not in seen_urls:

            seen_urls.add(url)

            unique_urls.append(url)

    if not unique_urls:

        raise RuntimeError(
            "Could not capture official "
            "Prudential ilpfunds.json request."
        )

    api_url = unique_urls[0]

    print(
        "\nPrudential API:"
    )

    print(
        api_url
    )

    response = page.request.get(
        api_url,
        timeout=PAGE_TIMEOUT_MS,
    )

    if not response.ok:

        raise RuntimeError(
            "Prudential ilpfunds.json returned "
            f"HTTP {response.status}"
        )

    try:

        data = response.json()

    except Exception as error:

        raise RuntimeError(
            "Could not parse Prudential "
            f"API JSON: {error}"
        ) from error

    if (
        not isinstance(data, list)
        or not data
        or not isinstance(data[0], dict)
    ):

        raise RuntimeError(
            "Unexpected Prudential ilpfunds.json "
            "response structure."
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
            "Prudential fundIdentifier is missing."
        )

    if not fund_name:

        raise RuntimeError(
            "Prudential fundName is missing."
        )

    if not inception_date:

        raise RuntimeError(
            "Prudential inceptionDate is missing."
        )

    # Validate inception date immediately.
    try:

        parse_prudential_date(
            inception_date
        )

    except ValueError as error:

        raise RuntimeError(
            "Invalid Prudential inceptionDate: "
            f"{inception_date}"
        ) from error

    result = {
        "apiUrl": api_url,

        "fundIdentifier":
            fund_identifier,

        "fundName":
            fund_name,

        "fundCode":
            clean_text(
                fund.get("fundCode")
            ),

        "fundCurrency":
            clean_text(
                fund.get("fundCurrency")
            ),

        "unitCurrency":
            clean_text(
                fund.get("unitCurrency")
            ),

        "assetClass":
            clean_text(
                fund.get("assetClass")
            ),

        "assetSubClass":
            clean_text(
                fund.get("assetSubClass")
            ),

        "riskClassification":
            clean_text(
                fund.get("riskClassification")
            ),

        "bidPrice":
            clean_text(
                fund.get("bidPrice")
            ),

        "offerPrice":
            clean_text(
                fund.get("offerPrice")
            ),

        "valuationDate":
            clean_text(
                fund.get("valuationDate")
            ),

        "inceptionDate":
            inception_date,

        "cumulativeYtd":
            clean_text(
                fund.get("cumulativeYtd")
            ),

        "cumulative1m":
            clean_text(
                fund.get("cumulative1m")
            ),

        "cumulative3m":
            clean_text(
                fund.get("cumulative3m")
            ),

        "cumulative6m":
            clean_text(
                fund.get("cumulative6m")
            ),

        "cumulative1y":
            clean_text(
                fund.get("cumulative1y")
            ),

        "cumulative3y":
            clean_text(
                fund.get("cumulative3y")
            ),

        "cumulative5y":
            clean_text(
                fund.get("cumulative5y")
            ),

        "annualised3y":
            clean_text(
                fund.get("annualised3y")
            ),

        "annualised5y":
            clean_text(
                fund.get("annualised5y")
            ),

        "annualised10y":
            clean_text(
                fund.get("annualised10y")
            ),

        "annualisedSinceLaunch":
            clean_text(
                fund.get("annualisedSinceLaunch")
            ),

        "factsheetUrl":
            clean_text(
                fund.get("factsheetUrl")
            ),

        "prospectusUrl":
            clean_text(
                fund.get("prospectusUrl")
            ),

        "productHighlightSheetUrl":
            clean_text(
                fund.get(
                    "productHighlightSheetUrl"
                )
            ),

        "annualReportUrl":
            clean_text(
                fund.get("annualReportUrl")
            ),

        "fundObjective":
            clean_text(
                fund.get("fundObjective")
            ),

        "investmentManager":
            clean_text(
                fund.get("investmentManager")
            ),

        "hasDividend":
            fund.get("hasDividend"),

        "dividendRate":
            clean_text(
                fund.get("dividendRate")
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
        f"\nPruAccess fund options found: {count}"
    )

    result: list[dict] = []

    for index in range(count):

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
                    clean_text(
                        option.get_attribute(
                            "value"
                        )
                    ),
            }
        )

    if not result:

        raise RuntimeError(
            "PruAccess fund option list is empty."
        )

    return result


# ============================================================
# EXACT PRUACCESS MATCH
# ============================================================

def find_exact_pruaccess_match(
    options: list[dict],
    target_name: str,
) -> dict:
    """
    Exact normalized match only.

    No fuzzy matching.
    """

    target_normalized = normalize_name(
        target_name
    )

    if not target_normalized:

        raise RuntimeError(
            "Excel PruAccess fund name is empty."
        )

    matches = []

    for option in options:

        option_name = normalize_name(
            option.get("text", "")
        )

        if option_name == target_normalized:

            matches.append(
                option
            )

    if not matches:

        raise RuntimeError(
            "Exact PruAccess fund-name match "
            f"failed for: {target_name}"
        )

    if len(matches) > 1:

        raise RuntimeError(
            "Multiple exact PruAccess fund-name "
            f"matches found for: {target_name}"
        )

    selected = matches[0]

    if not selected.get("value"):

        raise RuntimeError(
            "Exact PruAccess match has no fund ID "
            f"for: {target_name}"
        )

    return selected


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
            f"Could not find input: {selector}"
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
            f"Expected {value}, got {actual}."
        )


# ============================================================
# PRICE TABLE EXTRACTION
# ============================================================

def extract_price_rows(
    page,
) -> list[dict]:
    """
    Extract actual rows from the PruAccess price table.

    No deduplication occurs here.

    Validation is responsible for detecting duplicates.
    """

    tables = page.locator(
        "table"
    )

    table_count = tables.count()

    all_rows: list[dict] = []

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

            # PruAccess table date format:
            # 17-Sep-2026
            if not re.fullmatch(
                r"\d{1,2}-[A-Za-z]{3}-\d{4}",
                date_value,
            ):
                continue

            bid_value = values[1]

            if not is_valid_price(
                bid_value
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

    return all_rows


# ============================================================
# PAGE URL
# ============================================================

def build_page_url(
    base_url: str,
    fund_id: str,
    view_type: str,
    start_date: str,
    end_date: str,
    page_number: int,
) -> str:

    query = urlencode(
        {
            "fundId":
                fund_id,

            "viewType":
                view_type,

            "startDate":
                start_date,

            "endDate":
                end_date,

            "page.page":
                page_number,

            "page.size":
                PAGE_SIZE,
        }
    )

    return (
        f"{base_url}?{query}"
    )


# ============================================================
# PAGINATION DISCOVERY
# ============================================================

def discover_max_page(
    page,
) -> int | None:
    """
    Inspect pagination controls.

    Returns the largest discovered page number.

    Returns None when no reliable page number can be
    discovered, allowing sequential pagination fallback.
    """

    candidates: list[int] = []

    # --------------------------------------------------------
    # Look at links.
    # --------------------------------------------------------

    links = page.locator(
        "a"
    )

    link_count = links.count()

    for index in range(
        link_count
    ):

        link = links.nth(
            index
        )

        href = link.get_attribute(
            "href"
        )

        text = clean_text(
            link.inner_text()
        )

        combined = " ".join(
            [
                href or "",
                text,
            ]
        )

        for match in re.findall(
            r"(?:page\.page=|page=)(\d+)",
            combined,
            flags=re.IGNORECASE,
        ):

            try:

                candidates.append(
                    int(match)
                )

            except ValueError:
                pass

    # --------------------------------------------------------
    # Look at pagination text.
    # --------------------------------------------------------

    body_text = clean_text(
        page.locator(
            "body"
        ).inner_text()
    )

    for match in re.findall(
        r"(?:page\s*)?(\d+)\s*(?:of|/)\s*(\d+)",
        body_text,
        flags=re.IGNORECASE,
    ):

        try:

            candidates.append(
                int(match[1])
            )

        except (ValueError, IndexError):
            pass

    if not candidates:
        return None

    max_page = max(
        candidates
    )

    if max_page < 1:
        return None

    if max_page > MAX_PAGES:

        raise RuntimeError(
            "Discovered pagination exceeds "
            f"MAX_PAGES={MAX_PAGES}: {max_page}"
        )

    return max_page


# ============================================================
# FETCH PAGE WITH RETRIES
# ============================================================

def fetch_page_with_retries(
    page,
    url: str,
    page_number: int,
) -> tuple[list[dict], dict]:
    """
    Fetch one PruAccess page.

    Retries:
        PAGE_RETRY_COUNT

    A page with zero rows is treated as a temporary failure
    until all retries are exhausted.

    A page is successful only when actual price rows are found.
    """

    last_error = ""

    for attempt in range(
        1,
        PAGE_RETRY_COUNT + 1,
    ):

        print(
            f"\nPage {page_number} "
            f"attempt {attempt}/{PAGE_RETRY_COUNT}"
        )

        print(
            url
        )

        started = time.monotonic()

        try:

            page.goto(
                url,
                wait_until="domcontentloaded",
                timeout=PAGE_TIMEOUT_MS,
            )

            page.wait_for_timeout(
                PAGE_WAIT_MS
            )

            rows = extract_price_rows(
                page
            )

            elapsed = (
                time.monotonic()
                - started
            )

            if not rows:

                raise RuntimeError(
                    "No BID rows found on page."
                )

            diagnostic = {
                "page":
                    page_number,

                "url":
                    url,

                "attempts":
                    attempt,

                "rowCount":
                    len(rows),

                "firstDate":
                    rows[0]["date"],

                "lastDate":
                    rows[-1]["date"],

                "status":
                    "success",

                "elapsedSeconds":
                    round(
                        elapsed,
                        3,
                    ),
            }

            print(
                f"Rows found: {len(rows)}"
            )

            print(
                f"Date range: "
                f"{rows[0]['date']} → "
                f"{rows[-1]['date']}"
            )

            return (
                rows,
                diagnostic,
            )

        except Exception as error:

            last_error = clean_text(
                str(error)
            )

            elapsed = (
                time.monotonic()
                - started
            )

            print(
                f"Page {page_number} "
                f"attempt {attempt} failed: "
                f"{last_error}"
            )

            if attempt < PAGE_RETRY_COUNT:

                print(
                    f"Retrying in "
                    f"{PAGE_RETRY_DELAY_SECONDS} "
                    f"seconds..."
                )

                time.sleep(
                    PAGE_RETRY_DELAY_SECONDS
                )

    raise RuntimeError(
        f"PruAccess page {page_number} "
        f"failed after {PAGE_RETRY_COUNT} "
        f"attempts. Last error: "
        f"{last_error}"
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
    Extract every PruAccess page.

    Strategy:

    1. Fetch page 1.
    2. Attempt to discover maximum page.
    3. If maximum page is known, fetch every page 2..N.
    4. If maximum page is not known, continue sequentially.
    5. Every non-final page must contain PAGE_SIZE rows.
    6. Final page must contain 1..PAGE_SIZE rows.
    7. A failed page fails the entire fund.
    """

    print(
        "\n============================================================"
    )

    print(
        "STARTING COMPLETE PRUACCESS PAGINATION"
    )

    print(
        "============================================================"
    )

    all_rows: list[dict] = []

    page_diagnostics: list[dict] = []

    # --------------------------------------------------------
    # Page 1.
    # --------------------------------------------------------

    first_url = build_page_url(
        PRUACCESS_URL,
        fund_id,
        "TBL",
        start_date,
        end_date,
        1,
    )

    first_rows, first_diagnostic = (
        fetch_page_with_retries(
            page,
            first_url,
            1,
        )
    )

    all_rows.extend(
        first_rows
    )

    page_diagnostics.append(
        first_diagnostic
    )

    discovered_max_page = (
        discover_max_page(
            page
        )
    )

    if discovered_max_page:

        print(
            "\nPagination discovered:"
        )

        print(
            f"Maximum page: "
            f"{discovered_max_page}"
        )

    else:

        print(
            "\nCould not reliably discover "
            "maximum page number."
        )

        print(
            "Using sequential pagination."
        )

    # --------------------------------------------------------
    # Known maximum page.
    # --------------------------------------------------------

    if discovered_max_page:

        for page_number in range(
            2,
            discovered_max_page + 1,
        ):

            url = build_page_url(
                PRUACCESS_URL,
                fund_id,
                "TBL",
                start_date,
                end_date,
                page_number,
            )

            rows, diagnostic = (
                fetch_page_with_retries(
                    page,
                    url,
                    page_number,
                )
            )

            # ------------------------------------------------
            # Every page except final must contain PAGE_SIZE.
            # ------------------------------------------------

            if (
                page_number
                < discovered_max_page
                and len(rows) != PAGE_SIZE
            ):

                raise RuntimeError(
                    f"Pagination integrity failure: "
                    f"page {page_number} returned "
                    f"{len(rows)} rows; expected "
                    f"{PAGE_SIZE}."
                )

            if (
                page_number
                == discovered_max_page
                and not (
                    1
                    <= len(rows)
                    <= PAGE_SIZE
                )
            ):

                raise RuntimeError(
                    f"Final page {page_number} "
                    f"returned invalid row count: "
                    f"{len(rows)}."
                )

            all_rows.extend(
                rows
            )

            page_diagnostics.append(
                diagnostic
            )

    # --------------------------------------------------------
    # Sequential fallback.
    # --------------------------------------------------------

    else:

        page_number = 2

        while True:

            if page_number > MAX_PAGES:

                raise RuntimeError(
                    f"Pagination exceeded "
                    f"MAX_PAGES={MAX_PAGES}."
                )

            url = build_page_url(
                PRUACCESS_URL,
                fund_id,
                "TBL",
                start_date,
                end_date,
                page_number,
            )

            # ------------------------------------------------
            # We cannot use an empty page as success.
            # Instead, fetch it with retries. If it is empty
            # after all retries, the fund fails because we
            # cannot prove pagination completeness.
            # ------------------------------------------------

            try:

                rows, diagnostic = (
                    fetch_page_with_retries(
                        page,
                        url,
                        page_number,
                    )
                )

            except RuntimeError as error:

                # If page 1 was full and a later page cannot
                # be retrieved, this is a real extraction
                # failure, not a legitimate end condition.
                raise RuntimeError(
                    f"Could not retrieve sequential "
                    f"pagination page {page_number}: "
                    f"{error}"
                ) from error

            all_rows.extend(
                rows
            )

            page_diagnostics.append(
                diagnostic
            )

            if len(rows) < PAGE_SIZE:

                print(
                    "\nFinal page detected."
                )

                break

            page_number += 1

    # --------------------------------------------------------
    # Basic pagination integrity.
    # --------------------------------------------------------

    if not page_diagnostics:

        raise RuntimeError(
            "No pagination diagnostics were generated."
        )

    expected_pages = len(
        page_diagnostics
    )

    actual_pages = [
        diagnostic.get("page")
        for diagnostic in page_diagnostics
    ]

    expected_page_numbers = list(
        range(
            1,
            expected_pages + 1,
        )
    )

    if actual_pages != expected_page_numbers:

        raise RuntimeError(
            "Pagination page-number sequence is "
            f"invalid: {actual_pages}"
        )

    # --------------------------------------------------------
    # Ensure all rows are valid.
    # --------------------------------------------------------

    if not all_rows:

        raise RuntimeError(
            "No historical BID observations "
            "were extracted."
        )

    for index, row in enumerate(
        all_rows,
        start=1,
    ):

        if not isinstance(
            row,
            dict,
        ):

            raise RuntimeError(
                f"Historical row {index} "
                "is not an object."
            )

        if not clean_text(
            row.get("date")
        ):

            raise RuntimeError(
                f"Historical row {index} "
                "has no date."
            )

        if not is_valid_price(
            row.get("bidPrice", "")
        ):

            raise RuntimeError(
                f"Historical row {index} "
                "has invalid BID price: "
                f"{row.get('bidPrice')}"
            )

    # --------------------------------------------------------
    # Preserve every extracted row.
    #
    # Do NOT deduplicate here.
    # --------------------------------------------------------

    final_rows = sorted(
        all_rows,
        key=lambda row: parse_pruaccess_date(
            row["date"]
        ),
        reverse=True,
    )

    # --------------------------------------------------------
    # Confirm chronological range is sensible.
    # --------------------------------------------------------

    newest = parse_pruaccess_date(
        final_rows[0]["date"]
    )

    oldest = parse_pruaccess_date(
        final_rows[-1]["date"]
    )

    if newest < oldest:

        raise RuntimeError(
            "Historical date range is invalid."
        )

    print(
        "\n============================================================"
    )

    print(
        "PAGINATION COMPLETE"
    )

    print(
        "============================================================"
    )

    print(
        f"Pages extracted: "
        f"{len(page_diagnostics)}"
    )

    print(
        f"Historical observations: "
        f"{len(final_rows)}"
    )

    print(
        f"Newest: "
        f"{final_rows[0]['date']}"
    )

    print(
        f"Oldest: "
        f"{final_rows[-1]['date']}"
    )

    return (
        final_rows,
        page_diagnostics,
    )


# ============================================================
# SINGLE FUND EXTRACTION
# ============================================================

def extract_single_fund(
    context,
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
        f"\nPrudential URL:\n{prudential_url}"
    )

    print(
        f"\nExcel PruAccess name:\n"
        f"{excel_pruaccess_name}"
    )

    # ========================================================
    # Create isolated page.
    # ========================================================

    page = context.new_page()

    try:

        # ====================================================
        # PRUDENTIAL
        # ====================================================

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

        # ====================================================
        # DATE RANGE
        # ====================================================

        inception_raw = prudential.get(
            "inceptionDate"
        )

        if not inception_raw:

            raise RuntimeError(
                "Prudential inception date is missing."
            )

        inception_date = (
            parse_prudential_date(
                inception_raw
            )
        )

        start_date = pruaccess_date(
            inception_date
        )

        end_date = singapore_today_string()

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
        # EXACT PRUACCESS MATCH
        # ====================================================

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

        fund_id = clean_text(
            selected.get("value")
        )

        if not fund_id:

            raise RuntimeError(
                "Matched PruAccess fund "
                "has no fund ID."
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
            timeout=INITIAL_PAGE_TIMEOUT_MS,
        )

        page.wait_for_timeout(
            PRUACCESS_INITIAL_WAIT_MS
        )

        # ====================================================
        # SELECT FUND
        # ====================================================

        fund_selector = page.locator(
            "#fundName"
        )

        if fund_selector.count() == 0:

            raise RuntimeError(
                "PruAccess #fundName selector "
                "not found."
            )

        fund_selector.select_option(
            fund_id
        )

        selected_value = clean_text(
            fund_selector.input_value()
        )

        if selected_value != fund_id:

            raise RuntimeError(
                "PruAccess fund selection "
                "verification failed. "
                f"Expected {fund_id}, "
                f"got {selected_value}."
            )

        # ====================================================
        # TABLE VIEW
        # ====================================================

        view_selector = page.locator(
            "#viewType"
        )

        if view_selector.count() == 0:

            raise RuntimeError(
                "PruAccess #viewType selector "
                "not found."
            )

        view_selector.select_option(
            "TBL"
        )

        selected_view = clean_text(
            view_selector.input_value()
        )

        if selected_view != "TBL":

            raise RuntimeError(
                "Could not set PruAccess "
                "view type to TBL."
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
                clean_text(
                    fund_price_type.input_value()
                )
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

        csrf_value = clean_text(
            csrf.input_value()
        )

        if not csrf_value:

            raise RuntimeError(
                "PruAccess CSRF token is empty."
            )

        # ====================================================
        # VERIFY DATES
        # ====================================================

        actual_start = clean_text(
            page.locator(
                'input[name="startDate"]'
            ).input_value()
        )

        actual_end = clean_text(
            page.locator(
                'input[name="endDate"]'
            ).input_value()
        )

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
                bool(csrf_value),
        }

        # ====================================================
        # SUBMIT FORM
        # ====================================================

        print(
            "\n=== Submitting PruAccess form ==="
        )

        form = page.locator(
            "#fundForm"
        )

        if form.count() == 0:

            raise RuntimeError(
                "PruAccess #fundForm not found."
            )

        form.evaluate(
            """
            form => {
                form.submit();
            }
            """
        )

        page.wait_for_load_state(
            "domcontentloaded",
            timeout=PAGE_TIMEOUT_MS,
        )

        page.wait_for_timeout(
            PRUACCESS_RESULT_WAIT_MS
        )

        # ====================================================
        # Verify result contains actual price rows.
        # ====================================================

        first_result_rows = (
            extract_price_rows(
                page
            )
        )

        if not first_result_rows:

            raise RuntimeError(
                "PruAccess form submission did not "
                "produce historical BID rows."
            )

        print(
            f"\nInitial submitted result contains "
            f"{len(first_result_rows)} price rows."
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

        if not historical_rows:

            raise RuntimeError(
                "No historical BID observations "
                "were extracted."
            )

        # ====================================================
        # CREATE BID HISTORY
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

        # ====================================================
        # SUMMARY
        # ====================================================

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

        # ====================================================
        # CONSOLIDATED FUND RECORD
        # ====================================================

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

    finally:

        try:

            page.close()

        except Exception:

            pass


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

    directory_name = (
        f"{excel_row}_"
        f"{safe_filename("
            f"{fund_identifier or fund_code or fund_name}"
        )}"
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


# ============================================================
# SAVE FAILURE
# ============================================================

def save_failed_fund(
    failure: dict,
) -> None:

    excel_row = failure[
        "excelRow"
    ]

    failure_output_dir = (
        FUNDS_OUTPUT_DIR
        / f"{excel_row}_failed"
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

    run_started = utc_now()

    successful_funds: list[dict] = []

    failed_funds: list[dict] = []

    all_bid_history: dict = {}

    all_fund_records: list[dict] = []

    # ========================================================
    # PLAYWRIGHT
    # ========================================================

    with sync_playwright() as playwright:

        browser = playwright.chromium.launch(
            headless=True
        )

        # ====================================================
        # Load PruAccess options in a dedicated context.
        # ====================================================

        options_context = (
            browser.new_context(
                viewport={
                    "width": 1440,
                    "height": 1000,
                }
            )
        )

        options_page = (
            options_context.new_page()
        )

        try:

            print(
                "\n============================================================"
            )

            print(
                "LOADING PRUACCESS FUND OPTIONS"
            )

            print(
                "============================================================"
            )

            options_page.goto(
                PRUACCESS_URL,
                wait_until="domcontentloaded",
                timeout=INITIAL_PAGE_TIMEOUT_MS,
            )

            options_page.wait_for_timeout(
                PRUACCESS_INITIAL_WAIT_MS
            )

            pruaccess_options = (
                get_fund_options(
                    options_page
                )
            )

            save_json(
                OUTPUT_DIR
                / "fund_options.json",
                {
                    "source":
                        PRUACCESS_URL,

                    "retrievedAtUtc":
                        utc_iso(),

                    "count":
                        len(
                            pruaccess_options
                        ),

                    "options":
                        pruaccess_options,
                },
            )

        finally:

            options_context.close()

        # ====================================================
        # Process every Excel fund.
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

            # ------------------------------------------------
            # Every fund gets its own isolated context.
            # ------------------------------------------------

            fund_context = (
                browser.new_context(
                    viewport={
                        "width": 1440,
                        "height": 1000,
                    }
                )
            )

            try:

                result = extract_single_fund(
                    fund_context,
                    excel_fund,
                    pruaccess_options,
                )

                # --------------------------------------------
                # Successful fund.
                # --------------------------------------------

                all_fund_records.append(
                    result
                )

                successful_funds.append(
                    result
                )

                save_successful_fund(
                    result
                )

                prudential = result[
                    "prudential"
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
                    f"{len(result['pagination'])}"
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
                    f"Error: {error_text}"
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

                    "failedAtUtc":
                        utc_iso(),
                }

                failed_funds.append(
                    failure
                )

                save_failed_fund(
                    failure
                )

            finally:

                try:

                    fund_context.close()

                except Exception:

                    pass

        browser.close()

    # ========================================================
    # CONSOLIDATED FUND DATA
    # ========================================================

    run_finished = utc_now()

    all_funds = {
        "source":
            "Funds Links.xlsm",

        "generatedAtUtc":
            run_finished.isoformat()
            .replace(
                "+00:00",
                "Z",
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

        "funds":
            all_fund_records,

        "failedFunds":
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

    total_observations = sum(
        record[
            "bidHistory"
        ][
            "observationCount"
        ]
        for record in successful_funds
    )

    consolidated_bid_history = {
        "source":
            "PruAccess",

        "priceType":
            "BID",

        "generatedAtUtc":
            run_finished.isoformat()
            .replace(
                "+00:00",
                "Z",
            ),

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

    run_summary = {
        "status":
            (
                "success"
                if not failed_funds
                else "partial"
            ),

        "startedAtUtc":
            run_started.isoformat()
            .replace(
                "+00:00",
                "Z",
            ),

        "completedAtUtc":
            run_finished.isoformat()
            .replace(
                "+00:00",
                "Z",
            ),

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

                    "pagesExtracted":
                        len(
                            fund[
                                "pagination"
                            ]
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
            f"{len(fund['pagination'])} pages "
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
