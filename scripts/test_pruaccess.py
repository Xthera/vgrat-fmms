#!/usr/bin/env python3

"""
PruAccess All-Fund Historical BID Price Extractor

Master source:
    Funds Links.xlsm

Column A:
    Prudential fund URL

Column B:
    PruAccess fund name

Architecture:
    Funds Links.xlsm
            |
            v
    Prudential fund page/API
            |
            +--> Current fund information
            +--> Inception date
            +--> Fund identifier
            |
            v
    PruAccess Fund Performance
            |
            +--> Historical BID prices
            |
            v
    output_pruaccess/

Important:
    - Excel determines the fund universe.
    - Prudential determines fund identity/current data.
    - PruAccess provides historical BID prices.
    - Fund Price Type is NOT manipulated.
    - No synthetic/interpolated/estimated data is created.
    - A fund failure is recorded explicitly.
    - Individual pagination pages are retried independently.
"""

from __future__ import annotations

import json
import re
import time
from datetime import datetime, timezone
from pathlib import Path
from urllib.parse import urljoin

from openpyxl import load_workbook
from playwright.sync_api import (
    sync_playwright,
    TimeoutError as PlaywrightTimeoutError,
)


# ============================================================================
# CONFIGURATION
# ============================================================================

EXCEL_FILE = Path("Funds Links.xlsm")

OUTPUT_DIR = Path("output_pruaccess")

FUNDS_DIR = OUTPUT_DIR / "funds"

ALL_FUNDS_FILE = OUTPUT_DIR / "all_funds.json"
ALL_HISTORY_FILE = OUTPUT_DIR / "all_bid_history.json"
RUN_SUMMARY_FILE = OUTPUT_DIR / "run_summary.json"
FUND_OPTIONS_FILE = OUTPUT_DIR / "fund_options.json"


PRUDENTIAL_BASE_URL = (
    "https://www.prudential.com.sg"
)

PRUACCESS_BASE_URL = (
    "https://pruaccess.prudential.com.sg"
)

PRUACCESS_PERFORMANCE_URL = (
    "https://pruaccess.prudential.com.sg/"
    "prulinkfund/viewFundPerformance.do"
)


PAGE_SIZE = 20

PAGE_RETRY_COUNT = 4

PAGE_RETRY_DELAY_SECONDS = 5

PAGE_TIMEOUT_MS = 120000

INITIAL_PAGE_TIMEOUT_MS = 120000

BROWSER_HEADLESS = True


# ============================================================================
# GENERAL HELPERS
# ============================================================================

def clean_text(value) -> str:

    if value is None:
        return ""

    text = str(value)

    text = text.replace(
        "\xa0",
        " ",
    )

    text = re.sub(
        r"\s+",
        " ",
        text,
    )

    return text.strip()


def normalize_text(value) -> str:

    return clean_text(
        value
    ).casefold()


def save_json(
    path: Path,
    data,
):

    path.parent.mkdir(
        parents=True,
        exist_ok=True,
    )

    with path.open(
        "w",
        encoding="utf-8",
    ) as file:

        json.dump(
            data,
            file,
            indent=2,
            ensure_ascii=False,
        )

        file.write("\n")


def load_json(
    path: Path,
):

    with path.open(
        "r",
        encoding="utf-8",
    ) as file:

        return json.load(file)


def utc_now_iso() -> str:

    return (
        datetime.now(
            timezone.utc
        ).isoformat()
    )


def parse_price(value) -> float:

    text = clean_text(
        value
    )

    text = (
        text
        .replace(",", "")
        .replace("S$", "")
        .replace("$", "")
    )

    match = re.search(
        r"-?\d+(?:\.\d+)?",
        text,
    )

    if not match:

        raise ValueError(
            f"Unable to parse price: {value}"
        )

    return float(
        match.group(0)
    )


# ============================================================================
# EXCEL
# ============================================================================

def read_excel_funds() -> list[dict]:

    if not EXCEL_FILE.exists():

        raise FileNotFoundError(
            f"Excel file not found: "
            f"{EXCEL_FILE}"
        )

    workbook = load_workbook(
        EXCEL_FILE,
        read_only=True,
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

    return funds


# ============================================================================
# PRUDENTIAL API
# ============================================================================

def build_prudential_json_url(
    prudential_url: str,
    citicode: str,
) -> str:

    base_url = prudential_url.split(
        "?",
        1,
    )[0].rstrip("/")

    if base_url.endswith(
        ".html"
    ):

        base_url = base_url[:-5]

    return (
        base_url
        + "/_jcr_content.ilpfunds.json"
        + "?citicodes="
        + citicode
    )


def build_prudential_series_url(
    prudential_url: str,
    citicode: str,
    range_name: str = "1y",
    series_type: str = "1",
) -> str:

    base_url = prudential_url.split(
        "?",
        1,
    )[0].rstrip("/")

    if base_url.endswith(
        ".html"
    ):

        base_url = base_url[:-5]

    url = (
        base_url
        + "/_jcr_content.ilpseries.json"
        + f"?citicode={citicode}"
        + f"&range={range_name}"
        + f"&seriestype={series_type}"
    )

    return url


def extract_citicode_from_url(
    prudential_url: str,
) -> str:

    match = re.search(
        r"[?&]citicode=([^&#]+)",
        prudential_url,
        re.IGNORECASE,
    )

    if match:
        return clean_text(
            match.group(1)
        )

    match = re.search(
        r"[?&]citicodes=([^&#]+)",
        prudential_url,
        re.IGNORECASE,
    )

    if match:
        return clean_text(
            match.group(1)
        )

    return ""


def fetch_prudential_fund_data(
    page,
    prudential_url: str,
    citicode: str,
) -> dict:

    api_url = build_prudential_json_url(
        prudential_url,
        citicode,
    )

    response = page.request.get(
        api_url,
        timeout=INITIAL_PAGE_TIMEOUT_MS,
    )

    if not response.ok:

        raise RuntimeError(
            "Prudential fund API failed: "
            f"HTTP {response.status}"
        )

    data = response.json()

    if isinstance(
        data,
        list,
    ):

        if not data:

            raise RuntimeError(
                "Prudential fund API returned "
                "an empty list."
            )

        return data[0]

    if isinstance(
        data,
        dict,
    ):

        if isinstance(
            data.get("data"),
            list,
        ):

            if not data["data"]:

                raise RuntimeError(
                    "Prudential fund API data "
                    "list is empty."
                )

            return data["data"][0]

        return data

    raise RuntimeError(
        "Unexpected Prudential API response."
    )


# ============================================================================
# PRUDENTIAL PAGE / CITICODE
# ============================================================================

def extract_citicode_from_page(
    page,
    prudential_url: str,
) -> str:

    existing = extract_citicode_from_url(
        prudential_url
    )

    if existing:
        return existing

    page.goto(
        prudential_url,
        wait_until="domcontentloaded",
        timeout=INITIAL_PAGE_TIMEOUT_MS,
    )

    current_url = page.url

    existing = extract_citicode_from_url(
        current_url
    )

    if existing:
        return existing

    html = page.content()

    patterns = [
        r'"fundIdentifier"\s*:\s*"([^"]+)"',
        r'"citicode"\s*:\s*"([^"]+)"',
        r'"citicodes"\s*:\s*"([^"]+)"',
        r'citicode=([A-Za-z0-9]+)',
    ]

    for pattern in patterns:

        match = re.search(
            pattern,
            html,
            re.IGNORECASE,
        )

        if match:

            return clean_text(
                match.group(1)
            )

    raise RuntimeError(
        "Unable to determine Prudential "
        "citicode/fund identifier."
    )


# ============================================================================
# PRUACCESS FUND OPTIONS
# ============================================================================

def extract_pruaccess_options(
    page,
) -> list[dict]:

    page.goto(
        PRUACCESS_PERFORMANCE_URL,
        wait_until="domcontentloaded",
        timeout=INITIAL_PAGE_TIMEOUT_MS,
    )

    page.wait_for_selector(
        "#fundName",
        timeout=INITIAL_PAGE_TIMEOUT_MS,
    )

    options = page.locator(
        "#fundName option"
    )

    results = []

    count = options.count()

    for index in range(
        count
    ):

        option = options.nth(
            index
        )

        name = clean_text(
            option.inner_text()
        )

        value = clean_text(
            option.get_attribute(
                "value"
            )
        )

        if not name:
            continue

        results.append(
            {
                "name": name,
                "value": value,
            }
        )

    return results


def find_pruaccess_fund(
    options: list[dict],
    excel_name: str,
) -> dict:

    target = normalize_text(
        excel_name
    )

    if not target:

        raise RuntimeError(
            "Excel PruAccess fund name is empty."
        )

    # ------------------------------------------------------------------------
    # Exact normalized match.
    # ------------------------------------------------------------------------

    exact_matches = [
        option
        for option in options
        if normalize_text(
            option["name"]
        )
        == target
    ]

    if len(exact_matches) == 1:

        return exact_matches[0]

    if len(exact_matches) > 1:

        raise RuntimeError(
            "Multiple exact PruAccess matches "
            f"found for '{excel_name}'."
        )

    # ------------------------------------------------------------------------
    # Common naming difference:
    #
    # Excel:
    #     PruLink Global Equity Fund (SGD)
    #
    # PruAccess may expose:
    #     PruLink Global Equity Fund
    #
    # Do not use fuzzy matching across unrelated funds.
    # Only remove the explicit currency suffix.
    # ------------------------------------------------------------------------

    target_without_currency = re.sub(
        r"\s*\((?:SGD|USD|AUD|GBP|EUR|HKD|JPY|MYR|CNY)\)\s*$",
        "",
        target,
        flags=re.IGNORECASE,
    ).strip()

    currency_matches = [
        option
        for option in options
        if normalize_text(
            re.sub(
                r"\s*\((?:SGD|USD|AUD|GBP|EUR|HKD|JPY|MYR|CNY)\)\s*$",
                "",
                option["name"],
                flags=re.IGNORECASE,
            ).strip()
        )
        == target_without_currency
    ]

    if len(currency_matches) == 1:

        return currency_matches[0]

    if len(currency_matches) > 1:

        raise RuntimeError(
            "Multiple currency-normalized PruAccess "
            f"matches found for '{excel_name}'."
        )

    raise RuntimeError(
        "No exact PruAccess fund match found "
        f"for '{excel_name}'."
    )


# ============================================================================
# DATE HANDLING
# ============================================================================

def parse_prudential_inception_date(
    value,
) -> str:

    text = clean_text(
        value
    )

    if not text:
        raise RuntimeError(
            "Prudential inception date is empty."
        )

    formats = [
        "%d/%m/%Y",
        "%d-%m-%Y",
        "%Y-%m-%d",
        "%d/%m/%y",
    ]

    for fmt in formats:

        try:

            parsed = datetime.strptime(
                text,
                fmt,
            )

            return parsed.strftime(
                "%d-%b-%Y"
            )

        except ValueError:
            continue

    raise RuntimeError(
        "Unable to parse Prudential inception date: "
        f"{value}"
    )


# ============================================================================
# PRUACCESS FORM
# ============================================================================

def submit_pruaccess_search(
    page,
    fund_id: str,
    start_date: str,
    end_date: str,
):

    page.goto(
        PRUACCESS_PERFORMANCE_URL,
        wait_until="domcontentloaded",
        timeout=INITIAL_PAGE_TIMEOUT_MS,
    )

    page.wait_for_selector(
        "#fundName",
        timeout=INITIAL_PAGE_TIMEOUT_MS,
    )

    # ------------------------------------------------------------------------
    # Select fund by value.
    # ------------------------------------------------------------------------

    page.locator(
        "#fundName"
    ).select_option(
        fund_id
    )

    # ------------------------------------------------------------------------
    # Set dates.
    #
    # We deliberately do NOT manipulate fundPriceType.
    # It is greyed out on PruAccess.
    # ------------------------------------------------------------------------

    page.locator(
        'input[name="startDate"]'
    ).fill(
        start_date
    )

    page.locator(
        'input[name="endDate"]'
    ).fill(
        end_date
    )

    # ------------------------------------------------------------------------
    # Force table view through the existing form control.
    # ------------------------------------------------------------------------

    view_type = page.locator(
        'input[name="viewType"][value="TBL"]'
    )

    if view_type.count() > 0:

        try:
            view_type.check()
        except Exception:
            pass

    # ------------------------------------------------------------------------
    # Submit normal HTML form.
    # ------------------------------------------------------------------------

    search_button = page.locator(
        "#search"
    )

    if search_button.count() == 0:

        raise RuntimeError(
            "PruAccess Search button not found."
        )

    search_button.click()

    page.wait_for_load_state(
        "domcontentloaded",
        timeout=INITIAL_PAGE_TIMEOUT_MS,
    )

    page.wait_for_timeout(
        1000
    )


# ============================================================================
# PRUACCESS TABLE EXTRACTION
# ============================================================================

def extract_current_table_rows(
    page,
) -> list[dict]:

    rows = page.locator(
        "table tbody tr"
    )

    count = rows.count()

    results = []

    for index in range(
        count
    ):

        row = rows.nth(
            index
        )

        cells = row.locator(
            "td"
        )

        cell_count = cells.count()

        if cell_count < 3:
            continue

        date_value = clean_text(
            cells.nth(0).inner_text()
        )

        bid_price = clean_text(
            cells.nth(1).inner_text()
        )

        offer_price = clean_text(
            cells.nth(2).inner_text()
        )

        # --------------------------------------------------------------------
        # Skip non-data rows.
        # --------------------------------------------------------------------

        if not date_value:
            continue

        if not re.search(
            r"\d",
            date_value,
        ):
            continue

        results.append(
            {
                "date": date_value,
                "bidPrice": bid_price,
                "offerPrice": offer_price,
            }
        )

    return results


def extract_pagination_metadata(
    page,
) -> dict:

    pages = []

    # ------------------------------------------------------------------------
    # Locate pagination links.
    # ------------------------------------------------------------------------

    links = page.locator(
        "a"
    )

    link_count = links.count()

    page_numbers = set()

    for index in range(
        link_count
    ):

        link = links.nth(
            index
        )

        text = clean_text(
            link.inner_text()
        )

        href = clean_text(
            link.get_attribute(
                "href"
            )
        )

        if not href:
            continue

        match = re.search(
            r"[?&]page\.page=(\d+)",
            href,
            re.IGNORECASE,
        )

        if match:

            page_number = int(
                match.group(1)
            )

            page_numbers.add(
                page_number
            )

        elif text.isdigit():

            page_number = int(
                text
            )

            if page_number > 0:
                page_numbers.add(
                    page_number
                )

    # ------------------------------------------------------------------------
    # Determine current page.
    # ------------------------------------------------------------------------

    current_page = 1

    current_page_candidates = page.locator(
        ".pagination .active, "
        ".pagination li.active, "
        "a[aria-current='page']"
    )

    if current_page_candidates.count() > 0:

        current_text = clean_text(
            current_page_candidates
            .first
            .inner_text()
        )

        if current_text.isdigit():

            current_page = int(
                current_text
            )

    # ------------------------------------------------------------------------
    # Extract current table rows.
    # ------------------------------------------------------------------------

    rows = extract_current_table_rows(
        page
    )

    row_count = len(
        rows
    )

    if page_numbers:

        max_page = max(
            page_numbers
        )

    else:

        max_page = current_page

    return {
        "page": current_page,
        "rowCount": row_count,
        "maxPageDetected": max_page,
        "rows": rows,
    }


# ============================================================================
# INDIVIDUAL PAGINATION PAGE
# ============================================================================

def build_pagination_url(
    fund_id: str,
    view_type: str,
    start_date: str,
    end_date: str,
    page_number: int,
) -> str:

    return (
        PRUACCESS_PERFORMANCE_URL
        + f"?fundId={fund_id}"
        + f"&viewType={view_type}"
        + f"&startDate={start_date}"
        + f"&endDate={end_date}"
        + f"&page.page={page_number}"
        + f"&page.size={PAGE_SIZE}"
    )


def extract_single_pagination_page(
    page,
    fund_id: str,
    start_date: str,
    end_date: str,
    page_number: int,
) -> dict:

    url = build_pagination_url(
        fund_id,
        "TBL",
        start_date,
        end_date,
        page_number,
    )

    last_error = None

    for attempt in range(
        1,
        PAGE_RETRY_COUNT + 1,
    ):

        try:

            print(
                f"      Page {page_number} "
                f"(attempt {attempt}/"
                f"{PAGE_RETRY_COUNT})"
            )

            page.goto(
                url,
                wait_until="domcontentloaded",
                timeout=PAGE_TIMEOUT_MS,
            )

            page.wait_for_timeout(
                500
            )

            rows = extract_current_table_rows(
                page
            )

            if not rows:

                # A page with zero rows is potentially a transient
                # PruAccess response. Retry it.
                raise RuntimeError(
                    f"Page {page_number} returned "
                    "zero data rows."
                )

            return {
                "page": page_number,
                "url": url,
                "rowCount": len(rows),
                "rows": rows,
                "status": "success",
                "attempts": attempt,
            }

        except (
            PlaywrightTimeoutError,
            Exception,
        ) as error:

            last_error = error

            print(
                f"      Retry required for page "
                f"{page_number}: {error}"
            )

            if attempt < PAGE_RETRY_COUNT:

                time.sleep(
                    PAGE_RETRY_DELAY_SECONDS
                )

    raise RuntimeError(
        f"Page {page_number} failed after "
        f"{PAGE_RETRY_COUNT} attempts: "
        f"{last_error}"
    )


# ============================================================================
# DISCOVER TOTAL PAGES
# ============================================================================

def discover_total_pages(
    page,
    fund_id: str,
    start_date: str,
    end_date: str,
) -> int:

    url = build_pagination_url(
        fund_id,
        "TBL",
        start_date,
        end_date,
        1,
    )

    page.goto(
        url,
        wait_until="domcontentloaded",
        timeout=PAGE_TIMEOUT_MS,
    )

    page.wait_for_timeout(
        500
    )

    pagination = extract_pagination_metadata(
        page
    )

    rows = pagination.get(
        "rows",
        []
    )

    max_page = pagination.get(
        "maxPageDetected",
        1,
    )

    # ------------------------------------------------------------------------
    # If pagination exposes a max page, use it.
    # ------------------------------------------------------------------------

    if max_page > 1:
        return max_page

    # ------------------------------------------------------------------------
    # If first page is already partial, there is only one page.
    # ------------------------------------------------------------------------

    if len(rows) < PAGE_SIZE:
        return 1

    # ------------------------------------------------------------------------
    # Otherwise walk pages until a partial page is encountered.
    #
    # This is intentionally conservative. We don't assume a fixed number
    # of pages because different funds have different histories.
    # ------------------------------------------------------------------------

    page_number = 2

    while True:

        result = extract_single_pagination_page(
            page,
            fund_id,
            start_date,
            end_date,
            page_number,
        )

        row_count = result[
            "rowCount"
        ]

        if row_count < PAGE_SIZE:

            return page_number

        page_number += 1


# ============================================================================
# EXTRACT COMPLETE FUND HISTORY
# ============================================================================

def extract_complete_bid_history(
    page,
    fund_id: str,
    start_date: str,
    end_date: str,
) -> tuple[list[dict], list[dict]]:

    total_pages = discover_total_pages(
        page,
        fund_id,
        start_date,
        end_date,
    )

    print(
        f"    Total pages detected: "
        f"{total_pages}"
    )

    all_observations = []

    pagination_pages = []

    for page_number in range(
        1,
        total_pages + 1,
    ):

        result = extract_single_pagination_page(
            page,
            fund_id,
            start_date,
            end_date,
            page_number,
        )

        rows = result[
            "rows"
        ]

        # --------------------------------------------------------------------
        # Store BID only.
        # --------------------------------------------------------------------

        for row in rows:

            all_observations.append(
                {
                    "date": row[
                        "date"
                    ],
                    "bidPrice": row[
                        "bidPrice"
                    ],
                }
            )

        pagination_pages.append(
            {
                "page": page_number,
                "url": result[
                    "url"
                ],
                "rowCount": result[
                    "rowCount"
                ],
                "attempts": result[
                    "attempts"
                ],
                "status": result[
                    "status"
                ],
            }
        )

    return (
        all_observations,
        pagination_pages,
    )


# ============================================================================
# FUND EXTRACTION
# ============================================================================

def extract_single_fund(
    browser,
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

    fund_dir = None

    context = browser.new_context()

    page = context.new_page()

    try:

        print()
        print(
            "-" * 70
        )

        print(
            f"Excel row: {excel_row}"
        )

        print(
            f"PruAccess name: "
            f"{excel_pruaccess_name}"
        )

        print(
            f"Prudential URL: "
            f"{prudential_url}"
        )

        # ====================================================================
        # Identify Prudential fund
        # ====================================================================

        page.goto(
            prudential_url,
            wait_until="domcontentloaded",
            timeout=INITIAL_PAGE_TIMEOUT_MS,
        )

        citicode = extract_citicode_from_page(
            page,
            prudential_url,
        )

        print(
            f"  Citicode / identifier: "
            f"{citicode}"
        )

        prudential_fund = (
            fetch_prudential_fund_data(
                page,
                prudential_url,
                citicode,
            )
        )

        fund_identifier = clean_text(
            prudential_fund.get(
                "fundIdentifier"
            )
        )

        if not fund_identifier:

            fund_identifier = citicode

        fund_name = clean_text(
            prudential_fund.get(
                "fundName"
            )
        )

        if not fund_name:

            raise RuntimeError(
                "Prudential fundName is empty."
            )

        fund_code = clean_text(
            prudential_fund.get(
                "fundCode"
            )
        )

        inception_raw = clean_text(
            prudential_fund.get(
                "inceptionDate"
            )
        )

        start_date = (
            parse_prudential_inception_date(
                inception_raw
            )
        )

        # --------------------------------------------------------------------
        # Current Prudential prices.
        # --------------------------------------------------------------------

        current_bid_price = clean_text(
            prudential_fund.get(
                "bidPrice"
            )
        )

        current_offer_price = clean_text(
            prudential_fund.get(
                "offerPrice"
            )
        )

        # ====================================================================
        # PruAccess matching
        # ====================================================================

        pruaccess_match = find_pruaccess_fund(
            pruaccess_options,
            excel_pruaccess_name,
        )

        pruaccess_fund_name = clean_text(
            pruaccess_match[
                "name"
            ]
        )

        pruaccess_fund_id = clean_text(
            pruaccess_match[
                "value"
            ]
        )

        if not pruaccess_fund_id:

            raise RuntimeError(
                "Matched PruAccess fund has no ID."
            )

        print(
            f"  Matched PruAccess fund: "
            f"{pruaccess_fund_name}"
        )

        print(
            f"  PruAccess ID: "
            f"{pruaccess_fund_id}"
        )

        # ====================================================================
        # End date
        # ====================================================================

        request_date = datetime.now(
            timezone.utc
        ).astimezone().strftime(
            "%d-%b-%Y"
        )

        # ====================================================================
        # Create fund directory
        # ====================================================================

        fund_dir = get_fund_directory(
            excel_row,
            fund_identifier,
        )

        fund_dir.mkdir(
            parents=True,
            exist_ok=True,
        )

        # ====================================================================
        # Extract PruAccess history
        # ====================================================================

        print(
            f"  Historical range: "
            f"{start_date} -> "
            f"{request_date}"
        )

        submit_pruaccess_search(
            page,
            pruaccess_fund_id,
            start_date,
            request_date,
        )

        observations, pagination_pages = (
            extract_complete_bid_history(
                page,
                pruaccess_fund_id,
                start_date,
                request_date,
            )
        )

        if not observations:

            raise RuntimeError(
                "No historical BID observations "
                "were extracted."
            )

        # ====================================================================
        # Verify observation ordering
        # ====================================================================

        parsed_dates = []

        for observation in observations:

            date_value = clean_text(
                observation[
                    "date"
                ]
            )

            try:

                parsed_date = datetime.strptime(
                    date_value,
                    "%d-%b-%Y",
                )

            except ValueError:

                parsed_date = datetime.strptime(
                    date_value,
                    "%d-%b-%y",
                )

            parsed_dates.append(
                parsed_date
            )

        # PruAccess should return newest -> oldest.
        for index in range(
            1,
            len(parsed_dates),
        ):

            if (
                parsed_dates[index]
                >= parsed_dates[index - 1]
            ):

                raise RuntimeError(
                    "Historical BID observations "
                    "are not strictly descending "
                    f"at index {index}."
                )

        # ====================================================================
        # Verify duplicate dates
        # ====================================================================

        date_values = [
            clean_text(
                observation[
                    "date"
                ]
            )
            for observation in observations
        ]

        if len(
            date_values
        ) != len(
            set(date_values)
        ):

            raise RuntimeError(
                "Duplicate historical dates "
                "detected."
            )

        # ====================================================================
        # Verify numeric BID prices
        # ====================================================================

        for observation in observations:

            price_text = clean_text(
                observation[
                    "bidPrice"
                ]
            )

            price = parse_price(
                price_text
            )

            if price <= 0:

                raise RuntimeError(
                    "Invalid non-positive BID "
                    f"price: {price_text}"
                )

        # ====================================================================
        # Verify history reaches inception
        # ====================================================================

        oldest_date = parsed_dates[-1]

        inception_date = datetime.strptime(
            start_date,
            "%d-%b-%Y",
        )

        if (
            oldest_date.date()
            > inception_date.date()
        ):

            raise RuntimeError(
                "Historical BID data does not "
                "reach the Prudential inception "
                f"date. "
                f"Inception={start_date}; "
                f"oldest={observations[-1]['date']}."
            )

        # ====================================================================
        # Pagination total
        # ====================================================================

        pagination_row_total = sum(
            page_info[
                "rowCount"
            ]
            for page_info in pagination_pages
        )

        if (
            pagination_row_total
            != len(observations)
        ):

            raise RuntimeError(
                "Pagination row total does not "
                "match observation count. "
                f"Pagination={pagination_row_total}; "
                f"observations={len(observations)}."
            )

        # ====================================================================
        # Pagination full-page validation
        # ====================================================================

        if len(
            pagination_pages
        ) > 1:

            for page_info in pagination_pages[
                :-1
            ]:

                if (
                    page_info[
                        "rowCount"
                    ]
                    != PAGE_SIZE
                ):

                    raise RuntimeError(
                        "A non-final pagination page "
                        "does not contain "
                        f"{PAGE_SIZE} rows: "
                        f"page={page_info['page']}; "
                        f"rows={page_info['rowCount']}"
                    )

        final_page_count = (
            pagination_pages[-1][
                "rowCount"
            ]
        )

        if (
            final_page_count <= 0
            or final_page_count > PAGE_SIZE
        ):

            raise RuntimeError(
                "Invalid final pagination "
                f"row count: "
                f"{final_page_count}"
            )

        # ====================================================================
        # Build outputs
        # ====================================================================

        bid_history = {
            "source": "PruAccess",
            "priceType": "BID",
            "fundName": fund_name,
            "fundIdentifier": fund_identifier,
            "fundCode": fund_code,
            "pruAccessFundName": pruaccess_fund_name,
            "pruAccessFundId": pruaccess_fund_id,
            "startDate": start_date,
            "endDate": request_date,
            "observationCount": len(
                observations
            ),
            "observations": observations,
        }

        pagination_output = {
            "source": "PruAccess",
            "priceType": "BID",
            "fundIdentifier": fund_identifier,
            "pruAccessFundId": pruaccess_fund_id,
            "pageSize": PAGE_SIZE,
            "pagesExtracted": len(
                pagination_pages
            ),
            "rowCount": pagination_row_total,
            "pages": pagination_pages,
        }

        summary = {
            "status": "success",
            "excelRow": excel_row,
            "prudentialUrl": prudential_url,
            "excelPruAccessName": excel_pruaccess_name,
            "matchedPruAccessName": pruaccess_fund_name,
            "pruAccessFundId": pruaccess_fund_id,
            "fundIdentifier": fund_identifier,
            "fundName": fund_name,
            "fundCode": fund_code,
            "inceptionDate": inception_raw,
            "startDate": start_date,
            "endDate": request_date,
            "currentBidPrice": current_bid_price,
            "currentOfferPrice": current_offer_price,
            "historicalPriceType": "BID",
            "historicalObservationCount": len(
                observations
            ),
            "pagesExtracted": len(
                pagination_pages
            ),
            "newestObservation": observations[
                0
            ],
            "oldestObservation": observations[
                -1
            ],
            "source": "PruAccess",
        }

        save_json(
            fund_dir / "bid_history.json",
            bid_history,
        )

        save_json(
            fund_dir / "pagination.json",
            pagination_output,
        )

        save_json(
            fund_dir / "summary.json",
            summary,
        )

        save_json(
            fund_dir / "prudential_fund.json",
            prudential_fund,
        )

        print(
            f"  SUCCESS: "
            f"{len(observations)} BID observations"
        )

        print(
            f"  Newest: "
            f"{observations[0]['date']} "
            f"BID={observations[0]['bidPrice']}"
        )

        print(
            f"  Oldest: "
            f"{observations[-1]['date']} "
            f"BID={observations[-1]['bidPrice']}"
        )

        return summary

    except Exception as error:

        # ====================================================================
        # Explicit failure record
        # ====================================================================

        if fund_dir is None:

            fallback_identifier = (
                extract_citicode_from_url(
                    prudential_url
                )
                or f"row{excel_row}"
            )

            fund_dir = get_fund_directory(
                excel_row,
                fallback_identifier,
            )

        fund_dir.mkdir(
            parents=True,
            exist_ok=True,
        )

        failure = {
            "status": "failed",
            "excelRow": excel_row,
            "prudentialUrl": prudential_url,
            "excelPruAccessName": excel_pruaccess_name,
            "errorType": type(
                error
            ).__name__,
            "error": str(
                error
            ),
            "failedAtUtc": utc_now_iso(),
        }

        save_json(
            fund_dir / "failure.json",
            failure,
        )

        print(
            f"  FAILED: {error}"
        )

        return failure

    finally:

        context.close()


# ============================================================================
# FUND DIRECTORY
# ============================================================================

def get_fund_directory(
    excel_row: int,
    fund_identifier: str,
) -> Path:

    safe_identifier = re.sub(
        r"[^A-Za-z0-9._-]+",
        "_",
        clean_text(
            fund_identifier
        ),
    )

    return (
        FUNDS_DIR
        / f"{excel_row}_{safe_identifier}"
    )


# ============================================================================
# CONSOLIDATION
# ============================================================================

def build_all_bid_history(
    successful_results: list[dict],
) -> dict:

    consolidated = {
        "source": "PruAccess",
        "priceType": "BID",
        "fundCount": len(
            successful_results
        ),
        "totalHistoricalBidObservations": 0,
        "funds": {},
    }

    for summary in successful_results:

        identifier = summary[
            "fundIdentifier"
        ]

        fund_dir = get_fund_directory(
            summary[
                "excelRow"
            ],
            identifier,
        )

        history_file = (
            fund_dir
            / "bid_history.json"
        )

        history = load_json(
            history_file
        )

        observations = history[
            "observations"
        ]

        consolidated[
            "funds"
        ][identifier] = {
            "excelRow": summary[
                "excelRow"
            ],
            "fundName": summary[
                "fundName"
            ],
            "fundCode": summary[
                "fundCode"
            ],
            "pruAccessFundId": summary[
                "pruAccessFundId"
            ],
            "source": "PruAccess",
            "priceType": "BID",
            "observationCount": len(
                observations
            ),
            "observations": observations,
        }

        consolidated[
            "totalHistoricalBidObservations"
        ] += len(
            observations
        )

    return consolidated


def build_all_funds(
    successful_results: list[dict],
    failed_results: list[dict],
) -> dict:

    return {
        "generatedAtUtc": utc_now_iso(),
        "fundUniverseCount": (
            len(successful_results)
            + len(failed_results)
        ),
        "successfulFundCount": len(
            successful_results
        ),
        "failedFundCount": len(
            failed_results
        ),
        "funds": successful_results,
        "failedFunds": failed_results,
    }


# ============================================================================
# MAIN
# ============================================================================

def main() -> int:

    OUTPUT_DIR.mkdir(
        parents=True,
        exist_ok=True,
    )

    FUNDS_DIR.mkdir(
        parents=True,
        exist_ok=True,
    )

    print("=" * 70)
    print(
        "PruAccess All-Fund Historical BID Extraction"
    )
    print("=" * 70)

    print()
    print(
        f"Excel source: "
        f"{EXCEL_FILE}"
    )

    # ========================================================================
    # Read Excel universe
    # ========================================================================

    excel_funds = read_excel_funds()

    print(
        f"Fund universe: "
        f"{len(excel_funds)}"
    )

    if not excel_funds:

        raise RuntimeError(
            "No funds found in Funds Links.xlsm."
        )

    # ========================================================================
    # Start browser
    # ========================================================================

    successful_results = []
    failed_results = []

    with sync_playwright() as playwright:

        browser = playwright.chromium.launch(
            headless=BROWSER_HEADLESS
        )

        context = browser.new_context()

        page = context.new_page()

        # ====================================================================
        # Retrieve PruAccess fund options once.
        # ====================================================================

        print()
        print(
            "Loading PruAccess fund list..."
        )

        pruaccess_options = (
            extract_pruaccess_options(
                page
            )
        )

        print(
            f"PruAccess options found: "
            f"{len(pruaccess_options)}"
        )

        save_json(
            FUND_OPTIONS_FILE,
            {
                "source": PRUACCESS_PERFORMANCE_URL,
                "retrievedAtUtc": utc_now_iso(),
                "count": len(
                    pruaccess_options
                ),
                "options": pruaccess_options,
            },
        )

        context.close()

        # ====================================================================
        # Process every Excel fund.
        # ====================================================================

        for index, excel_fund in enumerate(
            excel_funds,
            start=1,
        ):

            print()
            print(
                "=" * 70
            )

            print(
                f"FUND {index}/"
                f"{len(excel_funds)}"
            )

            print(
                f"Excel row: "
                f"{excel_fund['excelRow']}"
            )

            try:

                result = extract_single_fund(
                    browser,
                    excel_fund,
                    pruaccess_options,
                )

                if result.get(
                    "status"
                ) == "success":

                    successful_results.append(
                        result
                    )

                else:

                    failed_results.append(
                        result
                    )

            except Exception as error:

                # This should rarely be reached because
                # extract_single_fund already catches errors.
                failure = {
                    "status": "failed",
                    "excelRow": excel_fund[
                        "excelRow"
                    ],
                    "prudentialUrl": excel_fund[
                        "prudentialUrl"
                    ],
                    "excelPruAccessName": (
                        excel_fund[
                            "pruAccessName"
                        ]
                    ),
                    "errorType": type(
                        error
                    ).__name__,
                    "error": str(
                        error
                    ),
                    "failedAtUtc": utc_now_iso(),
                }

                failed_results.append(
                    failure
                )

                print(
                    f"UNEXPECTED FAILURE: "
                    f"{error}"
                )

        browser.close()

    # ========================================================================
    # Consolidate
    # ========================================================================

    all_funds = build_all_funds(
        successful_results,
        failed_results,
    )

    all_bid_history = (
        build_all_bid_history(
            successful_results
        )
    )

    total_observations = (
        all_bid_history[
            "totalHistoricalBidObservations"
        ]
    )

    # ========================================================================
    # Run summary
    # ========================================================================

    run_status = (
        "success"
        if len(
            failed_results
        ) == 0
        else
        "partial"
    )

    run_summary = {
        "status": run_status,
        "startedAtUtc": None,
        "completedAtUtc": utc_now_iso(),
        "fundUniverseCount": len(
            excel_funds
        ),
        "successfulFundCount": len(
            successful_results
        ),
        "failedFundCount": len(
            failed_results
        ),
        "totalHistoricalBidObservations": (
            total_observations
        ),
        "historicalPriceType": "BID",
        "pageSize": PAGE_SIZE,
        "pageRetryCount": PAGE_RETRY_COUNT,
        "successfulFunds": successful_results,
        "failedFunds": failed_results,
    }

    # ========================================================================
    # Save consolidated files
    # ========================================================================

    save_json(
        ALL_FUNDS_FILE,
        all_funds,
    )

    save_json(
        ALL_HISTORY_FILE,
        all_bid_history,
    )

    save_json(
        RUN_SUMMARY_FILE,
        run_summary,
    )

    # ========================================================================
    # Final report
    # ========================================================================

    print()
    print("=" * 70)
    print(
        "EXTRACTION COMPLETE"
    )
    print("=" * 70)

    print(
        f"Fund universe: "
        f"{len(excel_funds)}"
    )

    print(
        f"Successful funds: "
        f"{len(successful_results)}"
    )

    print(
        f"Failed funds: "
        f"{len(failed_results)}"
    )

    print(
        f"Historical BID observations: "
        f"{total_observations}"
    )

    print()
    print(
        f"Output: "
        f"{OUTPUT_DIR}"
    )

    if failed_results:

        print()
        print(
            "FAILED FUNDS:"
        )

        for failure in failed_results:

            print(
                f"  Excel row "
                f"{failure.get('excelRow')}: "
                f"{failure.get('excelPruAccessName')}"
            )

            print(
                f"    Error: "
                f"{failure.get('error')}"
            )

    print(
        "=" * 70
    )

    # ------------------------------------------------------------------------
    # IMPORTANT:
    #
    # Extraction itself exits successfully even when some funds fail.
    # The validator is responsible for deciding whether the complete
    # 67-fund dataset is acceptable.
    # ------------------------------------------------------------------------

    return 0


if __name__ == "__main__":

    raise SystemExit(
        main()
    )
