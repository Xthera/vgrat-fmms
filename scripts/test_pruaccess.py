#!/usr/bin/env python3
"""
Extract Prudential fund information and complete historical BID prices
from PruAccess for every fund listed in Funds Links.xlsm.

SOURCE OF TRUTH
----------------
Funds Links.xlsm
    Column A = Prudential fund URL
    Column B = PruAccess fund name

PRUDENTIAL
----------
Current fund information is extracted from the official Prudential fund page
JSON endpoint.

PRUACCESS
---------
Historical daily BID prices are extracted from:
    https://pruaccess.prudential.com.sg/prulinkfund/viewFundPerformance.do

IMPORTANT
---------
- Excel determines the fund universe.
- Prudential determines fund identity and current fund information.
- PruAccess provides historical BID prices.
- Fund Price Type is NOT manipulated.
- No synthetic/interpolated/estimated historical data is created.
- If one page times out, that page is retried independently.
- Previously extracted pages are preserved.
- A fund is only marked failed after repeated page failures.
"""

from __future__ import annotations

import json
import re
import time
from datetime import datetime
from pathlib import Path
from urllib.parse import parse_qs, urlencode, urlparse, urlunparse

from openpyxl import load_workbook
from playwright.sync_api import (
    Browser,
    BrowserContext,
    Page,
    TimeoutError as PlaywrightTimeoutError,
    sync_playwright,
)


# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------

EXCEL_FILE = Path("Funds Links.xlsm")

OUTPUT_DIR = Path("output_pruaccess")
FUNDS_OUTPUT_DIR = OUTPUT_DIR / "funds"

PRUACCESS_URL = (
    "https://pruaccess.prudential.com.sg/prulinkfund/viewFundPerformance.do"
)

PAGE_SIZE = 20
MAX_PAGES = 1000

# Page-level retry settings.
PAGE_RETRY_COUNT = 4
PAGE_RETRY_DELAY_SECONDS = 5

# Navigation timeout for individual PruAccess pages.
PAGE_TIMEOUT_MS = 120000

# Initial Prudential page timeout.
PRUDENTIAL_TIMEOUT_MS = 120000

# Small delay between successful pagination requests.
PAGE_DELAY_SECONDS = 0.5


# ---------------------------------------------------------------------------
# Generic helpers
# ---------------------------------------------------------------------------

def utc_now_iso() -> str:
    return datetime.utcnow().isoformat() + "Z"


def save_json(path: Path, payload) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)

    with path.open("w", encoding="utf-8") as handle:
        json.dump(
            payload,
            handle,
            indent=2,
            ensure_ascii=False,
        )


def normalize_text(value: str | None) -> str:
    if value is None:
        return ""

    text = str(value)

    text = text.replace("\xa0", " ")
    text = re.sub(r"\s+", " ", text)
    text = text.strip()

    return text.casefold()


def clean_text(value: str | None) -> str:
    if value is None:
        return ""

    text = str(value)
    text = text.replace("\xa0", " ")
    text = re.sub(r"\s+", " ", text)

    return text.strip()


def safe_filename(value: str) -> str:
    value = re.sub(r"[^A-Za-z0-9._-]+", "_", value)
    return value[:120]


def extract_citicode_from_url(url: str) -> str:
    """
    Extract citicode/citicode-like query parameter if present.

    The Excel URLs may contain:
        ?citicode=D39F
    or
        ?citicodes=D39F
    """

    parsed = urlparse(url)
    query = parse_qs(parsed.query)

    for key in ("citicode", "citicodes"):
        values = query.get(key)

        if values:
            value = clean_text(values[0])

            if value:
                return value

    return ""


# ---------------------------------------------------------------------------
# Excel
# ---------------------------------------------------------------------------

def read_excel_funds() -> list[dict]:
    """
    Read every non-empty URL from Column A starting at row 2.

    Column B contains the corresponding PruAccess fund name.

    Excel determines the fund universe.
    """

    if not EXCEL_FILE.exists():
        raise FileNotFoundError(
            f"Excel file not found: {EXCEL_FILE}"
        )

    workbook = load_workbook(
        EXCEL_FILE,
        read_only=True,
        data_only=True,
    )

    worksheet = workbook.active

    funds: list[dict] = []

    for row_number in range(2, worksheet.max_row + 1):

        prudential_url = worksheet.cell(
            row=row_number,
            column=1,
        ).value

        pruaccess_name = worksheet.cell(
            row=row_number,
            column=2,
        ).value

        prudential_url = clean_text(prudential_url)
        pruaccess_name = clean_text(pruaccess_name)

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


# ---------------------------------------------------------------------------
# Prudential fund data
# ---------------------------------------------------------------------------

def get_prudential_fund_data(
    page: Page,
    prudential_url: str,
) -> dict:
    """
    Load the official Prudential fund page and extract the fund JSON data.

    The Prudential page exposes:

        _jcr_content.ilpfunds.json?citicodes=XXXX

    The citicode is taken from the Excel URL.
    """

    citicode = extract_citicode_from_url(prudential_url)

    if not citicode:
        raise RuntimeError(
            "Unable to extract citicode from Prudential URL."
        )

    parsed = urlparse(prudential_url)

    base_path = parsed.path

    if not base_path.endswith("/"):
        base_path = base_path + "/"

    json_path = (
        base_path.rstrip("/")
        + "/_jcr_content.ilpfunds.json"
    )

    json_url = urlunparse(
        (
            parsed.scheme,
            parsed.netloc,
            json_path,
            "",
            urlencode({"citicodes": citicode}),
            "",
        )
    )

    response = page.request.get(
        json_url,
        timeout=PRUDENTIAL_TIMEOUT_MS,
    )

    if not response.ok:
        raise RuntimeError(
            f"Prudential fund API returned HTTP {response.status}"
        )

    try:
        payload = response.json()
    except Exception as error:
        raise RuntimeError(
            f"Unable to parse Prudential JSON: {error}"
        ) from error

    if isinstance(payload, dict):
        data = payload.get("data")

        if isinstance(data, list):
            payload = data

    if not isinstance(payload, list):
        raise RuntimeError(
            "Unexpected Prudential fund JSON structure."
        )

    if not payload:
        raise RuntimeError(
            "Prudential fund API returned no fund records."
        )

    # Prefer exact citicode match.
    selected = None

    for item in payload:

        if not isinstance(item, dict):
            continue

        identifier = clean_text(
            item.get("fundIdentifier")
        )

        if normalize_text(identifier) == normalize_text(citicode):
            selected = item
            break

    if selected is None:
        selected = payload[0]

    return {
        "source": "Prudential",
        "sourceUrl": prudential_url,
        "citicode": citicode,
        "retrievedAtUtc": utc_now_iso(),
        "data": selected,
    }


# ---------------------------------------------------------------------------
# PruAccess fund options
# ---------------------------------------------------------------------------

def get_fund_options(page: Page) -> list[dict]:
    """
    Read all fund options from the PruAccess fund selector.
    """

    page.goto(
        PRUACCESS_URL,
        wait_until="domcontentloaded",
        timeout=PAGE_TIMEOUT_MS,
    )

    page.wait_for_selector(
        "#fundName",
        timeout=PAGE_TIMEOUT_MS,
    )

    options = page.locator(
        "#fundName option"
    )

    count = options.count()

    result: list[dict] = []

    for index in range(count):

        option = options.nth(index)

        value = clean_text(
            option.get_attribute("value")
        )

        text = clean_text(
            option.inner_text()
        )

        if not value and not text:
            continue

        result.append(
            {
                "value": value,
                "name": text,
            }
        )

    return result


def find_exact_pruaccess_match(
    options: list[dict],
    target_name: str,
) -> dict | None:
    """
    Exact normalized match only.

    We deliberately do not use fuzzy matching because the wrong PruAccess
    fund would result in incorrect historical prices.
    """

    target = normalize_text(target_name)

    if not target:
        return None

    for option in options:

        if normalize_text(option["name"]) == target:
            return option

    return None


# ---------------------------------------------------------------------------
# Read-only PruAccess fields
# ---------------------------------------------------------------------------

def set_readonly_input_value(
    page: Page,
    selector: str,
    value: str,
) -> None:
    """
    Set a readonly field using the DOM.

    This is used only for the date fields.

    Fund Price Type is deliberately NOT touched.
    """

    locator = page.locator(selector)

    if locator.count() == 0:
        raise RuntimeError(
            f"Required field not found: {selector}"
        )

    locator.evaluate(
        """
        (element, value) => {
            element.removeAttribute("readonly");
            element.value = value;

            element.dispatchEvent(
                new Event("input", { bubbles: true })
            );

            element.dispatchEvent(
                new Event("change", { bubbles: true })
            );
        }
        """,
        value,
    )


# ---------------------------------------------------------------------------
# Date handling
# ---------------------------------------------------------------------------

def convert_prudential_date_to_pruaccess(
    value: str,
) -> str:
    """
    Convert common Prudential date formats into:

        DD-MMM-YYYY

    Example:
        03/11/2021 -> 03-Nov-2021
    """

    value = clean_text(value)

    if not value:
        raise ValueError("Empty date.")

    formats = [
        "%d/%m/%Y",
        "%d-%m-%Y",
        "%Y-%m-%d",
        "%d/%m/%y",
        "%d-%b-%Y",
        "%d-%b-%y",
    ]

    for fmt in formats:

        try:
            parsed = datetime.strptime(value, fmt)

            return parsed.strftime("%d-%b-%Y")

        except ValueError:
            continue

    raise ValueError(
        f"Unable to parse Prudential date: {value}"
    )


def get_inception_date(
    prudential_fund: dict,
) -> str:
    data = prudential_fund["data"]

    candidates = [
        data.get("inceptionDate"),
        data.get("fundInceptionDate"),
        data.get("launchDate"),
    ]

    for candidate in candidates:

        if candidate:
            return convert_prudential_date_to_pruaccess(
                clean_text(candidate)
            )

    raise RuntimeError(
        "Prudential fund data does not contain an inception date."
    )


# ---------------------------------------------------------------------------
# PruAccess historical table parsing
# ---------------------------------------------------------------------------

def parse_decimal_price(value: str) -> str:
    """
    Preserve the actual displayed price precision.

    Example:
        1.11916 -> 1.11916
        $1.11916 -> 1.11916
    """

    value = clean_text(value)

    value = value.replace(",", "")
    value = value.replace("$", "")
    value = value.replace("S$", "")

    match = re.search(
        r"-?\d+(?:\.\d+)?",
        value,
    )

    if not match:
        raise ValueError(
            f"Unable to parse price: {value}"
        )

    return match.group(0)


def extract_price_rows(page: Page) -> list[dict]:
    """
    Extract rows from the PruAccess historical price table.

    Only BID prices are stored.

    Offer prices may appear in the HTML table but are intentionally ignored.
    """

    rows = page.locator("table tbody tr")

    row_count = rows.count()

    observations: list[dict] = []

    for index in range(row_count):

        row = rows.nth(index)

        cells = row.locator("td")

        cell_count = cells.count()

        if cell_count < 2:
            continue

        values = []

        for cell_index in range(cell_count):

            values.append(
                clean_text(
                    cells.nth(cell_index).inner_text()
                )
            )

        date_value = values[0]

        if not re.search(
            r"\d{1,2}[-/][A-Za-z0-9]{2,3}[-/]\d{2,4}",
            date_value,
        ):
            continue

        # PruAccess table format:
        #
        # Date
        # Bid Price
        # Offer Price
        #
        # Therefore the second cell is BID.
        bid_value = values[1]

        try:
            bid_price = parse_decimal_price(
                bid_value
            )
        except ValueError:
            continue

        observations.append(
            {
                "date": date_value,
                "bidPrice": bid_price,
            }
        )

    return observations


# ---------------------------------------------------------------------------
# Pagination URL
# ---------------------------------------------------------------------------

def build_page_url(
    fund_id: str,
    start_date: str,
    end_date: str,
    page_number: int,
) -> str:

    query = {
        "fundId": fund_id,
        "viewType": "TBL",
        "startDate": start_date,
        "endDate": end_date,
        "page.page": str(page_number),
        "page.size": str(PAGE_SIZE),
    }

    return (
        PRUACCESS_URL
        + "?"
        + urlencode(query)
    )


# ---------------------------------------------------------------------------
# Individual pagination page
# ---------------------------------------------------------------------------

def extract_single_pagination_page(
    page: Page,
    url: str,
    page_number: int,
) -> list[dict]:
    """
    Extract one pagination page.

    This function is intentionally isolated so it can be retried without
    losing any previously extracted pages.
    """

    last_error: Exception | None = None

    for attempt in range(
        1,
        PAGE_RETRY_COUNT + 1,
    ):

        try:

            print(
                f"      Page {page_number}: "
                f"attempt {attempt}/{PAGE_RETRY_COUNT}"
            )

            page.goto(
                url,
                wait_until="domcontentloaded",
                timeout=PAGE_TIMEOUT_MS,
            )

            # Wait for either table rows or page body.
            try:
                page.wait_for_selector(
                    "table tbody tr",
                    timeout=30000,
                )
            except PlaywrightTimeoutError:
                pass

            observations = extract_price_rows(page)

            if not observations:
                # Check whether this is genuinely an empty page.
                body_text = clean_text(
                    page.locator("body").inner_text()
                )

                if (
                    "No records" in body_text
                    or "No Record" in body_text
                    or "no records" in body_text.casefold()
                ):
                    return []

                raise RuntimeError(
                    f"Page {page_number} loaded but "
                    "no historical price rows were found."
                )

            return observations

        except (
            PlaywrightTimeoutError,
            RuntimeError,
            Exception,
        ) as error:

            last_error = error

            print(
                f"      Page {page_number} failed on "
                f"attempt {attempt}: {error}"
            )

            if attempt < PAGE_RETRY_COUNT:

                print(
                    f"      Waiting "
                    f"{PAGE_RETRY_DELAY_SECONDS}s before retry..."
                )

                time.sleep(
                    PAGE_RETRY_DELAY_SECONDS
                )

    raise RuntimeError(
        f"Page {page_number} failed after "
        f"{PAGE_RETRY_COUNT} attempts: {last_error}"
    )


# ---------------------------------------------------------------------------
# Complete pagination extraction
# ---------------------------------------------------------------------------

def extract_all_pages(
    page: Page,
    fund_id: str,
    start_date: str,
    end_date: str,
) -> dict:

    all_observations: list[dict] = []
    page_summaries: list[dict] = []

    seen_dates: set[str] = set()

    for page_number in range(
        1,
        MAX_PAGES + 1,
    ):

        url = build_page_url(
            fund_id=fund_id,
            start_date=start_date,
            end_date=end_date,
            page_number=page_number,
        )

        print(
            f"    Extracting page "
            f"{page_number}..."
        )

        try:

            observations = extract_single_pagination_page(
                page=page,
                url=url,
                page_number=page_number,
            )

        except Exception as error:

            raise RuntimeError(
                f"Pagination stopped at page "
                f"{page_number}: {error}"
            ) from error

        row_count = len(observations)

        page_summary = {
            "page": page_number,
            "url": url,
            "rowCount": row_count,
            "status": "success",
        }

        if observations:

            page_summary["firstDate"] = (
                observations[0]["date"]
            )

            page_summary["lastDate"] = (
                observations[-1]["date"]
            )

        page_summaries.append(
            page_summary
        )

        for observation in observations:

            date_value = observation["date"]

            if date_value in seen_dates:
                raise RuntimeError(
                    "Duplicate historical date detected: "
                    f"{date_value}"
                )

            seen_dates.add(date_value)

            all_observations.append(
                observation
            )

        print(
            f"      Page {page_number}: "
            f"{row_count} observations"
        )

        # A page smaller than PAGE_SIZE is the final page.
        if row_count < PAGE_SIZE:

            break

        time.sleep(
            PAGE_DELAY_SECONDS
        )

    else:

        raise RuntimeError(
            f"Reached MAX_PAGES={MAX_PAGES}."
        )

    return {
        "pageSize": PAGE_SIZE,
        "pagesExtracted": len(page_summaries),
        "observations": all_observations,
        "pages": page_summaries,
    }


# ---------------------------------------------------------------------------
# Single fund extraction
# ---------------------------------------------------------------------------

def extract_single_fund(
    page: Page,
    excel_fund: dict,
    pruaccess_options: list[dict],
) -> dict:

    excel_row = excel_fund["excelRow"]

    prudential_url = excel_fund[
        "prudentialUrl"
    ]

    excel_pruaccess_name = excel_fund[
        "pruAccessName"
    ]

    print()
    print("=" * 80)
    print(
        f"FUND EXCEL ROW {excel_row}"
    )
    print(
        f"Excel PruAccess Name: "
        f"{excel_pruaccess_name}"
    )
    print(
        f"Prudential URL: "
        f"{prudential_url}"
    )
    print("=" * 80)

    # -----------------------------------------------------------------------
    # Prudential
    # -----------------------------------------------------------------------

    prudential_fund = get_prudential_fund_data(
        page,
        prudential_url,
    )

    prudential_data = prudential_fund[
        "data"
    ]

    fund_name = clean_text(
        prudential_data.get("fundName")
    )

    fund_identifier = clean_text(
        prudential_data.get("fundIdentifier")
    )

    fund_code = clean_text(
        prudential_data.get("fundCode")
    )

    current_bid_price = clean_text(
        prudential_data.get("bidPrice")
    )

    current_offer_price = clean_text(
        prudential_data.get("offerPrice")
    )

    valuation_date = clean_text(
        prudential_data.get("valuationDate")
    )

    if not fund_name:
        raise RuntimeError(
            "Prudential fund name is missing."
        )

    if not fund_identifier:
        raise RuntimeError(
            "Prudential fund identifier is missing."
        )

    inception_date = get_inception_date(
        prudential_fund
    )

    # -----------------------------------------------------------------------
    # PruAccess exact fund matching
    # -----------------------------------------------------------------------

    pruaccess_match = find_exact_pruaccess_match(
        pruaccess_options,
        excel_pruaccess_name,
    )

    if pruaccess_match is None:
        raise RuntimeError(
            "No exact PruAccess fund-name match "
            f"for '{excel_pruaccess_name}'."
        )

    pruaccess_name = pruaccess_match[
        "name"
    ]

    pruaccess_fund_id = pruaccess_match[
        "value"
    ]

    if not pruaccess_fund_id:
        raise RuntimeError(
            "Matched PruAccess fund has no fund ID."
        )

    print(
        f"PruAccess Match: "
        f"{pruaccess_name}"
    )

    print(
        f"PruAccess Fund ID: "
        f"{pruaccess_fund_id}"
    )

    print(
        f"Prudential Fund: "
        f"{fund_name}"
    )

    print(
        f"Fund Identifier: "
        f"{fund_identifier}"
    )

    print(
        f"Fund Code: "
        f"{fund_code}"
    )

    print(
        f"Inception Date: "
        f"{inception_date}"
    )

    # -----------------------------------------------------------------------
    # PruAccess dates
    # -----------------------------------------------------------------------

    end_date = datetime.now().strftime(
        "%d-%b-%Y"
    )

    start_date = inception_date

    print(
        f"PruAccess Start Date: "
        f"{start_date}"
    )

    print(
        f"PruAccess End Date: "
        f"{end_date}"
    )

    # -----------------------------------------------------------------------
    # IMPORTANT:
    #
    # Fund Price Type is deliberately NOT manipulated.
    #
    # PruAccess default fund price type is BID.
    # -----------------------------------------------------------------------

    print(
        "Fund Price Type: "
        "NOT MODIFIED"
    )

    # -----------------------------------------------------------------------
    # Pagination
    # -----------------------------------------------------------------------

    pagination = extract_all_pages(
        page=page,
        fund_id=pruaccess_fund_id,
        start_date=start_date,
        end_date=end_date,
    )

    observations = pagination[
        "observations"
    ]

    pages = pagination[
        "pages"
    ]

    if not observations:
        raise RuntimeError(
            "PruAccess returned no historical BID observations."
        )

    # -----------------------------------------------------------------------
    # Validate observations
    # -----------------------------------------------------------------------

    newest_observation = observations[0]
    oldest_observation = observations[-1]

    # -----------------------------------------------------------------------
    # Build result
    # -----------------------------------------------------------------------

    result = {
        "status": "success",

        "excelRow": excel_row,

        "prudentialUrl": prudential_url,

        "excelPruAccessName": excel_pruaccess_name,

        "prudentialFundName": fund_name,

        "matchedPruAccessName": pruaccess_name,

        "pruAccessFundId": pruaccess_fund_id,

        "fundIdentifier": fund_identifier,

        "fundCode": fund_code,

        "inceptionDate": clean_text(
            prudential_data.get(
                "inceptionDate"
            )
        ),

        "startDate": start_date,

        "endDate": end_date,

        "currentBidPrice": current_bid_price,

        "currentOfferPrice": current_offer_price,

        "valuationDate": valuation_date,

        "fundPriceType": "BID",

        "historicalPriceType": "BID",

        "source": "PruAccess",

        "pageSize": PAGE_SIZE,

        "pagesExtracted": len(pages),

        "historicalObservationCount": len(
            observations
        ),

        "newestObservation": newest_observation,

        "oldestObservation": oldest_observation,

        "observations": observations,

        "pagination": pages,

        "prudentialFundData": prudential_data,

        "retrievedAtUtc": utc_now_iso(),
    }

    return result


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main() -> None:

    started_at = utc_now_iso()

    OUTPUT_DIR.mkdir(
        parents=True,
        exist_ok=True,
    )

    FUNDS_OUTPUT_DIR.mkdir(
        parents=True,
        exist_ok=True,
    )

    print(
        "Reading Excel fund universe..."
    )

    funds = read_excel_funds()

    print(
        f"Excel fund universe: "
        f"{len(funds)} funds"
    )

    if not funds:
        raise RuntimeError(
            "No funds found in Funds Links.xlsm."
        )

    successful_funds: list[dict] = []
    failed_funds: list[dict] = []

    with sync_playwright() as playwright:

        browser: Browser = playwright.chromium.launch(
            headless=True,
        )

        context: BrowserContext = browser.new_context(
            viewport={
                "width": 1440,
                "height": 1000,
            }
        )

        page = context.new_page()

        page.set_default_timeout(
            PAGE_TIMEOUT_MS
        )

        # -------------------------------------------------------------------
        # Load PruAccess options once.
        # -------------------------------------------------------------------

        print()
        print(
            "Loading PruAccess fund options..."
        )

        pruaccess_options = get_fund_options(
            page
        )

        print(
            f"PruAccess fund options found: "
            f"{len(pruaccess_options)}"
        )

        save_json(
            OUTPUT_DIR / "fund_options.json",
            {
                "retrievedAtUtc": utc_now_iso(),
                "count": len(
                    pruaccess_options
                ),
                "options": pruaccess_options,
            },
        )

        # -------------------------------------------------------------------
        # Process every Excel fund.
        # -------------------------------------------------------------------

        for fund_index, excel_fund in enumerate(
            funds,
            start=1,
        ):

            print()
            print(
                f"PROCESSING FUND "
                f"{fund_index}/{len(funds)}"
            )

            excel_row = excel_fund[
                "excelRow"
            ]

            try:

                result = extract_single_fund(
                    page=page,
                    excel_fund=excel_fund,
                    pruaccess_options=pruaccess_options,
                )

                successful_funds.append(
                    result
                )

                identifier = safe_filename(
                    result["fundIdentifier"]
                )

                fund_name_safe = safe_filename(
                    result["prudentialFundName"]
                )

                fund_dir = (
                    FUNDS_OUTPUT_DIR
                    / f"{excel_row}_{identifier}"
                )

                fund_dir.mkdir(
                    parents=True,
                    exist_ok=True,
                )

                # Complete fund result.
                save_json(
                    fund_dir / "summary.json",
                    {
                        key: value
                        for key, value in result.items()
                        if key not in (
                            "observations",
                            "pagination",
                        )
                    },
                )

                # BID historical observations.
                save_json(
                    fund_dir / "bid_history.json",
                    {
                        "fundName": result[
                            "prudentialFundName"
                        ],
                        "fundIdentifier": result[
                            "fundIdentifier"
                        ],
                        "fundCode": result[
                            "fundCode"
                        ],
                        "source": "PruAccess",
                        "priceType": "BID",
                        "observationCount": len(
                            result["observations"]
                        ),
                        "observations": result[
                            "observations"
                        ],
                    },
                )

                # Pagination details.
                save_json(
                    fund_dir / "pagination.json",
                    {
                        "fundName": result[
                            "prudentialFundName"
                        ],
                        "fundIdentifier": result[
                            "fundIdentifier"
                        ],
                        "fundCode": result[
                            "fundCode"
                        ],
                        "fundId": result[
                            "pruAccessFundId"
                        ],
                        "pageSize": PAGE_SIZE,
                        "pagesExtracted": result[
                            "pagesExtracted"
                        ],
                        "pages": result[
                            "pagination"
                        ],
                    },
                )

                # Prudential current fund information.
                save_json(
                    fund_dir / "prudential_fund.json",
                    result[
                        "prudentialFundData"
                    ],
                )

                print()
                print(
                    f"SUCCESS: "
                    f"{result['prudentialFundName']}"
                )

                print(
                    f"Historical observations: "
                    f"{len(result['observations'])}"
                )

                print(
                    f"Pages: "
                    f"{result['pagesExtracted']}"
                )

            except Exception as error:

                failure = {
                    "status": "failed",
                    "excelRow": excel_row,
                    "prudentialUrl": excel_fund[
                        "prudentialUrl"
                    ],
                    "excelPruAccessName": excel_fund[
                        "pruAccessName"
                    ],
                    "error": str(error),
                    "failedAtUtc": utc_now_iso(),
                }

                failed_funds.append(
                    failure
                )

                failure_dir = (
                    FUNDS_OUTPUT_DIR
                    / f"{excel_row}_FAILED"
                )

                failure_dir.mkdir(
                    parents=True,
                    exist_ok=True,
                )

                save_json(
                    failure_dir / "failure.json",
                    failure,
                )

                print()
                print(
                    f"FAILED FUND ROW {excel_row}: "
                    f"{error}"
                )

                # Continue with the next fund.
                continue

        browser.close()

    # -----------------------------------------------------------------------
    # Consolidated successful fund data.
    # -----------------------------------------------------------------------

    successful_summary = []

    for fund in successful_funds:

        successful_summary.append(
            {
                "excelRow": fund[
                    "excelRow"
                ],

                "fundName": fund[
                    "prudentialFundName"
                ],

                "fundIdentifier": fund[
                    "fundIdentifier"
                ],

                "fundCode": fund[
                    "fundCode"
                ],

                "pruAccessFundName": fund[
                    "matchedPruAccessName"
                ],

                "pruAccessFundId": fund[
                    "pruAccessFundId"
                ],

                "observationCount": len(
                    fund["observations"]
                ),

                "newestObservation": (
                    fund["observations"][0]
                    if fund["observations"]
                    else None
                ),

                "oldestObservation": (
                    fund["observations"][-1]
                    if fund["observations"]
                    else None
                ),
            }
        )

    failed_summary = []

    for failure in failed_funds:

        failed_summary.append(
            {
                "status": "failed",
                "excelRow": failure[
                    "excelRow"
                ],
                "prudentialUrl": failure[
                    "prudentialUrl"
                ],
                "excelPruAccessName": failure[
                    "excelPruAccessName"
                ],
                "error": failure[
                    "error"
                ],
            }
        )

    # -----------------------------------------------------------------------
    # All funds JSON.
    # -----------------------------------------------------------------------

    all_funds_payload = {
        "status": (
            "success"
            if not failed_funds
            else "partial"
        ),

        "retrievedAtUtc": utc_now_iso(),

        "excelFile": str(
            EXCEL_FILE
        ),

        "fundUniverseCount": len(
            funds
        ),

        "successfulFundCount": len(
            successful_funds
        ),

        "failedFundCount": len(
            failed_funds
        ),

        "funds": successful_summary,

        "failedFunds": failed_summary,
    }

    save_json(
        OUTPUT_DIR / "all_funds.json",
        all_funds_payload,
    )

    # -----------------------------------------------------------------------
    # Consolidated BID history.
    # -----------------------------------------------------------------------

    consolidated_bid_history = {
        "status": (
            "success"
            if not failed_funds
            else "partial"
        ),

        "source": "PruAccess",

        "priceType": "BID",

        "retrievedAtUtc": utc_now_iso(),

        "fundCount": len(
            successful_funds
        ),

        "funds": {},
    }

    total_observations = 0

    for fund in successful_funds:

        identifier = fund[
            "fundIdentifier"
        ]

        observations = fund[
            "observations"
        ]

        total_observations += len(
            observations
        )

        consolidated_bid_history[
            "funds"
        ][identifier] = {
            "fundName": fund[
                "prudentialFundName"
            ],

            "fundCode": fund[
                "fundCode"
            ],

            "pruAccessFundId": fund[
                "pruAccessFundId"
            ],

            "observationCount": len(
                observations
            ),

            "observations": observations,
        }

    consolidated_bid_history[
        "totalHistoricalBidObservations"
    ] = total_observations

    save_json(
        OUTPUT_DIR / "all_bid_history.json",
        consolidated_bid_history,
    )

    # -----------------------------------------------------------------------
    # Run summary.
    # -----------------------------------------------------------------------

    completed_at = utc_now_iso()

    run_summary = {
        "status": (
            "success"
            if not failed_funds
            else "partial"
        ),

        "startedAtUtc": started_at,

        "completedAtUtc": completed_at,

        "excelFile": str(
            EXCEL_FILE
        ),

        "fundUniverseCount": len(
            funds
        ),

        "successfulFundCount": len(
            successful_funds
        ),

        "failedFundCount": len(
            failed_funds
        ),

        "totalHistoricalBidObservations": total_observations,

        "historicalPriceType": "BID",

        "pageSize": PAGE_SIZE,

        "successfulFunds": successful_summary,

        "failedFunds": failed_summary,
    }

    save_json(
        OUTPUT_DIR / "run_summary.json",
        run_summary,
    )

    # -----------------------------------------------------------------------
    # Final console summary.
    # -----------------------------------------------------------------------

    print()
    print("=" * 80)
    print("EXTRACTION COMPLETE")
    print("=" * 80)

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
        f"Historical BID observations: "
        f"{total_observations}"
    )

    if failed_funds:

        print()
        print(
            "FAILED FUNDS:"
        )

        for failure in failed_funds:

            print(
                f"  Row {failure['excelRow']}: "
                f"{failure['excelPruAccessName']}"
            )

            print(
                f"    {failure['error']}"
            )

    else:

        print()
        print(
            "ALL 67 FUNDS EXTRACTED SUCCESSFULLY."
        )

    print()
    print(
        f"Output directory: "
        f"{OUTPUT_DIR}"
    )


if __name__ == "__main__":
    main()
