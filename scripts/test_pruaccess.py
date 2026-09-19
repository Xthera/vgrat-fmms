#!/usr/bin/env python3
"""
VGrat FMS - Prudential PruAccess Historical BID Price Extractor

SOURCE HIERARCHY
----------------
1. Funds Links.xlsm
   - Column A = Prudential fund URL
   - Column B = exact PruAccess fund name
   - Every populated row is processed.
   - No hardcoded fund count.

2. Prudential official fund page / ilpfunds.json
   - Authoritative fund identity and current fund information.
   - Used to obtain:
       fundIdentifier
       fundName
       fundCode
       currency
       asset class
       risk classification
       bid price
       offer price
       valuation date
       inception date
       etc.

3. PruAccess
   - Authoritative source for historical BID price history.
   - Historical OFFER prices are never used.
   - No synthetic / estimated / interpolated / fabricated values.

HARD RULES
----------
- Excel Column A is the source of truth for the fund universe.
- Process every populated Column A row.
- Never hardcode 67 funds.
- Excel Column B is the exact PruAccess fund name.
- No fuzzy fund matching.
- Only harmless normalization is allowed for matching.
- Prudential is authoritative for fund identity/current fund information.
- PruAccess is authoritative for historical BID prices.
- Historical price type must remain BID.
- Never substitute OFFER for BID.
- Never modify the PruAccess price-type selector.
- No synthetic data.
- No estimated data.
- No interpolation.
- No fabricated data.
- No carry-forward values.
- Historical start date comes from official Prudential inception date.
- Historical end date is the current Singapore calendar date.
- Every pagination page must be successfully extracted.
- Temporary page failures are retried.
- Permanent page failures fail the entire fund.
- Partial historical data is never marked successful.
- Returned rows are preserved exactly.
- No silent deduplication.
- Duplicate validation must detect duplicates.
- Historical rows must contain valid dates and BID prices.
- Complete historical range must be validated before success.
- Whenever this script changes, the FULL script must be provided.

OUTPUT
------
output_pruaccess/
    funds/
        <excel_row>_<identifier>/
            prudential_fund.json
            pre_submit.json
            pagination.json
            bid_history.json
            summary.json

        <excel_row>_failed/
            failure.json

    all_funds.json
    all_bid_history.json
    run_summary.json
"""

from __future__ import annotations

import json
import os
import re
import sys
import time
import traceback
from datetime import date, datetime
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple
from urllib.parse import parse_qsl, urlencode, urlparse, urlunparse
from zoneinfo import ZoneInfo

from openpyxl import load_workbook
from playwright.sync_api import (
    Browser,
    BrowserContext,
    Page,
    TimeoutError as PlaywrightTimeoutError,
    sync_playwright,
)


# ============================================================================
# CONFIGURATION
# ============================================================================

BASE_DIR = Path(__file__).resolve().parent.parent

EXCEL_FILE = BASE_DIR / "Funds Links.xlsm"

OUTPUT_DIR = BASE_DIR / "output_pruaccess"
FUNDS_OUTPUT_DIR = OUTPUT_DIR / "funds"

PRUDENTIAL_WAIT_MS = 3000
PRUACCESS_INITIAL_WAIT_MS = 1500
PRUACCESS_RESULT_WAIT_MS = 2500

# IMPORTANT:
# fetch_page_with_retries() uses this value.
PAGE_WAIT_MS = 2500

REQUEST_TIMEOUT_MS = 60000

MAX_RETRIES = 4
RETRY_DELAY_SECONDS = 2

PAGE_SIZE = 20

PRUACCESS_PER_PAGE_EXPECTED = PAGE_SIZE

SINGAPORE_TZ = ZoneInfo("Asia/Singapore")

PRUACCESS_URL = (
    "https://pruaccess.prudential.com.sg/prulinkfund/"
    "viewFundPerformance.do"
)

PRUDENTIAL_API_MARKER = "ilpfunds.json"

USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) "
    "Chrome/131.0.0.0 Safari/537.36"
)


# ============================================================================
# LOGGING
# ============================================================================

def log(message: str) -> None:
    timestamp = datetime.now(SINGAPORE_TZ).strftime("%Y-%m-%d %H:%M:%S")
    print(f"[{timestamp}] {message}", flush=True)


def log_section(title: str) -> None:
    print()
    print("=" * 80)
    print(title)
    print("=" * 80)
    print()


# ============================================================================
# JSON HELPERS
# ============================================================================

def json_default(value: Any) -> Any:
    if isinstance(value, (datetime, date)):
        return value.isoformat()

    if isinstance(value, Path):
        return str(value)

    return str(value)


def save_json(path: Path, data: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)

    with path.open("w", encoding="utf-8") as handle:
        json.dump(
            data,
            handle,
            indent=2,
            ensure_ascii=False,
            default=json_default,
        )


# ============================================================================
# GENERAL HELPERS
# ============================================================================

def safe_filename(value: str) -> str:
    value = str(value or "").strip()

    value = re.sub(r"[<>:\"/\\|?*\x00-\x1F]", "_", value)
    value = re.sub(r"\s+", "_", value)

    value = value.strip(" ._")

    if not value:
        value = "unknown"

    return value[:180]


def normalize_text(value: Any) -> str:
    """
    Harmless normalization for exact matching only.

    This is NOT fuzzy matching.
    """
    if value is None:
        return ""

    value = str(value)

    value = value.replace("\u00a0", " ")
    value = value.replace("\u2013", "-")
    value = value.replace("\u2014", "-")

    value = re.sub(r"\s+", " ", value)

    return value.strip().casefold()


def singapore_today() -> date:
    return datetime.now(SINGAPORE_TZ).date()


def format_pruaccess_date(value: date) -> str:
    """
    PruAccess date format expected by the form.
    Example:
        01-Jan-2020
    """
    return value.strftime("%d-%b-%Y")


def parse_prudential_date(value: Any) -> Optional[date]:
    if value is None:
        return None

    if isinstance(value, datetime):
        return value.date()

    if isinstance(value, date):
        return value

    text = str(value).strip()

    if not text:
        return None

    formats = [
        "%d/%m/%Y",
        "%Y-%m-%d",
        "%d-%b-%Y",
        "%d-%b-%y",
        "%d/%m/%y",
        "%Y/%m/%d",
    ]

    for fmt in formats:
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            continue

    return None


def parse_history_date(value: str) -> Optional[date]:
    value = str(value or "").strip()

    formats = [
        "%d-%b-%Y",
        "%d-%B-%Y",
        "%d/%m/%Y",
        "%d-%m-%Y",
        "%Y-%m-%d",
    ]

    for fmt in formats:
        try:
            return datetime.strptime(value, fmt).date()
        except ValueError:
            continue

    return None


def clean_price(value: str) -> Optional[str]:
    """
    Preserve the returned price text as much as possible while validating
    that it represents a numeric price.

    No rounding or transformation is performed.
    """
    if value is None:
        return None

    text = str(value).strip()

    if not text:
        return None

    text = text.replace(",", "")

    if not re.fullmatch(r"\d+(?:\.\d+)?", text):
        return None

    return text


# ============================================================================
# EXCEL
# ============================================================================

def load_excel_funds() -> List[Dict[str, Any]]:
    """
    Reads every populated Column A starting from row 2.

    Column A:
        Prudential fund URL

    Column B:
        Exact PruAccess fund name
    """

    if not EXCEL_FILE.exists():
        raise FileNotFoundError(
            f"Excel source file not found: {EXCEL_FILE}"
        )

    log(f"Loading Excel source: {EXCEL_FILE}")

    workbook = load_workbook(
        EXCEL_FILE,
        read_only=True,
        data_only=True,
        keep_vba=True,
    )

    try:
        worksheet = workbook.active

        funds: List[Dict[str, Any]] = []

        for row_number in range(2, worksheet.max_row + 1):
            url_value = worksheet.cell(row=row_number, column=1).value
            name_value = worksheet.cell(row=row_number, column=2).value

            if url_value is None:
                continue

            url = str(url_value).strip()

            if not url:
                continue

            pruaccess_name = (
                str(name_value).strip()
                if name_value is not None
                else ""
            )

            funds.append(
                {
                    "excelRow": row_number,
                    "prudentialUrl": url,
                    "pruAccessName": pruaccess_name,
                }
            )

        log(f"Excel fund universe: {len(funds)} populated rows")

        return funds

    finally:
        workbook.close()


# ============================================================================
# PRUDENTIAL OFFICIAL DATA
# ============================================================================

def extract_prudential_api_url(page: Page) -> Optional[str]:
    """
    Searches loaded page/network information for the official ilpfunds.json
    endpoint.

    The actual request URL is captured from network responses.
    """

    captured: List[str] = []

    def handle_response(response: Any) -> None:
        try:
            url = response.url

            if PRUDENTIAL_API_MARKER.lower() in url.lower():
                captured.append(url)
        except Exception:
            pass

    page.on("response", handle_response)

    try:
        page.wait_for_timeout(PRUDENTIAL_WAIT_MS)
    except Exception:
        pass

    if captured:
        return captured[-1]

    return None


def get_prudential_fund_data(
    context: BrowserContext,
    prudential_url: str,
) -> Dict[str, Any]:

    page = context.new_page()

    try:
        page.set_default_timeout(REQUEST_TIMEOUT_MS)

        captured_urls: List[str] = []

        def handle_response(response: Any) -> None:
            try:
                response_url = response.url

                if PRUDENTIAL_API_MARKER.lower() in response_url.lower():
                    captured_urls.append(response_url)

            except Exception:
                pass

        page.on("response", handle_response)

        log(f"Opening Prudential page: {prudential_url}")

        page.goto(
            prudential_url,
            wait_until="domcontentloaded",
            timeout=REQUEST_TIMEOUT_MS,
        )

        try:
            page.wait_for_load_state(
                "networkidle",
                timeout=REQUEST_TIMEOUT_MS,
            )
        except Exception:
            pass

        page.wait_for_timeout(PRUDENTIAL_WAIT_MS)

        api_url = (
            captured_urls[-1]
            if captured_urls
            else extract_prudential_api_url(page)
        )

        if not api_url:
            raise RuntimeError(
                "Could not capture Prudential ilpfunds.json API URL"
            )

        log(f"Prudential API: {api_url}")

        response = context.request.get(
            api_url,
            timeout=REQUEST_TIMEOUT_MS,
        )

        if not response.ok:
            raise RuntimeError(
                f"Prudential API returned HTTP {response.status}"
            )

        try:
            payload = response.json()
        except Exception as exc:
            raise RuntimeError(
                f"Could not decode Prudential API JSON: {exc}"
            ) from exc

        return parse_prudential_fund_payload(
            payload=payload,
            source_url=prudential_url,
            api_url=api_url,
        )

    finally:
        page.close()


def parse_prudential_fund_payload(
    payload: Any,
    source_url: str,
    api_url: str,
) -> Dict[str, Any]:

    if isinstance(payload, dict):
        if isinstance(payload.get("data"), list):
            data = payload["data"]
        elif isinstance(payload.get("funds"), list):
            data = payload["funds"]
        else:
            data = [payload]

    elif isinstance(payload, list):
        data = payload

    else:
        raise RuntimeError(
            "Unexpected Prudential API response structure"
        )

    if not data:
        raise RuntimeError(
            "Prudential API returned no fund records"
        )

    fund = data[0]

    if not isinstance(fund, dict):
        raise RuntimeError(
            "Prudential API first fund record is not an object"
        )

    def first_value(*keys: str) -> Any:
        for key in keys:
            if key in fund and fund[key] not in (None, ""):
                return fund[key]
        return None

    parsed = {
        "source": "Prudential",
        "sourceUrl": source_url,
        "apiUrl": api_url,

        "fundIdentifier": first_value(
            "fundIdentifier",
            "citicode",
            "citiCode",
            "fundId",
        ),

        "fundName": first_value(
            "fundName",
            "name",
            "fund_name",
        ),

        "fundCode": first_value(
            "fundCode",
            "code",
        ),

        "fundCurrency": first_value(
            "fundCurrency",
            "currency",
        ),

        "unitCurrency": first_value(
            "unitCurrency",
            "unit_currency",
        ),

        "assetClass": first_value(
            "assetClass",
            "asset_class",
        ),

        "assetSubClass": first_value(
            "assetSubClass",
            "asset_sub_class",
        ),

        "riskClassification": first_value(
            "riskClassification",
            "risk_classification",
            "risk",
        ),

        "bidPrice": first_value(
            "bidPrice",
            "bid",
        ),

        "offerPrice": first_value(
            "offerPrice",
            "offer",
        ),

        "valuationDate": first_value(
            "valuationDate",
            "valuation_date",
        ),

        "inceptionDate": first_value(
            "inceptionDate",
            "inception_date",
        ),

        "cumulativeReturn1Y": first_value(
            "cumulativeReturn1Y",
            "return1Y",
            "oneYearReturn",
        ),

        "annualisedReturn3Y": first_value(
            "annualisedReturn3Y",
            "return3Y",
            "threeYearReturn",
        ),

        "annualisedReturn5Y": first_value(
            "annualisedReturn5Y",
            "return5Y",
            "fiveYearReturn",
        ),

        "objective": first_value(
            "objective",
            "fundObjective",
        ),

        "investmentManager": first_value(
            "investmentManager",
            "manager",
        ),

        "dividendInfo": first_value(
            "dividendInfo",
            "dividend",
        ),

        "raw": fund,
    }

    return parsed


# ============================================================================
# PRUACCESS FUND OPTIONS
# ============================================================================

def get_pruaccess_fund_options(
    page: Page,
) -> List[Dict[str, str]]:

    page.set_default_timeout(REQUEST_TIMEOUT_MS)

    log(f"Opening PruAccess: {PRUACCESS_URL}")

    page.goto(
        PRUACCESS_URL,
        wait_until="domcontentloaded",
        timeout=REQUEST_TIMEOUT_MS,
    )

    try:
        page.wait_for_load_state(
            "networkidle",
            timeout=REQUEST_TIMEOUT_MS,
        )
    except Exception:
        pass

    page.wait_for_timeout(PRUACCESS_INITIAL_WAIT_MS)

    options = page.locator("#fundName option")

    count = options.count()

    if count == 0:
        raise RuntimeError(
            "PruAccess #fundName contains no options"
        )

    results: List[Dict[str, str]] = []

    for index in range(count):
        option = options.nth(index)

        text = option.inner_text().strip()
        value = option.get_attribute("value")

        if not text:
            continue

        results.append(
            {
                "text": text,
                "value": value or "",
            }
        )

    log(f"PruAccess fund options found: {len(results)}")

    return results


def match_exact_pruaccess_fund(
    excel_name: str,
    options: List[Dict[str, str]],
) -> Dict[str, str]:

    normalized_target = normalize_text(excel_name)

    if not normalized_target:
        raise RuntimeError(
            "Excel Column B PruAccess fund name is empty"
        )

    matches = [
        option
        for option in options
        if normalize_text(option["text"]) == normalized_target
    ]

    if len(matches) == 0:
        raise RuntimeError(
            f"No exact PruAccess fund match for Excel Column B: "
            f"{excel_name!r}"
        )

    if len(matches) > 1:
        raise RuntimeError(
            f"Multiple exact PruAccess fund matches for: "
            f"{excel_name!r}"
        )

    return matches[0]


# ============================================================================
# PRUACCESS FORM HELPERS
# ============================================================================

def set_readonly_input_value(
    page: Page,
    selector: str,
    value: str,
) -> None:

    locator = page.locator(selector)

    if locator.count() == 0:
        raise RuntimeError(
            f"Required PruAccess input not found: {selector}"
        )

    locator.first.evaluate(
        """
        (element, value) => {
            element.removeAttribute('readonly');
            element.value = value;

            element.dispatchEvent(
                new Event('input', { bubbles: true })
            );

            element.dispatchEvent(
                new Event('change', { bubbles: true })
            );
        }
        """,
        value,
    )


def get_csrf_token(page: Page) -> Optional[str]:

    selectors = [
        'input[name="_csrf"]',
        'input[name="csrf"]',
        'input[name="csrfToken"]',
    ]

    for selector in selectors:
        locator = page.locator(selector)

        if locator.count() > 0:
            value = locator.first.get_attribute("value")

            if value:
                return value

    return None


def select_required_value(
    page: Page,
    selector: str,
    value: str,
) -> None:

    locator = page.locator(selector)

    if locator.count() == 0:
        raise RuntimeError(
            f"Required PruAccess selector not found: {selector}"
        )

    locator.first.select_option(value=value)


def get_selected_value(
    page: Page,
    selector: str,
) -> Optional[str]:

    locator = page.locator(selector)

    if locator.count() == 0:
        return None

    try:
        return locator.first.input_value()
    except Exception:
        return None


def submit_pruaccess_form(
    page: Page,
    fund_id: str,
    start_date: str,
    end_date: str,
) -> Dict[str, Any]:

    select_required_value(
        page,
        "#fundName",
        fund_id,
    )

    select_required_value(
        page,
        "#viewType",
        "TBL",
    )

    # HARD RULE:
    # Do not modify #fundPriceType.
    price_type_value = get_selected_value(
        page,
        "#fundPriceType",
    )

    price_type_text = None

    try:
        price_type_text = page.locator(
            "#fundPriceType option:checked"
        ).inner_text()
    except Exception:
        pass

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

    csrf = get_csrf_token(page)

    if not csrf:
        raise RuntimeError(
            "PruAccess CSRF token not found"
        )

    form = page.locator("#fundForm")

    if form.count() == 0:
        raise RuntimeError(
            "PruAccess #fundForm not found"
        )

    form_action = form.get_attribute("action")
    form_method = form.get_attribute("method")

    pre_submit = {
        "fundId": fund_id,
        "startDate": start_date,
        "endDate": end_date,
        "viewType": "TBL",

        # Recorded only.
        # NEVER changed by this script.
        "priceTypeValue": price_type_value,
        "priceTypeText": price_type_text,

        "csrfPresent": True,
        "formAction": form_action,
        "formMethod": form_method,
    }

    log(
        "Submitting PruAccess form "
        f"fundId={fund_id} "
        f"start={start_date} "
        f"end={end_date} "
        f"viewType=TBL "
        f"priceType={price_type_text or price_type_value or 'UNKNOWN'}"
    )

    try:
        with page.expect_navigation(
            wait_until="domcontentloaded",
            timeout=REQUEST_TIMEOUT_MS,
        ):
            form.evaluate(
                """
                form => form.submit()
                """
            )
    except PlaywrightTimeoutError:
        # Some versions/pages update without a normal navigation.
        log(
            "PruAccess form submission navigation timed out; "
            "checking page content"
        )

    page.wait_for_timeout(PRUACCESS_RESULT_WAIT_MS)

    return pre_submit


# ============================================================================
# PRICE TABLE EXTRACTION
# ============================================================================

DATE_PATTERN = re.compile(
    r"^\d{1,2}-[A-Za-z]{3}-\d{4}$"
)

PRICE_PATTERN = re.compile(
    r"^\d+(?:,\d{3})*(?:\.\d+)?$"
)


def extract_price_rows(page: Page) -> List[Dict[str, Any]]:
    """
    Extracts BID rows from the rendered result table.

    Current PruAccess table structure is expected to contain:

        Date | Bid Price | ...

    We intentionally preserve the first two cells exactly as returned.

    This function does NOT deduplicate rows.
    """

    tables = page.locator("table")

    table_count = tables.count()

    candidates: List[Dict[str, Any]] = []

    for table_index in range(table_count):

        table = tables.nth(table_index)

        rows = table.locator("tr")
        row_count = rows.count()

        if row_count == 0:
            continue

        for row_index in range(row_count):

            row = rows.nth(row_index)

            cells = row.locator("th, td")
            cell_count = cells.count()

            if cell_count < 2:
                continue

            values: List[str] = []

            for cell_index in range(cell_count):
                try:
                    values.append(
                        cells.nth(cell_index).inner_text().strip()
                    )
                except Exception:
                    values.append("")

            date_text = values[0].strip()
            price_text = values[1].strip()

            if not DATE_PATTERN.match(date_text):
                continue

            if not PRICE_PATTERN.match(price_text.replace(",", "")):
                continue

            parsed_date = parse_history_date(date_text)

            if parsed_date is None:
                continue

            cleaned_price = clean_price(price_text)

            if cleaned_price is None:
                continue

            candidates.append(
                {
                    "date": date_text,
                    "dateParsed": parsed_date.isoformat(),
                    "bidPrice": cleaned_price,
                    "rawCells": values,
                    "tableIndex": table_index,
                    "rowIndex": row_index,
                }
            )

    return candidates


def verify_bid_result(
    page: Page,
    rows: List[Dict[str, Any]],
) -> None:

    if not rows:
        raise RuntimeError(
            "PruAccess result contains zero BID history rows"
        )

    selected_price_type = get_selected_value(
        page,
        "#fundPriceType",
    )

    selected_price_text = None

    try:
        selected_price_text = page.locator(
            "#fundPriceType option:checked"
        ).inner_text()
    except Exception:
        pass

    log(
        "PruAccess result price type: "
        f"{selected_price_text or selected_price_type or 'UNKNOWN'}"
    )

    # We do not modify this selector.
    #
    # If the page explicitly reports OFFER rather than BID, reject it.
    combined = normalize_text(
        f"{selected_price_type or ''} {selected_price_text or ''}"
    )

    if "offer" in combined:
        raise RuntimeError(
            "PruAccess result appears to be OFFER price data; "
            "historical BID data is required"
        )

    log(
        f"PruAccess initial result rows detected: {len(rows)}"
    )


# ============================================================================
# PAGINATION URL
# ============================================================================

def build_page_url(
    base_url: str,
    fund_id: str,
    view_type: str,
    start_date: str,
    end_date: str,
    page_number: int,
    page_size: int,
) -> str:

    parsed = urlparse(base_url)

    existing = dict(parse_qsl(parsed.query))

    existing.update(
        {
            "fundId": fund_id,
            "viewType": view_type,
            "startDate": start_date,
            "endDate": end_date,
            "page.page": str(page_number),
            "page.size": str(page_size),
        }
    )

    query = urlencode(existing)

    return urlunparse(
        (
            parsed.scheme,
            parsed.netloc,
            parsed.path,
            parsed.params,
            query,
            parsed.fragment,
        )
    )


# ============================================================================
# PAGINATION DISCOVERY
# ============================================================================

def discover_max_page(page: Page) -> int:
    """
    Attempts to discover the maximum page number from visible pagination.

    If none can be found, defaults to 1.

    This function does not fabricate additional pages.
    """

    page_numbers: List[int] = []

    # Links
    links = page.locator("a")

    for index in range(links.count()):
        link = links.nth(index)

        try:
            text = link.inner_text().strip()
        except Exception:
            text = ""

        href = link.get_attribute("href") or ""

        for source in (text, href):
            matches = re.findall(
                r"(?:page\.page=|page=)(\d+)",
                source,
                flags=re.IGNORECASE,
            )

            for match in matches:
                try:
                    page_numbers.append(int(match))
                except ValueError:
                    pass

        if text.isdigit():
            try:
                number = int(text)

                if 1 <= number <= 100000:
                    page_numbers.append(number)
            except ValueError:
                pass

    # Buttons
    buttons = page.locator("button")

    for index in range(buttons.count()):
        button = buttons.nth(index)

        try:
            text = button.inner_text().strip()
        except Exception:
            text = ""

        if text.isdigit():
            try:
                number = int(text)

                if 1 <= number <= 100000:
                    page_numbers.append(number)
            except ValueError:
                pass

    # Body
    try:
        body_text = page.locator("body").inner_text()

        matches = re.findall(
            r"(?:page\.page=|page=)(\d+)",
            body_text,
            flags=re.IGNORECASE,
        )

        for match in matches:
            try:
                page_numbers.append(int(match))
            except ValueError:
                pass

    except Exception:
        pass

    if not page_numbers:
        return 1

    return max(page_numbers)


# ============================================================================
# FETCH INDIVIDUAL PAGE
# ============================================================================

def fetch_page_with_retries(
    page: Page,
    url: str,
    page_number: int,
) -> Tuple[List[Dict[str, Any]], Dict[str, Any]]:

    last_error: Optional[str] = None

    for attempt in range(1, MAX_RETRIES + 1):

        try:
            log(
                f"Fetching PruAccess page {page_number} "
                f"(attempt {attempt}/{MAX_RETRIES})"
            )

            response = page.goto(
                url,
                wait_until="domcontentloaded",
                timeout=REQUEST_TIMEOUT_MS,
            )

            if response is not None:
                if response.status >= 400:
                    raise RuntimeError(
                        f"HTTP {response.status}"
                    )

            try:
                page.wait_for_load_state(
                    "networkidle",
                    timeout=REQUEST_TIMEOUT_MS,
                )
            except Exception:
                pass

            page.wait_for_timeout(PAGE_WAIT_MS)

            rows = extract_price_rows(page)

            if not rows:
                raise RuntimeError(
                    "Page loaded but contains zero price rows"
                )

            first_date = rows[0]["date"]
            last_date = rows[-1]["date"]

            log(
                f"Page {page_number}: "
                f"{len(rows)} rows "
                f"({first_date} -> {last_date})"
            )

            return (
                rows,
                {
                    "page": page_number,
                    "url": url,
                    "attempt": attempt,
                    "rowCount": len(rows),
                    "firstDate": first_date,
                    "lastDate": last_date,
                    "status": (
                        response.status
                        if response is not None
                        else None
                    ),
                },
            )

        except Exception as exc:

            last_error = str(exc)

            log(
                f"Page {page_number} attempt {attempt} failed: "
                f"{last_error}"
            )

            if attempt < MAX_RETRIES:
                time.sleep(RETRY_DELAY_SECONDS * attempt)

    raise RuntimeError(
        f"PruAccess page {page_number} failed after "
        f"{MAX_RETRIES} attempts: {last_error}"
    )


# ============================================================================
# PAGINATION EXTRACTION
# ============================================================================

def validate_page_row_count(
    page_number: int,
    row_count: int,
    known_max_page: Optional[int],
) -> None:

    if known_max_page is not None:

        if page_number < known_max_page:
            if row_count != PAGE_SIZE:
                raise RuntimeError(
                    f"Page {page_number} has {row_count} rows, "
                    f"but expected exactly {PAGE_SIZE} because it is "
                    f"not the final page"
                )

        else:
            if not (1 <= row_count <= PAGE_SIZE):
                raise RuntimeError(
                    f"Final page {page_number} has invalid row count: "
                    f"{row_count}"
                )

    else:

        if not (1 <= row_count <= PAGE_SIZE):
            raise RuntimeError(
                f"Page {page_number} has invalid row count: "
                f"{row_count}"
            )


def validate_history_rows(
    rows: List[Dict[str, Any]],
    requested_start: date,
    requested_end: date,
) -> Dict[str, Any]:

    if not rows:
        raise RuntimeError(
            "Historical BID result contains zero rows"
        )

    parsed_dates: List[date] = []

    duplicate_dates: Dict[str, int] = {}

    invalid_rows: List[Dict[str, Any]] = []

    for index, row in enumerate(rows):

        date_text = row.get("date")
        price_text = row.get("bidPrice")

        parsed_date = parse_history_date(
            str(date_text or "")
        )

        if parsed_date is None:
            invalid_rows.append(
                {
                    "index": index,
                    "reason": "invalid date",
                    "row": row,
                }
            )
            continue

        if clean_price(str(price_text or "")) is None:
            invalid_rows.append(
                {
                    "index": index,
                    "reason": "invalid BID price",
                    "row": row,
                }
            )
            continue

        parsed_dates.append(parsed_date)

        key = parsed_date.isoformat()

        duplicate_dates[key] = (
            duplicate_dates.get(key, 0) + 1
        )

    if invalid_rows:
        raise RuntimeError(
            "Historical BID data contains invalid rows: "
            f"{len(invalid_rows)}"
        )

    duplicates = {
        key: count
        for key, count in duplicate_dates.items()
        if count > 1
    }

    if duplicates:
        raise RuntimeError(
            "Historical BID data contains duplicate dates: "
            f"{duplicates}"
        )

    if not parsed_dates:
        raise RuntimeError(
            "Historical BID data contains no valid dates"
        )

    newest = max(parsed_dates)
    oldest = min(parsed_dates)

    if newest < oldest:
        raise RuntimeError(
            "Historical date ordering is invalid"
        )

    if oldest > requested_start:
        raise RuntimeError(
            "Historical BID data does not reach the requested "
            f"start boundary. Requested={requested_start.isoformat()}, "
            f"oldestReturned={oldest.isoformat()}"
        )

    if newest > requested_end:
        raise RuntimeError(
            "Historical BID data contains a date later than "
            f"requested end date. Requested={requested_end.isoformat()}, "
            f"newestReturned={newest.isoformat()}"
        )

    return {
        "rowCount": len(rows),
        "oldestDate": oldest.isoformat(),
        "newestDate": newest.isoformat(),
        "requestedStartDate": requested_start.isoformat(),
        "requestedEndDate": requested_end.isoformat(),
        "duplicateDates": {},
        "invalidRows": 0,
    }


def extract_all_pages(
    page: Page,
    fund_id: str,
    start_date: str,
    end_date: str,
    initial_rows: List[Dict[str, Any]],
) -> Tuple[List[Dict[str, Any]], Dict[str, Any]]:

    requested_start = parse_history_date(start_date)
    requested_end = parse_history_date(end_date)

    if requested_start is None:
        raise RuntimeError(
            f"Invalid requested start date: {start_date}"
        )

    if requested_end is None:
        raise RuntimeError(
            f"Invalid requested end date: {end_date}"
        )

    if requested_start > requested_end:
        raise RuntimeError(
            "Requested start date is after requested end date"
        )

    # ------------------------------------------------------------------------
    # IMPORTANT:
    #
    # The submitted result is page 1.
    #
    # We preserve it and use it as the first page rather than discarding it.
    # ------------------------------------------------------------------------

    all_rows: List[Dict[str, Any]] = []

    page_metadata: List[Dict[str, Any]] = []

    first_page_rows = initial_rows

    if not first_page_rows:
        raise RuntimeError(
            "Initial submitted PruAccess result contains zero rows"
        )

    all_rows.extend(first_page_rows)

    page_metadata.append(
        {
            "page": 1,
            "url": page.url,
            "rowCount": len(first_page_rows),
            "firstDate": first_page_rows[0]["date"],
            "lastDate": first_page_rows[-1]["date"],
            "source": "submitted_form_result",
        }
    )

    # Discover maximum page count from the submitted result.
    discovered_max_page = discover_max_page(page)

    log(
        f"PruAccess pagination discovery: "
        f"maxPage={discovered_max_page}"
    )

    current_page = 1

    # ------------------------------------------------------------------------
    # If page 1 has fewer than PAGE_SIZE rows, it is normally the final page.
    #
    # If pagination UI says there are additional pages despite fewer rows,
    # that is treated as an integrity problem rather than silently guessing.
    # ------------------------------------------------------------------------

    if discovered_max_page > 1 and len(first_page_rows) != PAGE_SIZE:
        raise RuntimeError(
            "PruAccess pagination indicates additional pages, "
            f"but page 1 contains only {len(first_page_rows)} rows"
        )

    # ------------------------------------------------------------------------
    # Fetch pages 2 onward.
    # ------------------------------------------------------------------------

    while current_page < discovered_max_page:

        current_page += 1

        url = build_page_url(
            base_url=PRUACCESS_URL,
            fund_id=fund_id,
            view_type="TBL",
            start_date=start_date,
            end_date=end_date,
            page_number=current_page,
            page_size=PAGE_SIZE,
        )

        rows, metadata = fetch_page_with_retries(
            page=page,
            url=url,
            page_number=current_page,
        )

        validate_page_row_count(
            page_number=current_page,
            row_count=len(rows),
            known_max_page=discovered_max_page,
        )

        all_rows.extend(rows)
        page_metadata.append(metadata)

    # ------------------------------------------------------------------------
    # Fallback discovery:
    #
    # If no pagination UI was discovered and page 1 is full, sequentially
    # continue until a page contains fewer than PAGE_SIZE rows.
    #
    # This does NOT silently accept a failed page.
    # ------------------------------------------------------------------------

    if discovered_max_page == 1 and len(first_page_rows) == PAGE_SIZE:

        log(
            "No explicit maximum page discovered and page 1 is full; "
            "continuing sequential pagination"
        )

        next_page = 2

        while True:

            url = build_page_url(
                base_url=PRUACCESS_URL,
                fund_id=fund_id,
                view_type="TBL",
                start_date=start_date,
                end_date=end_date,
                page_number=next_page,
                page_size=PAGE_SIZE,
            )

            rows, metadata = fetch_page_with_retries(
                page=page,
                url=url,
                page_number=next_page,
            )

            if len(rows) > PAGE_SIZE:
                raise RuntimeError(
                    f"Page {next_page} returned "
                    f"{len(rows)} rows, exceeding page size {PAGE_SIZE}"
                )

            all_rows.extend(rows)
            page_metadata.append(metadata)

            if len(rows) < PAGE_SIZE:
                break

            next_page += 1

            if next_page > 10000:
                raise RuntimeError(
                    "Pagination exceeded safety limit of 10,000 pages"
                )

    # ------------------------------------------------------------------------
    # Page validation
    # ------------------------------------------------------------------------

    if not page_metadata:
        raise RuntimeError(
            "No pagination metadata was collected"
        )

    expected_page_numbers = list(
        range(1, len(page_metadata) + 1)
    )

    actual_page_numbers = [
        int(item["page"])
        for item in page_metadata
        if "page" in item
    ]

    if actual_page_numbers != expected_page_numbers:
        raise RuntimeError(
            "Pagination page sequence is invalid: "
            f"expected={expected_page_numbers}, "
            f"actual={actual_page_numbers}"
        )

    # ------------------------------------------------------------------------
    # Preserve every returned row.
    #
    # No deduplication.
    # ------------------------------------------------------------------------

    validation = validate_history_rows(
        rows=all_rows,
        requested_start=requested_start,
        requested_end=requested_end,
    )

    # Sort only after validation.
    all_rows.sort(
        key=lambda row: parse_history_date(
            row["date"]
        ) or date.min,
        reverse=True,
    )

    # Verify sorted result.
    if all_rows:
        newest = parse_history_date(all_rows[0]["date"])
        oldest = parse_history_date(all_rows[-1]["date"])

        if newest is None or oldest is None:
            raise RuntimeError(
                "Unable to parse final sorted history boundaries"
            )

        if newest < oldest:
            raise RuntimeError(
                "Final history ordering is invalid"
            )

    pagination = {
        "pageSize": PAGE_SIZE,
        "pagesSuccessfullyFetched": len(page_metadata),
        "pages": page_metadata,
        "totalRows": len(all_rows),
        "validation": validation,
    }

    return all_rows, pagination


# ============================================================================
# FUND EXTRACTION
# ============================================================================

def extract_single_fund(
    browser: Browser,
    fund: Dict[str, Any],
    pruaccess_options: List[Dict[str, str]],
) -> Dict[str, Any]:

    excel_row = fund["excelRow"]
    prudential_url = fund["prudentialUrl"]
    pruaccess_name = fund["pruAccessName"]

    log_section(
        f"PROCESSING EXCEL ROW {excel_row}: {pruaccess_name}"
    )

    context = browser.new_context(
        user_agent=USER_AGENT,
        timezone_id="Asia/Singapore",
    )

    try:
        context.set_default_timeout(REQUEST_TIMEOUT_MS)

        # --------------------------------------------------------------------
        # STEP 1 - Prudential official data
        # --------------------------------------------------------------------

        prudential_fund = get_prudential_fund_data(
            context=context,
            prudential_url=prudential_url,
        )

        fund_identifier = prudential_fund.get(
            "fundIdentifier"
        )

        fund_code = prudential_fund.get(
            "fundCode"
        )

        fund_name = prudential_fund.get(
            "fundName"
        )

        inception_raw = prudential_fund.get(
            "inceptionDate"
        )

        inception_date = parse_prudential_date(
            inception_raw
        )

        if inception_date is None:
            raise RuntimeError(
                "Could not determine official Prudential "
                f"inception date: {inception_raw!r}"
            )

        end_date = singapore_today()

        if inception_date > end_date:
            raise RuntimeError(
                "Official Prudential inception date is after "
                f"Singapore current date: "
                f"{inception_date} > {end_date}"
            )

        start_date_text = format_pruaccess_date(
            inception_date
        )

        end_date_text = format_pruaccess_date(
            end_date
        )

        log(
            f"Prudential fund: {fund_name}"
        )

        log(
            f"Identifier: {fund_identifier}"
        )

        log(
            f"Code: {fund_code}"
        )

        log(
            f"Inception: {inception_date.isoformat()}"
        )

        log(
            f"PruAccess requested range: "
            f"{start_date_text} -> {end_date_text}"
        )

        # --------------------------------------------------------------------
        # STEP 2 - Exact PruAccess fund matching
        # --------------------------------------------------------------------

        matched_option = match_exact_pruaccess_fund(
            excel_name=pruaccess_name,
            options=pruaccess_options,
        )

        fund_id = matched_option["value"]

        if not fund_id:
            raise RuntimeError(
                "Exact PruAccess fund match has an empty option value"
            )

        log(
            f"Exact PruAccess match: "
            f"{matched_option['text']} "
            f"(value={fund_id})"
        )

        # --------------------------------------------------------------------
        # STEP 3 - Open isolated PruAccess page
        # --------------------------------------------------------------------

        page = context.new_page()

        try:
            page.set_default_timeout(
                REQUEST_TIMEOUT_MS
            )

            page.goto(
                PRUACCESS_URL,
                wait_until="domcontentloaded",
                timeout=REQUEST_TIMEOUT_MS,
            )

            try:
                page.wait_for_load_state(
                    "networkidle",
                    timeout=REQUEST_TIMEOUT_MS,
                )
            except Exception:
                pass

            page.wait_for_timeout(
                PRUACCESS_INITIAL_WAIT_MS
            )

            # ----------------------------------------------------------------
            # STEP 4 - Submit exact PruAccess search
            # ----------------------------------------------------------------

            pre_submit = submit_pruaccess_form(
                page=page,
                fund_id=fund_id,
                start_date=start_date_text,
                end_date=end_date_text,
            )

            # ----------------------------------------------------------------
            # STEP 5 - Extract initial submitted result
            # ----------------------------------------------------------------

            initial_rows = extract_price_rows(page)

            verify_bid_result(
                page=page,
                rows=initial_rows,
            )

            # ----------------------------------------------------------------
            # STEP 6 - Historical pagination
            # ----------------------------------------------------------------

            bid_history, pagination = extract_all_pages(
                page=page,
                fund_id=fund_id,
                start_date=start_date_text,
                end_date=end_date_text,
                initial_rows=initial_rows,
            )

            # ----------------------------------------------------------------
            # STEP 7 - Final integrity checks
            # ----------------------------------------------------------------

            if not bid_history:
                raise RuntimeError(
                    "No historical BID rows extracted"
                )

            if pagination["totalRows"] != len(bid_history):
                raise RuntimeError(
                    "Pagination row count does not match extracted "
                    "history length"
                )

            summary = {
                "status": "success",

                "excelRow": excel_row,

                "prudentialUrl": prudential_url,

                "excelPruAccessName": pruaccess_name,

                "matchedPruAccessName": matched_option["text"],

                "pruAccessFundId": fund_id,

                "fundIdentifier": fund_identifier,

                "fundCode": fund_code,

                "fundName": fund_name,

                "officialInceptionDate": (
                    inception_date.isoformat()
                ),

                "requestedStartDate": (
                    inception_date.isoformat()
                ),

                "requestedEndDate": (
                    end_date.isoformat()
                ),

                "historicalRowCount": len(bid_history),

                "oldestHistoricalDate": (
                    bid_history[-1]["date"]
                ),

                "newestHistoricalDate": (
                    bid_history[0]["date"]
                ),

                "priceType": "BID",

                "syntheticData": False,

                "estimatedData": False,

                "interpolatedData": False,

                "fabricatedData": False,

                "carryForwardData": False,

                "completePagination": True,

                "paginationPages": (
                    pagination["pagesSuccessfullyFetched"]
                ),

                "completedAt": datetime.now(
                    SINGAPORE_TZ
                ).isoformat(),
            }

            result = {
                "status": "success",
                "excel": fund,
                "prudential": prudential_fund,
                "pruaccess": {
                    "fundId": fund_id,
                    "fundName": matched_option["text"],
                    "startDate": start_date_text,
                    "endDate": end_date_text,
                    "priceType": "BID",
                },
                "preSubmit": pre_submit,
                "pagination": pagination,
                "bidHistory": bid_history,
                "summary": summary,
            }

            log(
                f"SUCCESS: row {excel_row} "
                f"{fund_name or pruaccess_name} "
                f"-> {len(bid_history)} historical BID rows"
            )

            return result

        finally:
            page.close()

    finally:
        context.close()


# ============================================================================
# OUTPUT
# ============================================================================

def save_successful_fund(
    result: Dict[str, Any],
) -> Path:

    excel = result["excel"]
    prudential = result["prudential"]

    excel_row = excel["excelRow"]

    fund_identifier = prudential.get(
        "fundIdentifier"
    )

    fund_code = prudential.get(
        "fundCode"
    )

    fund_name = prudential.get(
        "fundName"
    )

    # ------------------------------------------------------------------------
    # IMPORTANT:
    # Avoid nested f-string syntax.
    # ------------------------------------------------------------------------

    filename_source = (
        fund_identifier
        or fund_code
        or fund_name
    )

    directory_name = (
        f"{excel_row}_"
        f"{safe_filename(filename_source)}"
    )

    output_directory = (
        FUNDS_OUTPUT_DIR / directory_name
    )

    output_directory.mkdir(
        parents=True,
        exist_ok=True,
    )

    save_json(
        output_directory / "prudential_fund.json",
        prudential,
    )

    save_json(
        output_directory / "pre_submit.json",
        result["preSubmit"],
    )

    save_json(
        output_directory / "pagination.json",
        result["pagination"],
    )

    save_json(
        output_directory / "bid_history.json",
        result["bidHistory"],
    )

    save_json(
        output_directory / "summary.json",
        result["summary"],
    )

    return output_directory


def save_failed_fund(
    fund: Dict[str, Any],
    error: Exception,
) -> Path:

    excel_row = fund["excelRow"]

    output_directory = (
        FUNDS_OUTPUT_DIR
        / f"{excel_row}_failed"
    )

    output_directory.mkdir(
        parents=True,
        exist_ok=True,
    )

    failure = {
        "status": "failed",

        "excel": fund,

        "error": str(error),

        "errorType": type(error).__name__,

        "traceback": traceback.format_exc(),

        "failedAt": datetime.now(
            SINGAPORE_TZ
        ).isoformat(),
    }

    save_json(
        output_directory / "failure.json",
        failure,
    )

    return output_directory


def build_consolidated_bid_history(
    successful_results: List[Dict[str, Any]],
) -> Dict[str, Any]:

    consolidated: Dict[str, Any] = {}

    for result in successful_results:

        prudential = result["prudential"]
        excel = result["excel"]

        key = (
            prudential.get("fundIdentifier")
            or prudential.get("fundCode")
            or str(excel["excelRow"])
        )

        consolidated[str(key)] = {
            "excelRow": excel["excelRow"],

            "fundIdentifier": prudential.get(
                "fundIdentifier"
            ),

            "fundCode": prudential.get(
                "fundCode"
            ),

            "fundName": prudential.get(
                "fundName"
            ),

            "priceType": "BID",

            "history": result["bidHistory"],
        }

    return consolidated


# ============================================================================
# MAIN
# ============================================================================

def main() -> int:

    started_at = datetime.now(
        SINGAPORE_TZ
    )

    log_section(
        "VGrat FMS - PruAccess Historical BID Extractor"
    )

    log(
        f"Python: {sys.version}"
    )

    log(
        f"Excel source: {EXCEL_FILE}"
    )

    log(
        f"Output directory: {OUTPUT_DIR}"
    )

    # ------------------------------------------------------------------------
    # Prepare output directories
    # ------------------------------------------------------------------------

    OUTPUT_DIR.mkdir(
        parents=True,
        exist_ok=True,
    )

    FUNDS_OUTPUT_DIR.mkdir(
        parents=True,
        exist_ok=True,
    )

    # ------------------------------------------------------------------------
    # Load Excel source of truth
    # ------------------------------------------------------------------------

    funds = load_excel_funds()

    if not funds:
        raise RuntimeError(
            "Excel source contains no populated fund URLs"
        )

    # ------------------------------------------------------------------------
    # Validate duplicate Excel rows / URLs
    # ------------------------------------------------------------------------

    seen_urls: Dict[str, int] = {}

    for fund in funds:

        url = fund["prudentialUrl"]

        normalized_url = url.strip()

        if normalized_url in seen_urls:
            raise RuntimeError(
                "Duplicate Prudential URL detected in Excel: "
                f"rows {seen_urls[normalized_url]} and "
                f"{fund['excelRow']}"
            )

        seen_urls[normalized_url] = fund["excelRow"]

    # ------------------------------------------------------------------------
    # Start browser
    # ------------------------------------------------------------------------

    with sync_playwright() as playwright:

        browser = playwright.chromium.launch(
            headless=True,
        )

        try:

            # ---------------------------------------------------------------
            # Load PruAccess fund list once.
            # ---------------------------------------------------------------

            option_context = browser.new_context(
                user_agent=USER_AGENT,
                timezone_id="Asia/Singapore",
            )

            try:

                option_page = option_context.new_page()

                try:
                    pruaccess_options = (
                        get_pruaccess_fund_options(
                            option_page
                        )
                    )

                finally:
                    option_page.close()

            finally:
                option_context.close()

            # ---------------------------------------------------------------
            # Process every Excel fund.
            # ---------------------------------------------------------------

            successful_results: List[Dict[str, Any]] = []
            failed_results: List[Dict[str, Any]] = []

            for index, fund in enumerate(
                funds,
                start=1,
            ):

                log(
                    f"Fund {index}/{len(funds)} "
                    f"(Excel row {fund['excelRow']})"
                )

                try:

                    result = extract_single_fund(
                        browser=browser,
                        fund=fund,
                        pruaccess_options=pruaccess_options,
                    )

                    output_directory = (
                        save_successful_fund(
                            result
                        )
                    )

                    log(
                        f"Saved successful fund to: "
                        f"{output_directory}"
                    )

                    successful_results.append(
                        result
                    )

                except Exception as exc:

                    log(
                        f"FAILED Excel row {fund['excelRow']}: "
                        f"{exc}"
                    )

                    failure_directory = (
                        save_failed_fund(
                            fund=fund,
                            error=exc,
                        )
                    )

                    log(
                        f"Saved failure to: "
                        f"{failure_directory}"
                    )

                    failed_results.append(
                        {
                            "excel": fund,
                            "error": str(exc),
                            "errorType": type(exc).__name__,
                        }
                    )

            # ----------------------------------------------------------------
            # Consolidated outputs
            # ----------------------------------------------------------------

            all_funds = []

            for result in successful_results:

                all_funds.append(
                    {
                        "excel": result["excel"],
                        "prudential": result["prudential"],
                        "pruaccess": result["pruaccess"],
                        "summary": result["summary"],
                    }
                )

            all_bid_history = (
                build_consolidated_bid_history(
                    successful_results
                )
            )

            save_json(
                OUTPUT_DIR / "all_funds.json",
                all_funds,
            )

            save_json(
                OUTPUT_DIR / "all_bid_history.json",
                all_bid_history,
            )

            completed_at = datetime.now(
                SINGAPORE_TZ
            )

            run_summary = {
                "status": (
                    "success"
                    if not failed_results
                    else "partial"
                ),

                "startedAt": started_at.isoformat(),

                "completedAt": completed_at.isoformat(),

                "excelSource": str(EXCEL_FILE),

                "sourceHierarchy": [
                    "Funds Links.xlsm",
                    "Prudential official fund data",
                    "PruAccess historical BID prices",
                ],

                "rules": {
                    "excelIsSourceOfTruth": True,
                    "processEveryPopulatedExcelRow": True,
                    "hardcodedFundCount": False,
                    "exactPruAccessMatching": True,
                    "fuzzyMatching": False,
                    "historicalPriceType": "BID",
                    "offerUsedAsBid": False,
                    "syntheticData": False,
                    "estimatedData": False,
                    "interpolatedData": False,
                    "fabricatedData": False,
                    "carryForwardData": False,
                    "completePaginationRequired": True,
                    "partialHistoryAccepted": False,
                    "duplicateRowsSilentlyRemoved": False,
                },

                "excelFundCount": len(funds),

                "successfulFundCount": len(
                    successful_results
                ),

                "failedFundCount": len(
                    failed_results
                ),

                "successfulExcelRows": [
                    result["excel"]["excelRow"]
                    for result in successful_results
                ],

                "failedFunds": failed_results,
            }

            save_json(
                OUTPUT_DIR / "run_summary.json",
                run_summary,
            )

            # ----------------------------------------------------------------
            # Final console summary
            # ----------------------------------------------------------------

            log_section(
                "RUN COMPLETE"
            )

            log(
                f"Excel funds: "
                f"{len(funds)}"
            )

            log(
                f"Successful: "
                f"{len(successful_results)}"
            )

            log(
                f"Failed: "
                f"{len(failed_results)}"
            )

            log(
                f"Output: "
                f"{OUTPUT_DIR}"
            )

            if failed_results:

                log(
                    "FAILED FUND ROWS:"
                )

                for failure in failed_results:

                    excel = failure["excel"]

                    log(
                        f"  Row {excel['excelRow']}: "
                        f"{excel['pruAccessName']} "
                        f"-> {failure['error']}"
                    )

                log(
                    "Run completed with failures. "
                    "Failed funds were NOT included as successful "
                    "historical BID data."
                )

                return 1

            log(
                "All Excel funds completed successfully."
            )

            return 0

        finally:
            browser.close()


# ============================================================================
# ENTRY POINT
# ============================================================================

if __name__ == "__main__":
    try:
        exit_code = main()
        sys.exit(exit_code)

    except KeyboardInterrupt:
        log("Interrupted by user")
        sys.exit(130)

    except Exception as exc:
        log(
            f"FATAL ERROR: {exc}"
        )

        traceback.print_exc()

        sys.exit(1)
