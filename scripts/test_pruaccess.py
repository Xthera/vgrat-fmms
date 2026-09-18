```python
#!/usr/bin/env python3

"""
PruAccess diagnostic + full historical bid-price extraction.

Source of truth:
- Funds Links.xlsm
  - Column A: Prudential fund URL
  - Column B: PruAccess fund name

Workflow:
1. Read A2/B2 from Funds Links.xlsm.
2. Open the Prudential fund page.
3. Read ilpfunds.json for:
   - fund name
   - citicode
   - inception date
   - current bid/offer
   - published returns
4. Open PruAccess Fund Performance page.
5. Match Excel Column B against the PruAccess fund selector.
6. Select Table view.
7. Leave Fund Price Type untouched.
8. Submit:
      Start Date = Prudential inception date
      End Date   = today's date
9. Extract every historical table page.
10. Save actual BID history only.

No synthetic, estimated, interpolated, or simulated data is created.
"""

from __future__ import annotations

import json
import re
import time
from datetime import datetime
from pathlib import Path
from urllib.parse import urljoin

from openpyxl import load_workbook
from playwright.sync_api import sync_playwright


# ============================================================
# CONFIG
# ============================================================

EXCEL_FILE = Path("Funds Links.xlsm")
OUTPUT_DIR = Path("output_pruaccess")

PRUDENTIAL_BASE = "https://www.prudential.com.sg"
PRUACCESS_URL = "https://pruaccess.prudential.com.sg/prulinkfund/viewFundPerformance.do"

# How long to wait after pagination actions.
PAGE_WAIT_MS = 1200

# Safety limit.
# A normal fund should not require anywhere near this many pages.
MAX_PAGES = 1000


# ============================================================
# HELPERS
# ============================================================

def clean_text(value) -> str:
    if value is None:
        return ""

    return re.sub(r"\s+", " ", str(value)).strip()


def normalize_name(value: str) -> str:
    """
    Normalise fund names for exact comparison.

    We deliberately do not perform fuzzy matching.
    """
    value = clean_text(value).lower()

    value = value.replace("–", "-")
    value = value.replace("—", "-")

    value = re.sub(r"\s+", " ", value)

    return value.strip()


def parse_prudential_date(value: str) -> datetime:
    """
    Parse Prudential dates such as:
        03/11/2021
        16/09/2026
    """
    return datetime.strptime(value.strip(), "%d/%m/%Y")


def pruaccess_date(dt: datetime) -> str:
    """
    Convert datetime to the format used by PruAccess:
        03-Nov-2021
    """
    return dt.strftime("%d-%b-%Y")


def safe_filename(value: str) -> str:
    value = re.sub(r"[^A-Za-z0-9._-]+", "_", value)
    return value.strip("_")[:150]


def save_json(path: Path, data) -> None:
    path.write_text(
        json.dumps(data, indent=2, ensure_ascii=False),
        encoding="utf-8",
    )


# ============================================================
# EXCEL
# ============================================================

def read_excel_fund() -> tuple[str, str]:
    """
    Read A2 and B2.

    Column A:
        Prudential URL

    Column B:
        PruAccess fund name
    """

    if not EXCEL_FILE.exists():
        raise FileNotFoundError(
            f"Excel file not found: {EXCEL_FILE}"
        )

    wb = load_workbook(
        EXCEL_FILE,
        read_only=True,
        keep_vba=True,
        data_only=True,
    )

    ws = wb.active

    prudential_url = clean_text(ws["A2"].value)
    pruaccess_name = clean_text(ws["B2"].value)

    wb.close()

    if not prudential_url:
        raise ValueError("Excel A2 is empty.")

    if not pruaccess_name:
        raise ValueError("Excel B2 is empty.")

    return prudential_url, pruaccess_name


# ============================================================
# PRUDENTIAL API
# ============================================================

def get_prudential_fund_data(page, prudential_url: str) -> dict:
    """
    Open the Prudential fund page and capture ilpfunds.json.
    """

    print("\n=== Prudential fund page ===")
    print(prudential_url)

    response_urls = []

    def on_response(response):
        url = response.url

        if "ilpfunds.json" in url:
            response_urls.append(url)

    page.on("response", on_response)

    page.goto(
        prudential_url,
        wait_until="domcontentloaded",
        timeout=120000,
    )

    # Allow page scripts/network requests to settle.
    page.wait_for_timeout(3000)

    # Prefer the captured API request.
    api_url = None

    for url in response_urls:
        if "ilpfunds.json" in url:
            api_url = url
            break

    if not api_url:
        raise RuntimeError(
            "Could not capture Prudential ilpfunds.json request."
        )

    print("\nCaptured Prudential API:")
    print(api_url)

    response = page.request.get(
        api_url,
        timeout=120000,
    )

    if not response.ok:
        raise RuntimeError(
            f"Prudential API returned HTTP {response.status}"
        )

    data = response.json()

    if not isinstance(data, list) or not data:
        raise RuntimeError(
            "Unexpected Prudential ilpfunds.json response."
        )

    fund = data[0]

    return {
        "apiUrl": api_url,
        "fundIdentifier": fund.get("fundIdentifier"),
        "fundName": fund.get("fundName"),
        "fundCode": fund.get("fundCode"),
        "fundCurrency": fund.get("fundCurrency"),
        "unitCurrency": fund.get("unitCurrency"),
        "assetClass": fund.get("assetClass"),
        "assetSubClass": fund.get("assetSubClass"),
        "riskClassification": fund.get("riskClassification"),
        "bidPrice": fund.get("bidPrice"),
        "offerPrice": fund.get("offerPrice"),
        "valuationDate": fund.get("valuationDate"),
        "inceptionDate": fund.get("inceptionDate"),
        "cumulativeYtd": fund.get("cumulativeYtd"),
        "cumulative1m": fund.get("cumulative1m"),
        "cumulative3m": fund.get("cumulative3m"),
        "cumulative6m": fund.get("cumulative6m"),
        "cumulative1y": fund.get("cumulative1y"),
        "cumulative3y": fund.get("cumulative3y"),
        "cumulative5y": fund.get("cumulative5y"),
        "annualised3y": fund.get("annualised3y"),
        "annualised5y": fund.get("annualised5y"),
        "annualised10y": fund.get("annualised10y"),
        "annualisedSinceLaunch": fund.get("annualisedSinceLaunch"),
        "factsheetUrl": fund.get("factsheetUrl"),
        "fundObjective": fund.get("fundObjective"),
        "investmentManager": fund.get("investmentManager"),
        "hasDividend": fund.get("hasDividend"),
        "dividendRate": fund.get("dividendRate"),
        "raw": fund,
    }


# ============================================================
# PRUACCESS FUND OPTIONS
# ============================================================

def get_pruaccess_fund_options(page) -> list[dict]:
    """
    Read every option from #fundName.
    """

    options = page.locator("#fundName option")

    count = options.count()

    result = []

    for i in range(count):
        option = options.nth(i)

        result.append(
            {
                "text": clean_text(option.inner_text()),
                "value": option.get_attribute("value"),
            }
        )

    return result


# ============================================================
# TABLE EXTRACTION
# ============================================================

def extract_price_rows(page) -> list[dict]:
    """
    Extract historical price rows from the current PruAccess table.

    Expected table columns:

        Date
        Bid Price
        Offer Price

    We only retain BID price.

    Returns:
        [
            {
                "date": "17-Sep-2026",
                "bidPrice": "1.11916"
            }
        ]
    """

    rows = page.locator("table tr")

    count = rows.count()

    extracted = []

    for i in range(count):
        row = rows.nth(i)

        cells = row.locator("th, td")

        cell_count = cells.count()

        if cell_count < 2:
            continue

        values = []

        for j in range(cell_count):
            values.append(
                clean_text(cells.nth(j).inner_text())
            )

        if not values:
            continue

        # Detect a date-like first column.
        first = values[0]

        date_match = re.match(
            r"^\d{1,2}-[A-Za-z]{3}-\d{4}$",
            first,
        )

        if not date_match:
            continue

        bid = values[1]

        # Ignore malformed rows.
        if not re.match(r"^-?\d+(?:\.\d+)?$", bid):
            continue

        extracted.append(
            {
                "date": first,
                "bidPrice": bid,
            }
        )

    return extracted


# ============================================================
# PAGINATION
# ============================================================

def get_pagination_snapshot(page) -> list[dict]:
    """
    Inspect visible pagination controls.

    This does not assume a particular implementation.
    It records:
      - text
      - href
      - onclick
      - class
      - disabled state

    This helps us understand the actual PruAccess pagination.
    """

    candidates = page.locator(
        "a, button, input[type='submit'], input[type='button']"
    )

    count = candidates.count()

    result = []

    for i in range(count):
        element = candidates.nth(i)

        try:
            text = clean_text(element.inner_text())
        except Exception:
            text = ""

        try:
            value = element.get_attribute("value")
        except Exception:
            value = None

        if not text:
            text = clean_text(value)

        try:
            href = element.get_attribute("href")
        except Exception:
            href = None

        try:
            onclick = element.get_attribute("onclick")
        except Exception:
            onclick = None

        try:
            class_name = element.get_attribute("class")
        except Exception:
            class_name = None

        try:
            disabled = element.is_disabled()
        except Exception:
            disabled = False

        lowered = text.lower()

        interesting = (
            lowered in {
                "1",
                "2",
                "3",
                "4",
                "5",
                "»",
                "last",
                "first",
                "next",
                "previous",
                "prev",
                "last →",
                "last→",
                "»",
            }
            or "page" in lowered
            or "last" in lowered
            or "next" in lowered
        )

        if interesting:
            result.append(
                {
                    "text": text,
                    "value": value,
                    "href": href,
                    "onclick": onclick,
                    "class": class_name,
                    "disabled": disabled,
                }
            )

    return result


def table_signature(rows: list[dict]) -> str:
    """
    Signature used to detect whether pagination actually changed
    the table.
    """

    if not rows:
        return ""

    return "|".join(
        f"{row['date']}={row['bidPrice']}"
        for row in rows
    )


def click_pagination_control(
    page,
    target_text: str,
    current_signature: str,
) -> bool:
    """
    Try to activate a pagination control.

    Returns True if the table changed.

    We try the most conservative methods first.
    """

    # --------------------------------------------------------
    # Find visible element by exact text.
    # --------------------------------------------------------

    locators = [
        page.get_by_role(
            "link",
            name=target_text,
            exact=True,
        ),
        page.get_by_role(
            "button",
            name=target_text,
            exact=True,
        ),
    ]

    for locator in locators:
        try:
            if locator.count() == 0:
                continue

            element = locator.first

            if not element.is_visible():
                continue

            if element.is_disabled():
                continue

            before_url = page.url

            element.click(
                timeout=10000,
                no_wait_after=True,
            )

            page.wait_for_timeout(PAGE_WAIT_MS)

            # If navigation occurred, wait for DOM.
            if page.url != before_url:
                try:
                    page.wait_for_load_state(
                        "domcontentloaded",
                        timeout=10000,
                    )
                except Exception:
                    pass

                page.wait_for_timeout(500)

            new_rows = extract_price_rows(page)
            new_signature = table_signature(new_rows)

            if new_signature and new_signature != current_signature:
                return True

        except Exception:
            continue

    return False


def extract_all_pages(page) -> list[dict]:
    """
    Extract all PruAccess pages.

    Strategy:
      1. Extract current table.
      2. Inspect pagination.
      3. Prefer "Last" when available.
      4. If Last is unavailable, walk Next/page numbers.
      5. Stop when no new table appears.
      6. Protect against loops using table signatures.

    The actual pagination controls are recorded for diagnostics.
    """

    print("\n=== Extracting historical pages ===")

    all_rows: dict[tuple[str, str], dict] = {}

    visited_signatures = set()

    pagination_snapshots = []

    for page_number in range(1, MAX_PAGES + 1):

        current_rows = extract_price_rows(page)

        signature = table_signature(current_rows)

        print(
            f"Page attempt {page_number}: "
            f"{len(current_rows)} rows"
        )

        if not current_rows:
            print("No table rows detected. Stopping.")
            break

        if signature in visited_signatures:
            print(
                "Table signature already seen. "
                "Pagination loop detected. Stopping."
            )
            break

        visited_signatures.add(signature)

        for row in current_rows:
            key = (
                row["date"],
                row["bidPrice"],
            )

            all_rows[key] = row

        snapshot = get_pagination_snapshot(page)

        pagination_snapshots.append(
            {
                "pageAttempt": page_number,
                "url": page.url,
                "controls": snapshot,
                "rows": len(current_rows),
                "firstDate": current_rows[0]["date"],
                "lastDate": current_rows[-1]["date"],
            }
        )

        # ----------------------------------------------------
        # Try NEXT first.
        #
        # This walks naturally through every page and avoids
        # making assumptions about how "Last" is implemented.
        # ----------------------------------------------------

        next_candidates = [
            "Next",
            "Next →",
            "Next→",
            "›",
            "»",
        ]

        moved = False

        for target in next_candidates:
            if click_pagination_control(
                page,
                target,
                signature,
            ):
                moved = True
                print(f"Moved to next page using: {target}")
                break

        if moved:
            continue

        # ----------------------------------------------------
        # If no Next control worked, try numbered pages.
        # ----------------------------------------------------

        for target in [
            str(page_number + 1),
        ]:
            if click_pagination_control(
                page,
                target,
                signature,
            ):
                moved = True
                print(f"Moved to page: {target}")
                break

        if moved:
            continue

        # ----------------------------------------------------
        # No further pagination.
        # ----------------------------------------------------

        print("No further pagination control detected.")
        break

    # Sort newest → oldest.
    def date_key(row):
        try:
            return datetime.strptime(
                row["date"],
                "%d-%b-%Y",
            )
        except Exception:
            return datetime.min

    final_rows = sorted(
        all_rows.values(),
        key=date_key,
        reverse=True,
    )

    return final_rows, pagination_snapshots


# ============================================================
# MAIN
# ============================================================

def main():

    OUTPUT_DIR.mkdir(
        parents=True,
        exist_ok=True,
    )

    print("============================================================")
    print("PruAccess Historical Bid Price Diagnostic")
    print("============================================================")

    # --------------------------------------------------------
    # Read Excel
    # --------------------------------------------------------

    prudential_url, excel_pruaccess_name = read_excel_fund()

    print("\nExcel A2:")
    print(prudential_url)

    print("\nExcel B2:")
    print(excel_pruaccess_name)

    # --------------------------------------------------------
    # Browser
    # --------------------------------------------------------

    with sync_playwright() as p:

        browser = p.chromium.launch(
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

        prudential = get_prudential_fund_data(
            page,
            prudential_url,
        )

        print("\n=== Prudential fund data ===")
        print(
            json.dumps(
                prudential,
                indent=2,
                ensure_ascii=False,
            )
        )

        # Save Prudential diagnostic.
        save_json(
            OUTPUT_DIR / "prudential_fund.json",
            prudential,
        )

        # ----------------------------------------------------
        # Inception date
        # ----------------------------------------------------

        inception_raw = prudential.get(
            "inceptionDate"
        )

        if not inception_raw:
            raise RuntimeError(
                "Prudential API did not provide inceptionDate."
            )

        inception_dt = parse_prudential_date(
            inception_raw
        )

        start_date = pruaccess_date(
            inception_dt
        )

        end_date = datetime.now().strftime(
            "%d-%b-%Y"
        )

        print("\nPruAccess start date:")
        print(start_date)

        print("\nPruAccess end date:")
        print(end_date)

        # ====================================================
        # PRUACCESS
        # ====================================================

        print("\n=== Opening PruAccess ===")

        page.goto(
            PRUACCESS_URL,
            wait_until="domcontentloaded",
            timeout=120000,
        )

        page.wait_for_timeout(1500)

        # ----------------------------------------------------
        # Save initial visible text
        # ----------------------------------------------------

        (OUTPUT_DIR / "pre_submit_visible_text.txt").write_text(
            page.locator("body").inner_text(),
            encoding="utf-8",
        )

        # ----------------------------------------------------
        # Fund options
        # ----------------------------------------------------

        options = get_pruaccess_fund_options(page)

        print(
            f"\nPruAccess fund options found: {len(options)}"
        )

        save_json(
            OUTPUT_DIR / "fund_options.json",
            options,
        )

        target_normalized = normalize_name(
            excel_pruaccess_name
        )

        matches = []

        for option in options:
            option_normalized = normalize_name(
                option["text"]
            )

            if option_normalized == target_normalized:
                matches.append(option)

        if not matches:

            # Print close diagnostic information, but do NOT
            # fuzzy-select a fund.
            print(
                "\nERROR: No exact PruAccess fund-name match."
            )

            print(
                "\nExcel B2 normalized:"
            )
            print(target_normalized)

            print(
                "\nAvailable PruAccess names:"
            )

            for option in options:
                print(
                    f"- {option['text']} "
                    f"[{option['value']}]"
                )

            raise RuntimeError(
                "Exact PruAccess fund-name match failed."
            )

        if len(matches) > 1:
            raise RuntimeError(
                "Multiple exact PruAccess fund-name matches found."
            )

        selected_option = matches[0]

        print("\nMatched PruAccess fund:")
        print(selected_option)

        # ----------------------------------------------------
        # Select fund
        # ----------------------------------------------------

        page.locator("#fundName").select_option(
            selected_option["value"]
        )

        # ----------------------------------------------------
        # Table view
        # ----------------------------------------------------

        page.locator("#viewType").select_option("TBL")

        # ----------------------------------------------------
        # DO NOT TOUCH FUND PRICE TYPE
        # ----------------------------------------------------

        fund_price_type = page.locator(
            "#fundPriceType"
        )

        fund_price_type_value = (
            fund_price_type.input_value()
            if fund_price_type.count()
            else None
        )

        print("\nFund Price Type:")
        print(
            f"{fund_price_type_value} "
            "(left untouched)"
        )

        # ----------------------------------------------------
        # Dates
        # ----------------------------------------------------

        page.locator(
            'input[name="startDate"]'
        ).fill(start_date)

        page.locator(
            'input[name="endDate"]'
        ).fill(end_date)

        # ----------------------------------------------------
        # CSRF
        # ----------------------------------------------------

        csrf = page.locator(
            'input[name="_csrf"]'
        )

        csrf_value = (
            csrf.input_value()
            if csrf.count()
            else None
        )

        if not csrf_value:
            raise RuntimeError(
                "Could not find PruAccess CSRF token."
            )

        # ----------------------------------------------------
        # Save pre-submit diagnostic
        # ----------------------------------------------------

        pre_submit = {
            "prudentialUrl": prudential_url,
            "excelPruAccessName": excel_pruaccess_name,
            "matchedPruAccessName": selected_option["text"],
            "matchedPruAccessValue": selected_option["value"],
            "startDate": start_date,
            "endDate": end_date,
            "viewType": "TBL",
            "fundPriceType": fund_price_type_value,
            "csrfPresent": bool(csrf_value),
        }

        save_json(
            OUTPUT_DIR / "pre_submit.json",
            pre_submit,
        )

        # ====================================================
        # SUBMIT
        # ====================================================

        print("\n=== Submitting PruAccess form ===")

        # Preserve the previously proven submission method.
        page.locator("#fundForm").evaluate(
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

        page.wait_for_timeout(2500)

        # ----------------------------------------------------
        # Save returned page
        # ----------------------------------------------------

        html = page.content()

        (OUTPUT_DIR / "result.html").write_text(
            html,
            encoding="utf-8",
        )

        (OUTPUT_DIR / "result_visible_text.txt").write_text(
            page.locator("body").inner_text(),
            encoding="utf-8",
        )

        print("\nReturned URL:")
        print(page.url)

        # ----------------------------------------------------
        # Extract first/current page
        # ----------------------------------------------------

        first_rows = extract_price_rows(page)

        print(
            f"\nCurrent page contains "
            f"{len(first_rows)} historical rows."
        )

        if not first_rows:
            raise RuntimeError(
                "PruAccess returned no historical price rows."
            )

        # ----------------------------------------------------
        # Extract every page
        # ----------------------------------------------------

        historical_rows, pagination = extract_all_pages(
            page
        )

        print("\n============================================================")
        print("EXTRACTION COMPLETE")
        print("============================================================")

        print(
            f"Total unique BID observations: "
            f"{len(historical_rows)}"
        )

        if historical_rows:
            print(
                f"Newest observation: "
                f"{historical_rows[0]}"
            )

            print(
                f"Oldest observation: "
                f"{historical_rows[-1]}"
            )

        # ====================================================
        # SAVE BID HISTORY
        # ====================================================

        bid_history = {
            "source": "PruAccess",
            "fundName": prudential.get("fundName"),
            "pruAccessFundName": selected_option["text"],
            "fundIdentifier": prudential.get("fundIdentifier"),
            "fundCode": prudential.get("fundCode"),
            "currency": prudential.get("fundCurrency"),
            "startDate": start_date,
            "endDate": end_date,
            "priceType": "BID",
            "observationCount": len(historical_rows),
            "observations": historical_rows,
        }

        save_json(
            OUTPUT_DIR / "bid_history.json",
            bid_history,
        )

        # ====================================================
        # SAVE PAGINATION DIAGNOSTICS
        # ====================================================

        save_json(
            OUTPUT_DIR / "pagination.json",
            pagination,
        )

        # ====================================================
        # SUMMARY
        # ====================================================

        summary = {
            "status": "success",
            "excelFile": str(EXCEL_FILE),
            "prudentialUrl": prudential_url,
            "excelPruAccessName": excel_pruaccess_name,
            "matchedPruAccessName": selected_option["text"],
            "matchedPruAccessValue": selected_option["value"],
            "prudentialFundName": prudential.get("fundName"),
            "fundIdentifier": prudential.get("fundIdentifier"),
            "fundCode": prudential.get("fundCode"),
            "inceptionDate": inception_raw,
            "startDate": start_date,
            "endDate": end_date,
            "currentBidPrice": prudential.get("bidPrice"),
            "currentOfferPrice": prudential.get("offerPrice"),
            "fundPriceType": fund_price_type_value,
            "priceTypeExtracted": "BID",
            "historicalObservationCount": len(
                historical_rows
            ),
            "paginationAttempts": len(pagination),
            "oldestObservation": (
                historical_rows[-1]
                if historical_rows
                else None
            ),
            "newestObservation": (
                historical_rows[0]
                if historical_rows
                else None
            ),
        }

        save_json(
            OUTPUT_DIR / "summary.json",
            summary,
        )

        browser.close()

    print("\nFiles created:")
    for path in sorted(OUTPUT_DIR.iterdir()):
        print(f" - {path}")

    print("\nDone.")


if __name__ == "__main__":
    main()
```
