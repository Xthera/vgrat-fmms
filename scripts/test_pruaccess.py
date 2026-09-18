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
3. Read ilpfunds.json for fund information.
4. Open PruAccess Fund Performance page.
5. Match Excel Column B against the PruAccess fund selector.
6. Select Table view.
7. Leave Fund Price Type untouched.
8. Set:
      Start Date = Prudential inception date
      End Date   = today's date
9. Submit the PruAccess form.
10. Extract all historical BID prices through pagination.
11. Save the raw BID history.

Important:
- No synthetic data.
- No estimated data.
- No interpolation.
- No fuzzy fund matching.
- Offer prices are not used for the historical BID dataset.
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
OUTPUT_DIR = Path("output_pruaccess")

PRUACCESS_URL = (
    "https://pruaccess.prudential.com.sg/"
    "prulinkfund/viewFundPerformance.do"
)

PAGE_WAIT_MS = 1200
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
    Normalise names for exact comparison only.
    """

    value = clean_text(value).lower()

    value = value.replace("–", "-")
    value = value.replace("—", "-")

    value = re.sub(r"\s+", " ", value)

    return value.strip()


def parse_prudential_date(value: str) -> datetime:
    return datetime.strptime(
        value.strip(),
        "%d/%m/%Y",
    )


def pruaccess_date(dt: datetime) -> str:
    return dt.strftime("%d-%b-%Y")


def save_json(path: Path, data) -> None:
    path.write_text(
        json.dumps(
            data,
            indent=2,
            ensure_ascii=False,
        ),
        encoding="utf-8",
    )


# ============================================================
# EXCEL
# ============================================================

def read_excel_fund() -> tuple[str, str]:

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

    prudential_url = clean_text(
        ws["A2"].value
    )

    pruaccess_name = clean_text(
        ws["B2"].value
    )

    wb.close()

    if not prudential_url:
        raise ValueError(
            "Excel A2 is empty."
        )

    if not pruaccess_name:
        raise ValueError(
            "Excel B2 is empty."
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

    print("\n=== Prudential fund page ===")
    print(prudential_url)

    captured_api_urls = []

    def on_response(response):

        if "ilpfunds.json" in response.url:
            captured_api_urls.append(
                response.url
            )

    page.on(
        "response",
        on_response,
    )

    page.goto(
        prudential_url,
        wait_until="domcontentloaded",
        timeout=120000,
    )

    page.wait_for_timeout(3000)

    if not captured_api_urls:
        raise RuntimeError(
            "Could not capture Prudential ilpfunds.json request."
        )

    api_url = captured_api_urls[0]

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

    result = {
        "apiUrl": api_url,
        "fundIdentifier": fund.get(
            "fundIdentifier"
        ),
        "fundName": fund.get(
            "fundName"
        ),
        "fundCode": fund.get(
            "fundCode"
        ),
        "fundCurrency": fund.get(
            "fundCurrency"
        ),
        "unitCurrency": fund.get(
            "unitCurrency"
        ),
        "assetClass": fund.get(
            "assetClass"
        ),
        "assetSubClass": fund.get(
            "assetSubClass"
        ),
        "riskClassification": fund.get(
            "riskClassification"
        ),
        "bidPrice": fund.get(
            "bidPrice"
        ),
        "offerPrice": fund.get(
            "offerPrice"
        ),
        "valuationDate": fund.get(
            "valuationDate"
        ),
        "inceptionDate": fund.get(
            "inceptionDate"
        ),
        "cumulativeYtd": fund.get(
            "cumulativeYtd"
        ),
        "cumulative1m": fund.get(
            "cumulative1m"
        ),
        "cumulative3m": fund.get(
            "cumulative3m"
        ),
        "cumulative6m": fund.get(
            "cumulative6m"
        ),
        "cumulative1y": fund.get(
            "cumulative1y"
        ),
        "cumulative3y": fund.get(
            "cumulative3y"
        ),
        "cumulative5y": fund.get(
            "cumulative5y"
        ),
        "annualised3y": fund.get(
            "annualised3y"
        ),
        "annualised5y": fund.get(
            "annualised5y"
        ),
        "annualised10y": fund.get(
            "annualised10y"
        ),
        "annualisedSinceLaunch": fund.get(
            "annualisedSinceLaunch"
        ),
        "factsheetUrl": fund.get(
            "factsheetUrl"
        ),
        "fundObjective": fund.get(
            "fundObjective"
        ),
        "investmentManager": fund.get(
            "investmentManager"
        ),
        "hasDividend": fund.get(
            "hasDividend"
        ),
        "dividendRate": fund.get(
            "dividendRate"
        ),
        "raw": fund,
    }

    return result


# ============================================================
# PRUACCESS FUND OPTIONS
# ============================================================

def get_pruaccess_fund_options(page) -> list[dict]:

    options = page.locator(
        "#fundName option"
    )

    count = options.count()

    result = []

    for i in range(count):

        option = options.nth(i)

        result.append(
            {
                "text": clean_text(
                    option.inner_text()
                ),
                "value": option.get_attribute(
                    "value"
                ),
            }
        )

    return result


# ============================================================
# READONLY DATE HANDLING
# ============================================================

def set_readonly_input_value(
    page,
    selector: str,
    value: str,
) -> None:
    """
    PruAccess date inputs are readonly.

    We therefore set the DOM value using JavaScript
    instead of Playwright.fill().

    We also dispatch input/change events so any page
    JavaScript listening for those events sees the new value.
    """

    locator = page.locator(selector)

    if locator.count() == 0:
        raise RuntimeError(
            f"Could not find date input: {selector}"
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

    actual_value = locator.input_value()

    print(
        f"Set {selector}: {actual_value}"
    )

    if actual_value != value:
        raise RuntimeError(
            f"Failed to set {selector}. "
            f"Expected {value}, got {actual_value}"
        )


# ============================================================
# TABLE EXTRACTION
# ============================================================

def extract_price_rows(page) -> list[dict]:

    rows = page.locator(
        "table tr"
    )

    count = rows.count()

    extracted = []

    for i in range(count):

        row = rows.nth(i)

        cells = row.locator(
            "th, td"
        )

        cell_count = cells.count()

        if cell_count < 2:
            continue

        values = []

        for j in range(cell_count):

            values.append(
                clean_text(
                    cells.nth(j).inner_text()
                )
            )

        if not values:
            continue

        date_value = values[0]

        date_match = re.match(
            r"^\d{1,2}-[A-Za-z]{3}-\d{4}$",
            date_value,
        )

        if not date_match:
            continue

        bid_value = values[1]

        if not re.match(
            r"^-?\d+(?:\.\d+)?$",
            bid_value,
        ):
            continue

        extracted.append(
            {
                "date": date_value,
                "bidPrice": bid_value,
            }
        )

    return extracted


# ============================================================
# PAGINATION DIAGNOSTICS
# ============================================================

def get_pagination_snapshot(page) -> list[dict]:

    candidates = page.locator(
        "a, button, "
        "input[type='submit'], "
        "input[type='button']"
    )

    count = candidates.count()

    result = []

    interesting_words = {
        "first",
        "last",
        "next",
        "previous",
        "prev",
        "»",
        "›",
        "«",
        "‹",
    }

    for i in range(count):

        element = candidates.nth(i)

        try:
            text = clean_text(
                element.inner_text()
            )
        except Exception:
            text = ""

        try:
            value = element.get_attribute(
                "value"
            )
        except Exception:
            value = None

        if not text:
            text = clean_text(value)

        try:
            href = element.get_attribute(
                "href"
            )
        except Exception:
            href = None

        try:
            onclick = element.get_attribute(
                "onclick"
            )
        except Exception:
            onclick = None

        try:
            class_name = element.get_attribute(
                "class"
            )
        except Exception:
            class_name = None

        try:
            disabled = element.is_disabled()
        except Exception:
            disabled = False

        lowered = text.lower()

        interesting = (
            lowered in interesting_words
            or "last" in lowered
            or "next" in lowered
            or "previous" in lowered
            or "prev" in lowered
        )

        # Include numeric pagination links too.
        if text.isdigit():
            interesting = True

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


def table_signature(
    rows: list[dict],
) -> str:

    return "|".join(
        f"{row['date']}={row['bidPrice']}"
        for row in rows
    )


# ============================================================
# PAGINATION
# ============================================================

def click_pagination_control(
    page,
    target_text: str,
    current_signature: str,
) -> bool:
    """
    Try a pagination control and confirm that the
    historical table actually changed.
    """

    locator_candidates = [
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

    for locator in locator_candidates:

        try:

            if locator.count() == 0:
                continue

            element = locator.first

            if not element.is_visible():
                continue

            if element.is_disabled():
                continue

            old_url = page.url

            element.click(
                timeout=10000,
                no_wait_after=True,
            )

            page.wait_for_timeout(
                PAGE_WAIT_MS
            )

            if page.url != old_url:

                try:
                    page.wait_for_load_state(
                        "domcontentloaded",
                        timeout=10000,
                    )
                except Exception:
                    pass

                page.wait_for_timeout(
                    500
                )

            new_rows = extract_price_rows(
                page
            )

            new_signature = table_signature(
                new_rows
            )

            if (
                new_signature
                and new_signature != current_signature
            ):
                return True

        except Exception:
            continue

    return False


def extract_all_pages(page):

    print(
        "\n=== Extracting historical pages ==="
    )

    all_rows = {}

    visited_signatures = set()

    pagination_snapshots = []

    for page_attempt in range(
        1,
        MAX_PAGES + 1,
    ):

        current_rows = extract_price_rows(
            page
        )

        print(
            f"Page attempt {page_attempt}: "
            f"{len(current_rows)} rows"
        )

        if not current_rows:

            print(
                "No historical rows detected. "
                "Stopping."
            )

            break

        signature = table_signature(
            current_rows
        )

        if signature in visited_signatures:

            print(
                "Pagination loop detected. "
                "Stopping."
            )

            break

        visited_signatures.add(
            signature
        )

        for row in current_rows:

            key = (
                row["date"],
                row["bidPrice"],
            )

            all_rows[key] = row

        controls = get_pagination_snapshot(
            page
        )

        pagination_snapshots.append(
            {
                "pageAttempt": page_attempt,
                "url": page.url,
                "controls": controls,
                "rowCount": len(current_rows),
                "firstDate": current_rows[0]["date"],
                "lastDate": current_rows[-1]["date"],
            }
        )

        # ----------------------------------------------------
        # Try Next
        # ----------------------------------------------------

        moved = False

        next_controls = [
            "Next",
            "Next →",
            "Next→",
            "»",
            "›",
        ]

        for target in next_controls:

            if click_pagination_control(
                page,
                target,
                signature,
            ):

                print(
                    f"Moved to next page using: "
                    f"{target}"
                )

                moved = True
                break

        if moved:
            continue

        # ----------------------------------------------------
        # Try next numeric page.
        # ----------------------------------------------------

        next_number = str(
            page_attempt + 1
        )

        if click_pagination_control(
            page,
            next_number,
            signature,
        ):

            print(
                f"Moved to page "
                f"{next_number}"
            )

            continue

        # ----------------------------------------------------
        # Nothing else worked.
        # ----------------------------------------------------

        print(
            "No further pagination control "
            "detected."
        )

        break

    # --------------------------------------------------------
    # Sort newest → oldest
    # --------------------------------------------------------

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

    return (
        final_rows,
        pagination_snapshots,
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
        "PruAccess Historical Bid Price Diagnostic"
    )

    print(
        "============================================================"
    )

    # ========================================================
    # EXCEL
    # ========================================================

    prudential_url, excel_pruaccess_name = (
        read_excel_fund()
    )

    print(
        "\nExcel A2:"
    )

    print(
        prudential_url
    )

    print(
        "\nExcel B2:"
    )

    print(
        excel_pruaccess_name
    )

    # ========================================================
    # BROWSER
    # ========================================================

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

        print(
            "\n=== Prudential fund data ==="
        )

        print(
            json.dumps(
                prudential,
                indent=2,
                ensure_ascii=False,
            )
        )

        save_json(
            OUTPUT_DIR
            / "prudential_fund.json",
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
                "Prudential API did not provide "
                "inceptionDate."
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
        # PRUACCESS
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

        # ----------------------------------------------------
        # Save initial page
        # ----------------------------------------------------

        (
            OUTPUT_DIR
            / "pre_submit_visible_text.txt"
        ).write_text(
            page.locator("body").inner_text(),
            encoding="utf-8",
        )

        # ----------------------------------------------------
        # Fund options
        # ----------------------------------------------------

        options = get_pruaccess_fund_options(
            page
        )

        print(
            f"\nPruAccess fund options found: "
            f"{len(options)}"
        )

        save_json(
            OUTPUT_DIR
            / "fund_options.json",
            options,
        )

        # ----------------------------------------------------
        # Exact fund-name match
        # ----------------------------------------------------

        target_normalized = normalize_name(
            excel_pruaccess_name
        )

        matches = []

        for option in options:

            if normalize_name(
                option["text"]
            ) == target_normalized:

                matches.append(
                    option
                )

        if not matches:

            print(
                "\nERROR: No exact PruAccess "
                "fund-name match."
            )

            print(
                "\nExcel B2 normalized:"
            )

            print(
                target_normalized
            )

            raise RuntimeError(
                "Exact PruAccess fund-name "
                "match failed."
            )

        if len(matches) > 1:

            raise RuntimeError(
                "Multiple exact PruAccess "
                "fund-name matches found."
            )

        selected_option = matches[0]

        print(
            "\nMatched PruAccess fund:"
        )

        print(
            selected_option
        )

        # ----------------------------------------------------
        # Select fund
        # ----------------------------------------------------

        page.locator(
            "#fundName"
        ).select_option(
            selected_option["value"]
        )

        # ----------------------------------------------------
        # Table view
        # ----------------------------------------------------

        page.locator(
            "#viewType"
        ).select_option(
            "TBL"
        )

        # ----------------------------------------------------
        # Fund price type
        #
        # IMPORTANT:
        # Leave it untouched.
        # ----------------------------------------------------

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
        # DATES
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
                "Could not find PruAccess "
                "CSRF token."
            )

        csrf_value = csrf.input_value()

        if not csrf_value:

            raise RuntimeError(
                "PruAccess CSRF token is empty."
            )

        # ====================================================
        # VERIFY FORM VALUES BEFORE SUBMIT
        # ====================================================

        actual_start = page.locator(
            'input[name="startDate"]'
        ).input_value()

        actual_end = page.locator(
            'input[name="endDate"]'
        ).input_value()

        print(
            "\nVerified form values:"
        )

        print(
            f"Start Date: {actual_start}"
        )

        print(
            f"End Date:   {actual_end}"
        )

        if actual_start != start_date:

            raise RuntimeError(
                "Start Date verification failed."
            )

        if actual_end != end_date:

            raise RuntimeError(
                "End Date verification failed."
            )

        # ====================================================
        # PRE-SUBMIT DIAGNOSTIC
        # ====================================================

        pre_submit = {
            "prudentialUrl": prudential_url,
            "excelPruAccessName": (
                excel_pruaccess_name
            ),
            "matchedPruAccessName": (
                selected_option["text"]
            ),
            "matchedPruAccessValue": (
                selected_option["value"]
            ),
            "startDate": start_date,
            "endDate": end_date,
            "viewType": "TBL",
            "fundPriceType": fund_price_type_value,
            "csrfPresent": bool(
                csrf_value
            ),
        }

        save_json(
            OUTPUT_DIR
            / "pre_submit.json",
            pre_submit,
        )

        # ====================================================
        # SUBMIT FORM
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
        # SAVE RESULT PAGE
        # ====================================================

        html = page.content()

        (
            OUTPUT_DIR
            / "result.html"
        ).write_text(
            html,
            encoding="utf-8",
        )

        (
            OUTPUT_DIR
            / "result_visible_text.txt"
        ).write_text(
            page.locator("body").inner_text(),
            encoding="utf-8",
        )

        print(
            "\nReturned URL:"
        )

        print(
            page.url
        )

        # ====================================================
        # FIRST PAGE
        # ====================================================

        first_rows = extract_price_rows(
            page
        )

        print(
            f"\nCurrent page contains "
            f"{len(first_rows)} historical rows."
        )

        if not first_rows:

            raise RuntimeError(
                "PruAccess returned no "
                "historical price rows."
            )

        # ====================================================
        # ALL PAGES
        # ====================================================

        historical_rows, pagination = (
            extract_all_pages(page)
        )

        # ====================================================
        # RESULT
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
            f"Total unique BID observations: "
            f"{len(historical_rows)}"
        )

        if historical_rows:

            print(
                "\nNewest observation:"
            )

            print(
                historical_rows[0]
            )

            print(
                "\nOldest observation:"
            )

            print(
                historical_rows[-1]
            )

        # ====================================================
        # BID HISTORY
        # ====================================================

        bid_history = {
            "source": "PruAccess",
            "fundName": prudential.get(
                "fundName"
            ),
            "pruAccessFundName": (
                selected_option["text"]
            ),
            "fundIdentifier": (
                prudential.get(
                    "fundIdentifier"
                )
            ),
            "fundCode": (
                prudential.get(
                    "fundCode"
                )
            ),
            "currency": (
                prudential.get(
                    "fundCurrency"
                )
            ),
            "startDate": start_date,
            "endDate": end_date,
            "priceType": "BID",
            "observationCount": len(
                historical_rows
            ),
            "observations": historical_rows,
        }

        save_json(
            OUTPUT_DIR
            / "bid_history.json",
            bid_history,
        )

        # ====================================================
        # PAGINATION DIAGNOSTICS
        # ====================================================

        save_json(
            OUTPUT_DIR
            / "pagination.json",
            pagination,
        )

        # ====================================================
        # SUMMARY
        # ====================================================

        summary = {
            "status": "success",
            "excelFile": str(
                EXCEL_FILE
            ),
            "prudentialUrl": (
                prudential_url
            ),
            "excelPruAccessName": (
                excel_pruaccess_name
            ),
            "matchedPruAccessName": (
                selected_option["text"]
            ),
            "matchedPruAccessValue": (
                selected_option["value"]
            ),
            "prudentialFundName": (
                prudential.get(
                    "fundName"
                )
            ),
            "fundIdentifier": (
                prudential.get(
                    "fundIdentifier"
                )
            ),
            "fundCode": (
                prudential.get(
                    "fundCode"
                )
            ),
            "inceptionDate": (
                inception_raw
            ),
            "startDate": start_date,
            "endDate": end_date,
            "currentBidPrice": (
                prudential.get(
                    "bidPrice"
                )
            ),
            "currentOfferPrice": (
                prudential.get(
                    "offerPrice"
                )
            ),
            "fundPriceType": (
                fund_price_type_value
            ),
            "priceTypeExtracted": "BID",
            "historicalObservationCount": (
                len(historical_rows)
            ),
            "paginationAttempts": (
                len(pagination)
            ),
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
            OUTPUT_DIR
            / "summary.json",
            summary,
        )

        browser.close()

    # ========================================================
    # FILE LIST
    # ========================================================

    print(
        "\nFiles created:"
    )

    for path in sorted(
        OUTPUT_DIR.iterdir()
    ):

        print(
            f" - {path}"
        )

    print(
        "\nDone."
    )


if __name__ == "__main__":
    main()
