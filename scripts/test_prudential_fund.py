import asyncio
import json
import re
from pathlib import Path
from playwright.async_api import async_playwright


# ============================================================
# TEST URL
# ============================================================

TEST_URL = (
    "https://www.prudential.com.sg/content/"
    "prudential-aem-lbu/pacs/en/products/wealth/ilp/"
    "prulink-funds/"
    "prulink-strategicinvest-income-fund-distribution.html"
)


# ============================================================
# OUTPUT
# ============================================================

OUTPUT_DIR = Path("output")
OUTPUT_DIR.mkdir(
    parents=True,
    exist_ok=True
)


# ============================================================
# EXISTING PRUDENTIAL JSON SOURCES
#
# DO NOT REMOVE OR CHANGE THESE.
# ============================================================

IMPORTANT_PARTS = [
    "ilpfunds.json",
    "ilpseries.json",
]


# ============================================================
# NEW:
# DIVIDEND / DISTRIBUTION KEYWORDS
# ============================================================

DIVIDEND_KEYWORDS = [
    "dividend",
    "distribution",
    "distributions",
    "distributionrate",
    "distribution_rate",
    "dividendrate",
    "dividend_rate",
    "dividendamount",
    "dividend_amount",
    "distributionamount",
    "distribution_amount",
    "payout",
    "payoutamount",
    "payout_amount",
    "distributiondate",
    "distribution_date",
    "dividenddate",
    "dividend_date",
    "exdate",
    "ex_date",
    "recorddate",
    "record_date",
    "paymentdate",
    "payment_date",
]


# ============================================================
# NEW:
# TEXT INDICATORS THAT SUGGEST DISTRIBUTION FUND
# ============================================================

DISTRIBUTION_TEXT_PATTERNS = [
    r"\bdistribution\s+class\b",
    r"\bdistribution\s+fund\b",
    r"\bdividend\b",
    r"\bdividends\b",
    r"\bdistribution\b",
    r"\bdistributions\b",
]


# ============================================================
# NEW:
# RATE FIELD KEYWORDS
#
# These are intentionally broad because we do not yet know
# Prudential's exact API field naming.
# ============================================================

RATE_KEYWORDS = [
    "rate",
    "amount",
    "percentage",
    "percent",
    "payout",
    "distribution",
    "dividend",
]


# ============================================================
# RECURSIVE JSON SEARCH
# ============================================================

def find_dividend_fields(
    value,
    path="$",
    matches=None
):

    if matches is None:
        matches = []

    # --------------------------------------------------------
    # Dictionary
    # --------------------------------------------------------

    if isinstance(value, dict):

        for key, child in value.items():

            key_text = str(key).lower()

            matched_keywords = [
                keyword
                for keyword in DIVIDEND_KEYWORDS
                if keyword in key_text
            ]

            if matched_keywords:

                matches.append({
                    "path": f"{path}.{key}",
                    "key": key,
                    "matchedKeywords": matched_keywords,
                    "value": child
                })

            find_dividend_fields(
                child,
                f"{path}.{key}",
                matches
            )

    # --------------------------------------------------------
    # List
    # --------------------------------------------------------

    elif isinstance(value, list):

        for index, child in enumerate(value):

            find_dividend_fields(
                child,
                f"{path}[{index}]",
                matches
            )

    return matches


# ============================================================
# LOAD JSON
# ============================================================

def load_json_file(path):

    try:

        text = path.read_text(
            encoding="utf-8"
        )

        return json.loads(text)

    except Exception as e:

        print(
            "Could not parse JSON:",
            path,
            repr(e)
        )

        return None


# ============================================================
# NORMALIZE TEXT
# ============================================================

def normalize_text(value):

    if value is None:
        return ""

    return re.sub(
        r"\s+",
        " ",
        str(value)
    ).strip()


# ============================================================
# DETECT DISTRIBUTION FUND FROM TEXT
# ============================================================

def detect_distribution_from_text(text):

    normalized = normalize_text(
        text
    )

    matches = []

    for pattern in DISTRIBUTION_TEXT_PATTERNS:

        found = re.search(
            pattern,
            normalized,
            flags=re.IGNORECASE
        )

        if found:

            matches.append({
                "pattern": pattern,
                "matchedText": found.group(0)
            })

    return matches


# ============================================================
# SEARCH TEXT FOR POSSIBLE DIVIDEND RATE
#
# IMPORTANT:
# This does NOT calculate anything.
#
# It only captures a number when the page explicitly places
# a percentage/rate/amount close to dividend/distribution text.
# ============================================================

def find_dividend_rates_in_text(text):

    normalized = normalize_text(
        text
    )

    results = []

    # --------------------------------------------------------
    # Examples:
    #
    # Dividend rate: 1.25%
    # Distribution rate 2.50%
    # Dividend: 0.0125
    # Distribution amount: 0.50
    # --------------------------------------------------------

    patterns = [

        (
            "percentage_rate",
            r"(?i)"
            r"(?:dividend|distribution|payout)"
            r"(?:\s+\w+){0,4}"
            r"\s*(?:rate|percentage|percent)"
            r"\s*[:\-]?\s*"
            r"([0-9]+(?:\.[0-9]+)?)\s*%"
        ),

        (
            "rate_before_label",
            r"(?i)"
            r"([0-9]+(?:\.[0-9]+)?)\s*%"
            r"\s*"
            r"(?:dividend|distribution|payout)"
            r"(?:\s+\w+){0,4}"
            r"\s*(?:rate|percentage|percent)"
        ),

        (
            "amount",
            r"(?i)"
            r"(?:dividend|distribution|payout)"
            r"(?:\s+\w+){0,4}"
            r"\s*(?:amount|rate)"
            r"\s*[:\-]?\s*"
            r"([0-9]+(?:\.[0-9]+)?)"
        ),

    ]

    for name, pattern in patterns:

        for match in re.finditer(
            pattern,
            normalized
        ):

            results.append({
                "type": name,
                "matchedText": match.group(0),
                "value": match.group(1)
            })

    return results


# ============================================================
# DETECT DISTRIBUTION FROM URL
# ============================================================

def detect_distribution_from_url(url):

    lower_url = url.lower()

    indicators = []

    if "distribution" in lower_url:

        indicators.append(
            "URL contains 'distribution'"
        )

    if "dividend" in lower_url:

        indicators.append(
            "URL contains 'dividend'"
        )

    return indicators


# ============================================================
# MAIN
# ============================================================

async def main():

    captured_json = []
    captured_requests = []

    async with async_playwright() as p:

        browser = await p.chromium.launch(
            headless=True
        )

        page = await browser.new_page(
            viewport={
                "width": 1440,
                "height": 1000
            }
        )

        # ====================================================
        # EXISTING RESPONSE HANDLER
        #
        # THIS IS KEPT FUNCTIONALLY THE SAME.
        # ====================================================

        async def handle_response(response):

            url = response.url

            if not any(
                part in url
                for part in IMPORTANT_PARTS
            ):
                return

            print()
            print(
                "FOUND PRUDENTIAL JSON:"
            )
            print(url)
            print(
                "STATUS:",
                response.status
            )

            try:

                body = await response.text()

                filename = (
                    f"response_"
                    f"{len(captured_json) + 1}"
                    f".json"
                )

                output_file = (
                    OUTPUT_DIR /
                    filename
                )

                output_file.write_text(
                    body,
                    encoding="utf-8"
                )

                captured_json.append({
                    "filename":
                        filename,

                    "url":
                        url,

                    "status":
                        response.status,

                    "contentType":
                        response.headers.get(
                            "content-type",
                            ""
                        )
                })

            except Exception as e:

                print(
                    "Could not read response:",
                    repr(e)
                )

        # ====================================================
        # EXISTING REQUEST HANDLER
        #
        # ALSO KEPT.
        # ====================================================

        async def handle_request(request):

            url = request.url

            if any(
                part in url
                for part in IMPORTANT_PARTS
            ):

                captured_requests.append({
                    "method":
                        request.method,

                    "url":
                        request.url,

                    "postData":
                        request.post_data
                })

        page.on(
            "response",
            handle_response
        )

        page.on(
            "request",
            handle_request
        )

        # ====================================================
        # OPEN PAGE
        # ====================================================

        print(
            "Opening Prudential page..."
        )

        print(
            TEST_URL
        )

        await page.goto(
            TEST_URL,
            wait_until="domcontentloaded",
            timeout=120000
        )

        # ====================================================
        # WAIT FOR JAVASCRIPT
        # ====================================================

        print(
            "Waiting for Prudential JavaScript..."
        )

        await page.wait_for_timeout(
            15000
        )

        # ====================================================
        # SCROLL
        # ====================================================

        await page.evaluate(
            """
            window.scrollTo(
                0,
                document.body.scrollHeight
            );
            """
        )

        await page.wait_for_timeout(
            5000
        )

        # ====================================================
        # SAVE RENDERED HTML
        # ====================================================

        html = await page.content()

        (
            OUTPUT_DIR /
            "rendered.html"
        ).write_text(
            html,
            encoding="utf-8"
        )

        # ====================================================
        # SAVE VISIBLE TEXT
        # ====================================================

        visible_text = await page.locator(
            "body"
        ).inner_text()

        (
            OUTPUT_DIR /
            "visible_text.txt"
        ).write_text(
            visible_text,
            encoding="utf-8"
        )

        # ====================================================
        # SAVE REQUEST LIST
        # ====================================================

        (
            OUTPUT_DIR /
            "requests.json"
        ).write_text(
            json.dumps(
                captured_requests,
                indent=2,
                ensure_ascii=False
            ),
            encoding="utf-8"
        )

        # ====================================================
        # EXISTING SUMMARY
        # ====================================================

        summary = {
            "testUrl":
                TEST_URL,

            "jsonResponsesCaptured":
                len(captured_json),

            "responses":
                captured_json
        }

        (
            OUTPUT_DIR /
            "summary.json"
        ).write_text(
            json.dumps(
                summary,
                indent=2,
                ensure_ascii=False
            ),
            encoding="utf-8"
        )

        # ====================================================
        # NEW DIVIDEND ANALYSIS
        # ====================================================

        print()
        print(
            "========================================"
        )
        print(
            "DIVIDEND / DISTRIBUTION ANALYSIS"
        )
        print(
            "========================================"
        )

        # ----------------------------------------------------
        # URL detection
        # ----------------------------------------------------

        url_indicators = (
            detect_distribution_from_url(
                TEST_URL
            )
        )

        # ----------------------------------------------------
        # Visible page text detection
        # ----------------------------------------------------

        text_indicators = (
            detect_distribution_from_text(
                visible_text
            )
        )

        # ----------------------------------------------------
        # Search visible text for rates
        # ----------------------------------------------------

        visible_text_rates = (
            find_dividend_rates_in_text(
                visible_text
            )
        )

        # ----------------------------------------------------
        # Search captured JSON
        # ----------------------------------------------------

        json_analysis = []

        all_json_dividend_matches = []

        for item in captured_json:

            json_file = (
                OUTPUT_DIR /
                item["filename"]
            )

            data = load_json_file(
                json_file
            )

            if data is None:

                continue

            matches = find_dividend_fields(
                data
            )

            json_analysis.append({

                "filename":
                    item["filename"],

                "url":
                    item["url"],

                "matches":
                    matches

            })

            all_json_dividend_matches.extend(
                matches
            )

        # ====================================================
        # DETERMINE WHETHER THIS APPEARS TO BE A
        # DISTRIBUTION / DIVIDEND FUND
        # ====================================================

        is_distribution_fund = (
            len(url_indicators) > 0
            or len(text_indicators) > 0
            or len(all_json_dividend_matches) > 0
        )

        # ====================================================
        # FIND POSSIBLE OFFICIAL RATE
        # ====================================================

        possible_rates = []

        # From visible page

        for item in visible_text_rates:

            possible_rates.append({

                "source":
                    "visible_page_text",

                **item

            })

        # From JSON

        for item in all_json_dividend_matches:

            key_lower = str(
                item["key"]
            ).lower()

            # Only treat it as a possible rate if
            # the field itself looks rate/amount related.

            if any(
                keyword in key_lower
                for keyword in RATE_KEYWORDS
            ):

                value = item["value"]

                # Do not automatically label arbitrary
                # dividend-related objects as rates.

                if isinstance(
                    value,
                    (
                        str,
                        int,
                        float
                    )
                ):

                    possible_rates.append({

                        "source":
                            "prudential_json",

                        "path":
                            item["path"],

                        "key":
                            item["key"],

                        "value":
                            value

                    })

        # ====================================================
        # REMOVE DUPLICATES
        # ====================================================

        unique_rates = []

        seen_rates = set()

        for item in possible_rates:

            fingerprint = json.dumps(
                item,
                sort_keys=True,
                ensure_ascii=False
            )

            if fingerprint in seen_rates:

                continue

            seen_rates.add(
                fingerprint
            )

            unique_rates.append(
                item
            )

        # ====================================================
        # FINAL DIVIDEND RESULT
        # ====================================================

        dividend_analysis = {

            "testUrl":
                TEST_URL,

            "isDistributionOrDividendFund":
                is_distribution_fund,

            "detection": {

                "urlIndicators":
                    url_indicators,

                "visiblePageIndicators":
                    text_indicators,

                "jsonDividendFieldCount":
                    len(
                        all_json_dividend_matches
                    )

            },

            "possibleOfficialRates":
                unique_rates,

            "jsonDividendFields":
                json_analysis,

            "rules": {

                "rateCalculated":
                    False,

                "rateInferred":
                    False,

                "rateEstimated":
                    False,

                "rateSynthetic":
                    False

            }

        }

        (
            OUTPUT_DIR /
            "dividend_analysis.json"
        ).write_text(
            json.dumps(
                dividend_analysis,
                indent=2,
                ensure_ascii=False
            ),
            encoding="utf-8"
        )

        # ====================================================
        # PRINT DIVIDEND RESULTS
        # ====================================================

        print()
        print(
            "Distribution / Dividend fund:",
            is_distribution_fund
        )

        print(
            "URL indicators:",
            len(url_indicators)
        )

        print(
            "Visible page indicators:",
            len(text_indicators)
        )

        print(
            "JSON dividend fields:",
            len(all_json_dividend_matches)
        )

        print(
            "Possible official rates:",
            len(unique_rates)
        )

        # ----------------------------------------------------
        # Print detected rates
        # ----------------------------------------------------

        if unique_rates:

            print()
            print(
                "POSSIBLE OFFICIAL DIVIDEND "
                "OR DISTRIBUTION VALUES:"
            )

            for item in unique_rates:

                print(
                    json.dumps(
                        item,
                        ensure_ascii=False
                    )
                )

        else:

            print()
            print(
                "No explicit dividend/distribution "
                "rate was found."
            )

            print(
                "No rate has been calculated or inferred."
            )

        # ====================================================
        # FINAL EXISTING OUTPUT
        # ====================================================

        print()
        print(
            "========================================"
        )
        print(
            "PRUDENTIAL JSON CAPTURE COMPLETE"
        )
        print(
            "========================================"
        )

        print(
            "JSON responses:",
            len(captured_json)
        )

        for item in captured_json:

            print(
                item["filename"],
                "->",
                item["url"]
            )

        print()
        print(
            "Files saved to output/"
        )

        print(
            "New dividend analysis:"
        )

        print(
            "output/dividend_analysis.json"
        )

        await browser.close()


if __name__ == "__main__":

    asyncio.run(main())
