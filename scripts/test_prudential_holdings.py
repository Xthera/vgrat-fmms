#!/usr/bin/env python3

"""
Prudential official factsheet Top 10 holdings extractor.

TEST FUND
=========

PRULink ActiveInvest Portfolio - Balanced (Accumulation)

SOURCE
======

Prudential Singapore official fund page
and the official Prudential fund factsheet PDF.

PURPOSE
=======

This is an isolated test script.

It does NOT modify:

    scripts/test_pruaccess.py
    scripts/test_prudential_fund.py
    build_fund_data.py

WORKFLOW
========

1. Open the official Prudential fund page.
2. Locate the official Fund Factsheet link.
3. Download the factsheet PDF.
4. Extract PDF text using pypdf.
5. Locate the "Top 10 holdings" section.
6. Extract populated holding name + percentage rows.
7. Return up to 10 holdings.
8. Save all verification files.

HARD RULES
==========

- Official Prudential source only.
- Do not use third-party holdings data.
- Extract only holdings explicitly published by Prudential.
- Maximum 10 holdings.
- If Prudential publishes fewer than 10 populated holdings,
  return exactly those published holdings.
- Never invent holdings.
- Never infer holdings.
- Never fill missing holdings.
- Never fabricate percentages.
- If the holdings section cannot be extracted, the test fails.

OUTPUT
======

scripts/output_holdings/
    factsheet.pdf
    factsheet_text.txt
    top_holdings.json
    summary.json
"""


from __future__ import annotations

import asyncio
import json
import re
from io import BytesIO
from pathlib import Path
from typing import Any
from urllib.parse import urljoin

from playwright.async_api import (
    Browser,
    Page,
    async_playwright,
)
from pypdf import PdfReader


# ============================================================
# TEST FUND
# ============================================================

TEST_URL = (
    "https://www.prudential.com.sg/products/wealth-accumulation/"
    "ilp/prulink-funds/"
    "prulink-activeinvest-portfolio-balanced-accumulation"
)


# ============================================================
# OUTPUT
# ============================================================

OUTPUT_DIR = (
    Path("scripts")
    / "output_holdings"
)

OUTPUT_DIR.mkdir(
    parents=True,
    exist_ok=True,
)


# ============================================================
# SETTINGS
# ============================================================

BROWSER_HEADLESS = True

PAGE_TIMEOUT_MS = 120000

FACTSHEET_WAIT_MS = 5000

MAX_HOLDINGS = 10


# ============================================================
# HELPERS
# ============================================================

def clean_text(
    value: Any,
) -> str:

    if value is None:

        return ""

    return " ".join(
        str(value)
        .replace("\xa0", " ")
        .replace("\u2022", " ")
        .split()
    ).strip()


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


# ============================================================
# FIND OFFICIAL FACTSHEET LINK
# ============================================================

async def find_factsheet_url(
    page: Page,
) -> str:

    # --------------------------------------------------------
    # Primary method:
    #
    # Look for links containing "factsheet".
    # --------------------------------------------------------

    links = await page.locator(
        "a"
    ).evaluate_all(
        """
        elements => elements.map(element => ({
            text: (element.innerText || element.textContent || "").trim(),
            href: element.href || ""
        }))
        """
    )

    candidates = []

    for item in links:

        text = clean_text(
            item.get(
                "text"
            )
        )

        href = clean_text(
            item.get(
                "href"
            )
        )

        combined = (
            f"{text} {href}"
        ).casefold()

        if (
            "factsheet"
            in combined
        ):

            candidates.append(
                {
                    "text":
                        text,

                    "href":
                        href,
                }
            )

    if not candidates:

        raise RuntimeError(
            "Could not find an official Prudential "
            "Fund Factsheet link on the fund page."
        )

    # --------------------------------------------------------
    # Prefer a PDF link.
    # --------------------------------------------------------

    pdf_candidates = [
        item
        for item
        in candidates
        if ".pdf"
        in item["href"].casefold()
    ]

    selected = (
        pdf_candidates[0]
        if pdf_candidates
        else candidates[0]
    )

    factsheet_url = urljoin(
        TEST_URL,
        selected["href"],
    )

    if not factsheet_url:

        raise RuntimeError(
            "Factsheet link was found but its URL is empty."
        )

    print()
    print(
        "Official factsheet found:"
    )

    print(
        factsheet_url
    )

    return factsheet_url


# ============================================================
# DOWNLOAD FACTSHEET
# ============================================================

async def download_factsheet(
    page: Page,
    factsheet_url: str,
) -> bytes:

    print()
    print(
        "Downloading official Prudential factsheet..."
    )

    response = await page.request.get(
        factsheet_url,
        timeout=PAGE_TIMEOUT_MS,
        fail_on_status_code=False,
    )

    print(
        f"Factsheet HTTP status: "
        f"{response.status}"
    )

    if response.status != 200:

        raise RuntimeError(
            "Prudential factsheet download failed. "
            f"HTTP {response.status}"
        )

    content_type = clean_text(
        response.headers.get(
            "content-type",
            ""
        )
    )

    print(
        f"Content-Type: "
        f"{content_type}"
    )

    body = await response.body()

    if not body:

        raise RuntimeError(
            "Prudential factsheet returned an empty response."
        )

    # --------------------------------------------------------
    # Basic PDF signature validation.
    # --------------------------------------------------------

    if not body.startswith(
        b"%PDF"
    ):

        raise RuntimeError(
            "Downloaded factsheet does not appear to be a PDF."
        )

    print(
        f"Downloaded PDF bytes: "
        f"{len(body)}"
    )

    return body


# ============================================================
# EXTRACT PDF TEXT
# ============================================================

def extract_pdf_text(
    pdf_bytes: bytes,
) -> tuple[str, int]:

    reader = PdfReader(
        BytesIO(
            pdf_bytes
        )
    )

    page_count = len(
        reader.pages
    )

    if page_count == 0:

        raise RuntimeError(
            "Factsheet PDF contains no pages."
        )

    page_text = []

    for page_number, page in enumerate(
        reader.pages,
        start=1,
    ):

        try:

            text = page.extract_text()

        except Exception as error:

            raise RuntimeError(
                f"Could not extract text from "
                f"factsheet PDF page {page_number}: "
                f"{error}"
            ) from error

        if text:

            page_text.append(
                text
            )

    full_text = "\n".join(
        page_text
    )

    full_text = (
        full_text
        .replace("\r\n", "\n")
        .replace("\r", "\n")
    )

    if not clean_text(
        full_text
    ):

        raise RuntimeError(
            "Factsheet PDF text extraction returned empty text."
        )

    return (
        full_text,
        page_count,
    )


# ============================================================
# EXTRACT DATA-AS-AT DATE
# ============================================================

def extract_data_as_at(
    text: str,
) -> str | None:

    patterns = [
        r"All data as at\s+([0-9]{1,2}\s+[A-Za-z]{3,9}\s+[0-9]{4})",
        r"All data as at\s+([0-9]{1,2}\s+[A-Za-z]+\s+[0-9]{4})",
    ]

    for pattern in patterns:

        match = re.search(
            pattern,
            text,
            re.IGNORECASE,
        )

        if match:

            return clean_text(
                match.group(1)
            )

    return None


# ============================================================
# EXTRACT FACTSHEET DOCUMENT DATE
# ============================================================

def extract_factsheet_document_date(
    text: str,
) -> str | None:

    months = (
        "January|February|March|April|May|June|July|"
        "August|September|October|November|December"
    )

    pattern = (
        rf"\b({months})\s+20\d{{2}}\b"
    )

    # Prefer occurrences near the fund name.

    matches = re.findall(
        pattern,
        text,
        re.IGNORECASE,
    )

    if not matches:

        return None

    for match in re.finditer(
        pattern,
        text,
        re.IGNORECASE,
    ):

        value = clean_text(
            match.group(0)
        )

        return value

    return None


# ============================================================
# FIND TOP 10 HOLDINGS SECTION
# ============================================================

def extract_holdings_section(
    text: str,
) -> str:

    normalized_text = (
        text
        .replace(
            "\u00a0",
            " ",
        )
        .replace(
            "\r\n",
            "\n",
        )
        .replace(
            "\r",
            "\n",
        )
    )

    heading_match = re.search(
        r"\bTop\s+10\s+holdings\b",
        normalized_text,
        re.IGNORECASE,
    )

    if not heading_match:

        raise RuntimeError(
            'Could not find the "Top 10 holdings" '
            "section in the Prudential factsheet."
        )

    start = (
        heading_match.end()
    )

    # --------------------------------------------------------
    # The holdings section normally ends at the Source line.
    # --------------------------------------------------------

    end_patterns = [
        r"\n\s*Source\s*:",
        r"\n\s*Inception date\s*:",
        r"\n\s*Important Information\b",
        r"\n\s*Page\s+\d+\s+of\s+\d+",
    ]

    end = len(
        normalized_text
    )

    for pattern in end_patterns:

        match = re.search(
            pattern,
            normalized_text[
                start:
            ],
            re.IGNORECASE,
        )

        if match:

            candidate_end = (
                start
                + match.start()
            )

            if candidate_end < end:

                end = candidate_end

    section = normalized_text[
        start:end
    ]

    if not clean_text(
        section
    ):

        raise RuntimeError(
            'The "Top 10 holdings" section was found '
            "but contained no text."
        )

    return section


# ============================================================
# PARSE HOLDINGS
# ============================================================

def parse_holdings(
    section: str,
) -> list[dict[str, Any]]:

    lines = [
        clean_text(line)
        for line
        in section.splitlines()
    ]

    lines = [
        line
        for line
        in lines
        if line
    ]

    holdings = []

    pending_name = None

    # --------------------------------------------------------
    # Patterns
    # --------------------------------------------------------

    combined_pattern = re.compile(
        r"^(?P<name>.+?)\s+"
        r"(?P<weight>\d+(?:\.\d+)?)"
        r"%$"
    )

    weight_only_pattern = re.compile(
        r"^(?P<weight>\d+(?:\.\d+)?)%$"
    )

    for line in lines:

        # ----------------------------------------------------
        # Ignore obvious non-holding lines.
        # ----------------------------------------------------

        if re.search(
            r"^Source\s*:",
            line,
            re.IGNORECASE,
        ):

            break

        if re.search(
            r"^Inception date\s*:",
            line,
            re.IGNORECASE,
        ):

            break

        # ----------------------------------------------------
        # Name + percentage on one line.
        # ----------------------------------------------------

        combined_match = (
            combined_pattern.match(
                line
            )
        )

        if combined_match:

            name = clean_text(
                combined_match.group(
                    "name"
                )
            )

            weight_text = (
                combined_match.group(
                    "weight"
                )
            )

            if (
                name
                and not re.fullmatch(
                    r"\d+(?:\.\d+)?",
                    name,
                )
            ):

                holdings.append(
                    {
                        "rank":
                            len(
                                holdings
                            )
                            + 1,

                        "name":
                            name,

                        "weightPercent":
                            float(
                                weight_text
                            ),

                        "weightText":
                            f"{weight_text}%",
                    }
                )

                pending_name = None

                if len(
                    holdings
                ) >= MAX_HOLDINGS:

                    break

                continue

        # ----------------------------------------------------
        # Percentage-only line.
        # ----------------------------------------------------

        weight_only_match = (
            weight_only_pattern.match(
                line
            )
        )

        if weight_only_match:

            if pending_name:

                weight_text = (
                    weight_only_match.group(
                        "weight"
                    )
                )

                holdings.append(
                    {
                        "rank":
                            len(
                                holdings
                            )
                            + 1,

                        "name":
                            pending_name,

                        "weightPercent":
                            float(
                                weight_text
                            ),

                        "weightText":
                            f"{weight_text}%",
                    }
                )

                pending_name = None

                if len(
                    holdings
                ) >= MAX_HOLDINGS:

                    break

            continue

        # ----------------------------------------------------
        # Potential holding name.
        #
        # Do not treat generic labels as holdings.
        # ----------------------------------------------------

        if line.casefold() in {
            "top 10 holdings",
            "holdings",
            "source",
        }:

            continue

        if re.fullmatch(
            r"[0-9.\-]+",
            line,
        ):

            continue

        pending_name = line

    return holdings


# ============================================================
# VALIDATE HOLDINGS
# ============================================================

def validate_holdings(
    holdings: list[dict[str, Any]],
) -> None:

    if not holdings:

        raise RuntimeError(
            "No holdings were extracted from the official "
            "Prudential Top 10 holdings section."
        )

    if len(holdings) > MAX_HOLDINGS:

        raise RuntimeError(
            "Extracted more than 10 holdings."
        )

    names = [
        clean_text(
            holding["name"]
        ).casefold()
        for holding
        in holdings
    ]

    if len(
        names
    ) != len(
        set(names)
    ):

        raise RuntimeError(
            "Duplicate holding names were extracted."
        )

    for expected_rank, holding in enumerate(
        holdings,
        start=1,
    ):

        if (
            holding["rank"]
            != expected_rank
        ):

            raise RuntimeError(
                "Holding ranks are not sequential."
            )

        weight = holding.get(
            "weightPercent"
        )

        if not isinstance(
            weight,
            (int, float),
        ):

            raise RuntimeError(
                f"Invalid holding percentage for "
                f"{holding['name']!r}."
            )

        if weight < 0:

            raise RuntimeError(
                f"Negative holding percentage detected "
                f"for {holding['name']!r}."
            )


# ============================================================
# MAIN
# ============================================================

async def main():

    print(
        "============================================================"
    )

    print(
        "PRUDENTIAL TOP 10 HOLDINGS EXTRACTION TEST"
    )

    print(
        "============================================================"
    )

    print()
    print(
        "Test URL:"
    )

    print(
        TEST_URL
    )

    async with async_playwright() as playwright:

        browser: Browser = (
            await playwright.chromium.launch(
                headless=BROWSER_HEADLESS
            )
        )

        page = await browser.new_page(
            viewport={
                "width":
                    1440,

                "height":
                    1000,
            }
        )

        # ====================================================
        # OPEN FUND PAGE
        # ====================================================

        print()
        print(
            "Opening Prudential fund page..."
        )

        response = await page.goto(
            TEST_URL,
            wait_until="domcontentloaded",
            timeout=PAGE_TIMEOUT_MS,
        )

        if response:

            print(
                f"Fund page HTTP status: "
                f"{response.status}"
            )

        await page.wait_for_timeout(
            FACTSHEET_WAIT_MS
        )

        print(
            f"Page title: "
            f"{await page.title()}"
        )

        # ====================================================
        # FIND FACTSHEET
        # ====================================================

        factsheet_url = (
            await find_factsheet_url(
                page
            )
        )

        # ====================================================
        # DOWNLOAD FACTSHEET
        # ====================================================

        pdf_bytes = (
            await download_factsheet(
                page,
                factsheet_url,
            )
        )

        factsheet_file = (
            OUTPUT_DIR
            / "factsheet.pdf"
        )

        factsheet_file.write_bytes(
            pdf_bytes
        )

        print()
        print(
            f"Saved factsheet: "
            f"{factsheet_file}"
        )

        # ====================================================
        # EXTRACT PDF TEXT
        # ====================================================

        (
            pdf_text,
            page_count,
        ) = extract_pdf_text(
            pdf_bytes
        )

        text_file = (
            OUTPUT_DIR
            / "factsheet_text.txt"
        )

        text_file.write_text(
            pdf_text,
            encoding="utf-8",
        )

        print()
        print(
            f"Factsheet pages: "
            f"{page_count}"
        )

        print(
            f"Saved extracted text: "
            f"{text_file}"
        )

        # ====================================================
        # HOLDINGS SECTION
        # ====================================================

        holdings_section = (
            extract_holdings_section(
                pdf_text
            )
        )

        section_file = (
            OUTPUT_DIR
            / "top_holdings_section.txt"
        )

        section_file.write_text(
            holdings_section,
            encoding="utf-8",
        )

        # ====================================================
        # PARSE
        # ====================================================

        holdings = parse_holdings(
            holdings_section
        )

        # ====================================================
        # VALIDATE
        # ====================================================

        validate_holdings(
            holdings
        )

        # ====================================================
        # OTHER FACTSHEET METADATA
        # ====================================================

        data_as_at = (
            extract_data_as_at(
                pdf_text
            )
        )

        document_date = (
            extract_factsheet_document_date(
                pdf_text
            )
        )

        # ====================================================
        # RESULT
        # ====================================================

        result = {
            "status":
                "success",

            "source":
                "Prudential Singapore",

            "fundPageUrl":
                TEST_URL,

            "factsheetUrl":
                factsheet_url,

            "factsheetFile":
                str(
                    factsheet_file
                ),

            "factsheetPageCount":
                page_count,

            "factsheetDataAsAt":
                data_as_at,

            "factsheetDocumentDate":
                document_date,

            "holdingsSection":
                "Top 10 holdings",

            "topHoldingsCount":
                len(
                    holdings
                ),

            "topHoldings":
                holdings,

            "rules":
                {
                    "officialPrudentialSourceOnly":
                        True,

                    "maximumHoldings":
                        MAX_HOLDINGS,

                    "publishedCountUsedExactly":
                        True,

                    "noForcedTenEntries":
                        True,

                    "noInferredHoldings":
                        True,

                    "noFabricatedPercentages":
                        True,

                    "thirdPartyHoldingsData":
                        False,
                },
        }

        save_json(
            OUTPUT_DIR
            / "top_holdings.json",
            result,
        )

        # ====================================================
        # SUMMARY
        # ====================================================

        summary = {
            "status":
                "success",

            "fundPageUrl":
                TEST_URL,

            "factsheetUrl":
                factsheet_url,

            "factsheetPageCount":
                page_count,

            "factsheetDataAsAt":
                data_as_at,

            "factsheetDocumentDate":
                document_date,

            "topHoldingsCount":
                len(
                    holdings
                ),

            "topHoldings":
                holdings,
        }

        save_json(
            OUTPUT_DIR
            / "summary.json",
            summary,
        )

        # ====================================================
        # CONSOLE OUTPUT
        # ====================================================

        print()
        print(
            "============================================================"
        )

        print(
            "TOP HOLDINGS EXTRACTION SUCCESS"
        )

        print(
            "============================================================"
        )

        print(
            f"Factsheet: "
            f"{factsheet_url}"
        )

        print(
            f"Data as at: "
            f"{data_as_at or 'Not identified'}"
        )

        print(
            f"Holdings extracted: "
            f"{len(holdings)}"
        )

        print()

        for holding in holdings:

            print(
                f"{holding['rank']}. "
                f"{holding['name']} "
                f"- "
                f"{holding['weightText']}"
            )

        print()
        print(
            "Files saved to:"
        )

        print(
            f"  {OUTPUT_DIR}"
        )

        await browser.close()


# ============================================================
# ENTRY POINT
# ============================================================

if __name__ == "__main__":

    asyncio.run(
        main()
    )
