#!/usr/bin/env python3

"""
Prudential Singapore official fund factsheet + Top 10 holdings extractor.

MASTER SOURCE
=============

Excel:
    Funds Links.xlsm

Column A:
    Prudential official fund URL

Column B:
    Exact PruAccess fund name

HARD RULES
==========

1. Excel Column A controls the fund universe.
2. Every populated Prudential URL in Column A is processed.
3. No hardcoded fund count.
4. Only official Prudential Singapore sources are allowed.
5. Third-party holdings sources are not allowed.
6. The official Prudential fund page is used to locate the official factsheet.
7. Generic Prudential Fund Reports pages are NOT accepted as a factsheet.
8. The final downloaded document MUST be a PDF.
9. No inferred holding names.
10. No fabricated holding names.
11. No inferred percentages.
12. No calculated percentages.
13. No fabricated percentages.
14. No forced ten holdings.
15. If Prudential publishes fewer than 10 holdings, store exactly that count.
16. Holdings remain in the order published by Prudential.
17. Holding rank is the primary row boundary.
18. Published percentage is optional.
19. Missing published percentage is stored as null.
20. Duplicate percentages are allowed.
21. Duplicate holding names are not allowed.
22. If a Top 10 holdings section exists but no identifiable holding names
    can be extracted, the fund fails.
23. If no Top 10 holdings section exists at all, classify as
    no_holdings_section.
24. No synthetic data.
25. No estimated data.
26. No interpolation.
27. No carry-forward.
28. This script does not modify PruAccess settings or test_pruaccess.py.
29. This script is a collector/tester only.
"""

import json
import re
import sys
import time
from datetime import datetime, timezone
from io import BytesIO
from pathlib import Path
from typing import Any, Optional
from urllib.parse import urljoin, urlparse

import openpyxl
import requests
from bs4 import BeautifulSoup
from pypdf import PdfReader


# ============================================================================
# PATHS / CONSTANTS
# ============================================================================

SCRIPT_DIR = Path(__file__).resolve().parent
REPO_DIR = SCRIPT_DIR.parent

EXCEL_FILE = REPO_DIR / "Funds Links.xlsm"

OUTPUT_DIR = REPO_DIR / "output_holdings"
FUNDS_OUTPUT_DIR = OUTPUT_DIR / "funds"

MAX_HOLDINGS = 10

REQUEST_TIMEOUT = 60

USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) "
    "Chrome/153.0.0.0 Safari/537.36"
)

ALLOWED_HOSTS = {
    "www.prudential.com.sg",
    "prudential.com.sg",
}

GENERIC_REPORT_PATH_MARKERS = (
    "/about-us/about-us/financial-listings/fund-reports/",
    "/financial-listings/fund-reports/",
    "/fund-reports/",
)

PDF_EXTENSIONS = (
    ".pdf",
)


# ============================================================================
# GENERAL HELPERS
# ============================================================================

def utc_now_iso() -> str:
    return datetime.now(timezone.utc).isoformat()


def clean_text(value: Any) -> str:
    if value is None:
        return ""

    text = str(value)

    text = text.replace("\xa0", " ")
    text = text.replace("\u200b", "")
    text = text.replace("\u2013", "-")
    text = text.replace("\u2014", "-")
    text = text.replace("\u2212", "-")

    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\n[ \t]+", "\n", text)

    return text.strip()


def normalize_space(value: str) -> str:
    return re.sub(r"\s+", " ", clean_text(value)).strip()


def safe_filename(value: str) -> str:
    value = clean_text(value)

    value = re.sub(r'[<>:"/\\|?*]', "_", value)
    value = re.sub(r"\s+", " ", value)
    value = value.strip(" .")

    if not value:
        value = "unknown_fund"

    return value[:180]


def write_json(path: Path, data: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)

    with path.open("w", encoding="utf-8") as handle:
        json.dump(
            data,
            handle,
            ensure_ascii=False,
            indent=2,
        )


def write_text(path: Path, text: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)

    with path.open("w", encoding="utf-8") as handle:
        handle.write(text)


# ============================================================================
# EXCEL
# ============================================================================

def read_excel_funds() -> list[dict[str, Any]]:
    if not EXCEL_FILE.exists():
        raise FileNotFoundError(
            f"Excel file not found: {EXCEL_FILE}"
        )

    workbook = openpyxl.load_workbook(
        EXCEL_FILE,
        read_only=True,
        data_only=True,
    )

    worksheet = workbook.active

    funds: list[dict[str, Any]] = []

    for row_number in range(2, worksheet.max_row + 1):
        url_value = worksheet.cell(row=row_number, column=1).value
        name_value = worksheet.cell(row=row_number, column=2).value

        url = clean_text(url_value)
        pruaccess_name = clean_text(name_value)

        if not url:
            continue

        funds.append(
            {
                "excelRow": row_number,
                "url": url,
                "pruAccessFundName": pruaccess_name,
            }
        )

    workbook.close()

    return funds


# ============================================================================
# HTTP SESSION
# ============================================================================

def create_session() -> requests.Session:
    session = requests.Session()

    session.headers.update(
        {
            "User-Agent": USER_AGENT,
            "Accept": (
                "text/html,application/xhtml+xml,application/xml;"
                "q=0.9,image/avif,image/webp,*/*;q=0.8"
            ),
            "Accept-Language": "en-SG,en;q=0.9",
            "Connection": "keep-alive",
        }
    )

    return session


# ============================================================================
# URL VALIDATION
# ============================================================================

def is_official_prudential_url(url: str) -> bool:
    try:
        parsed = urlparse(url)

        if parsed.scheme.lower() not in ("http", "https"):
            return False

        hostname = (parsed.hostname or "").lower()

        return hostname in ALLOWED_HOSTS

    except Exception:
        return False


def is_generic_fund_reports_url(url: str) -> bool:
    lowered = url.lower()

    for marker in GENERIC_REPORT_PATH_MARKERS:
        if marker in lowered:
            return True

    return False


def is_pdf_url(url: str) -> bool:
    try:
        parsed = urlparse(url)
        path = parsed.path.lower()

        return path.endswith(PDF_EXTENSIONS)

    except Exception:
        return False


# ============================================================================
# FACTSHEET URL DISCOVERY
# ============================================================================

def score_factsheet_candidate(
    url: str,
    anchor_text: str = "",
    surrounding_text: str = "",
) -> int:
    """
    Score an official Prudential URL as a possible fund factsheet.

    Higher is better.

    The most important rule is that the URL must NOT be the generic
    Fund Reports page.
    """

    if not url:
        return -10_000

    if not is_official_prudential_url(url):
        return -10_000

    if is_generic_fund_reports_url(url):
        return -10_000

    score = 0

    lowered_url = url.lower()
    lowered_anchor = normalize_space(anchor_text).lower()
    lowered_surrounding = normalize_space(surrounding_text).lower()

    # Actual PDF is strongly preferred.
    if is_pdf_url(url):
        score += 100

    # Explicit factsheet wording is highly important.
    if "factsheet" in lowered_anchor:
        score += 100

    if "fund factsheet" in lowered_anchor:
        score += 30

    if "factsheet" in lowered_url:
        score += 80

    if "factsheet" in lowered_surrounding:
        score += 20

    # Common official Prudential factsheet storage paths.
    if "/factsheets/" in lowered_url:
        score += 50

    if "/fact-sheets/" in lowered_url:
        score += 50

    if "/ebrochures/" in lowered_url:
        score += 20

    if "/media/" in lowered_url:
        score += 10

    # Penalise obvious non-factsheet document types.
    if "fund-information-booklet" in lowered_url:
        score -= 60

    if "information-booklet" in lowered_url:
        score -= 60

    if "product-highlights" in lowered_url:
        score -= 60

    if "productsummary" in lowered_url:
        score -= 60

    if "product-summary" in lowered_url:
        score -= 60

    if "fund-report" in lowered_url:
        score -= 100

    if "report" in lowered_url and "factsheet" not in lowered_url:
        score -= 20

    return score


def collect_factsheet_candidates(
    page_url: str,
    html: str,
) -> list[dict[str, Any]]:
    """
    Find possible official Prudential factsheet URLs.

    Multiple discovery strategies are deliberately used because Prudential
    has changed page structures over time.
    """

    soup = BeautifulSoup(html, "html.parser")

    candidates: dict[str, dict[str, Any]] = {}

    def add_candidate(
        url: str,
        anchor_text: str = "",
        surrounding_text: str = "",
        source: str = "",
    ) -> None:
        if not url:
            return

        absolute_url = urljoin(page_url, url)

        if not is_official_prudential_url(absolute_url):
            return

        if is_generic_fund_reports_url(absolute_url):
            return

        score = score_factsheet_candidate(
            absolute_url,
            anchor_text,
            surrounding_text,
        )

        if score <= -1000:
            return

        key = absolute_url.split("#", 1)[0]

        existing = candidates.get(key)

        candidate = {
            "url": key,
            "score": score,
            "anchorText": clean_text(anchor_text),
            "surroundingText": normalize_space(surrounding_text),
            "source": source,
        }

        if existing is None or score > existing["score"]:
            candidates[key] = candidate

    # ------------------------------------------------------------------
    # Strategy 1:
    # Inspect every anchor and prioritise "View factsheet".
    # ------------------------------------------------------------------

    for anchor in soup.find_all("a"):
        href = anchor.get("href")

        if not href:
            continue

        anchor_text = normalize_space(anchor.get_text(" ", strip=True))

        parent = anchor.parent

        surrounding_text = ""

        if parent is not None:
            surrounding_text = normalize_space(
                parent.get_text(" ", strip=True)
            )

        add_candidate(
            href,
            anchor_text,
            surrounding_text,
            "anchor",
        )

    # ------------------------------------------------------------------
    # Strategy 2:
    # Search raw HTML for official PDF URLs.
    #
    # This handles cases where the PDF URL is embedded in JavaScript,
    # data attributes, JSON, or other page metadata instead of a normal
    # <a href=""> element.
    # ------------------------------------------------------------------

    pdf_patterns = [
        r'https?://www\.prudential\.com\.sg/[^"\']+?\.pdf(?:\?[^"\']*)?',
        r'https?://prudential\.com\.sg/[^"\']+?\.pdf(?:\?[^"\']*)?',
        r'["\']([^"\']+?\.pdf(?:\?[^"\']*)?)["\']',
    ]

    for pattern in pdf_patterns:
        for match in re.finditer(
            pattern,
            html,
            flags=re.IGNORECASE,
        ):
            value = match.group(1) if match.lastindex else match.group(0)

            value = (
                value
                .replace("\\/", "/")
                .replace("\\u0026", "&")
            )

            add_candidate(
                value,
                "embedded PDF URL",
                "",
                "raw-html",
            )

    # ------------------------------------------------------------------
    # Strategy 3:
    # Search specifically around the words "View factsheet" and
    # "Fund Factsheet".
    # ------------------------------------------------------------------

    lowered_html = html.lower()

    for keyword in (
        "view factsheet",
        "fund factsheet",
        "factsheet",
    ):
        start = 0

        while True:
            position = lowered_html.find(keyword, start)

            if position < 0:
                break

            window_start = max(0, position - 5000)
            window_end = min(
                len(html),
                position + 5000,
            )

            window = html[window_start:window_end]

            for match in re.finditer(
                r'(?:href|src|data-href|data-url|url)\s*=\s*["\']([^"\']+)["\']',
                window,
                flags=re.IGNORECASE,
            ):
                add_candidate(
                    match.group(1),
                    keyword,
                    keyword,
                    "factsheet-window",
                )

            start = position + len(keyword)

    # ------------------------------------------------------------------
    # Sort strongest candidate first.
    # ------------------------------------------------------------------

    result = list(candidates.values())

    result.sort(
        key=lambda item: (
            item["score"],
            len(item["url"]),
        ),
        reverse=True,
    )

    return result


def find_factsheet_url(
    session: requests.Session,
    fund_url: str,
) -> dict[str, Any]:
    """
    Fetch the Prudential fund page and locate its official factsheet.

    IMPORTANT:
    A generic Prudential Fund Reports URL is never accepted.

    The returned URL is only a candidate until download_url() verifies
    that the final response is actually a PDF.
    """

    if not is_official_prudential_url(fund_url):
        raise RuntimeError(
            f"Fund URL is not an official Prudential Singapore URL: {fund_url}"
        )

    response = session.get(
        fund_url,
        timeout=REQUEST_TIMEOUT,
        allow_redirects=True,
    )

    response.raise_for_status()

    final_page_url = response.url

    if not is_official_prudential_url(final_page_url):
        raise RuntimeError(
            "Prudential fund page redirected outside the official "
            f"Prudential domain: {final_page_url}"
        )

    content_type = (
        response.headers.get("Content-Type") or ""
    ).lower()

    if "html" not in content_type and "text" not in content_type:
        raise RuntimeError(
            "Expected Prudential fund page HTML but received "
            f"content-type='{content_type}' from {final_page_url}"
        )

    html = response.text

    candidates = collect_factsheet_candidates(
        final_page_url,
        html,
    )

    if not candidates:
        raise RuntimeError(
            "Could not locate an official Prudential fund factsheet "
            f"from {final_page_url}"
        )

    # Only return the strongest candidate here.
    # download_url() performs the definitive PDF verification.
    best = candidates[0]

    if is_generic_fund_reports_url(best["url"]):
        raise RuntimeError(
            "Factsheet discovery selected a generic Fund Reports URL, "
            "which is not allowed: "
            f"{best['url']}"
        )

    return {
        "fundPageUrl": fund_url,
        "finalFundPageUrl": final_page_url,
        "factsheetUrl": best["url"],
        "factsheetCandidateScore": best["score"],
        "factsheetCandidateSource": best["source"],
        "factsheetAnchorText": best["anchorText"],
        "candidateCount": len(candidates),
        "candidates": candidates[:20],
    }


# ============================================================================
# DOWNLOAD PDF
# ============================================================================

def download_url(
    session: requests.Session,
    url: str,
    expected_pdf: bool = False,
) -> tuple[bytes, requests.Response]:
    """
    Download a URL.

    If expected_pdf=True, the final response MUST be a PDF.

    Content-Type alone is not trusted because some Prudential endpoints
    may omit or misreport it. The PDF magic header is also checked.
    """

    if not is_official_prudential_url(url):
        raise RuntimeError(
            f"Refusing to download non-Prudential URL: {url}"
        )

    if is_generic_fund_reports_url(url):
        raise RuntimeError(
            f"Refusing to download generic Fund Reports URL: {url}"
        )

    response = session.get(
        url,
        timeout=REQUEST_TIMEOUT,
        allow_redirects=True,
    )

    response.raise_for_status()

    final_url = response.url

    if not is_official_prudential_url(final_url):
        raise RuntimeError(
            "Download redirected outside official Prudential domain: "
            f"{final_url}"
        )

    content_type = (
        response.headers.get("Content-Type") or ""
    ).lower()

    content = response.content

    is_pdf_magic = content.startswith(b"%PDF")

    if expected_pdf:
        if not is_pdf_magic:
            raise RuntimeError(
                "Expected PDF but received "
                f"content-type='{content_type}' from {final_url}"
            )

        if is_generic_fund_reports_url(final_url):
            raise RuntimeError(
                "Expected factsheet PDF but final URL is the generic "
                f"Fund Reports page: {final_url}"
            )

    return content, response


def download_factsheet(
    session: requests.Session,
    discovery: dict[str, Any],
) -> dict[str, Any]:
    """
    Download the discovered factsheet.

    If the strongest candidate fails PDF validation, try other official
    candidates rather than accepting an HTML page.
    """

    candidates = discovery.get("candidates") or []

    failures: list[str] = []

    for candidate in candidates:
        url = candidate["url"]

        if is_generic_fund_reports_url(url):
            continue

        try:
            pdf_bytes, response = download_url(
                session,
                url,
                expected_pdf=True,
            )

            return {
                "url": url,
                "finalUrl": response.url,
                "contentType": (
                    response.headers.get("Content-Type") or ""
                ),
                "bytes": len(pdf_bytes),
                "content": pdf_bytes,
                "candidateScore": candidate.get("score"),
                "candidateSource": candidate.get("source"),
                "anchorText": candidate.get("anchorText"),
            }

        except Exception as exc:
            failures.append(
                f"{url}: {type(exc).__name__}: {exc}"
            )

    details = "\n".join(failures[:10])

    raise RuntimeError(
        "Could not download any candidate as an official Prudential "
        f"factsheet PDF.\n{details}"
    )


# ============================================================================
# FUND PAGE METADATA
# ============================================================================

def extract_fund_page_metadata(
    session: requests.Session,
    fund_url: str,
) -> dict[str, Any]:
    response = session.get(
        fund_url,
        timeout=REQUEST_TIMEOUT,
        allow_redirects=True,
    )

    response.raise_for_status()

    soup = BeautifulSoup(
        response.text,
        "html.parser",
    )

    page_text = normalize_space(
        soup.get_text(" ", strip=True)
    )

    metadata: dict[str, Any] = {
        "requestedUrl": fund_url,
        "finalUrl": response.url,
    }

    title = soup.title.get_text(
        " ",
        strip=True,
    ) if soup.title else ""

    metadata["pageTitle"] = clean_text(title)

    # Fund code
    fund_code_match = re.search(
        r"\bFund code\b\s*([A-Z0-9]+)",
        page_text,
        flags=re.IGNORECASE,
    )

    if fund_code_match:
        metadata["fundCode"] = clean_text(
            fund_code_match.group(1)
        )
    else:
        metadata["fundCode"] = None

    # Inception date
    inception_match = re.search(
        r"Inception date\s+"
        r"(\d{1,2}\s+[A-Za-z]+\s+\d{4})",
        page_text,
        flags=re.IGNORECASE,
    )

    if inception_match:
        metadata["inceptionDate"] = clean_text(
            inception_match.group(1)
        )
    else:
        metadata["inceptionDate"] = None

    # Currency
    currency_match = re.search(
        r"\bCurrency\b\s+([A-Z]{3})",
        page_text,
        flags=re.IGNORECASE,
    )

    if currency_match:
        metadata["currency"] = clean_text(
            currency_match.group(1)
        )
    else:
        metadata["currency"] = None

    # Risk classification
    risk_match = re.search(
        r"Risk classification\s+(.+?)"
        r"\s+Currency\b",
        page_text,
        flags=re.IGNORECASE,
    )

    if risk_match:
        metadata["riskClassification"] = clean_text(
            risk_match.group(1)
        )
    else:
        metadata["riskClassification"] = None

    # Asset class
    asset_match = re.search(
        r"Asset class type.*?"
        r"(Money Market|Fixed Income|Equity|Multi-Asset)",
        page_text,
        flags=re.IGNORECASE,
    )

    if asset_match:
        metadata["assetClass"] = clean_text(
            asset_match.group(1)
        )
    else:
        metadata["assetClass"] = None

    return metadata


# ============================================================================
# PDF TEXT EXTRACTION
# ============================================================================

def extract_pdf_text(pdf_bytes: bytes) -> dict[str, str]:
    """
    Extract PDF text using both default and layout modes.

    Both versions are retained because Prudential factsheets can have
    different text positioning depending on the PDF.
    """

    reader = PdfReader(BytesIO(pdf_bytes))

    default_parts: list[str] = []
    layout_parts: list[str] = []

    for page in reader.pages:
        try:
            text = page.extract_text() or ""
        except Exception:
            text = ""

        default_parts.append(text)

        try:
            layout_text = page.extract_text(
                extraction_mode="layout"
            ) or ""
        except Exception:
            layout_text = ""

        layout_parts.append(layout_text)

    return {
        "default": "\n".join(default_parts),
        "layout": "\n".join(layout_parts),
    }


# ============================================================================
# TOP HOLDINGS SECTION
# ============================================================================

def extract_top_holdings_section(
    text: str,
) -> Optional[str]:
    """
    Locate the official "Top 10 holdings" section.

    The section ends at the next major factsheet section.

    Returns:
        None
            if no Top 10 holdings heading exists.

        string
            section text if the heading exists.
    """

    if not text:
        return None

    normalized = text.replace("\r\n", "\n")
    normalized = normalized.replace("\r", "\n")

    heading_match = re.search(
        r"\bTop\s*10\s+holdings\b",
        normalized,
        flags=re.IGNORECASE,
    )

    if not heading_match:
        return None

    start = heading_match.end()

    remainder = normalized[start:]

    # Possible next section headings in Prudential factsheets.
    stop_patterns = [
        r"\n\s*Top\s+10\s+countries\b",
        r"\n\s*Top\s+10\s+sectors\b",
        r"\n\s*Asset\s+allocation\b",
        r"\n\s*Geographical\s+allocation\b",
        r"\n\s*Sector\s+allocation\b",
        r"\n\s*Investment\s+manager",
        r"\n\s*Underlying\s+fund",
        r"\n\s*Performance\b",
        r"\n\s*Fund\s+details\b",
        r"\n\s*Important\s+information\b",
        r"\n\s*Disclaimer\b",
        r"\n\s*Page\s+\d+\s+of\s+\d+",
    ]

    stop_positions: list[int] = []

    for pattern in stop_patterns:
        match = re.search(
            pattern,
            remainder,
            flags=re.IGNORECASE,
        )

        if match:
            stop_positions.append(match.start())

    if stop_positions:
        end = min(stop_positions)
        section = remainder[:end]
    else:
        section = remainder

    section = section.strip()

    return section


# ============================================================================
# HOLDING PARSING
# ============================================================================

def extract_percentage_from_line(
    line: str,
) -> tuple[Optional[float], Optional[str]]:
    """
    Extract a published percentage from a holding line.

    No percentage is calculated or inferred.

    Examples:
        "ABC FUND 12.3%" -> (12.3, "12.3%")
        "ABC FUND"        -> (None, None)
    """

    match = re.search(
        r"(?<![\d.])"
        r"(\d+(?:\.\d+)?)"
        r"\s*%"
        r"(?!\w)",
        line,
    )

    if not match:
        return None, None

    raw = match.group(0)

    try:
        value = float(match.group(1))
    except ValueError:
        return None, None

    return value, clean_text(raw)


def clean_holding_name(name: str) -> str:
    name = clean_text(name)

    name = re.sub(
        r"\s+",
        " ",
        name,
    )

    # Remove percentage at the end only when it is clearly a published
    # percentage belonging to the same holding.
    name = re.sub(
        r"\s+\d+(?:\.\d+)?\s*%\s*$",
        "",
        name,
    )

    # Remove leading/trailing punctuation introduced by PDF extraction.
    name = name.strip(" -–—|:;,")

    return name.strip()


def parse_rank_line(
    line: str,
) -> Optional[tuple[int, str]]:
    """
    Detect a holding rank at the beginning of a line.

    Supported forms include:
        1 ABC FUND
        1. ABC FUND
        1) ABC FUND
        1 - ABC FUND
        1 ABC FUND 12.3%

    A rank is the primary holding boundary.
    """

    match = re.match(
        r"^\s*"
        r"(\d{1,2})"
        r"(?:[.)]|[-–—:])?"
        r"\s+"
        r"(.+?)"
        r"\s*$",
        line,
    )

    if not match:
        return None

    try:
        rank = int(match.group(1))
    except ValueError:
        return None

    if rank < 1 or rank > MAX_HOLDINGS:
        return None

    remainder = clean_text(match.group(2))

    if not remainder:
        return None

    return rank, remainder


def parse_holdings(
    section_text: str,
) -> list[dict[str, Any]]:
    """
    Parse ranked holdings.

    IMPORTANT:
    Rank defines the holding boundary.

    Percentage is optional.

    Therefore:

        1 FUND A
        2 FUND B 10.5%

    produces:

        FUND A -> null percentage
        FUND B -> 10.5%

    No percentage is inferred for FUND A.
    """

    if not section_text:
        return []

    raw_lines = section_text.replace(
        "\r",
        "\n",
    ).split("\n")

    lines: list[str] = []

    for raw_line in raw_lines:
        line = clean_text(raw_line)

        if not line:
            continue

        lines.append(line)

    holdings: list[dict[str, Any]] = []

    current: Optional[dict[str, Any]] = None

    expected_rank = 1

    for line in lines:
        parsed_rank = parse_rank_line(line)

        if parsed_rank is not None:
            rank, remainder = parsed_rank

            # A new rank closes the previous holding.
            if current is not None:
                current_name = clean_holding_name(
                    current["rawName"]
                )

                if not current_name:
                    raise RuntimeError(
                        "A holding rank was found but the holding name "
                        "could not be identified."
                    )

                current["name"] = current_name

                holdings.append(current)

            if rank != expected_rank:
                raise RuntimeError(
                    "Holding ranks are not sequential. "
                    f"Expected rank {expected_rank}, found rank {rank}."
                )

            percentage, percentage_text = (
                extract_percentage_from_line(remainder)
            )

            name_part = re.sub(
                r"\s+\d+(?:\.\d+)?\s*%\s*$",
                "",
                remainder,
            )

            name_part = clean_holding_name(name_part)

            if not name_part:
                raise RuntimeError(
                    f"Rank {rank} has no identifiable holding name."
                )

            current = {
                "rank": rank,
                "rawName": name_part,
                "name": name_part,
                "weightPercent": percentage,
                "weightText": percentage_text,
            }

            expected_rank += 1

            if rank == MAX_HOLDINGS:
                # Do not parse beyond the tenth published holding.
                continue

            continue

        # Non-rank line.
        #
        # It belongs to the current holding only. This preserves
        # multi-line holding names.
        if current is None:
            # Ignore introductory/header text inside the section.
            continue

        # If this is a line containing a percentage and the current
        # holding has no percentage yet, accept the published value.
        #
        # We do NOT calculate anything.
        percentage, percentage_text = (
            extract_percentage_from_line(line)
        )

        if (
            percentage is not None
            and current.get("weightPercent") is None
        ):
            current["weightPercent"] = percentage
            current["weightText"] = percentage_text

            name_part = re.sub(
                r"\s+\d+(?:\.\d+)?\s*%\s*$",
                "",
                line,
            )

            name_part = clean_holding_name(name_part)

            if name_part:
                current["rawName"] = (
                    current["rawName"] + " " + name_part
                )

        else:
            current["rawName"] = (
                current["rawName"] + " " + line
            )

    # Final holding.
    if current is not None:
        current_name = clean_holding_name(
            current["rawName"]
        )

        if not current_name:
            raise RuntimeError(
                "The final holding has no identifiable name."
            )

        current["name"] = current_name

        holdings.append(current)

    # Hard limit.
    holdings = holdings[:MAX_HOLDINGS]

    # Remove internal parser-only field.
    for holding in holdings:
        holding.pop("rawName", None)

    return holdings


# ============================================================================
# HOLDING VALIDATION
# ============================================================================

def validate_holdings(
    holdings: list[dict[str, Any]],
) -> None:
    """
    Apply all holding hard rules.
    """

    if not holdings:
        raise RuntimeError(
            "Top 10 holdings section exists but no identifiable holdings "
            "were extracted."
        )

    if len(holdings) > MAX_HOLDINGS:
        raise RuntimeError(
            f"Extracted {len(holdings)} holdings, exceeding maximum "
            f"of {MAX_HOLDINGS}."
        )

    expected_rank = 1

    seen_names: set[str] = set()

    for holding in holdings:
        rank = holding.get("rank")
        name = clean_text(holding.get("name"))

        if rank != expected_rank:
            raise RuntimeError(
                "Holding rank sequence is invalid. "
                f"Expected {expected_rank}, found {rank}."
            )

        if not name:
            raise RuntimeError(
                f"Holding rank {rank} has no name."
            )

        normalized_name = re.sub(
            r"\s+",
            " ",
            name.lower(),
        ).strip()

        if normalized_name in seen_names:
            raise RuntimeError(
                "Duplicate holding name detected: "
                f"{name}"
            )

        seen_names.add(normalized_name)

        weight = holding.get("weightPercent")

        if weight is not None:
            if not isinstance(weight, (int, float)):
                raise RuntimeError(
                    f"Invalid published percentage for {name}: "
                    f"{weight}"
                )

            if weight < 0 or weight > 100:
                raise RuntimeError(
                    f"Published percentage outside 0-100 range for "
                    f"{name}: {weight}"
                )

        expected_rank += 1


# ============================================================================
# PROCESS ONE FUND
# ============================================================================

def process_fund(
    session: requests.Session,
    fund: dict[str, Any],
) -> dict[str, Any]:
    excel_row = fund["excelRow"]
    fund_url = fund["url"]
    pruaccess_name = fund["pruAccessFundName"]

    started = time.time()

    fund_metadata = extract_fund_page_metadata(
        session,
        fund_url,
    )

    discovery = find_factsheet_url(
        session,
        fund_url,
    )

    factsheet = download_factsheet(
        session,
        discovery,
    )

    pdf_bytes = factsheet["content"]

    pdf_texts = extract_pdf_text(
        pdf_bytes,
    )

    # Prefer default extraction, but also test layout extraction.
    candidates: list[dict[str, Any]] = []

    for extraction_mode in (
        "default",
        "layout",
    ):
        text = pdf_texts.get(extraction_mode) or ""

        section = extract_top_holdings_section(
            text,
        )

        if section is None:
            continue

        try:
            holdings = parse_holdings(
                section,
            )

            validate_holdings(
                holdings,
            )

            candidates.append(
                {
                    "mode": extraction_mode,
                    "section": section,
                    "holdings": holdings,
                }
            )

        except Exception:
            continue

    if not candidates:
        # Distinguish between:
        #   no holdings section
        # and
        #   holdings section exists but parsing failed.
        default_section = extract_top_holdings_section(
            pdf_texts.get("default", ""),
        )

        layout_section = extract_top_holdings_section(
            pdf_texts.get("layout", ""),
        )

        if (
            default_section is None
            and layout_section is None
        ):
            return {
                "status": "no_holdings_section",
                "excelRow": excel_row,
                "url": fund_url,
                "pruAccessFundName": pruaccess_name,
                "fundPage": fund_metadata,
                "factsheet": {
                    "url": factsheet["url"],
                    "finalUrl": factsheet["finalUrl"],
                    "contentType": factsheet["contentType"],
                    "bytes": factsheet["bytes"],
                },
                "holdings": [],
                "elapsedSeconds": round(
                    time.time() - started,
                    3,
                ),
            }

        raise RuntimeError(
            "Top 10 holdings section was found, but the holdings "
            "could not be parsed and validated."
        )

    # Select the candidate with the greatest number of holdings.
    # If tied, prefer default extraction.
    candidates.sort(
        key=lambda candidate: (
            len(candidate["holdings"]),
            1 if candidate["mode"] == "default" else 0,
        ),
        reverse=True,
    )

    selected = candidates[0]

    holdings = selected["holdings"]

    fund_slug = safe_filename(
        pruaccess_name
        or fund_metadata.get("pageTitle")
        or f"excel_row_{excel_row}"
    )

    fund_output_dir = FUNDS_OUTPUT_DIR / fund_slug
    fund_output_dir.mkdir(
        parents=True,
        exist_ok=True,
    )

    pdf_path = fund_output_dir / "factsheet.pdf"
    pdf_path.write_bytes(pdf_bytes)

    write_text(
        fund_output_dir / "factsheet-default.txt",
        pdf_texts.get("default", ""),
    )

    write_text(
        fund_output_dir / "factsheet-layout.txt",
        pdf_texts.get("layout", ""),
    )

    write_text(
        fund_output_dir / "top-holdings-section.txt",
        selected["section"],
    )

    result = {
        "status": "success",
        "excelRow": excel_row,
        "url": fund_url,
        "pruAccessFundName": pruaccess_name,
        "fundPage": fund_metadata,
        "factsheet": {
            "url": factsheet["url"],
            "finalUrl": factsheet["finalUrl"],
            "contentType": factsheet["contentType"],
            "bytes": factsheet["bytes"],
            "candidateScore": factsheet["candidateScore"],
            "candidateSource": factsheet["candidateSource"],
            "anchorText": factsheet["anchorText"],
            "parserMode": selected["mode"],
        },
        "holdings": holdings,
        "publishedHoldingCount": len(holdings),
        "publishedPercentageCount": sum(
            1
            for holding in holdings
            if holding.get("weightPercent") is not None
        ),
        "missingPublishedPercentageCount": sum(
            1
            for holding in holdings
            if holding.get("weightPercent") is None
        ),
        "elapsedSeconds": round(
            time.time() - started,
            3,
        ),
    }

    write_json(
        fund_output_dir / "holdings.json",
        result,
    )

    write_json(
        fund_output_dir / "factsheet-metadata.json",
        {
            "fund": fund,
            "fundPage": fund_metadata,
            "factsheet": {
                "url": factsheet["url"],
                "finalUrl": factsheet["finalUrl"],
                "contentType": factsheet["contentType"],
                "bytes": factsheet["bytes"],
                "candidateScore": factsheet["candidateScore"],
                "candidateSource": factsheet["candidateSource"],
                "anchorText": factsheet["anchorText"],
            },
        },
    )

    return result


# ============================================================================
# RUN SUMMARY
# ============================================================================

def build_run_summary(
    started_at: str,
    completed_at: str,
    funds: list[dict[str, Any]],
    results: list[dict[str, Any]],
) -> dict[str, Any]:
    successful = [
        result
        for result in results
        if result.get("status") == "success"
    ]

    no_holdings_section = [
        result
        for result in results
        if result.get("status") == "no_holdings_section"
    ]

    failed = [
        result
        for result in results
        if result.get("status") == "failed"
    ]

    total_holdings = sum(
        len(result.get("holdings") or [])
        for result in successful
    )

    total_published_percentages = sum(
        result.get("publishedPercentageCount", 0)
        for result in successful
    )

    total_missing_percentages = sum(
        result.get("missingPublishedPercentageCount", 0)
        for result in successful
    )

    return {
        "status": (
            "success"
            if not failed
            else "failed"
        ),
        "startedAtUtc": started_at,
        "completedAtUtc": completed_at,
        "excelFile": EXCEL_FILE.name,
        "fundUniverseCount": len(funds),
        "successfulFundCount": len(successful),
        "noHoldingsSectionFundCount": len(
            no_holdings_section
        ),
        "failedFundCount": len(failed),
        "totalHoldings": total_holdings,
        "totalPublishedPercentages": total_published_percentages,
        "totalMissingPublishedPercentages": total_missing_percentages,
        "failedFunds": failed,
        "noHoldingsSectionFunds": no_holdings_section,
        "rules": {
            "excelColumnAControlsUniverse": True,
            "officialPrudentialSourcesOnly": True,
            "thirdPartySourcesAllowed": False,
            "maximumHoldings": MAX_HOLDINGS,
            "publishedHoldingCountIsAuthoritative": True,
            "holdingRankDefinesRowBoundary": True,
            "publishedPercentageOptional": True,
            "missingPublishedPercentageStoredAsNull": True,
            "percentageCalculationAllowed": False,
            "percentageInferenceAllowed": False,
            "holdingNameInferenceAllowed": False,
            "duplicatePercentagesAllowed": True,
            "duplicateHoldingNamesAllowed": False,
            "syntheticDataAllowed": False,
            "estimatedDataAllowed": False,
            "interpolationAllowed": False,
            "forcedTenHoldings": False,
            "genericFundReportsPageRejected": True,
            "finalFactsheetMustBePdf": True,
        },
    }


# ============================================================================
# MAIN
# ============================================================================

def main() -> int:
    started_at = utc_now_iso()

    OUTPUT_DIR.mkdir(
        parents=True,
        exist_ok=True,
    )

    FUNDS_OUTPUT_DIR.mkdir(
        parents=True,
        exist_ok=True,
    )

    print("=" * 80)
    print("PRUDENTIAL OFFICIAL FUND FACTSHEET + HOLDINGS TEST")
    print("=" * 80)
    print(f"Excel file: {EXCEL_FILE}")
    print(f"Output directory: {OUTPUT_DIR}")
    print()

    try:
        funds = read_excel_funds()
    except Exception as exc:
        print(
            f"ERROR reading Excel: {type(exc).__name__}: {exc}",
            file=sys.stderr,
        )
        return 1

    print(
        f"Excel fund universe: {len(funds)}"
    )

    if not funds:
        print(
            "ERROR: No populated Prudential URLs were found in "
            "Excel Column A.",
            file=sys.stderr,
        )
        return 1

    session = create_session()

    results: list[dict[str, Any]] = []

    for index, fund in enumerate(
        funds,
        start=1,
    ):
        print()
        print("-" * 80)
        print(
            f"[{index}/{len(funds)}] "
            f"Excel row {fund['excelRow']}"
        )
        print(
            f"PruAccess name: "
            f"{fund['pruAccessFundName']}"
        )
        print(
            f"URL: {fund['url']}"
        )

        try:
            result = process_fund(
                session,
                fund,
            )

            results.append(result)

            if result["status"] == "success":
                print(
                    "SUCCESS: "
                    f"{len(result['holdings'])} holdings"
                )

                print(
                    "Published percentages: "
                    f"{result['publishedPercentageCount']}"
                )

                print(
                    "Missing published percentages: "
                    f"{result['missingPublishedPercentageCount']}"
                )

                print(
                    "Factsheet: "
                    f"{result['factsheet']['finalUrl']}"
                )

            elif result["status"] == "no_holdings_section":
                print(
                    "NO HOLDINGS SECTION"
                )

        except Exception as exc:
            failure = {
                "status": "failed",
                "excelRow": fund["excelRow"],
                "url": fund["url"],
                "pruAccessFundName": fund["pruAccessFundName"],
                "reason": (
                    f"{type(exc).__name__}: {exc}"
                ),
            }

            results.append(failure)

            print(
                "FAILED: "
                f"{failure['reason']}",
                file=sys.stderr,
            )

    completed_at = utc_now_iso()

    summary = build_run_summary(
        started_at,
        completed_at,
        funds,
        results,
    )

    write_json(
        OUTPUT_DIR / "run_summary.json",
        summary,
    )

    write_json(
        OUTPUT_DIR / "all_holdings.json",
        {
            "generatedAtUtc": completed_at,
            "fundUniverseCount": len(funds),
            "funds": results,
        },
    )

    print()
    print("=" * 80)
    print("RUN SUMMARY")
    print("=" * 80)
    print(
        f"Status: {summary['status']}"
    )
    print(
        f"Fund universe: {summary['fundUniverseCount']}"
    )
    print(
        f"Successful: {summary['successfulFundCount']}"
    )
    print(
        "No holdings section: "
        f"{summary['noHoldingsSectionFundCount']}"
    )
    print(
        f"Failed: {summary['failedFundCount']}"
    )
    print(
        f"Total holdings: {summary['totalHoldings']}"
    )
    print(
        "Published percentages: "
        f"{summary['totalPublishedPercentages']}"
    )
    print(
        "Missing published percentages: "
        f"{summary['totalMissingPublishedPercentages']}"
    )
    print(
        f"Summary: {OUTPUT_DIR / 'run_summary.json'}"
    )
    print(
        f"All holdings: {OUTPUT_DIR / 'all_holdings.json'}"
    )
    print("=" * 80)

    # A failed fund causes the workflow to fail.
    if summary["failedFundCount"] > 0:
        return 1

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
