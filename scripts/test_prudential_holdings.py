#!/usr/bin/env python3

"""
Prudential Singapore ILP fund Top 10 holdings extractor.

MASTER SOURCE
=============

Funds Links.xlsm

Excel Column A:
    Official Prudential Singapore fund URL.
    This column controls the fund universe.

Excel Column B:
    Exact PruAccess fund name.
    Stored as the fund's PruAccess reference name.

SOURCE RULES
============

1. Official Prudential Singapore sources only.
2. Third-party holdings sources are not allowed.
3. Factsheets must be official Prudential Singapore PDF documents.
4. The extractor must identify the actual Prudential factsheet link.
5. Generic page links, CSS, JS, images, unrelated Prudential pages,
   external domains, etc. are rejected.
6. "View factsheet" is the primary factsheet link.
7. "Fund Factsheet" document links are accepted as a secondary method.
8. The final downloaded document must actually be a PDF.
9. The PDF must contain a Top 10 Holdings section to extract holdings.
10. Prudential's published holding count is authoritative.
11. Maximum holdings = 10.
12. Fewer than 10 published holdings is valid.
13. Holding rank defines the holding row boundary.
14. Published percentage is optional.
15. Missing published percentage is stored as null.
16. Percentages are never calculated.
17. Percentages are never inferred.
18. Holding names are never inferred.
19. No synthetic data.
20. No estimated data.
21. No interpolation.
22. Duplicate holding names are rejected.
23. Duplicate percentages are allowed.
24. If the Top 10 Holdings section exists but cannot be parsed
    confidently, the fund fails.
25. If no Top 10 Holdings section exists, the fund is classified
    as noHoldingsSection.
26. The script does not modify PruAccess extraction.
27. The script does not modify data.json.
28. The script is a holdings collector/tester only.

IMPORTANT PARSING RULE
======================

Some Prudential factsheets publish:

    1. HOLDING NAME
    2. HOLDING NAME
    3. HOLDING NAME

without publishing a percentage beside every holding.

Therefore:

    Rank = holding boundary.

A new rank finalizes the previous holding.

A missing percentage is represented as:

    "weightPercent": null
    "weightText": null

No percentage is calculated or inferred.
"""

from __future__ import annotations

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
# CONFIGURATION
# ============================================================================

EXCEL_FILE = Path("Funds Links.xlsm")

OUTPUT_DIR = Path("output_holdings")
PDF_DIR = OUTPUT_DIR / "pdfs"
TEXT_DIR = OUTPUT_DIR / "text"
JSON_DIR = OUTPUT_DIR / "funds"
METADATA_DIR = OUTPUT_DIR / "metadata"

RUN_SUMMARY_FILE = OUTPUT_DIR / "run_summary.json"
ALL_HOLDINGS_FILE = OUTPUT_DIR / "all_holdings.json"

MAX_HOLDINGS = 10

REQUEST_TIMEOUT_SECONDS = 45
REQUEST_RETRIES = 3
REQUEST_RETRY_DELAY_SECONDS = 2

OFFICIAL_HOST = "www.prudential.com.sg"

USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) "
    "Chrome/153.0.0.0 Safari/537.36"
)


# ============================================================================
# BASIC HELPERS
# ============================================================================

def utc_now_iso() -> str:
    return datetime.now(timezone.utc).isoformat()


def clean_text(value: Any) -> str:
    if value is None:
        return ""

    text = str(value)

    text = text.replace("\xa0", " ")
    text = text.replace("\u200b", "")
    text = text.replace("\ufeff", "")

    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\n[ \t]+", "\n", text)

    return text.strip()


def safe_filename(value: str, maximum_length: int = 180) -> str:
    value = clean_text(value)

    value = re.sub(r"[^\w\s.-]", "_", value, flags=re.UNICODE)
    value = re.sub(r"\s+", "_", value)

    value = value.strip("._")

    if not value:
        value = "fund"

    return value[:maximum_length]


def normalise_space(text: str) -> str:
    text = text.replace("\xa0", " ")
    text = text.replace("\u200b", "")
    text = text.replace("\ufeff", "")
    text = re.sub(r"[ \t]+", " ", text)
    return text.strip()


def normalise_name(text: str) -> str:
    text = normalise_space(text)

    text = re.sub(r"\s+", " ", text)

    return text.strip(" -–—:;|")


def host_is_official(url: str) -> bool:
    try:
        parsed = urlparse(url)

        if parsed.scheme.lower() not in {"http", "https"}:
            return False

        host = (parsed.hostname or "").lower()

        return (
            host == OFFICIAL_HOST
            or host.endswith("." + OFFICIAL_HOST)
        )

    except Exception:
        return False


def is_probable_pdf_url(url: str) -> bool:
    parsed = urlparse(url)

    path = (parsed.path or "").lower()

    return (
        path.endswith(".pdf")
        or ".pdf/" in path
        or "/pdf/" in path
        or "/factsheet" in path
        or "/factsheets/" in path
    )


def looks_like_pdf_bytes(content: bytes) -> bool:
    if not content:
        return False

    return content[:5] == b"%PDF-"


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
        keep_links=True,
    )

    worksheet = workbook.active

    funds: list[dict[str, Any]] = []

    for excel_row, row in enumerate(
        worksheet.iter_rows(min_row=2, values_only=True),
        start=2,
    ):
        url = clean_text(row[0] if len(row) > 0 else "")
        pru_access_name = clean_text(row[1] if len(row) > 1 else "")

        if not url:
            continue

        funds.append(
            {
                "excelRow": excel_row,
                "url": url,
                "pruAccessFundName": pru_access_name,
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
                "text/html,application/xhtml+xml,"
                "application/xml;q=0.9,"
                "image/avif,image/webp,*/*;q=0.8"
            ),
            "Accept-Language": "en-SG,en;q=0.9",
            "Connection": "keep-alive",
        }
    )

    return session


def request_with_retries(
    session: requests.Session,
    url: str,
    *,
    method: str = "GET",
    **kwargs: Any,
) -> requests.Response:
    last_error: Optional[Exception] = None

    for attempt in range(1, REQUEST_RETRIES + 1):
        try:
            response = session.request(
                method,
                url,
                timeout=REQUEST_TIMEOUT_SECONDS,
                allow_redirects=True,
                **kwargs,
            )

            response.raise_for_status()

            return response

        except Exception as exc:
            last_error = exc

            if attempt < REQUEST_RETRIES:
                time.sleep(REQUEST_RETRY_DELAY_SECONDS)

    raise RuntimeError(
        f"Request failed after {REQUEST_RETRIES} attempts: "
        f"{url}: {last_error}"
    )


# ============================================================================
# FUND PAGE
# ============================================================================

def download_fund_page(
    session: requests.Session,
    fund_url: str,
) -> tuple[str, str]:
    if not host_is_official(fund_url):
        raise RuntimeError(
            f"Fund URL is not on official Prudential Singapore domain: "
            f"{fund_url}"
        )

    response = request_with_retries(
        session,
        fund_url,
    )

    final_url = response.url

    if not host_is_official(final_url):
        raise RuntimeError(
            f"Fund page redirected outside official Prudential domain: "
            f"{final_url}"
        )

    content_type = (
        response.headers.get("Content-Type", "")
        .split(";")[0]
        .strip()
        .lower()
    )

    if content_type not in {
        "",
        "text/html",
        "application/xhtml+xml",
    }:
        raise RuntimeError(
            f"Expected HTML fund page but received "
            f"content-type='{response.headers.get('Content-Type', '')}' "
            f"from {final_url}"
        )

    return response.text, final_url


# ============================================================================
# FACTSHEET DISCOVERY
# ============================================================================

def _anchor_text(anchor) -> str:
    return normalise_space(
        anchor.get_text(" ", strip=True)
    )


def _anchor_href(anchor) -> str:
    href = anchor.get("href")

    if not href:
        return ""

    return str(href).strip()


def _candidate_url_from_anchor(
    anchor,
    base_url: str,
) -> str:
    href = _anchor_href(anchor)

    if not href:
        return ""

    if href.startswith("#"):
        return ""

    absolute_url = urljoin(base_url, href)

    return absolute_url


def _is_obviously_non_document_asset(url: str) -> bool:
    parsed = urlparse(url)

    path = (parsed.path or "").lower()

    rejected_extensions = {
        ".png",
        ".jpg",
        ".jpeg",
        ".gif",
        ".svg",
        ".webp",
        ".ico",
        ".css",
        ".js",
        ".json",
        ".xml",
        ".woff",
        ".woff2",
        ".ttf",
        ".eot",
        ".mp4",
        ".webm",
        ".mov",
        ".zip",
        ".doc",
        ".docx",
        ".xls",
        ".xlsx",
        ".ppt",
        ".pptx",
    }

    for extension in rejected_extensions:
        if path.endswith(extension):
            return True

    return False


def _score_factsheet_anchor(
    anchor,
    candidate_url: str,
) -> int:
    """
    Score only links that have factsheet/document semantics.

    The important change from the failed run is that we do NOT treat
    every href on the page as a factsheet candidate.
    """

    if not candidate_url:
        return -1000

    if not host_is_official(candidate_url):
        return -1000

    if _is_obviously_non_document_asset(candidate_url):
        return -1000

    text = _anchor_text(anchor).lower()
    href = candidate_url.lower()

    score = 0

    # Highest confidence:
    # The current Prudential fund page has a dedicated "View factsheet"
    # link immediately beside the price section.
    if "view factsheet" in text:
        score += 1000

    # Fund Documents section.
    if "fund factsheet" in text:
        score += 900

    if text.strip() == "factsheet":
        score += 850

    if "fund fact sheet" in text:
        score += 850

    if "factsheet" in text:
        score += 700

    if "factsheet" in href:
        score += 600

    if "/factsheets/" in href:
        score += 600

    if "/factsheet" in href:
        score += 500

    if href.endswith(".pdf"):
        score += 250

    # Explicitly penalise known non-factsheet documents.
    bad_document_terms = (
        "product highlights",
        "product highlights sheet",
        "fund information booklet",
        "information booklet",
        "fund report",
        "annual report",
        "privacy",
        "terms",
        "policy",
        "application",
    )

    for term in bad_document_terms:
        if term in text:
            score -= 500

    for term in bad_document_terms:
        if term in href:
            score -= 300

    return score


def find_factsheet_candidates(
    html: str,
    page_url: str,
) -> list[str]:
    """
    Find official Prudential factsheet links.

    IMPORTANT:
    We deliberately do NOT collect every href.

    Candidate selection is based on:
      - View factsheet
      - Fund Factsheet
      - factsheet semantics
      - official Prudential domain
      - document/PDF semantics

    This prevents CSS/JS/image/HTML/unrelated links from entering
    the PDF downloader.
    """

    soup = BeautifulSoup(html, "html.parser")

    scored: list[tuple[int, int, str]] = []

    for index, anchor in enumerate(soup.find_all("a")):
        candidate_url = _candidate_url_from_anchor(
            anchor,
            page_url,
        )

        if not candidate_url:
            continue

        if not host_is_official(candidate_url):
            continue

        score = _score_factsheet_anchor(
            anchor,
            candidate_url,
        )

        if score <= 0:
            continue

        scored.append(
            (
                score,
                index,
                candidate_url,
            )
        )

    # ----------------------------------------------------------------------
    # Additional strict fallback:
    #
    # If the page exposes a PDF factsheet URL in HTML but the anchor text is
    # weak, only accept URLs that themselves clearly identify a factsheet.
    # ----------------------------------------------------------------------

    for match in re.finditer(
        r"""https?://[^"'<>\\\s]+""",
        html,
        flags=re.IGNORECASE,
    ):
        raw_url = match.group(0)

        raw_url = raw_url.rstrip(
            ".,);]}>'\""
        )

        if not host_is_official(raw_url):
            continue

        if not is_probable_pdf_url(raw_url):
            continue

        if _is_obviously_non_document_asset(raw_url):
            continue

        score = 400

        if "factsheet" in raw_url.lower():
            score += 300

        if "/factsheets/" in raw_url.lower():
            score += 200

        scored.append(
            (
                score,
                10_000 + match.start(),
                raw_url,
            )
        )

    # Highest score first.
    scored.sort(
        key=lambda item: (
            -item[0],
            item[1],
        )
    )

    candidates: list[str] = []
    seen: set[str] = set()

    for _, _, url in scored:
        normalized = url.strip()

        if normalized in seen:
            continue

        seen.add(normalized)

        candidates.append(normalized)

    return candidates


# ============================================================================
# PDF DOWNLOAD
# ============================================================================

def download_official_factsheet(
    session: requests.Session,
    candidate_url: str,
) -> tuple[str, bytes, str]:
    """
    Download and validate an official Prudential factsheet.

    Returns:
        final_url
        PDF bytes
        content_type
    """

    if not host_is_official(candidate_url):
        raise RuntimeError(
            f"Factsheet candidate is outside official Prudential domain: "
            f"{candidate_url}"
        )

    if _is_obviously_non_document_asset(candidate_url):
        raise RuntimeError(
            f"Rejected non-document asset: {candidate_url}"
        )

    response = request_with_retries(
        session,
        candidate_url,
    )

    final_url = response.url

    if not host_is_official(final_url):
        raise RuntimeError(
            f"Download redirected outside official Prudential domain: "
            f"{final_url}"
        )

    content_type_header = response.headers.get(
        "Content-Type",
        "",
    )

    content_type = (
        content_type_header
        .split(";")[0]
        .strip()
        .lower()
    )

    content = response.content

    if not looks_like_pdf_bytes(content):
        raise RuntimeError(
            f"Expected PDF but received "
            f"content-type='{content_type_header}' "
            f"from {final_url}"
        )

    if content_type not in {
        "",
        "application/pdf",
        "binary/octet-stream",
    }:
        raise RuntimeError(
            f"Expected PDF but received "
            f"content-type='{content_type_header}' "
            f"from {final_url}"
        )

    return (
        final_url,
        content,
        content_type_header,
    )


def find_and_download_factsheet(
    session: requests.Session,
    html: str,
    page_url: str,
) -> tuple[str, bytes, list[dict[str, Any]]]:
    candidates = find_factsheet_candidates(
        html,
        page_url,
    )

    if not candidates:
        raise RuntimeError(
            "No official Prudential factsheet candidate was found "
            "from the fund page."
        )

    errors: list[dict[str, Any]] = []

    for candidate in candidates:
        try:
            final_url, pdf_bytes, content_type = (
                download_official_factsheet(
                    session,
                    candidate,
                )
            )

            return (
                final_url,
                pdf_bytes,
                errors,
            )

        except Exception as exc:
            errors.append(
                {
                    "url": candidate,
                    "error": f"{type(exc).__name__}: {exc}",
                }
            )

    error_lines = []

    for item in errors:
        error_lines.append(
            f"{item['url']}: {item['error']}"
        )

    raise RuntimeError(
        "Could not download any candidate as an official "
        "Prudential factsheet PDF.\n"
        + "\n".join(error_lines)
    )


# ============================================================================
# FUND PAGE METADATA
# ============================================================================

def extract_fund_page_metadata(
    html: str,
    page_url: str,
) -> dict[str, Any]:
    soup = BeautifulSoup(html, "html.parser")

    metadata: dict[str, Any] = {
        "pageUrl": page_url,
    }

    title = soup.title.get_text(
        " ",
        strip=True,
    ) if soup.title else ""

    metadata["pageTitle"] = clean_text(title)

    # ----------------------------------------------------------------------
    # Try to identify the fund heading.
    # ----------------------------------------------------------------------

    fund_name = ""

    for heading in soup.find_all(
        ["h1", "h2", "h3"]
    ):
        text = clean_text(
            heading.get_text(" ", strip=True)
        )

        if not text:
            continue

        if "fund" in text.lower():
            fund_name = text
            break

    metadata["fundName"] = fund_name

    # ----------------------------------------------------------------------
    # Extract common fund fields from visible page text.
    #
    # These are metadata only. They are not used to infer holdings.
    # ----------------------------------------------------------------------

    full_text = soup.get_text(
        "\n",
        strip=True,
    )

    def extract_labeled_value(
        label: str,
        aliases: Optional[list[str]] = None,
    ) -> Optional[str]:
        labels = [label]

        if aliases:
            labels.extend(aliases)

        for candidate_label in labels:
            pattern = (
                rf"(?im)^\s*"
                rf"{re.escape(candidate_label)}"
                rf"\s*$"
                rf"\n\s*([^\n]+)"
            )

            match = re.search(
                pattern,
                full_text,
            )

            if match:
                value = clean_text(
                    match.group(1)
                )

                if value:
                    return value

        return None

    metadata["assetClass"] = extract_labeled_value(
        "Asset class type"
    )

    metadata["riskClassification"] = (
        extract_labeled_value(
            "Risk classification"
        )
    )

    metadata["currency"] = extract_labeled_value(
        "Currency"
    )

    metadata["inceptionDate"] = (
        extract_labeled_value(
            "Inception date"
        )
    )

    metadata["fundCode"] = extract_labeled_value(
        "Fund code"
    )

    return metadata


# ============================================================================
# PDF TEXT EXTRACTION
# ============================================================================

def extract_pdf_text(
    pdf_bytes: bytes,
) -> tuple[str, list[str]]:
    reader = PdfReader(
        BytesIO(pdf_bytes)
    )

    page_texts: list[str] = []

    for page in reader.pages:
        text = page.extract_text() or ""
        page_texts.append(text)

    combined = "\n".join(
        page_texts
    )

    return combined, page_texts


# ============================================================================
# TOP HOLDINGS SECTION
# ============================================================================

def extract_top_holdings_section(
    text: str,
) -> tuple[Optional[str], dict[str, Any]]:
    """
    Extract the section beginning at "Top 10 Holdings".

    The parser intentionally does not require percentages.

    Section termination is based on the next strong document heading,
    not on the appearance/disappearance of percentages.
    """

    normalized = text.replace(
        "\r\n",
        "\n",
    ).replace(
        "\r",
        "\n",
    )

    lines = [
        clean_text(line)
        for line in normalized.split("\n")
    ]

    lines = [
        line
        for line in lines
        if line
    ]

    start_index: Optional[int] = None

    for index, line in enumerate(lines):
        compact = re.sub(
            r"[^a-z0-9]+",
            " ",
            line.lower(),
        ).strip()

        if compact in {
            "top 10 holdings",
            "top 10 holding",
            "top holdings",
            "top 10 holdings 3",
            "top 10 holdings 2",
            "top 10 holdings 1",
        }:
            start_index = index
            break

        if re.match(
            r"^top\s+10\s+holdings(?:\b|[^a-z])",
            line,
            flags=re.IGNORECASE,
        ):
            start_index = index
            break

    if start_index is None:
        return (
            None,
            {
                "found": False,
                "startLine": None,
                "endLine": None,
            },
        )

    section_lines: list[str] = []

    # Strong section headings which commonly follow Top 10 Holdings.
    stop_patterns = [
        r"^country allocation\b",
        r"^sector allocation\b",
        r"^geographical allocation\b",
        r"^asset allocation\b",
        r"^portfolio allocation\b",
        r"^investment objective\b",
        r"^fund details\b",
        r"^performance\b",
        r"^calendar year performance\b",
        r"^important information\b",
        r"^important notes\b",
        r"^benchmark\b",
        r"^risk classification\b",
    ]

    for index in range(
        start_index + 1,
        len(lines),
    ):
        line = lines[index]

        should_stop = False

        for pattern in stop_patterns:
            if re.search(
                pattern,
                line,
                flags=re.IGNORECASE,
            ):
                should_stop = True
                break

        if should_stop:
            break

        section_lines.append(line)

    section = "\n".join(
        section_lines
    ).strip()

    return (
        section if section else None,
        {
            "found": True,
            "startLine": start_index + 1,
            "endLine": (
                start_index
                + len(section_lines)
            ),
        },
    )


# ============================================================================
# HOLDING PARSING
# ============================================================================

def extract_percentage_from_line(
    line: str,
) -> tuple[Optional[float], Optional[str]]:
    """
    Extract an explicitly published percentage.

    No percentage is calculated or inferred.
    """

    match = re.search(
        r"(?<!\d)"
        r"([+-]?\d+(?:\.\d+)?)"
        r"\s*%"
        r"(?!\d)",
        line,
    )

    if not match:
        return None, None

    raw = match.group(1)

    try:
        value = float(raw)
    except ValueError:
        return None, None

    return value, f"{raw}%"


def strip_percentage_from_line(
    line: str,
) -> str:
    return re.sub(
        r"\s*[+-]?\d+(?:\.\d+)?\s*%\s*",
        " ",
        line,
    ).strip()


def clean_holding_name(
    text: str,
) -> str:
    text = normalise_name(text)

    # Remove common footnote markers that can appear after a holding.
    text = re.sub(
        r"\s*[†‡*]+\s*$",
        "",
        text,
    )

    text = re.sub(
        r"\s+\d+\s*$",
        "",
        text,
    )

    return normalise_name(text)


def rank_from_line(
    line: str,
) -> Optional[int]:
    """
    Recognise a holding rank when it is explicitly present.

    Examples:
        1 NVIDIA CORPORATION
        1. NVIDIA CORPORATION
        1) NVIDIA CORPORATION
        1 NVIDIA CORPORATION 9.5%
    """

    match = re.match(
        r"^\s*(\d{1,2})\s*[\.\):\-]?\s+(.+?)\s*$",
        line,
    )

    if not match:
        return None

    rank = int(match.group(1))

    if 1 <= rank <= MAX_HOLDINGS:
        return rank

    return None


def looks_like_non_holding_line(
    line: str,
) -> bool:
    lower = line.lower()

    excluded_prefixes = (
        "source:",
        "all data as at",
        "data as at",
        "benchmark",
        "performance",
        "offer-bid",
        "bid-bid",
        "calendar year",
        "inception date",
        "investment objective",
        "fund details",
        "risk classification",
        "subscription method",
        "manager of the fund",
        "funds under management",
        "underlying fund size",
        "financial year end",
        "initial investment charge",
        "continuing investment charge",
        "country allocation",
        "sector allocation",
        "asset allocation",
        "important information",
        "important notes",
    )

    if lower.startswith(excluded_prefixes):
        return True

    # A pure percentage line is not a holding name.
    if re.fullmatch(
        r"[+-]?\d+(?:\.\d+)?\s*%",
        line,
    ):
        return True

    return False


def parse_holdings(
    section: str,
) -> list[dict[str, Any]]:
    """
    Parse holdings using rank as the primary row boundary.

    IMPORTANT:

    Percentage is optional.

    A new rank always finalizes the previous holding.

    This directly fixes the previous failure:

        "A new holding rank appeared before the previous holding
         received a published percentage."

    That condition is no longer an error.

    Missing percentage:
        weightPercent = None
        weightText = None
    """

    raw_lines = section.splitlines()

    lines: list[str] = []

    for raw_line in raw_lines:
        line = clean_text(raw_line)

        if not line:
            continue

        if looks_like_non_holding_line(line):
            continue

        lines.append(line)

    holdings: list[dict[str, Any]] = []

    current: Optional[dict[str, Any]] = None

    def finalize_current() -> None:
        nonlocal current

        if current is None:
            return

        name = clean_holding_name(
            current.get("holdingName", "")
        )

        if not name:
            raise RuntimeError(
                f"Holding rank {current.get('rank')} "
                f"does not have an identifiable holding name."
            )

        current["holdingName"] = name

        holdings.append(current)

        current = None

    for line in lines:
        rank = rank_from_line(line)

        if rank is not None:
            # --------------------------------------------------------------
            # New rank = hard row boundary.
            # --------------------------------------------------------------

            finalize_current()

            remainder = re.sub(
                r"^\s*\d{1,2}\s*[\.\):\-]?\s+",
                "",
                line,
            ).strip()

            percentage, weight_text = (
                extract_percentage_from_line(
                    remainder
                )
            )

            name_without_percentage = (
                strip_percentage_from_line(
                    remainder
                )
            )

            name_without_percentage = (
                clean_holding_name(
                    name_without_percentage
                )
            )

            if not name_without_percentage:
                raise RuntimeError(
                    f"Holding rank {rank} "
                    f"does not have an identifiable name."
                )

            current = {
                "rank": rank,
                "holdingName": name_without_percentage,
                "weightPercent": percentage,
                "weightText": weight_text,
            }

            continue

        # ------------------------------------------------------------------
        # Non-ranked continuation line.
        # ------------------------------------------------------------------

        if current is None:
            # Do not invent a rank.
            continue

        percentage, weight_text = (
            extract_percentage_from_line(
                line
            )
        )

        name_part = strip_percentage_from_line(
            line
        )

        name_part = clean_holding_name(
            name_part
        )

        if name_part:
            current_name = current.get(
                "holdingName",
                "",
            )

            current["holdingName"] = (
                f"{current_name} {name_part}"
            ).strip()

        if (
            percentage is not None
            and current.get("weightPercent") is None
        ):
            current["weightPercent"] = percentage
            current["weightText"] = weight_text

    finalize_current()

    # ----------------------------------------------------------------------
    # Validate rank sequence.
    # ----------------------------------------------------------------------

    if not holdings:
        raise RuntimeError(
            "Top 10 Holdings section was found, "
            "but no holding rows could be identified."
        )

    if len(holdings) > MAX_HOLDINGS:
        raise RuntimeError(
            f"Parsed {len(holdings)} holdings, "
            f"which exceeds maximum {MAX_HOLDINGS}."
        )

    expected_rank = 1

    for holding in holdings:
        rank = holding["rank"]

        if rank != expected_rank:
            raise RuntimeError(
                "Holding ranks are not sequential. "
                f"Expected rank {expected_rank}, "
                f"found rank {rank}."
            )

        expected_rank += 1

    return holdings


# ============================================================================
# HOLDINGS VALIDATION
# ============================================================================

def validate_holdings(
    holdings: list[dict[str, Any]],
) -> dict[str, Any]:
    if not holdings:
        raise RuntimeError(
            "No holdings were extracted."
        )

    if len(holdings) > MAX_HOLDINGS:
        raise RuntimeError(
            f"Too many holdings extracted: "
            f"{len(holdings)} > {MAX_HOLDINGS}"
        )

    names_seen: set[str] = set()

    missing_percentages = 0
    published_percentages = 0

    for holding in holdings:
        name = clean_holding_name(
            holding.get(
                "holdingName",
                "",
            )
        )

        if not name:
            raise RuntimeError(
                "Holding name is empty."
            )

        normalized_name = re.sub(
            r"\s+",
            " ",
            name.lower(),
        ).strip()

        if normalized_name in names_seen:
            raise RuntimeError(
                f"Duplicate holding name detected: {name}"
            )

        names_seen.add(
            normalized_name
        )

        percentage = holding.get(
            "weightPercent"
        )

        if percentage is None:
            missing_percentages += 1
        else:
            published_percentages += 1

            if not isinstance(
                percentage,
                (int, float),
            ):
                raise RuntimeError(
                    f"Invalid published percentage for "
                    f"{name}: {percentage!r}"
                )

    return {
        "holdingCount": len(holdings),
        "publishedPercentageCount": (
            published_percentages
        ),
        "missingPublishedPercentageCount": (
            missing_percentages
        ),
    }


# ============================================================================
# OUTPUT
# ============================================================================

def ensure_output_directories() -> None:
    for directory in (
        OUTPUT_DIR,
        PDF_DIR,
        TEXT_DIR,
        JSON_DIR,
        METADATA_DIR,
    ):
        directory.mkdir(
            parents=True,
            exist_ok=True,
        )


def write_json(
    path: Path,
    data: Any,
) -> None:
    path.parent.mkdir(
        parents=True,
        exist_ok=True,
    )

    with path.open(
        "w",
        encoding="utf-8",
    ) as handle:
        json.dump(
            data,
            handle,
            indent=2,
            ensure_ascii=False,
        )


def write_text(
    path: Path,
    text: str,
) -> None:
    path.parent.mkdir(
        parents=True,
        exist_ok=True,
    )

    path.write_text(
        text,
        encoding="utf-8",
    )


# ============================================================================
# PROCESS ONE FUND
# ============================================================================

def process_fund(
    session: requests.Session,
    fund: dict[str, Any],
) -> dict[str, Any]:
    excel_row = fund["excelRow"]
    fund_url = fund["url"]
    pru_access_name = fund["pruAccessFundName"]

    started_at = utc_now_iso()

    result: dict[str, Any] = {
        "status": "failed",
        "excelRow": excel_row,
        "url": fund_url,
        "pruAccessFundName": pru_access_name,
        "startedAtUtc": started_at,
    }

    fund_key = safe_filename(
        pru_access_name
        or f"excel-row-{excel_row}"
    )

    try:
        # --------------------------------------------------------------
        # Fund page
        # --------------------------------------------------------------

        html, final_page_url = (
            download_fund_page(
                session,
                fund_url,
            )
        )

        page_metadata = (
            extract_fund_page_metadata(
                html,
                final_page_url,
            )
        )

        result["page"] = page_metadata

        # --------------------------------------------------------------
        # Factsheet discovery/download
        # --------------------------------------------------------------

        (
            factsheet_url,
            pdf_bytes,
            candidate_errors,
        ) = find_and_download_factsheet(
            session,
            html,
            final_page_url,
        )

        result["factsheetUrl"] = factsheet_url

        if candidate_errors:
            result["factsheetCandidateErrors"] = (
                candidate_errors
            )

        # --------------------------------------------------------------
        # Save PDF
        # --------------------------------------------------------------

        pdf_path = (
            PDF_DIR
            / f"{fund_key}.pdf"
        )

        pdf_path.write_bytes(
            pdf_bytes
        )

        result["factsheetPdfPath"] = str(
            pdf_path
        )

        result["factsheetPdfBytes"] = (
            len(pdf_bytes)
        )

        # --------------------------------------------------------------
        # Extract PDF text
        # --------------------------------------------------------------

        pdf_text, page_texts = (
            extract_pdf_text(
                pdf_bytes
            )
        )

        text_path = (
            TEXT_DIR
            / f"{fund_key}.txt"
        )

        write_text(
            text_path,
            pdf_text,
        )

        result["factsheetTextPath"] = str(
            text_path
        )

        result["pdfPageCount"] = len(
            page_texts
        )

        # --------------------------------------------------------------
        # Top holdings section
        # --------------------------------------------------------------

        (
            holdings_section,
            section_metadata,
        ) = extract_top_holdings_section(
            pdf_text
        )

        result["topHoldingsSection"] = (
            section_metadata
        )

        if holdings_section is None:
            result["status"] = (
                "noHoldingsSection"
            )

            result["completedAtUtc"] = (
                utc_now_iso()
            )

            metadata_path = (
                METADATA_DIR
                / f"{fund_key}.json"
            )

            write_json(
                metadata_path,
                result,
            )

            return result

        # Save isolated section for debugging.
        section_path = (
            TEXT_DIR
            / f"{fund_key}_top_holdings.txt"
        )

        write_text(
            section_path,
            holdings_section,
        )

        result["topHoldingsSectionPath"] = (
            str(section_path)
        )

        # --------------------------------------------------------------
        # Parse holdings
        # --------------------------------------------------------------

        holdings = parse_holdings(
            holdings_section
        )

        validation = validate_holdings(
            holdings
        )

        # --------------------------------------------------------------
        # Final holdings result
        # --------------------------------------------------------------

        result["holdings"] = holdings

        result["summary"] = {
            "publishedHoldingCount": (
                validation["holdingCount"]
            ),
            "publishedPercentageCount": (
                validation[
                    "publishedPercentageCount"
                ]
            ),
            "missingPublishedPercentageCount": (
                validation[
                    "missingPublishedPercentageCount"
                ]
            ),
        }

        result["status"] = "success"
        result["completedAtUtc"] = (
            utc_now_iso()
        )

        # --------------------------------------------------------------
        # Per-fund JSON
        # --------------------------------------------------------------

        fund_json = {
            "status": "success",
            "excelRow": excel_row,
            "url": fund_url,
            "finalFundPageUrl": final_page_url,
            "pruAccessFundName": (
                pru_access_name
            ),
            "fundPageMetadata": (
                page_metadata
            ),
            "factsheetUrl": factsheet_url,
            "factsheetPdfBytes": len(
                pdf_bytes
            ),
            "holdings": holdings,
            "summary": result["summary"],
            "rules": {
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

        per_fund_path = (
            JSON_DIR
            / f"{fund_key}.json"
        )

        write_json(
            per_fund_path,
            fund_json,
        )

        result["jsonPath"] = str(
            per_fund_path
        )

        metadata_path = (
            METADATA_DIR
            / f"{fund_key}.json"
        )

        write_json(
            metadata_path,
            result,
        )

        return result

    except Exception as exc:
        result["status"] = "failed"

        result["reason"] = (
            f"{type(exc).__name__}: {exc}"
        )

        result["completedAtUtc"] = (
            utc_now_iso()
        )

        metadata_path = (
            METADATA_DIR
            / f"{fund_key}.json"
        )

        write_json(
            metadata_path,
            result,
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
        if result.get("status")
        == "noHoldingsSection"
    ]

    failed = [
        result
        for result in results
        if result.get("status") == "failed"
    ]

    total_holdings = 0
    total_published_percentages = 0
    total_missing_percentages = 0

    for result in successful:
        summary = result.get(
            "summary",
            {},
        )

        total_holdings += int(
            summary.get(
                "publishedHoldingCount",
                0,
            )
        )

        total_published_percentages += int(
            summary.get(
                "publishedPercentageCount",
                0,
            )
        )

        total_missing_percentages += int(
            summary.get(
                "missingPublishedPercentageCount",
                0,
            )
        )

    failed_funds = []

    for result in failed:
        failed_funds.append(
            {
                "excelRow": result.get(
                    "excelRow"
                ),
                "url": result.get(
                    "url"
                ),
                "pruAccessFundName": result.get(
                    "pruAccessFundName"
                ),
                "reason": result.get(
                    "reason"
                ),
            }
        )

    summary = {
        "status": (
            "success"
            if not failed
            else "failed"
        ),
        "startedAtUtc": started_at,
        "completedAtUtc": completed_at,
        "excelFile": str(
            EXCEL_FILE
        ),
        "fundUniverseCount": len(funds),
        "successfulFundCount": len(
            successful
        ),
        "noHoldingsSectionFundCount": len(
            no_holdings_section
        ),
        "failedFundCount": len(
            failed
        ),
        "totalHoldings": total_holdings,
        "totalPublishedPercentages": (
            total_published_percentages
        ),
        "totalMissingPublishedPercentages": (
            total_missing_percentages
        ),
        "failedFunds": failed_funds,
        "noHoldingsSectionFunds": [
            {
                "excelRow": result.get(
                    "excelRow"
                ),
                "url": result.get(
                    "url"
                ),
                "pruAccessFundName": result.get(
                    "pruAccessFundName"
                ),
            }
            for result in no_holdings_section
        ],
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

    return summary


# ============================================================================
# MAIN
# ============================================================================

def main() -> int:
    started_at = utc_now_iso()

    ensure_output_directories()

    print(
        "============================================================"
    )
    print(
        "Prudential Singapore Top Holdings Extractor"
    )
    print(
        "============================================================"
    )
    print(
        f"Excel file: {EXCEL_FILE}"
    )
    print(
        "Official source: prudential.com.sg"
    )
    print(
        f"Maximum holdings: {MAX_HOLDINGS}"
    )
    print(
        "Factsheet discovery: strict official factsheet links"
    )
    print(
        "Percentage calculation: DISABLED"
    )
    print(
        "Percentage inference: DISABLED"
    )
    print(
        "Holding-name inference: DISABLED"
    )
    print(
        "============================================================"
    )

    try:
        funds = read_excel_funds()

    except Exception as exc:
        print(
            f"ERROR reading Excel: "
            f"{type(exc).__name__}: {exc}",
            file=sys.stderr,
        )
        return 1

    print(
        f"Excel fund universe: {len(funds)}"
    )

    if not funds:
        print(
            "ERROR: No populated fund URLs found "
            "in Excel Column A.",
            file=sys.stderr,
        )
        return 1

    session = create_session()

    results: list[dict[str, Any]] = []

    for index, fund in enumerate(
        funds,
        start=1,
    ):
        print(
            ""
        )

        print(
            f"[{index}/{len(funds)}] "
            f"Excel row {fund['excelRow']}"
        )

        print(
            f"  PruAccess name: "
            f"{fund['pruAccessFundName']}"
        )

        print(
            f"  URL: {fund['url']}"
        )

        result = process_fund(
            session,
            fund,
        )

        results.append(
            result
        )

        status = result.get(
            "status"
        )

        if status == "success":
            summary = result.get(
                "summary",
                {},
            )

            print(
                "  STATUS: SUCCESS"
            )

            print(
                "  Factsheet: "
                f"{result.get('factsheetUrl')}"
            )

            print(
                "  Holdings: "
                f"{summary.get('publishedHoldingCount', 0)}"
            )

            print(
                "  Published percentages: "
                f"{summary.get('publishedPercentageCount', 0)}"
            )

            print(
                "  Missing percentages: "
                f"{summary.get('missingPublishedPercentageCount', 0)}"
            )

        elif status == "noHoldingsSection":
            print(
                "  STATUS: NO HOLDINGS SECTION"
            )

            print(
                "  Factsheet: "
                f"{result.get('factsheetUrl')}"
            )

        else:
            print(
                "  STATUS: FAILED"
            )

            print(
                f"  REASON: {result.get('reason')}",
                file=sys.stderr,
            )

    completed_at = utc_now_iso()

    run_summary = build_run_summary(
        started_at,
        completed_at,
        funds,
        results,
    )

    write_json(
        RUN_SUMMARY_FILE,
        run_summary,
    )

    all_holdings: list[dict[str, Any]] = []

    for result in results:
        if result.get("status") != "success":
            continue

        all_holdings.append(
            {
                "excelRow": result.get(
                    "excelRow"
                ),
                "url": result.get(
                    "url"
                ),
                "pruAccessFundName": result.get(
                    "pruAccessFundName"
                ),
                "factsheetUrl": result.get(
                    "factsheetUrl"
                ),
                "holdings": result.get(
                    "holdings",
                    [],
                ),
                "summary": result.get(
                    "summary",
                    {},
                ),
            }
        )

    write_json(
        ALL_HOLDINGS_FILE,
        {
            "status": run_summary[
                "status"
            ],
            "generatedAtUtc": completed_at,
            "fundUniverseCount": len(
                funds
            ),
            "successfulFundCount": run_summary[
                "successfulFundCount"
            ],
            "noHoldingsSectionFundCount": run_summary[
                "noHoldingsSectionFundCount"
            ],
            "failedFundCount": run_summary[
                "failedFundCount"
            ],
            "funds": all_holdings,
        },
    )

    print("")
    print(
        "============================================================"
    )
    print(
        "RUN COMPLETE"
    )
    print(
        "============================================================"
    )

    print(
        f"Fund universe: "
        f"{run_summary['fundUniverseCount']}"
    )

    print(
        f"Successful: "
        f"{run_summary['successfulFundCount']}"
    )

    print(
        f"No holdings section: "
        f"{run_summary['noHoldingsSectionFundCount']}"
    )

    print(
        f"Failed: "
        f"{run_summary['failedFundCount']}"
    )

    print(
        f"Total holdings: "
        f"{run_summary['totalHoldings']}"
    )

    print(
        f"Published percentages: "
        f"{run_summary['totalPublishedPercentages']}"
    )

    print(
        f"Missing published percentages: "
        f"{run_summary['totalMissingPublishedPercentages']}"
    )

    print(
        f"Run summary: {RUN_SUMMARY_FILE}"
    )

    print(
        f"All holdings: {ALL_HOLDINGS_FILE}"
    )

    print(
        "============================================================"
    )

    if run_summary["failedFundCount"] > 0:
        return 1

    return 0


if __name__ == "__main__":
    raise SystemExit(
        main()
    )
