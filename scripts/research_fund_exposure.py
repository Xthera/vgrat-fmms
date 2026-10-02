#!/usr/bin/env python3

"""
VGrat FMS - AI FUND EXPOSURE RESEARCH

Purpose
-------
Research geography and sector exposure for Prudential Singapore
ILP funds using LOCAL Hugging Face AI.

The pipeline supports:

    1. Direct funds
       - The fund invests directly in securities.
       - AI analyses its own holdings/allocation/strategy.

    2. Fund-of-funds
       - AI identifies underlying funds.
       - The pipeline researches the available underlying-fund
         documents/pages.
       - AI classifies each underlying fund.
       - Results are aggregated.

    3. Mixed structures
       - Both direct portfolio evidence and underlying-fund
         evidence may be used.

Important
---------
AI is the classifier.

The Python code only:
    - retrieves documents/pages
    - extracts text
    - identifies candidate linked documents
    - asks AI structured questions
    - validates AI output
    - aggregates already-classified AI results

No search engine is used.
No DuckDuckGo is used.
No external search API is used.

Publishing rule
---------------
    - Complete successful run -> replace fund_exposure.json
    - Any failed fund -> preserve existing fund_exposure.json
    - Any model/AI/source/validation failure -> preserve existing file
    - Partial results are NEVER published

Input
-----
    Funds Links.xlsx
    or
    Funds Links.xlsm

Column A:
    Prudential URL

Column B:
    Exact Prudential Fund Name

Categories
----------
    scripts/exposure_categories.txt

Output
------
    data/fund_exposure.json

Testing
-------
    python scripts/research_fund_exposure.py --limit 1

Full run
--------
    python scripts/research_fund_exposure.py
"""

from __future__ import annotations

import argparse
import io
import json
import os
import re
import sys
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Dict, List, Optional, Set, Tuple
from urllib.parse import urljoin, urlparse, urldefrag

import requests
from bs4 import BeautifulSoup
from openpyxl import load_workbook
from pypdf import PdfReader

import torch
from transformers import AutoModelForCausalLM, AutoTokenizer


# ============================================================
# CONFIGURATION
# ============================================================

MODEL_NAME = os.getenv(
    "HF_MODEL",
    "Qwen/Qwen3-1.7B",
)

HF_TOKEN = os.getenv(
    "HF_TOKEN",
    "",
).strip()

SCRIPT_DIR = Path(__file__).resolve().parent
REPOSITORY_ROOT = SCRIPT_DIR.parent

CATEGORIES_FILE = (
    SCRIPT_DIR / "exposure_categories.txt"
)

OUTPUT_PATH = (
    REPOSITORY_ROOT
    / "data"
    / "fund_exposure.json"
)

REQUEST_TIMEOUT = int(
    os.getenv(
        "REQUEST_TIMEOUT",
        "60",
    )
)

MAX_PDF_PAGES = int(
    os.getenv(
        "MAX_PDF_PAGES",
        "60",
    )
)

MAX_SOURCE_CHARS = int(
    os.getenv(
        "MAX_SOURCE_CHARS",
        "100000",
    )
)

MAX_AI_SOURCE_CHARS = int(
    os.getenv(
        "MAX_AI_SOURCE_CHARS",
        "24000",
    )
)

MAX_DISCOVERY_DOCUMENTS = int(
    os.getenv(
        "MAX_DISCOVERY_DOCUMENTS",
        "12",
    )
)

MAX_UNDERLYING_FUNDS = int(
    os.getenv(
        "MAX_UNDERLYING_FUNDS",
        "10",
    )
)

MAX_CLASSIFICATION_DOCUMENTS = int(
    os.getenv(
        "MAX_CLASSIFICATION_DOCUMENTS",
        "10",
    )
)

MAX_NEW_TOKENS = int(
    os.getenv(
        "MAX_NEW_TOKENS",
        "600",
    )
)

# Number of characters shown in diagnostics.
SOURCE_PREVIEW_CHARS = int(
    os.getenv(
        "SOURCE_PREVIEW_CHARS",
        "5000",
    )
)

USER_AGENT = (
    "Mozilla/5.0 (X11; Linux x86_64) "
    "AppleWebKit/537.36 "
    "(KHTML, like Gecko) "
    "Chrome/131.0.0.0 Safari/537.36"
)

HEADERS = {
    "User-Agent": USER_AGENT,
    "Accept": (
        "text/html,application/xhtml+xml,"
        "application/xml;q=0.9,"
        "application/pdf;q=0.9,"
        "*/*;q=0.8"
    ),
    "Accept-Language": "en-US,en;q=0.9",
}


# ============================================================
# GLOBAL AI OBJECTS
# ============================================================

TOKENIZER = None
MODEL = None


# ============================================================
# LOGGING
# ============================================================

def log(message: str = "") -> None:
    print(
        message,
        flush=True,
    )


def utc_now() -> str:
    return datetime.now(
        timezone.utc
    ).isoformat()


# ============================================================
# TEXT HELPERS
# ============================================================

def clean_text(
    text: str,
) -> str:

    if not text:
        return ""

    text = text.replace(
        "\x00",
        " ",
    )

    text = text.replace(
        "\r\n",
        "\n",
    )

    text = text.replace(
        "\r",
        "\n",
    )

    # Remove excessive horizontal whitespace.
    text = re.sub(
        r"[ \t]+",
        " ",
        text,
    )

    # Remove excessive blank lines.
    text = re.sub(
        r"\n{3,}",
        "\n\n",
        text,
    )

    return text.strip()


def truncate_text(
    text: str,
    maximum: int,
) -> str:

    if len(text) <= maximum:
        return text

    return (
        text[:maximum]
        + "\n\n[DOCUMENT TRUNCATED]"
    )


def normalize_space(
    text: str,
) -> str:

    return re.sub(
        r"\s+",
        " ",
        str(text or ""),
    ).strip()


def normalize_name(
    text: str,
) -> str:

    text = normalize_space(
        text
    )

    text = text.lower()

    text = re.sub(
        r"[^a-z0-9]+",
        " ",
        text,
    )

    return text.strip()


# ============================================================
# WORKBOOK
# ============================================================

def find_funds_excel_file(
    repository_root: Path,
) -> Path:

    candidates = [
        repository_root / "Funds Links.xlsx",
        repository_root / "Funds Links.xlsm",
    ]

    for candidate in candidates:

        if candidate.exists():
            return candidate

    raise FileNotFoundError(
        "Could not find Funds Links.xlsx or "
        f"Funds Links.xlsm in {repository_root}"
    )


def load_funds(
    workbook_path: Path,
) -> List[Dict[str, str]]:

    log(
        f"Loading workbook: {workbook_path}"
    )

    workbook = load_workbook(
        workbook_path,
        read_only=True,
        data_only=True,
    )

    try:

        worksheet = workbook.active

        funds: List[
            Dict[str, str]
        ] = []

        for row_number, row in enumerate(
            worksheet.iter_rows(
                min_row=2,
                values_only=True,
            ),
            start=2,
        ):

            url = (
                row[0]
                if len(row) >= 1
                else None
            )

            fund_name = (
                row[1]
                if len(row) >= 2
                else None
            )

            if (
                url is None
                and fund_name is None
            ):
                continue

            url = str(
                url or ""
            ).strip()

            fund_name = str(
                fund_name or ""
            ).strip()

            if not fund_name:

                log(
                    f"WARNING: Row {row_number} "
                    "has no fund name. Skipping."
                )

                continue

            if not url:

                raise RuntimeError(
                    f"Row {row_number} "
                    f"({fund_name}) has no Prudential URL."
                )

            funds.append(
                {
                    "fundName": fund_name,
                    "prudentialUrl": url,
                }
            )

        if not funds:

            raise RuntimeError(
                "No funds found in workbook."
            )

        log(
            f"Funds loaded: {len(funds)}"
        )

        return funds

    finally:

        workbook.close()


# ============================================================
# CATEGORY VOCABULARY
# ============================================================

def load_categories(
    categories_path: Path,
) -> Tuple[
    List[str],
    List[str],
]:

    if not categories_path.exists():

        raise FileNotFoundError(
            f"Category file not found: "
            f"{categories_path}"
        )

    geography: List[str] = []
    sector: List[str] = []

    current_section: Optional[
        str
    ] = None

    with categories_path.open(
        "r",
        encoding="utf-8",
    ) as file:

        for raw_line in file:

            line = raw_line.strip()

            if not line:
                continue

            if line.startswith("#"):
                continue

            upper = line.upper()

            if upper == "[GEOGRAPHY]":

                current_section = (
                    "geography"
                )

                continue

            if upper == "[SECTOR]":

                current_section = (
                    "sector"
                )

                continue

            if (
                current_section
                == "geography"
            ):

                geography.append(
                    line
                )

            elif (
                current_section
                == "sector"
            ):

                sector.append(
                    line
                )

    geography = list(
        dict.fromkeys(
            geography
        )
    )

    sector = list(
        dict.fromkeys(
            sector
        )
    )

    if not geography:

        raise RuntimeError(
            "No geography categories found."
        )

    if not sector:

        raise RuntimeError(
            "No sector categories found."
        )

    return (
        geography,
        sector,
    )


# ============================================================
# HTTP
# ============================================================

def get_response(
    url: str,
) -> requests.Response:

    response = requests.get(
        url,
        headers=HEADERS,
        timeout=REQUEST_TIMEOUT,
        allow_redirects=True,
    )

    response.raise_for_status()

    return response


# ============================================================
# URL HELPERS
# ============================================================

def canonicalize_url(
    base_url: str,
    href: str,
) -> str:

    absolute = urljoin(
        base_url,
        href,
    )

    absolute = urldefrag(
        absolute
    )[0]

    return absolute.strip()


def is_http_url(
    url: str,
) -> bool:

    parsed = urlparse(
        url
    )

    return parsed.scheme.lower() in {
        "http",
        "https",
    }


def same_domain(
    first_url: str,
    second_url: str,
) -> bool:

    first_host = (
        urlparse(first_url)
        .netloc
        .lower()
        .split(":")[0]
    )

    second_host = (
        urlparse(second_url)
        .netloc
        .lower()
        .split(":")[0]
    )

    return (
        first_host == second_host
    )


# ============================================================
# PDF EXTRACTION
# ============================================================

def extract_pdf_text(
    raw_bytes: bytes,
) -> str:

    reader = PdfReader(
        io.BytesIO(raw_bytes)
    )

    page_count = min(
        len(reader.pages),
        MAX_PDF_PAGES,
    )

    chunks: List[str] = []

    for page_number in range(
        page_count
    ):

        try:

            page_text = (
                reader.pages[
                    page_number
                ].extract_text()
                or ""
            )

            page_text = clean_text(
                page_text
            )

            if page_text:

                chunks.append(
                    "\n"
                    f"--- PAGE "
                    f"{page_number + 1} ---\n"
                    f"{page_text}"
                )

        except Exception as exc:

            log(
                f"        PDF page "
                f"{page_number + 1} "
                f"extraction failed: {exc}"
            )

    return clean_text(
        "\n".join(
            chunks
        )
    )


# ============================================================
# HTML EXTRACTION
# ============================================================

def extract_html(
    html: str,
) -> Tuple[
    str,
    List[Dict[str, str]],
]:

    soup = BeautifulSoup(
        html,
        "html.parser",
    )

    # Remove content that is almost never useful
    # for portfolio exposure analysis.
    for tag in soup(
        [
            "script",
            "style",
            "noscript",
            "svg",
            "iframe",
        ]
    ):

        tag.decompose()

    links: List[
        Dict[str, str]
    ] = []

    for anchor in soup.find_all(
        "a"
    ):

        href = anchor.get(
            "href"
        )

        if not href:
            continue

        href = href.strip()

        text = normalize_space(
            anchor.get_text(
                " ",
                strip=True,
            )
        )

        links.append(
            {
                "href": href,
                "text": text,
            }
        )

    text = clean_text(
        soup.get_text(
            "\n",
            strip=True,
        )
    )

    return (
        text,
        links,
    )


# ============================================================
# DOCUMENT TYPE DETECTION
# ============================================================

DOCUMENT_KEYWORDS = [
    "fund",
    "funds",
    "portfolio",
    "holdings",
    "holding",
    "factsheet",
    "fact sheet",
    "fund fact",
    "investment",
    "allocation",
    "asset",
    "equity",
    "bond",
    "income",
    "manager",
    "morningstar",
    "fund profile",
    "fund information",
    "investment objective",
    "investment strategy",
]


def link_relevance_score(
    link_text: str,
    url: str,
) -> int:

    combined = (
        normalize_space(
            link_text
        )
        + " "
        + normalize_space(
            url
        )
    ).lower()

    score = 0

    for keyword in DOCUMENT_KEYWORDS:

        if keyword in combined:

            score += 2

    # PDFs are particularly valuable.
    if ".pdf" in combined:
        score += 8

    # URLs with likely fund-document patterns.
    for keyword in [
        "factsheet",
        "fact-sheet",
        "fundfactsheet",
        "fund-fact",
        "fundprofile",
        "fund-profile",
        "funddocument",
        "fund-document",
        "funddetails",
        "fund-details",
        "portfolio",
        "holdings",
    ]:

        if keyword in combined:

            score += 5

    return score


# ============================================================
# SOURCE DOCUMENT
# ============================================================

class SourceDocument:

    def __init__(
        self,
        url: str,
        title: str,
        text: str,
        source_type: str,
    ) -> None:

        self.url = url
        self.title = title
        self.text = text
        self.source_type = source_type

    def to_dict(
        self,
    ) -> Dict[str, Any]:

        return {
            "url": self.url,
            "title": self.title,
            "sourceType": self.source_type,
            "characters": len(
                self.text
            ),
        }


# ============================================================
# SOURCE COLLECTION
# ============================================================

def fetch_single_document(
    url: str,
) -> Optional[
    SourceDocument
]:

    try:

        response = get_response(
            url
        )

    except Exception as exc:

        log(
            f"        Document fetch failed: "
            f"{url} -> {exc}"
        )

        return None

    final_url = response.url

    content_type = (
        response.headers.get(
            "Content-Type",
            "",
        )
        .lower()
    )

    # --------------------------------------------------------
    # PDF
    # --------------------------------------------------------

    if (
        "application/pdf"
        in content_type
        or final_url.lower()
        .split("?")[0]
        .endswith(".pdf")
    ):

        text = extract_pdf_text(
            response.content
        )

        if not text:

            return None

        return SourceDocument(
            url=final_url,
            title=Path(
                urlparse(
                    final_url
                ).path
            ).name,
            text=truncate_text(
                text,
                MAX_SOURCE_CHARS,
            ),
            source_type="pdf",
        )

    # --------------------------------------------------------
    # HTML
    # --------------------------------------------------------

    try:

        html_text, _ = extract_html(
            response.text
        )

    except Exception:

        return None

    if not html_text:

        return None

    return SourceDocument(
        url=final_url,
        title=(
            urlparse(
                final_url
            ).path
            or final_url
        ),
        text=truncate_text(
            html_text,
            MAX_SOURCE_CHARS,
        ),
        source_type="html",
    )


def collect_prudential_sources(
    root_url: str,
) -> Tuple[
    List[SourceDocument],
    List[Dict[str, str]],
]:

    log(
        f"        Fetching source: "
        f"{root_url}"
    )

    root_document = (
        fetch_single_document(
            root_url
        )
    )

    if root_document is None:

        raise RuntimeError(
            "Unable to extract usable "
            "Prudential source."
        )

    documents: List[
        SourceDocument
    ] = [
        root_document
    ]

    discovered_links: List[
        Dict[str, str]
    ] = []

    # --------------------------------------------------------
    # If root is HTML, inspect its links.
    # --------------------------------------------------------

    if (
        root_document.source_type
        == "html"
    ):

        try:

            response = get_response(
                root_url
            )

            html_text, links = (
                extract_html(
                    response.text
                )
            )

            del html_text

        except Exception:

            links = []

        scored_links: List[
            Tuple[
                int,
                Dict[str, str],
            ]
        ] = []

        seen_urls: Set[str] = set()

        for link in links:

            absolute = canonicalize_url(
                root_document.url,
                link.get(
                    "href",
                    "",
                ),
            )

            if not absolute:
                continue

            if not is_http_url(
                absolute
            ):
                continue

            if absolute in seen_urls:
                continue

            seen_urls.add(
                absolute
            )

            text = link.get(
                "text",
                "",
            )

            score = link_relevance_score(
                text,
                absolute,
            )

            # Same-domain links are preferred.
            if same_domain(
                root_document.url,
                absolute,
            ):

                score += 5

            # Any explicit PDF receives a strong preference.
            if ".pdf" in absolute.lower():

                score += 10

            scored_links.append(
                (
                    score,
                    {
                        "url": absolute,
                        "text": text,
                    },
                )
            )

        scored_links.sort(
            key=lambda item: item[0],
            reverse=True,
        )

        discovered_links = [
            item[1]
            for item in scored_links
        ]

        log(
            f"        Candidate linked "
            f"documents: "
            f"{len(discovered_links)}"
        )

        # ----------------------------------------------------
        # Fetch the most relevant documents.
        # ----------------------------------------------------

        for (
            score,
            link,
        ) in scored_links[
            :MAX_DISCOVERY_DOCUMENTS
        ]:

            url = link["url"]

            log(
                f"        Candidate "
                f"document "
                f"(score {score}): "
                f"{url}"
            )

            document = (
                fetch_single_document(
                    url
                )
            )

            if document is None:
                continue

            if any(
                existing.url
                == document.url
                for existing in documents
            ):

                continue

            documents.append(
                document
            )

    # Deduplicate.
    unique_documents: List[
        SourceDocument
    ] = []

    seen_document_urls: Set[
        str
    ] = set()

    for document in documents:

        if (
            document.url
            in seen_document_urls
        ):

            continue

        seen_document_urls.add(
            document.url
        )

        unique_documents.append(
            document
        )

    return (
        unique_documents,
        discovered_links,
    )


# ============================================================
# EXPOSURE SECTION EXTRACTION
# ============================================================

EXPOSURE_SECTION_PATTERNS = [
    r"asset allocation",
    r"asset mix",
    r"asset allocation by",
    r"geographical allocation",
    r"geographic allocation",
    r"geographical exposure",
    r"geographic exposure",
    r"country allocation",
    r"country exposure",
    r"regional allocation",
    r"regional exposure",
    r"sector allocation",
    r"sector exposure",
    r"industry allocation",
    r"industry exposure",
    r"top holdings",
    r"top holding",
    r"holdings",
    r"portfolio allocation",
    r"portfolio holdings",
    r"equity allocation",
    r"equities",
    r"fixed income",
    r"bond allocation",
    r"bond holdings",
    r"underlying fund",
    r"underlying funds",
    r"underlying investment",
    r"investment objective",
    r"investment strategy",
    r"investment approach",
    r"investment policy",
    r"portfolio",
]


def extract_relevant_sections(
    source_text: str,
    maximum: int = MAX_AI_SOURCE_CHARS,
) -> str:

    lines = source_text.splitlines()

    if not lines:
        return ""

    matched_indexes: List[
        int
    ] = []

    for index, line in enumerate(
        lines
    ):

        lower = line.lower()

        for pattern in (
            EXPOSURE_SECTION_PATTERNS
        ):

            if re.search(
                pattern,
                lower,
            ):

                start = max(
                    0,
                    index - 10,
                )

                end = min(
                    len(lines),
                    index + 40,
                )

                matched_indexes.extend(
                    range(
                        start,
                        end,
                    )
                )

                break

    if not matched_indexes:
        return ""

    unique_indexes = sorted(
        set(
            matched_indexes
        )
    )

    chunks: List[str] = []

    current_chunk: List[
        str
    ] = []

    previous_index: Optional[
        int
    ] = None

    for index in unique_indexes:

        if (
            previous_index is not None
            and index
            > previous_index + 1
        ):

            if current_chunk:

                chunks.append(
                    "\n".join(
                        current_chunk
                    )
                )

            current_chunk = []

        current_chunk.append(
            lines[index]
        )

        previous_index = index

    if current_chunk:

        chunks.append(
            "\n".join(
                current_chunk
            )
        )

    relevant = clean_text(
        "\n\n--- RELEVANT SECTION ---\n\n"
        .join(
            chunks
        )
    )

    return truncate_text(
        relevant,
        maximum,
    )


def remove_irrelevant_footer(
    source_text: str,
) -> str:

    """
    Remove common legal/footer content.

    This is preprocessing only.
    It does NOT classify anything.
    """

    lines = source_text.splitlines()

    useful_lines: List[
        str
    ] = []

    footer_patterns = [
        r"past performance",
        r"not necessarily indicative",
        r"not financial advice",
        r"not be relied upon",
        r"offer or solicitation",
        r"possible loss of",
        r"monetary authority",
        r"copyright \d{4}",
        r"all rights reserved",
        r"back to top",
        r"visit your local prudential",
        r"follow us singapore",
        r"privacy",
        r"terms and conditions",
        r"important notice",
        r"disclaimer",
    ]

    footer_started = False

    for line in lines:

        lower = line.lower()

        if any(
            re.search(
                pattern,
                lower,
            )
            for pattern in footer_patterns
        ):

            # Do not automatically discard the entire document
            # at the first occurrence because some documents
            # contain disclaimers in the middle.
            continue

        if (
            "back to top"
            in lower
        ):

            footer_started = True

        if footer_started:
            continue

        useful_lines.append(
            line
        )

    return clean_text(
        "\n".join(
            useful_lines
        )
    )


def prepare_ai_source(
    source_text: str,
) -> str:

    source_text = remove_irrelevant_footer(
        source_text
    )

    relevant = extract_relevant_sections(
        source_text
    )

    if relevant:

        return relevant

    return truncate_text(
        source_text,
        MAX_AI_SOURCE_CHARS,
    )


# ============================================================
# LINK / DOCUMENT MATCHING
# ============================================================

def candidate_text_for_document(
    document: SourceDocument,
) -> str:

    return (
        document.title
        + " "
        + document.url
        + " "
        + document.text[:5000]
    )


def document_matches_underlying_name(
    document: SourceDocument,
    underlying_name: str,
) -> bool:

    name = normalize_name(
        underlying_name
    )

    if not name:
        return False

    candidate = normalize_name(
        candidate_text_for_document(
            document
        )
    )

    name_tokens = [
        token
        for token in name.split()
        if len(token) >= 3
    ]

    if not name_tokens:
        return False

    # Exact normalized phrase.
    if name in candidate:
        return True

    # Strong token overlap.
    matches = sum(
        1
        for token in name_tokens
        if token in candidate
    )

    required = max(
        2,
        min(
            len(name_tokens),
            4,
        ),
    )

    return (
        matches >= required
    )


def rank_documents_for_underlying_fund(
    documents: List[SourceDocument],
    underlying_name: str,
) -> List[SourceDocument]:

    ranked: List[
        Tuple[
            int,
            SourceDocument,
        ]
    ] = []

    target = normalize_name(
        underlying_name
    )

    target_tokens = [
        token
        for token in target.split()
        if len(token) >= 3
    ]

    for document in documents:

        candidate = normalize_name(
            candidate_text_for_document(
                document
            )
        )

        score = 0

        if target and target in candidate:

            score += 100

        for token in target_tokens:

            if token in candidate:
                score += 8

        if (
            document.source_type
            == "pdf"
        ):

            score += 10

        relevant = (
            extract_relevant_sections(
                document.text,
                maximum=12000,
            )
        )

        if relevant:

            score += 15

        ranked.append(
            (
                score,
                document,
            )
        )

    ranked.sort(
        key=lambda item: item[0],
        reverse=True,
    )

    return [
        item[1]
        for item in ranked
        if item[0] > 0
    ]


# ============================================================
# AI MODEL
# ============================================================

def load_ai_model() -> None:

    global TOKENIZER
    global MODEL

    log("")
    log(
        "=" * 70
    )
    log(
        "LOADING LOCAL AI MODEL"
    )
    log(
        "=" * 70
    )

    log(
        f"Model: {MODEL_NAME}"
    )

    token_kwargs: Dict[
        str,
        Any,
    ] = {}

    if HF_TOKEN:

        token_kwargs["token"] = (
            HF_TOKEN
        )

        log(
            "Hugging Face authentication: "
            "enabled"
        )

    else:

        log(
            "Hugging Face authentication: "
            "not configured"
        )

    TOKENIZER = (
        AutoTokenizer.from_pretrained(
            MODEL_NAME,
            **token_kwargs,
        )
    )

    MODEL = (
        AutoModelForCausalLM.from_pretrained(
            MODEL_NAME,
            torch_dtype="auto",
            device_map="auto",
            trust_remote_code=True,
            **token_kwargs,
        )
    )

    MODEL.eval()

    log(
        "AI model loaded successfully."
    )


# ============================================================
# AI JSON GENERATION
# ============================================================

def generate_ai_response(
    prompt: str,
    max_new_tokens: Optional[
        int
    ] = None,
) -> str:

    if TOKENIZER is None:

        raise RuntimeError(
            "Tokenizer is not loaded."
        )

    if MODEL is None:

        raise RuntimeError(
            "Model is not loaded."
        )

    messages = [
        {
            "role": "system",
            "content": (
                "You are a precise financial "
                "research analyst. "
                "Return only valid JSON."
            ),
        },
        {
            "role": "user",
            "content": prompt,
        },
    ]

    try:

        inputs = (
            TOKENIZER.apply_chat_template(
                messages,
                tokenize=True,
                add_generation_prompt=True,
                enable_thinking=False,
                return_dict=True,
                return_tensors="pt",
            )
        )

    except TypeError:

        inputs = (
            TOKENIZER.apply_chat_template(
                messages,
                tokenize=True,
                add_generation_prompt=True,
                return_dict=True,
                return_tensors="pt",
            )
        )

    try:

        inputs = inputs.to(
            MODEL.device
        )

    except Exception:
        pass

    input_tokens = (
        inputs[
            "input_ids"
        ].shape[-1]
    )

    log(
        f"        Input tokens: "
        f"{input_tokens}"
    )

    generation_tokens = (
        max_new_tokens
        if max_new_tokens is not None
        else MAX_NEW_TOKENS
    )

    with torch.inference_mode():

        generated = MODEL.generate(
            **inputs,
            max_new_tokens=(
                generation_tokens
            ),
            do_sample=False,
            pad_token_id=(
                TOKENIZER.eos_token_id
            ),
        )

    generated_tokens = (
        generated[
            0,
            input_tokens:
        ]
    )

    response = TOKENIZER.decode(
        generated_tokens,
        skip_special_tokens=True,
    )

    return response.strip()


# ============================================================
# JSON PARSING
# ============================================================

def extract_json_object(
    response: str,
) -> Optional[
    Dict[str, Any]
]:

    if not response:
        return None

    cleaned = response.strip()

    cleaned = re.sub(
        r"^```(?:json)?\s*",
        "",
        cleaned,
        flags=re.IGNORECASE,
    )

    cleaned = re.sub(
        r"\s*```$",
        "",
        cleaned,
    )

    try:

        parsed = json.loads(
            cleaned
        )

        if isinstance(
            parsed,
            dict,
        ):

            return parsed

    except json.JSONDecodeError:
        pass

    # Find the first complete-looking JSON object.
    match = re.search(
        r"\{.*\}",
        cleaned,
        flags=re.DOTALL,
    )

    if not match:
        return None

    try:

        parsed = json.loads(
            match.group(0)
        )

        if isinstance(
            parsed,
            dict,
        ):

            return parsed

    except json.JSONDecodeError:
        return None

    return None


# ============================================================
# AI STAGE 1
# STRUCTURE DETECTION
# ============================================================

def build_structure_prompt(
    fund_name: str,
    source_text: str,
) -> str:

    return f"""
Analyze this investment-fund document.

Fund name:
{fund_name}

Your task is ONLY to determine the fund structure.

Determine:

1. Is this a fund-of-funds or multi-fund structure?
2. If yes, identify the underlying fund names explicitly
   mentioned in the supplied document.
3. If no, return an empty underlyingFunds list.
4. Determine whether the supplied document contains direct
   portfolio/holdings/allocation evidence for the fund itself.

IMPORTANT:

- Use ONLY the supplied document.
- Do NOT use outside knowledge.
- Do NOT infer underlying funds from the fund name.
- Do NOT invent fund names.
- Only include an underlying fund if the document explicitly
  identifies it.
- Preserve the underlying fund name as written.
- Maximum {MAX_UNDERLYING_FUNDS} underlying funds.
- Return ONLY JSON.
- No markdown.
- No explanation.

Required JSON:

{{
  "structure": "direct",
  "underlyingFunds": [],
  "hasDirectPortfolioEvidence": false
}}

Allowed structure values:

- "direct"
- "fund_of_funds"
- "mixed"

DOCUMENT:

{source_text}
""".strip()


def validate_structure_result(
    parsed: Dict[str, Any],
) -> Tuple[
    bool,
    str,
    str,
    List[str],
    bool,
]:

    if not isinstance(
        parsed,
        dict,
    ):

        return (
            False,
            "Structure result is not an object.",
            "",
            [],
            False,
        )

    structure = parsed.get(
        "structure"
    )

    underlying = parsed.get(
        "underlyingFunds"
    )

    direct_evidence = parsed.get(
        "hasDirectPortfolioEvidence"
    )

    if structure not in {
        "direct",
        "fund_of_funds",
        "mixed",
    }:

        return (
            False,
            "Invalid structure value.",
            "",
            [],
            False,
        )

    if not isinstance(
        underlying,
        list,
    ):

        return (
            False,
            "underlyingFunds is not a list.",
            "",
            [],
            False,
        )

    if len(underlying) > (
        MAX_UNDERLYING_FUNDS
    ):

        return (
            False,
            "Too many underlying funds.",
            "",
            [],
            False,
        )

    clean_underlying: List[
        str
    ] = []

    for item in underlying:

        if not isinstance(
            item,
            str,
        ):

            return (
                False,
                "Underlying fund is not a string.",
                "",
                [],
                False,
            )

        item = normalize_space(
            item
        )

        if not item:
            continue

        if item not in clean_underlying:

            clean_underlying.append(
                item
            )

    if not isinstance(
        direct_evidence,
        bool,
    ):

        return (
            False,
            "hasDirectPortfolioEvidence "
            "must be boolean.",
            "",
            [],
            False,
        )

    # If the AI says fund_of_funds but gives no names,
    # treat it as a structure-analysis failure rather than
    # allowing a false classification.
    if (
        structure
        in {
            "fund_of_funds",
            "mixed",
        }
        and not clean_underlying
        and not direct_evidence
    ):

        return (
            False,
            "Fund structure requires underlying "
            "funds or direct portfolio evidence.",
            "",
            [],
            False,
        )

    return (
        True,
        "",
        structure,
        clean_underlying,
        direct_evidence,
    )


# ============================================================
# AI STAGE 2
# EXPOSURE CLASSIFICATION
# ============================================================

def build_exposure_prompt(
    subject_name: str,
    source_text: str,
    geography_categories: List[str],
    sector_categories: List[str],
    context_label: str,
) -> str:

    geography_text = "\n".join(
        f"- {item}"
        for item in geography_categories
    )

    sector_text = "\n".join(
        f"- {item}"
        for item in sector_categories
    )

    return f"""
You are analyzing actual investment exposure for a fund.

Subject:
{subject_name}

Research context:
{context_label}

Determine the supported:

1. Geographic exposure
2. Sector exposure

EVIDENCE RULES:

- Use ONLY the supplied document.
- Do NOT use outside knowledge.
- Do NOT use the fund name as evidence.
- Do NOT assume that the fund's domicile is its investment geography.
- Do NOT infer a sector from the fund name.
- Do NOT invent holdings.
- Do NOT invent countries or sectors.
- Use explicit portfolio evidence where available.
- Accept evidence from:
  * geographical allocation
  * country allocation
  * regional allocation
  * sector allocation
  * industry allocation
  * top holdings
  * portfolio holdings
  * asset allocation
  * investment objective
  * investment strategy
  * underlying-fund description
  * portfolio manager description
- If an underlying fund is explicitly described as investing in
  a particular region or sector, that documented description
  may be used.
- Select AT MOST 3 geography categories.
- Select AT MOST 3 sector categories.
- Return fewer than 3 when evidence supports fewer.
- Percentages must NOT be returned.
- Category names must match the canonical vocabulary EXACTLY.
- Do not create category names.
- "Global" may only be used when the supplied document itself
  supports broad global exposure.
- "Other" should only be used when the source clearly supports
  an exposure that cannot reasonably map to another supplied
  canonical category.
- If there is no supported geography, return [].
- If there is no supported sector, return [].
- Return ONLY JSON.
- No markdown.
- No explanation.

CANONICAL GEOGRAPHY:

{geography_text}

CANONICAL SECTOR:

{sector_text}

OUTPUT:

{{
  "geography": [],
  "sector": []
}}

DOCUMENT:

{source_text}
""".strip()


def validate_exposure_result(
    parsed: Dict[str, Any],
    geography_categories: List[str],
    sector_categories: List[str],
) -> Tuple[
    bool,
    str,
    List[str],
    List[str],
]:

    if not isinstance(
        parsed,
        dict,
    ):

        return (
            False,
            "Exposure result is not an object.",
            [],
            [],
        )

    geography = parsed.get(
        "geography"
    )

    sector = parsed.get(
        "sector"
    )

    if not isinstance(
        geography,
        list,
    ):

        return (
            False,
            "AI geography is not a list.",
            [],
            [],
        )

    if not isinstance(
        sector,
        list,
    ):

        return (
            False,
            "AI sector is not a list.",
            [],
            [],
        )

    if len(geography) > 3:

        return (
            False,
            "More than 3 geography categories.",
            [],
            [],
        )

    if len(sector) > 3:

        return (
            False,
            "More than 3 sector categories.",
            [],
            [],
        )

    geography_set = set(
        geography_categories
    )

    sector_set = set(
        sector_categories
    )

    clean_geography: List[
        str
    ] = []

    clean_sector: List[
        str
    ] = []

    for category in geography:

        if not isinstance(
            category,
            str,
        ):

            return (
                False,
                "Geography contains non-string value.",
                [],
                [],
            )

        category = normalize_space(
            category
        )

        if not category:
            continue

        if "%" in category:

            return (
                False,
                "Percentage found in geography.",
                [],
                [],
            )

        if category not in geography_set:

            return (
                False,
                f"Invalid geography category: "
                f"{category}",
                [],
                [],
            )

        if category not in clean_geography:

            clean_geography.append(
                category
            )

    for category in sector:

        if not isinstance(
            category,
            str,
        ):

            return (
                False,
                "Sector contains non-string value.",
                [],
                [],
            )

        category = normalize_space(
            category
        )

        if not category:
            continue

        if "%" in category:

            return (
                False,
                "Percentage found in sector.",
                [],
                [],
            )

        if category not in sector_set:

            return (
                False,
                f"Invalid sector category: "
                f"{category}",
                [],
                [],
            )

        if category not in clean_sector:

            clean_sector.append(
                category
            )

    # An individual source is allowed to have one dimension
    # but not the other.
    #
    # Both empty means that this particular source did not
    # provide usable evidence.
    if (
        not clean_geography
        and not clean_sector
    ):

        return (
            False,
            "AI found no valid geography or sector exposure.",
            [],
            [],
        )

    return (
        True,
        "",
        clean_geography,
        clean_sector,
    )


# ============================================================
# AI CLASSIFICATION HELPERS
# ============================================================

def classify_source(
    subject_name: str,
    source_text: str,
    geography_categories: List[str],
    sector_categories: List[str],
    context_label: str,
) -> Dict[str, Any]:

    ai_source = prepare_ai_source(
        source_text
    )

    prompt = build_exposure_prompt(
        subject_name=subject_name,
        source_text=ai_source,
        geography_categories=(
            geography_categories
        ),
        sector_categories=(
            sector_categories
        ),
        context_label=context_label,
    )

    try:

        response = generate_ai_response(
            prompt
        )

    except Exception as exc:

        return {
            "status": "ai_error",
            "error": str(exc),
            "geography": [],
            "sector": [],
        }

    log("")
    log(
        "        Exposure AI response:"
    )

    log(
        response
    )

    parsed = extract_json_object(
        response
    )

    if parsed is None:

        return {
            "status": "ai_error",
            "error": (
                "AI response could not "
                "be parsed as JSON."
            ),
            "geography": [],
            "sector": [],
        }

    (
        valid,
        error,
        geography,
        sector,
    ) = validate_exposure_result(
        parsed,
        geography_categories,
        sector_categories,
    )

    if not valid:

        return {
            "status": "ai_error",
            "error": error,
            "geography": geography,
            "sector": sector,
        }

    return {
        "status": "success",
        "error": "",
        "geography": geography,
        "sector": sector,
    }


# ============================================================
# STRUCTURE ANALYSIS
# ============================================================

def analyze_structure(
    fund_name: str,
    source_text: str,
) -> Dict[str, Any]:

    ai_source = prepare_ai_source(
        source_text
    )

    prompt = build_structure_prompt(
        fund_name=fund_name,
        source_text=ai_source,
    )

    try:

        response = generate_ai_response(
            prompt,
            max_new_tokens=500,
        )

    except Exception as exc:

        return {
            "status": "ai_error",
            "error": str(exc),
            "structure": "",
            "underlyingFunds": [],
            "hasDirectPortfolioEvidence": False,
        }

    log("")
    log(
        "      Structure AI response:"
    )

    log(
        response
    )

    parsed = extract_json_object(
        response
    )

    if parsed is None:

        return {
            "status": "ai_error",
            "error": (
                "Structure AI response "
                "was not valid JSON."
            ),
            "structure": "",
            "underlyingFunds": [],
            "hasDirectPortfolioEvidence": False,
        }

    (
        valid,
        error,
        structure,
        underlying,
        direct_evidence,
    ) = validate_structure_result(
        parsed
    )

    if not valid:

        return {
            "status": "ai_error",
            "error": error,
            "structure": structure,
            "underlyingFunds": underlying,
            "hasDirectPortfolioEvidence": (
                direct_evidence
            ),
        }

    return {
        "status": "success",
        "error": "",
        "structure": structure,
        "underlyingFunds": underlying,
        "hasDirectPortfolioEvidence": (
            direct_evidence
        ),
    }


# ============================================================
# AGGREGATION
# ============================================================

def aggregate_exposures(
    classifications: List[
        Dict[str, Any]
    ],
    geography_categories: List[str],
    sector_categories: List[str],
) -> Tuple[
    List[str],
    List[str],
]:

    geography_order = {
        category: index
        for index, category
        in enumerate(
            geography_categories
        )
    }

    sector_order = {
        category: index
        for index, category
        in enumerate(
            sector_categories
        )
    }

    geography_score: Dict[
        str,
        float,
    ] = {}

    sector_score: Dict[
        str,
        float,
    ] = {}

    for classification in (
        classifications
    ):

        if (
            classification.get(
                "status"
            )
            != "success"
        ):
            continue

        # Direct classification and underlying classification
        # each provide one evidence unit.
        #
        # A category mentioned by multiple independent sources
        # receives a higher score.
        #
        # We intentionally do NOT use percentages because the
        # requested output is categorical.
        for category in classification.get(
            "geography",
            [],
        ):

            geography_score[
                category
            ] = (
                geography_score.get(
                    category,
                    0.0,
                )
                + 1.0
            )

        for category in classification.get(
            "sector",
            [],
        ):

            sector_score[
                category
            ] = (
                sector_score.get(
                    category,
                    0.0,
                )
                + 1.0
            )

    ranked_geography = sorted(
        geography_score.keys(),
        key=lambda category: (
            -geography_score[
                category
            ],
            geography_order.get(
                category,
                9999,
            ),
        ),
    )

    ranked_sector = sorted(
        sector_score.keys(),
        key=lambda category: (
            -sector_score[
                category
            ],
            sector_order.get(
                category,
                9999,
            ),
        ),
    )

    return (
        ranked_geography[:3],
        ranked_sector[:3],
    )


# ============================================================
# FUND PROCESSING
# ============================================================

def process_fund(
    fund: Dict[str, str],
    geography_categories: List[str],
    sector_categories: List[str],
    index: int,
    total: int,
) -> Dict[str, Any]:

    fund_name = fund[
        "fundName"
    ]

    prudential_url = fund[
        "prudentialUrl"
    ]

    log("")
    log(
        "=" * 70
    )

    log(
        f"[{index}/{total}] "
        f"{fund_name}"
    )

    log(
        "=" * 70
    )

    # --------------------------------------------------------
    # SOURCE COLLECTION
    # --------------------------------------------------------

    try:

        (
            documents,
            discovered_links,
        ) = collect_prudential_sources(
            prudential_url
        )

    except Exception as exc:

        log(
            f"        Source failed: "
            f"{exc}"
        )

        return {
            "fundName": fund_name,
            "prudentialUrl": prudential_url,
            "researchMode": "unresolved",
            "underlyingFunds": [],
            "geography": [],
            "sector": [],
            "status": "source_error",
            "error": str(exc),
            "sources": [
                prudential_url
            ],
        }

    if not documents:

        return {
            "fundName": fund_name,
            "prudentialUrl": prudential_url,
            "researchMode": "unresolved",
            "underlyingFunds": [],
            "geography": [],
            "sector": [],
            "status": "source_error",
            "error": (
                "No usable source documents."
            ),
            "sources": [
                prudential_url
            ],
        }

    log("")
    log(
        f"        Documents collected: "
        f"{len(documents)}"
    )

    for document in documents:

        log(
            f"          - "
            f"{document.source_type}: "
            f"{document.url}"
        )

    # --------------------------------------------------------
    # BUILD COMBINED STRUCTURE SOURCE
    # --------------------------------------------------------

    combined_source_parts: List[
        str
    ] = []

    for document in documents:

        prepared = prepare_ai_source(
            document.text
        )

        if not prepared:
            continue

        combined_source_parts.append(
            "\n".join(
                [
                    (
                        "--- DOCUMENT ---"
                    ),
                    (
                        f"URL: "
                        f"{document.url}"
                    ),
                    (
                        f"TITLE: "
                        f"{document.title}"
                    ),
                    prepared,
                ]
            )
        )

    combined_source = truncate_text(
        clean_text(
            "\n\n".join(
                combined_source_parts
            )
        ),
        MAX_AI_SOURCE_CHARS,
    )

    if not combined_source:

        return {
            "fundName": fund_name,
            "prudentialUrl": prudential_url,
            "researchMode": "unresolved",
            "underlyingFunds": [],
            "geography": [],
            "sector": [],
            "status": "source_error",
            "error": (
                "No usable text after preprocessing."
            ),
            "sources": [
                document.url
                for document in documents
            ],
        }

    # --------------------------------------------------------
    # DIAGNOSTIC PREVIEW
    # --------------------------------------------------------

    log("")
    log(
        "        SOURCE PREVIEW:"
    )

    log(
        "        "
        + "-" * 60
    )

    preview = combined_source[
        :SOURCE_PREVIEW_CHARS
    ]

    for line in preview.splitlines():

        log(
            f"        {line}"
        )

    log(
        "        "
        + "-" * 60
    )

    # --------------------------------------------------------
    # STAGE 1 - STRUCTURE
    # --------------------------------------------------------

    log("")
    log(
        "      Stage 1 - "
        "detecting fund structure..."
    )

    structure_result = analyze_structure(
        fund_name=fund_name,
        source_text=combined_source,
    )

    if (
        structure_result.get(
            "status"
        )
        != "success"
    ):

        error = structure_result.get(
            "error",
            "Unknown structure AI error.",
        )

        log(
            f"        Structure analysis failed: "
            f"{error}"
        )

        return {
            "fundName": fund_name,
            "prudentialUrl": prudential_url,
            "researchMode": "unresolved",
            "underlyingFunds": [],
            "geography": [],
            "sector": [],
            "status": "ai_error",
            "error": error,
            "sources": [
                document.url
                for document in documents
            ],
        }

    structure = structure_result[
        "structure"
    ]

    underlying_funds = (
        structure_result[
            "underlyingFunds"
        ]
    )

    has_direct_evidence = (
        structure_result[
            "hasDirectPortfolioEvidence"
        ]
    )

    log("")
    log(
        f"        Detected structure: "
        f"{structure}"
    )

    log(
        f"        Underlying funds: "
        f"{underlying_funds}"
    )

    log(
        f"        Direct portfolio evidence: "
        f"{has_direct_evidence}"
    )

    # --------------------------------------------------------
    # STAGE 2 - DIRECT FUND CLASSIFICATION
    # --------------------------------------------------------

    classifications: List[
        Dict[str, Any]
    ] = []

    source_records: List[
        Dict[str, Any]
    ] = []

    if (
        structure
        in {
            "direct",
            "mixed",
        }
        or (
            not underlying_funds
            and has_direct_evidence
        )
    ):

        log("")
        log(
            "      Stage 2 - "
            "researching direct portfolio exposure..."
        )

        # Prefer the strongest documents first.
        direct_documents = sorted(
            documents,
            key=lambda document: (
                0
                if document.source_type
                == "pdf"
                else 1
            )
        )

        direct_documents = (
            direct_documents[
                :MAX_CLASSIFICATION_DOCUMENTS
            ]
        )

        direct_successes = 0

        for document in (
            direct_documents
        ):

            prepared = (
                prepare_ai_source(
                    document.text
                )
            )

            if not prepared:
                continue

            classification = (
                classify_source(
                    subject_name=fund_name,
                    source_text=prepared,
                    geography_categories=(
                        geography_categories
                    ),
                    sector_categories=(
                        sector_categories
                    ),
                    context_label=(
                        "Direct fund portfolio "
                        "research"
                    ),
                )
            )

            source_records.append(
                {
                    "subject": fund_name,
                    "sourceUrl": document.url,
                    "sourceType": document.source_type,
                    "classification": classification,
                }
            )

            if (
                classification.get(
                    "status"
                )
                == "success"
            ):

                classifications.append(
                    classification
                )

                direct_successes += 1

        log(
            f"        Direct successful "
            f"classifications: "
            f"{direct_successes}"
        )

    # --------------------------------------------------------
    # STAGE 3 - UNDERLYING FUND RESEARCH
    # --------------------------------------------------------

    underlying_records: List[
        Dict[str, Any]
    ] = []

    if underlying_funds:

        log("")
        log(
            "      Stage 3 - "
            "researching underlying funds..."
        )

        for underlying_name in (
            underlying_funds[
                :MAX_UNDERLYING_FUNDS
            ]
        ):

            log("")
            log(
                f"        Underlying fund: "
                f"{underlying_name}"
            )

            ranked_documents = (
                rank_documents_for_underlying_fund(
                    documents,
                    underlying_name,
                )
            )

            # If we cannot match a specific document,
            # still allow the AI to inspect the combined
            # Prudential source because the underlying fund
            # may be described inside the root page.
            selected_documents = (
                ranked_documents[
                    :3
                ]
            )

            if selected_documents:

                log(
                    f"        Matched documents: "
                    f"{len(selected_documents)}"
                )

            else:

                log(
                    "        No dedicated "
                    "underlying document matched; "
                    "using Prudential source."
                )

                selected_documents = []

            underlying_text_parts: List[
                str
            ] = []

            if selected_documents:

                for document in (
                    selected_documents
                ):

                    prepared = (
                        prepare_ai_source(
                            document.text
                        )
                    )

                    if prepared:

                        underlying_text_parts.append(
                            "\n".join(
                                [
                                    (
                                        "--- UNDERLYING "
                                        "FUND DOCUMENT ---"
                                    ),
                                    (
                                        f"URL: "
                                        f"{document.url}"
                                    ),
                                    prepared,
                                ]
                            )
                        )

            else:

                underlying_text_parts.append(
                    "\n".join(
                        [
                            (
                                "--- PRUDENTIAL "
                                "SOURCE ---"
                            ),
                            combined_source,
                        ]
                    )
                )

            underlying_source = truncate_text(
                clean_text(
                    "\n\n".join(
                        underlying_text_parts
                    )
                ),
                MAX_AI_SOURCE_CHARS,
            )

            if not underlying_source:

                log(
                    "        No usable source "
                    "for underlying fund."
                )

                continue

            classification = (
                classify_source(
                    subject_name=(
                        underlying_name
                    ),
                    source_text=(
                        underlying_source
                    ),
                    geography_categories=(
                        geography_categories
                    ),
                    sector_categories=(
                        sector_categories
                    ),
                    context_label=(
                        "Underlying fund research "
                        "for a Prudential "
                        "fund-of-funds structure"
                    ),
                )
            )

            underlying_record = {
                "fundName": underlying_name,
                "matchedSources": [
                    document.url
                    for document
                    in selected_documents
                ],
                "classification": (
                    classification
                ),
            }

            underlying_records.append(
                underlying_record
            )

            if (
                classification.get(
                    "status"
                )
                == "success"
            ):

                classifications.append(
                    classification
                )

                log(
                    f"        Underlying "
                    f"geography: "
                    f"{classification.get('geography')}"
                )

                log(
                    f"        Underlying "
                    f"sector: "
                    f"{classification.get('sector')}"
                )

            else:

                log(
                    "        Underlying "
                    "classification produced "
                    "no usable exposure."
                )

    # --------------------------------------------------------
    # AGGREGATE
    # --------------------------------------------------------

    geography, sector = (
        aggregate_exposures(
            classifications=(
                classifications
            ),
            geography_categories=(
                geography_categories
            ),
            sector_categories=(
                sector_categories
            ),
        )
    )

    log("")
    log(
        "      Aggregated exposure:"
    )

    log(
        f"        Geography: "
        f"{geography}"
    )

    log(
        f"        Sector: "
        f"{sector}"
    )

    # --------------------------------------------------------
    # FINAL FUND-LEVEL SUCCESS
    # --------------------------------------------------------

    if (
        not geography
        and not sector
    ):

        log(
            "        Exposure analysis failed: "
            "no supported geography or sector."
        )

        return {
            "fundName": fund_name,
            "prudentialUrl": prudential_url,
            "researchMode": structure,
            "underlyingFunds": (
                underlying_funds
            ),
            "geography": [],
            "sector": [],
            "status": "ai_error",
            "error": (
                "No geography or sector exposure "
                "could be supported from the "
                "available source documents."
            ),
            "sources": [
                document.url
                for document in documents
            ],
            "underlyingResearch": (
                underlying_records
            ),
        }

    # --------------------------------------------------------
    # SUCCESS
    # --------------------------------------------------------

    return {
        "fundName": fund_name,
        "prudentialUrl": prudential_url,
        "researchMode": structure,
        "underlyingFunds": (
            underlying_funds
        ),
        "geography": geography,
        "sector": sector,
        "status": "success",
        "error": "",
        "sources": [
            document.url
            for document in documents
        ],
        "discoveredLinks": [
            {
                "url": item.get(
                    "url",
                    "",
                ),
                "text": item.get(
                    "text",
                    "",
                ),
            }
            for item in discovered_links[
                :MAX_DISCOVERY_DOCUMENTS
            ]
        ],
        "underlyingResearch": (
            underlying_records
        ),
        "sourceClassifications": (
            source_records
        ),
    }


# ============================================================
# AI HEALTH CHECK
# ============================================================

def ai_health_check(
    geography_categories: List[str],
    sector_categories: List[str],
) -> bool:

    log("")
    log(
        "=" * 70
    )
    log(
        "AI HEALTH CHECK"
    )
    log(
        "=" * 70
    )

    test_prompt = f"""
Return ONLY this JSON object:

{{
  "status": "ok"
}}

Do not add any other text.
""".strip()

    try:

        response = generate_ai_response(
            test_prompt,
            max_new_tokens=100,
        )

        log(
            f"Health-check response: "
            f"{response}"
        )

        parsed = extract_json_object(
            response
        )

        if parsed is None:

            log(
                "ERROR: Health check "
                "returned invalid JSON."
            )

            return False

        if parsed.get(
            "status"
        ) != "ok":

            log(
                "ERROR: Health check "
                "returned unexpected JSON."
            )

            return False

        log(
            "AI health check PASSED."
        )

        return True

    except Exception as exc:

        log(
            f"ERROR: AI health check "
            f"failed: {exc}"
        )

        return False


# ============================================================
# COMPLETE OUTPUT VALIDATION
# ============================================================

def validate_complete_result(
    output: Dict[str, Any],
    expected_count: int,
    geography_categories: List[str],
    sector_categories: List[str],
) -> Tuple[
    bool,
    List[str],
]:

    errors: List[str] = []

    funds = output.get(
        "funds"
    )

    if not isinstance(
        funds,
        list,
    ):

        return (
            False,
            [
                "Output funds is not a list."
            ],
        )

    if len(funds) != expected_count:

        errors.append(
            f"Expected {expected_count} funds, "
            f"got {len(funds)}."
        )

    geography_set = set(
        geography_categories
    )

    sector_set = set(
        sector_categories
    )

    seen_fund_names: Set[
        str
    ] = set()

    for index, fund in enumerate(
        funds,
        start=1,
    ):

        if not isinstance(
            fund,
            dict,
        ):

            errors.append(
                f"Fund #{index} is not an object."
            )

            continue

        name = fund.get(
            "fundName",
            f"Fund #{index}",
        )

        if name in seen_fund_names:

            errors.append(
                f"{name}: duplicate fund."
            )

        seen_fund_names.add(
            name
        )

        status = fund.get(
            "status"
        )

        if status != "success":

            errors.append(
                f"{name}: "
                f"status={status!r}"
            )

        geography = fund.get(
            "geography",
            [],
        )

        sector = fund.get(
            "sector",
            [],
        )

        if not isinstance(
            geography,
            list,
        ):

            errors.append(
                f"{name}: geography "
                "is not a list."
            )

            continue

        if not isinstance(
            sector,
            list,
        ):

            errors.append(
                f"{name}: sector "
                "is not a list."
            )

            continue

        # Successful fund cannot have both empty.
        if (
            status == "success"
            and not geography
            and not sector
        ):

            errors.append(
                f"{name}: no geography "
                "or sector exposure."
            )

        if len(geography) > 3:

            errors.append(
                f"{name}: more than "
                "3 geography categories."
            )

        if len(sector) > 3:

            errors.append(
                f"{name}: more than "
                "3 sector categories."
            )

        for category in geography:

            if not isinstance(
                category,
                str,
            ):

                errors.append(
                    f"{name}: invalid "
                    "geography type."
                )

                continue

            if "%" in category:

                errors.append(
                    f"{name}: percentage "
                    "in geography."
                )

            if category not in geography_set:

                errors.append(
                    f"{name}: invalid "
                    f"geography {category!r}."
                )

        for category in sector:

            if not isinstance(
                category,
                str,
            ):

                errors.append(
                    f"{name}: invalid "
                    "sector type."
                )

                continue

            if "%" in category:

                errors.append(
                    f"{name}: percentage "
                    "in sector."
                )

            if category not in sector_set:

                errors.append(
                    f"{name}: invalid "
                    f"sector {category!r}."
                )

    return (
        len(errors) == 0,
        errors,
    )


# ============================================================
# ATOMIC PUBLISH
# ============================================================

def publish_output_atomically(
    output: Dict[str, Any],
    output_path: Path,
) -> None:

    output_path.parent.mkdir(
        parents=True,
        exist_ok=True,
    )

    temporary_path = (
        output_path.with_suffix(
            output_path.suffix
            + ".tmp"
        )
    )

    try:

        with temporary_path.open(
            "w",
            encoding="utf-8",
        ) as file:

            json.dump(
                output,
                file,
                indent=2,
                ensure_ascii=False,
            )

            file.write(
                "\n"
            )

            file.flush()

            os.fsync(
                file.fileno()
            )

        # Verify JSON after writing.
        with temporary_path.open(
            "r",
            encoding="utf-8",
        ) as file:

            verified = json.load(
                file
            )

        if not isinstance(
            verified,
            dict,
        ):

            raise RuntimeError(
                "Temporary output "
                "is not an object."
            )

        if not isinstance(
            verified.get(
                "funds"
            ),
            list,
        ):

            raise RuntimeError(
                "Temporary output funds "
                "is not a list."
            )

        # Atomic replacement.
        os.replace(
            temporary_path,
            output_path,
        )

    except Exception:

        try:

            if temporary_path.exists():

                temporary_path.unlink()

        except Exception:
            pass

        raise


# ============================================================
# ARGUMENTS
# ============================================================

def parse_args() -> argparse.Namespace:

    parser = argparse.ArgumentParser(
        description=(
            "VGrat FMS AI fund geography "
            "and sector research."
        )
    )

    parser.add_argument(
        "--limit",
        type=int,
        default=None,
        help=(
            "Process only first N funds."
        ),
    )

    return parser.parse_args()


# ============================================================
# MAIN
# ============================================================

def main() -> int:

    args = parse_args()

    log("")
    log(
        "=" * 70
    )

    log(
        "VGRAT FMS - AI FUND EXPOSURE RESEARCH"
    )

    log(
        "=" * 70
    )

    log(
        f"Model: {MODEL_NAME}"
    )

    log(
        "Search engine: DISABLED"
    )

    log(
        "Analysis: LOCAL AI ONLY"
    )

    log(
        "Structure: DIRECT + FUND-OF-FUNDS"
    )

    log(
        f"Output: {OUTPUT_PATH}"
    )

    log(
        "=" * 70
    )

    # --------------------------------------------------------
    # Categories
    # --------------------------------------------------------

    try:

        (
            geography_categories,
            sector_categories,
        ) = load_categories(
            CATEGORIES_FILE
        )

    except Exception as exc:

        log(
            f"ERROR: Category loading failed: "
            f"{exc}"
        )

        log(
            "Existing fund_exposure.json "
            "has NOT been modified."
        )

        return 1

    log(
        f"Geography categories: "
        f"{len(geography_categories)}"
    )

    log(
        f"Sector categories: "
        f"{len(sector_categories)}"
    )

    # --------------------------------------------------------
    # Workbook
    # --------------------------------------------------------

    try:

        workbook_path = (
            find_funds_excel_file(
                REPOSITORY_ROOT
            )
        )

        funds = load_funds(
            workbook_path
        )

    except Exception as exc:

        log(
            f"ERROR: Fund loading failed: "
            f"{exc}"
        )

        log(
            "Existing fund_exposure.json "
            "has NOT been modified."
        )

        return 1

    # --------------------------------------------------------
    # Limit
    # --------------------------------------------------------

    if args.limit is not None:

        if args.limit <= 0:

            log(
                "ERROR: --limit must be > 0."
            )

            return 1

        funds = funds[
            :args.limit
        ]

        log(
            f"TEST LIMIT: "
            f"{len(funds)} funds"
        )

    total = len(
        funds
    )

    if total == 0:

        log(
            "ERROR: No funds selected."
        )

        return 1

    # --------------------------------------------------------
    # Model
    # --------------------------------------------------------

    try:

        load_ai_model()

    except Exception as exc:

        log(
            f"ERROR: Model loading failed: "
            f"{exc}"
        )

        log(
            "Existing fund_exposure.json "
            "has NOT been modified."
        )

        return 1

    # --------------------------------------------------------
    # Health check
    # --------------------------------------------------------

    if not ai_health_check(
        geography_categories,
        sector_categories,
    ):

        log(
            "AI health check FAILED."
        )

        log(
            "Existing fund_exposure.json "
            "has NOT been modified."
        )

        return 1

    # --------------------------------------------------------
    # Research
    # --------------------------------------------------------

    results: List[
        Dict[str, Any]
    ] = []

    successful = 0
    failed = 0

    for index, fund in enumerate(
        funds,
        start=1,
    ):

        result = process_fund(
            fund=fund,
            geography_categories=(
                geography_categories
            ),
            sector_categories=(
                sector_categories
            ),
            index=index,
            total=total,
        )

        results.append(
            result
        )

        if (
            result.get(
                "status"
            )
            == "success"
        ):

            successful += 1

        else:

            failed += 1

    # --------------------------------------------------------
    # Build result in memory.
    # --------------------------------------------------------

    output: Dict[str, Any] = {
        "generatedAt": utc_now(),
        "model": MODEL_NAME,
        "analysis": (
            "Local AI classification using "
            "Prudential source documents and "
            "underlying-fund documents where "
            "available."
        ),
        "fundCount": total,
        "funds": results,
        "summary": {
            "totalFunds": total,
            "successful": successful,
            "failed": failed,
        },
    }

    # --------------------------------------------------------
    # FINAL VALIDATION
    # --------------------------------------------------------

    log("")
    log(
        "=" * 70
    )

    log(
        "FINAL VALIDATION"
    )

    log(
        "=" * 70
    )

    log(
        f"Total funds: {total}"
    )

    log(
        f"Successful: {successful}"
    )

    log(
        f"Failed: {failed}"
    )

    valid, errors = (
        validate_complete_result(
            output=output,
            expected_count=total,
            geography_categories=(
                geography_categories
            ),
            sector_categories=(
                sector_categories
            ),
        )
    )

    if not valid:

        log("")
        log(
            "FINAL VALIDATION FAILED."
        )

        log(
            "The existing fund_exposure.json "
            "will NOT be replaced."
        )

        log("")
        log(
            "Failure details:"
        )

        for error in errors:

            log(
                f"  - {error}"
            )

        return 1

    # --------------------------------------------------------
    # Atomic publish
    # --------------------------------------------------------

    try:

        publish_output_atomically(
            output,
            OUTPUT_PATH,
        )

    except Exception as exc:

        log("")
        log(
            f"ERROR: Output publishing failed: "
            f"{exc}"
        )

        log(
            "Existing fund_exposure.json "
            "was preserved."
        )

        return 1

    # --------------------------------------------------------
    # SUCCESS
    # --------------------------------------------------------

    log("")
    log(
        "=" * 70
    )

    log(
        "RESEARCH SUCCESS"
    )

    log(
        "=" * 70
    )

    log(
        f"Successfully researched: "
        f"{successful}/{total}"
    )

    log(
        f"Published: {OUTPUT_PATH}"
    )

    log(
        "The previous output was replaced "
        "because the complete run succeeded."
    )

    log(
        "=" * 70
    )

    return 0


if __name__ == "__main__":

    sys.exit(
        main()
    )
