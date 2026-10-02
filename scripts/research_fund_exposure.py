#!/usr/bin/env python3

"""
VGrat FMS - AI FUND EXPOSURE RESEARCH

AI-only geography and sector research.

Source:
    Prudential URL supplied in Funds Links.xlsx / Funds Links.xlsm

Analysis:
    Local Hugging Face Qwen model

Search engines:
    NONE

Publishing rule:
    - Complete successful run -> replace fund_exposure.json
    - Any failed fund -> preserve existing fund_exposure.json
    - Any AI/model/validation failure -> preserve existing file
    - Partial results are NEVER published
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
from typing import Any, Dict, List, Optional, Tuple
from urllib.parse import urljoin

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
        "40",
    )
)

MAX_SOURCE_CHARS = int(
    os.getenv(
        "MAX_SOURCE_CHARS",
        "60000",
    )
)

MAX_AI_SOURCE_CHARS = int(
    os.getenv(
        "MAX_AI_SOURCE_CHARS",
        "30000",
    )
)

MAX_NEW_TOKENS = int(
    os.getenv(
        "MAX_NEW_TOKENS",
        "800",
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

    text = re.sub(
        r"[ \t]+",
        " ",
        text,
    )

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

        funds: List[Dict[str, str]] = []

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
) -> Tuple[List[str], List[str]]:

    if not categories_path.exists():
        raise FileNotFoundError(
            f"Category file not found: "
            f"{categories_path}"
        )

    geography: List[str] = []
    sector: List[str] = []

    current_section: Optional[str] = None

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
                current_section = "geography"
                continue

            if upper == "[SECTOR]":
                current_section = "sector"
                continue

            if current_section == "geography":
                geography.append(line)

            elif current_section == "sector":
                sector.append(line)

    geography = list(
        dict.fromkeys(geography)
    )

    sector = list(
        dict.fromkeys(sector)
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
        "\n".join(chunks)
    )


# ============================================================
# HTML EXTRACTION
# ============================================================

def extract_html(
    html: str,
) -> Tuple[str, List[str]]:

    soup = BeautifulSoup(
        html,
        "html.parser",
    )

    for tag in soup(
        [
            "script",
            "style",
            "noscript",
            "svg",
        ]
    ):
        tag.decompose()

    links: List[str] = []

    for anchor in soup.find_all("a"):

        href = anchor.get(
            "href"
        )

        if not href:
            continue

        href = href.strip()

        links.append(href)

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
# PRUDENTIAL SOURCE COLLECTION
# ============================================================

def fetch_prudential_document(
    url: str,
) -> Tuple[str, List[str]]:

    log(
        f"        Fetching source: {url}"
    )

    response = get_response(
        url
    )

    final_url = response.url

    content_type = (
        response.headers.get(
            "Content-Type",
            "",
        )
        .lower()
    )

    source_urls = [
        final_url
    ]

    # --------------------------------------------------------
    # Direct PDF
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
            raise RuntimeError(
                "PDF contained no extractable text."
            )

        return (
            truncate_text(
                text,
                MAX_SOURCE_CHARS,
            ),
            source_urls,
        )

    # --------------------------------------------------------
    # HTML
    # --------------------------------------------------------

    html_text, links = extract_html(
        response.text
    )

    pdf_links: List[str] = []

    for link in links:

        absolute = urljoin(
            final_url,
            link,
        )

        lower = absolute.lower()

        if (
            ".pdf" in lower
            or "factsheet" in lower
            or "fact-sheet" in lower
            or "fund-fact" in lower
            or "fundfactsheet" in lower
        ):

            if absolute not in pdf_links:
                pdf_links.append(
                    absolute
                )

    pdf_texts: List[str] = []

    for pdf_url in pdf_links[:10]:

        try:

            log(
                f"        Linked document: "
                f"{pdf_url}"
            )

            pdf_response = get_response(
                pdf_url
            )

            pdf_text = extract_pdf_text(
                pdf_response.content
            )

            if pdf_text:

                source_urls.append(
                    pdf_response.url
                )

                pdf_texts.append(
                    pdf_text
                )

        except Exception as exc:

            log(
                f"        Linked PDF failed: "
                f"{exc}"
            )

    parts: List[str] = []

    if html_text:

        parts.append(
            "--- PRUDENTIAL WEB PAGE ---\n"
            + html_text
        )

    for number, pdf_text in enumerate(
        pdf_texts,
        start=1,
    ):

        parts.append(
            f"--- LINKED PRUDENTIAL "
            f"DOCUMENT {number} ---\n"
            + pdf_text
        )

    combined = clean_text(
        "\n\n".join(parts)
    )

    if not combined:
        raise RuntimeError(
            "No usable text was extracted "
            "from the Prudential source."
        )

    return (
        truncate_text(
            combined,
            MAX_SOURCE_CHARS,
        ),
        list(
            dict.fromkeys(
                source_urls
            )
        ),
    )


# ============================================================
# RELEVANT SECTION EXTRACTION
# ============================================================

EXPOSURE_SECTION_PATTERNS = [
    r"asset allocation",
    r"asset mix",
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
    r"underlying fund",
    r"underlying funds",
    r"investment objective",
    r"investment strategy",
    r"investment approach",
]


def extract_relevant_sections(
    source_text: str,
) -> str:
    """
    Extract document regions around financially relevant
    headings/phrases.

    IMPORTANT:
        This function does NOT classify geography or sector.

    It only reduces the document presented to the AI so that
    Qwen can focus on useful financial information.
    """

    lines = source_text.splitlines()

    if not lines:
        return ""

    matched_indexes: List[int] = []

    for index, line in enumerate(lines):

        lower = line.lower()

        for pattern in EXPOSURE_SECTION_PATTERNS:

            if re.search(
                pattern,
                lower,
            ):

                # Capture context around the matching line.
                start = max(
                    0,
                    index - 12,
                )

                end = min(
                    len(lines),
                    index + 45,
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
        set(matched_indexes)
    )

    chunks: List[str] = []

    current_chunk: List[str] = []
    previous_index: Optional[int] = None

    for index in unique_indexes:

        if (
            previous_index is not None
            and index > previous_index + 1
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
        "\n\n--- RELEVANT DOCUMENT SECTION ---\n\n"
        .join(chunks)
    )

    return truncate_text(
        relevant,
        MAX_AI_SOURCE_CHARS,
    )


def prepare_ai_source(
    source_text: str,
) -> str:
    """
    Prefer financially relevant document sections.

    If none can be identified, provide the complete source.
    """

    relevant = extract_relevant_sections(
        source_text
    )

    if relevant:

        log(
            f"        Relevant source characters: "
            f"{len(relevant)}"
        )

        return relevant

    log(
        "        No specific allocation section "
        "detected; using full extracted source."
    )

    return truncate_text(
        source_text,
        MAX_AI_SOURCE_CHARS,
    )


# ============================================================
# AI MODEL
# ============================================================

def load_ai_model() -> None:

    global TOKENIZER
    global MODEL

    log("")
    log("=" * 70)
    log("LOADING LOCAL AI MODEL")
    log("=" * 70)

    log(
        f"Model: {MODEL_NAME}"
    )

    token_kwargs: Dict[str, Any] = {}

    if HF_TOKEN:

        token_kwargs["token"] = HF_TOKEN

        log(
            "Hugging Face authentication: enabled"
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
# PROMPT
# ============================================================

def build_prompt(
    fund_name: str,
    source_text: str,
    geography_categories: List[str],
    sector_categories: List[str],
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
You are a professional financial fund research analyst.

Analyze the supplied Prudential fund documentation for:

Fund:
{fund_name}

Determine the fund's actual:

1. Geographic exposure
2. Sector exposure

STRICT RULES:

- Analyze ONLY the supplied document.
- Do not use outside knowledge.
- Do not use the fund name as evidence.
- Do not guess.
- Do not invent holdings.
- Do not assume that a company or industry is present unless the
  supplied document supports it.
- Use the investment strategy, underlying funds, holdings,
  country allocation, regional allocation, asset allocation,
  sector allocation and related portfolio information when available.
- If the document provides evidence for an underlying fund's
  geography or sector exposure, you may use that evidence.
- Select up to 3 geography categories.
- Select up to 3 sector categories.
- If only 1 or 2 categories are supported, return only those.
- Do not return percentages.
- Do not include percentages in category names.
- Use ONLY the canonical category names below.
- Do not create new category names.
- Return ONLY JSON.
- No explanation.
- No markdown.

CANONICAL GEOGRAPHY CATEGORIES:

{geography_text}

CANONICAL SECTOR CATEGORIES:

{sector_text}

OUTPUT FORMAT:

{{
  "geography": [],
  "sector": []
}}

DOCUMENT:

{source_text}
""".strip()


# ============================================================
# AI INFERENCE
# ============================================================

def generate_ai_response(
    prompt: str,
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

    with torch.inference_mode():

        generated = MODEL.generate(
            **inputs,
            max_new_tokens=MAX_NEW_TOKENS,
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
) -> Optional[Dict[str, Any]]:

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
# AI RESULT VALIDATION
# ============================================================

def validate_ai_result(
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
            "AI result is not an object.",
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

    clean_geography: List[str] = []
    clean_sector: List[str] = []

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

        category = category.strip()

        if not category:

            return (
                False,
                "Geography contains empty category.",
                [],
                [],
            )

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

        category = category.strip()

        if not category:

            return (
                False,
                "Sector contains empty category.",
                [],
                [],
            )

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

    # --------------------------------------------------------
    # Empty result is a FAILED research result.
    # --------------------------------------------------------

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
        "-" * 70
    )

    log(
        f"[{index}/{total}] "
        f"{fund_name}"
    )

    log(
        "-" * 70
    )

    log("")
    log(
        f"      Researching: "
        f"{fund_name}"
    )

    # --------------------------------------------------------
    # SOURCE
    # --------------------------------------------------------

    try:

        source_text, source_urls = (
            fetch_prudential_document(
                prudential_url
            )
        )

    except Exception as exc:

        log(
            f"        Source failed: "
            f"{exc}"
        )

        return {
            "fundName": fund_name,
            "prudentialUrl": prudential_url,
            "geography": [],
            "sector": [],
            "status": "source_error",
            "error": str(exc),
            "sources": [
                prudential_url
            ],
        }

    log(
        f"        Source characters: "
        f"{len(source_text)}"
    )

    # --------------------------------------------------------
    # PREPARE RELEVANT AI INPUT
    # --------------------------------------------------------

    ai_source = prepare_ai_source(
        source_text
    )

    # Diagnostic preview.
    log("")
    log(
        "        SOURCE PREVIEW:"
    )
    log(
        "        "
        + "-" * 60
    )

    preview = ai_source[:4000]

    for line in preview.splitlines():

        log(
            f"        {line}"
        )

    log(
        "        "
        + "-" * 60
    )

    # --------------------------------------------------------
    # PROMPT
    # --------------------------------------------------------

    prompt = build_prompt(
        fund_name=fund_name,
        source_text=ai_source,
        geography_categories=geography_categories,
        sector_categories=sector_categories,
    )

    # --------------------------------------------------------
    # AI
    # --------------------------------------------------------

    log("")
    log(
        "      Running local AI..."
    )

    try:

        response = generate_ai_response(
            prompt
        )

    except Exception as exc:

        log(
            f"        AI failed: "
            f"{exc}"
        )

        return {
            "fundName": fund_name,
            "prudentialUrl": prudential_url,
            "geography": [],
            "sector": [],
            "status": "ai_error",
            "error": str(exc),
            "sources": source_urls,
        }

    log("")
    log(
        "      AI response:"
    )

    log(
        response
    )

    # --------------------------------------------------------
    # PARSE
    # --------------------------------------------------------

    parsed = extract_json_object(
        response
    )

    if parsed is None:

        return {
            "fundName": fund_name,
            "prudentialUrl": prudential_url,
            "geography": [],
            "sector": [],
            "status": "ai_error",
            "error": (
                "AI response could not "
                "be parsed as JSON."
            ),
            "sources": source_urls,
        }

    # --------------------------------------------------------
    # VALIDATE
    # --------------------------------------------------------

    (
        valid,
        error,
        geography,
        sector,
    ) = validate_ai_result(
        parsed,
        geography_categories,
        sector_categories,
    )

    if not valid:

        log(
            f"        ERROR: "
            f"{error}"
        )

        return {
            "fundName": fund_name,
            "prudentialUrl": prudential_url,
            "geography": geography,
            "sector": sector,
            "status": "ai_error",
            "error": error,
            "sources": source_urls,
        }

    # --------------------------------------------------------
    # SUCCESS
    # --------------------------------------------------------

    log("")
    log(
        f"        Geography: "
        f"{geography}"
    )

    log(
        f"        Sector: "
        f"{sector}"
    )

    log(
        "        Status: success"
    )

    return {
        "fundName": fund_name,
        "prudentialUrl": prudential_url,
        "geography": geography,
        "sector": sector,
        "status": "success",
        "sources": source_urls,
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

    prompt = build_prompt(
        fund_name="HEALTH CHECK",
        source_text=(
            "This is a model health check. "
            "Return JSON using the required schema. "
            "There is no actual fund exposure "
            "information in this test document."
        ),
        geography_categories=geography_categories,
        sector_categories=sector_categories,
    )

    try:

        response = generate_ai_response(
            prompt
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

        if not isinstance(
            parsed.get("geography"),
            list,
        ):

            return False

        if not isinstance(
            parsed.get("sector"),
            list,
        ):

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
                f"{name}: geography is not a list."
            )

            continue

        if not isinstance(
            sector,
            list,
        ):

            errors.append(
                f"{name}: sector is not a list."
            )

            continue

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

    temporary_path = output_path.with_suffix(
        output_path.suffix + ".tmp"
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

            file.write("\n")

            file.flush()

            os.fsync(
                file.fileno()
            )

        # Verify temporary JSON.
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
                "Temporary output is not an object."
            )

        if not isinstance(
            verified.get("funds"),
            list,
        ):

            raise RuntimeError(
                "Temporary output funds "
                "is not a list."
            )

        # ----------------------------------------------------
        # ONLY successful point reaches here.
        # ----------------------------------------------------

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

    parser = argparse.ArgumentParser()

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
        f"Output: {OUTPUT_PATH}"
    )

    log(
        "=" * 70
    )

    # --------------------------------------------------------
    # Categories
    # --------------------------------------------------------

    try:

        geography_categories, sector_categories = (
            load_categories(
                CATEGORIES_FILE
            )
        )

    except Exception as exc:

        log(
            f"ERROR: Category loading failed: {exc}"
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
    # Funds
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
            f"ERROR: Fund loading failed: {exc}"
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

    total = len(funds)

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
            f"ERROR: Model loading failed: {exc}"
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

        if result.get(
            "status"
        ) == "success":

            successful += 1

        else:

            failed += 1

    # --------------------------------------------------------
    # Build result in memory.
    # --------------------------------------------------------

    output: Dict[str, Any] = {
        "generatedAt": utc_now(),
        "model": MODEL_NAME,
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
