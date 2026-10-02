#!/usr/bin/env python3

"""
VGrat FMS - AI FUND EXPOSURE RESEARCH

Purpose
-------
Research geography and sector exposure for every fund in the controlling
Funds Links workbook using a local Hugging Face Qwen model.

IMPORTANT
---------
This script does NOT use DuckDuckGo or any search engine.

The analysis source is the Prudential URL/document supplied in the workbook.
The AI performs the geography and sector classification.

Transactional publishing:
    - A completely successful run replaces fund_exposure.json.
    - Any failure preserves the existing fund_exposure.json.
    - Partial results are NEVER published.

Expected output:

{
  "generatedAt": "...",
  "model": "Qwen/Qwen3-1.7B",
  "fundCount": 67,
  "funds": [
    {
      "fundName": "...",
      "prudentialUrl": "...",
      "geography": [
        "United States",
        "Europe"
      ],
      "sector": [
        "Technology"
      ],
      "status": "success",
      "sources": [
        "..."
      ]
    }
  ],
  "summary": {
    "totalFunds": 67,
    "successful": 67,
    "failed": 0
  }
}
"""

from __future__ import annotations

import argparse
import json
import os
import re
import sys
import time
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

HF_TOKEN = os.getenv("HF_TOKEN", "").strip()

SCRIPT_DIR = Path(__file__).resolve().parent
REPOSITORY_ROOT = SCRIPT_DIR.parent

CATEGORIES_FILE = SCRIPT_DIR / "exposure_categories.txt"
OUTPUT_PATH = REPOSITORY_ROOT / "data" / "fund_exposure.json"

REQUEST_TIMEOUT = int(os.getenv("REQUEST_TIMEOUT", "60"))

MAX_SOURCE_CHARS = int(
    os.getenv("MAX_SOURCE_CHARS", "50000")
)

MAX_PDF_PAGES = int(
    os.getenv("MAX_PDF_PAGES", "30")
)

MAX_NEW_TOKENS = int(
    os.getenv("MAX_NEW_TOKENS", "700")
)

TEMPERATURE = float(
    os.getenv("AI_TEMPERATURE", "0.0")
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
        "application/pdf;q=0.8,*/*;q=0.7"
    ),
    "Accept-Language": "en-US,en;q=0.9",
}


# ============================================================
# GLOBAL MODEL OBJECTS
# ============================================================

TOKENIZER = None
MODEL = None


# ============================================================
# GENERAL HELPERS
# ============================================================

def log(message: str = "") -> None:
    print(message, flush=True)


def utc_now() -> str:
    return datetime.now(timezone.utc).isoformat()


def clean_text(text: str) -> str:
    """
    Normalise extracted document text.
    """
    if not text:
        return ""

    text = text.replace("\x00", " ")
    text = text.replace("\r\n", "\n")
    text = text.replace("\r", "\n")

    # Remove excessive whitespace while preserving useful line breaks.
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\n{3,}", "\n\n", text)

    return text.strip()


def truncate_text(text: str, max_chars: int = MAX_SOURCE_CHARS) -> str:
    """
    Keep prompt size manageable.
    """
    if len(text) <= max_chars:
        return text

    return text[:max_chars] + "\n\n[DOCUMENT TRUNCATED]"


# ============================================================
# WORKBOOK
# ============================================================

def find_funds_excel_file(repository_root: Path) -> Path:
    """
    Prefer Funds Links.xlsx.

    Fallback:
        Funds Links.xlsm
    """

    candidates = [
        repository_root / "Funds Links.xlsx",
        repository_root / "Funds Links.xlsm",
    ]

    for candidate in candidates:
        if candidate.exists():
            return candidate

    raise FileNotFoundError(
        "Could not find Funds Links.xlsx or Funds Links.xlsm "
        f"in {repository_root}"
    )


def load_funds(workbook_path: Path) -> List[Dict[str, str]]:
    """
    Load controlling fund universe.

    Column A:
        Prudential URL

    Column B:
        Exact fund name
    """

    log(f"Loading workbook: {workbook_path}")

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
            url = row[0] if len(row) >= 1 else None
            fund_name = row[1] if len(row) >= 2 else None

            if url is None and fund_name is None:
                continue

            url = str(url or "").strip()
            fund_name = str(fund_name or "").strip()

            if not fund_name:
                log(
                    f"WARNING: Row {row_number} has no fund name. "
                    "Skipping."
                )
                continue

            if not url:
                raise RuntimeError(
                    f"Row {row_number} ({fund_name}) has no Prudential URL."
                )

            funds.append(
                {
                    "fundName": fund_name,
                    "prudentialUrl": url,
                }
            )

        if not funds:
            raise RuntimeError(
                "No funds were found in the controlling workbook."
            )

        log(f"Funds loaded: {len(funds)}")

        return funds

    finally:
        workbook.close()


# ============================================================
# CATEGORY VOCABULARY
# ============================================================

def load_categories(
    categories_path: Path,
) -> Tuple[List[str], List[str]]:
    """
    Read canonical geography and sector vocabulary.

    Expected format:

    [GEOGRAPHY]
    United States
    Europe
    ...

    [SECTOR]
    Technology
    Financials
    ...
    """

    if not categories_path.exists():
        raise FileNotFoundError(
            f"Category file not found: {categories_path}"
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

    if not geography:
        raise RuntimeError(
            "No geography categories found in exposure_categories.txt."
        )

    if not sector:
        raise RuntimeError(
            "No sector categories found in exposure_categories.txt."
        )

    # Remove duplicates while preserving order.
    geography = list(dict.fromkeys(geography))
    sector = list(dict.fromkeys(sector))

    return geography, sector


# ============================================================
# HTTP / DOCUMENT FETCHING
# ============================================================

def fetch_url(
    url: str,
) -> Tuple[str, str]:
    """
    Download a URL.

    Returns:
        (content_type, body)

    body is text for HTML and raw bytes encoded as latin-1 for PDF
    handling by the caller.

    Raises on failure.
    """

    log(f"        Fetching: {url}")

    response = requests.get(
        url,
        headers=HEADERS,
        timeout=REQUEST_TIMEOUT,
        allow_redirects=True,
    )

    response.raise_for_status()

    content_type = (
        response.headers.get(
            "Content-Type",
            "",
        )
        .lower()
    )

    if (
        "application/pdf" in content_type
        or response.url.lower().split("?")[0].endswith(".pdf")
    ):
        return (
            "pdf",
            response.content.decode(
                "latin-1",
                errors="ignore",
            ),
        )

    return (
        "html",
        response.text,
    )


def extract_pdf_text_from_bytes(
    raw_bytes: bytes,
) -> str:
    """
    Extract text from PDF bytes.
    """

    import io

    reader = PdfReader(
        io.BytesIO(raw_bytes)
    )

    pages = min(
        len(reader.pages),
        MAX_PDF_PAGES,
    )

    chunks: List[str] = []

    for index in range(pages):
        try:
            page_text = reader.pages[index].extract_text() or ""

            if page_text.strip():
                chunks.append(
                    f"\n--- PAGE {index + 1} ---\n"
                    f"{page_text}"
                )

        except Exception as exc:
            log(
                f"        PDF page {index + 1} extraction failed: "
                f"{exc}"
            )

    return clean_text(
        "\n".join(chunks)
    )


def extract_html_text(
    html: str,
) -> Tuple[str, List[str]]:
    """
    Extract visible text and linked PDF/document URLs.
    """

    soup = BeautifulSoup(
        html,
        "html.parser",
    )

    for element in soup(
        [
            "script",
            "style",
            "noscript",
            "svg",
        ]
    ):
        element.decompose()

    links: List[str] = []

    for anchor in soup.find_all("a"):
        href = anchor.get("href")

        if not href:
            continue

        href = href.strip()

        absolute = urljoin(
            "",
            href,
        )

        links.append(absolute)

    text = clean_text(
        soup.get_text(
            "\n",
            strip=True,
        )
    )

    return text, links


def fetch_prudential_document(
    url: str,
) -> Tuple[str, List[str]]:
    """
    Fetch the Prudential source page and any directly linked PDF.

    No search engine is used.

    Returns:
        source_text
        source_urls
    """

    try:
        response = requests.get(
            url,
            headers=HEADERS,
            timeout=REQUEST_TIMEOUT,
            allow_redirects=True,
        )

        response.raise_for_status()

    except Exception as exc:
        raise RuntimeError(
            f"Prudential URL fetch failed: {exc}"
        ) from exc

    final_url = response.url

    content_type = (
        response.headers.get(
            "Content-Type",
            "",
        )
        .lower()
    )

    source_urls = [
        final_url,
    ]

    # --------------------------------------------------------
    # Direct PDF
    # --------------------------------------------------------

    if (
        "application/pdf" in content_type
        or final_url.lower().split("?")[0].endswith(".pdf")
    ):
        text = extract_pdf_text_from_bytes(
            response.content
        )

        if not text:
            raise RuntimeError(
                "Prudential PDF was downloaded but contained "
                "no extractable text."
            )

        return (
            truncate_text(text),
            source_urls,
        )

    # --------------------------------------------------------
    # HTML
    # --------------------------------------------------------

    html = response.text

    html_text, links = extract_html_text(
        html
    )

    # --------------------------------------------------------
    # Find linked PDFs.
    #
    # We intentionally only follow documents linked directly
    # from the supplied Prudential page. No web search.
    # --------------------------------------------------------

    pdf_links: List[str] = []

    for link in links:
        lower = link.lower()

        if (
            ".pdf" in lower
            or "pdf" in lower
            or "factsheet" in lower
            or "fund-fact" in lower
        ):
            absolute = urljoin(
                final_url,
                link,
            )

            if absolute not in pdf_links:
                pdf_links.append(absolute)

    pdf_texts: List[str] = []

    for pdf_url in pdf_links[:5]:
        try:
            log(
                f"        Linked document: {pdf_url}"
            )

            pdf_response = requests.get(
                pdf_url,
                headers=HEADERS,
                timeout=REQUEST_TIMEOUT,
                allow_redirects=True,
            )

            pdf_response.raise_for_status()

            pdf_content_type = (
                pdf_response.headers.get(
                    "Content-Type",
                    "",
                )
                .lower()
            )

            if (
                "application/pdf" not in pdf_content_type
                and not pdf_response.url.lower()
                .split("?")[0]
                .endswith(".pdf")
            ):
                continue

            pdf_text = extract_pdf_text_from_bytes(
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
                f"        Linked document failed: {exc}"
            )

    combined_parts: List[str] = []

    if html_text:
        combined_parts.append(
            "--- PRUDENTIAL WEB PAGE ---\n"
            + html_text
        )

    for index, pdf_text in enumerate(
        pdf_texts,
        start=1,
    ):
        combined_parts.append(
            f"--- LINKED PRUDENTIAL DOCUMENT {index} ---\n"
            + pdf_text
        )

    combined = clean_text(
        "\n\n".join(combined_parts)
    )

    if not combined:
        raise RuntimeError(
            "Prudential source produced no usable text."
        )

    return (
        truncate_text(combined),
        list(dict.fromkeys(source_urls)),
    )


# ============================================================
# AI MODEL
# ============================================================

def load_ai_model() -> None:
    """
    Load Qwen locally.
    """

    global TOKENIZER
    global MODEL

    log("")
    log("=" * 70)
    log("LOADING LOCAL AI MODEL")
    log("=" * 70)
    log(f"Model: {MODEL_NAME}")

    token_kwargs: Dict[str, Any] = {}

    if HF_TOKEN:
        token_kwargs["token"] = HF_TOKEN
        log("Hugging Face authentication: enabled")
    else:
        log(
            "Hugging Face authentication: not configured "
            "(public model access)"
        )

    TOKENIZER = AutoTokenizer.from_pretrained(
        MODEL_NAME,
        **token_kwargs,
    )

    MODEL = AutoModelForCausalLM.from_pretrained(
        MODEL_NAME,
        torch_dtype="auto",
        device_map="auto",
        trust_remote_code=True,
        **token_kwargs,
    )

    MODEL.eval()

    log("AI model loaded successfully.")


def build_prompt(
    fund_name: str,
    source_text: str,
    geography_categories: List[str],
    sector_categories: List[str],
) -> str:
    """
    Build strict JSON-only classification prompt.
    """

    geography_text = "\n".join(
        f"- {category}"
        for category in geography_categories
    )

    sector_text = "\n".join(
        f"- {category}"
        for category in sector_categories
    )

    return f"""
You are a financial-fund research analyst.

Analyze ONLY the supplied source document.

Fund:
{fund_name}

Your task is to identify the fund's actual:

1. Geographic exposure
2. Sector exposure

IMPORTANT RULES:

- Use ONLY information supported by the supplied document.
- Do NOT use your general knowledge.
- Do NOT infer exposure merely from the fund name.
- Do NOT invent holdings.
- Do NOT guess.
- Do NOT calculate or return percentages.
- Do NOT return percentages in category names.
- Return at most 3 geography categories.
- Return at most 3 sector categories.
- Return fewer than 3 when fewer are supported.
- If no valid geography can be established, return [].
- If no valid sector can be established, return [].
- You MUST use the canonical category names supplied below.
- Do not create new category names.
- Do not combine categories.
- Do not add explanations.
- Return ONLY valid JSON.

CANONICAL GEOGRAPHY CATEGORIES:

{geography_text}

CANONICAL SECTOR CATEGORIES:

{sector_text}

Required JSON format:

{{
  "geography": [],
  "sector": []
}}

SOURCE DOCUMENT:

{source_text}
""".strip()


def generate_ai_response(
    prompt: str,
) -> str:
    """
    Run local Qwen inference.
    """

    if TOKENIZER is None or MODEL is None:
        raise RuntimeError(
            "AI model is not loaded."
        )

    messages = [
        {
            "role": "system",
            "content": (
                "You are a precise financial research "
                "classification model. "
                "Return only valid JSON."
            ),
        },
        {
            "role": "user",
            "content": prompt,
        },
    ]

    # --------------------------------------------------------
    # Qwen3 chat template.
    # --------------------------------------------------------

    try:
        inputs = TOKENIZER.apply_chat_template(
            messages,
            tokenize=True,
            add_generation_prompt=True,
            enable_thinking=False,
            return_dict=True,
            return_tensors="pt",
        )

    except TypeError:
        inputs = TOKENIZER.apply_chat_template(
            messages,
            tokenize=True,
            add_generation_prompt=True,
            return_tensors="pt",
        )

    # --------------------------------------------------------
    # Move tensors to model device.
    # --------------------------------------------------------

    if hasattr(inputs, "to"):
        try:
            inputs = inputs.to(
                MODEL.device
            )
        except Exception:
            pass

    input_length = (
        inputs["input_ids"].shape[-1]
    )

    log(
        f"        Input tokens: {input_length}"
    )

    with torch.inference_mode():
        generated = MODEL.generate(
            **inputs,
            max_new_tokens=MAX_NEW_TOKENS,
            do_sample=False,
            temperature=TEMPERATURE,
            pad_token_id=(
                TOKENIZER.eos_token_id
            ),
        )

    generated_tokens = generated[
        0,
        input_length:,
    ]

    response = TOKENIZER.decode(
        generated_tokens,
        skip_special_tokens=True,
    )

    return response.strip()


# ============================================================
# AI RESPONSE PARSING
# ============================================================

def extract_json_object(
    response: str,
) -> Optional[Dict[str, Any]]:
    """
    Extract JSON object from model response.

    The model is asked to return JSON only, but this also handles
    accidental markdown fences.
    """

    if not response:
        return None

    cleaned = response.strip()

    # Remove markdown code fences.
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

    # Direct parse.
    try:
        parsed = json.loads(
            cleaned
        )

        if isinstance(parsed, dict):
            return parsed

    except json.JSONDecodeError:
        pass

    # Search for the first JSON object.
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

        if isinstance(parsed, dict):
            return parsed

    except json.JSONDecodeError:
        return None

    return None


def validate_ai_result(
    parsed: Dict[str, Any],
    geography_categories: List[str],
    sector_categories: List[str],
) -> Tuple[bool, str, List[str], List[str]]:
    """
    Validate AI output.

    Returns:
        valid
        error
        geography
        sector
    """

    if not isinstance(parsed, dict):
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
            "AI returned more than 3 geography categories.",
            [],
            [],
        )

    if len(sector) > 3:
        return (
            False,
            "AI returned more than 3 sector categories.",
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
                "Geography contains a non-string category.",
                [],
                [],
            )

        category = category.strip()

        if not category:
            return (
                False,
                "Geography contains an empty category.",
                [],
                [],
            )

        if "%" in category:
            return (
                False,
                f"Percentage found in geography: {category}",
                [],
                [],
            )

        if category not in geography_set:
            return (
                False,
                f"Invalid geography category: {category}",
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
                "Sector contains a non-string category.",
                [],
                [],
            )

        category = category.strip()

        if not category:
            return (
                False,
                "Sector contains an empty category.",
                [],
                [],
            )

        if "%" in category:
            return (
                False,
                f"Percentage found in sector: {category}",
                [],
                [],
            )

        if category not in sector_set:
            return (
                False,
                f"Invalid sector category: {category}",
                [],
                [],
            )

        if category not in clean_sector:
            clean_sector.append(
                category
            )

    # --------------------------------------------------------
    # Empty / empty is a failed research result.
    # --------------------------------------------------------

    if not clean_geography and not clean_sector:
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

    fund_name = fund["fundName"]
    prudential_url = fund["prudentialUrl"]

    log("")
    log("-" * 70)
    log(
        f"[{index}/{total}] {fund_name}"
    )
    log("-" * 70)

    log("")
    log(
        f"      Researching: {fund_name}"
    )

    # --------------------------------------------------------
    # Fetch Prudential source.
    # --------------------------------------------------------

    try:
        source_text, source_urls = (
            fetch_prudential_document(
                prudential_url
            )
        )

    except Exception as exc:
        log(
            f"        source failed: {exc}"
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
        f"        Source characters: {len(source_text)}"
    )

    # --------------------------------------------------------
    # AI analysis.
    # --------------------------------------------------------

    prompt = build_prompt(
        fund_name=fund_name,
        source_text=source_text,
        geography_categories=geography_categories,
        sector_categories=sector_categories,
    )

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
            f"        AI failed: {exc}"
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
    log(response)

    # --------------------------------------------------------
    # Parse AI response.
    # --------------------------------------------------------

    parsed = extract_json_object(
        response
    )

    if parsed is None:
        log(
            "        ERROR: AI response could not "
            "be parsed as JSON."
        )

        return {
            "fundName": fund_name,
            "prudentialUrl": prudential_url,
            "geography": [],
            "sector": [],
            "status": "ai_error",
            "error": (
                "AI response could not be parsed as JSON."
            ),
            "sources": source_urls,
        }

    # --------------------------------------------------------
    # Validate.
    # --------------------------------------------------------

    (
        valid,
        validation_error,
        geography,
        sector,
    ) = validate_ai_result(
        parsed,
        geography_categories,
        sector_categories,
    )

    if not valid:
        log(
            f"        ERROR: {validation_error}"
        )

        return {
            "fundName": fund_name,
            "prudentialUrl": prudential_url,
            "geography": geography,
            "sector": sector,
            "status": "ai_error",
            "error": validation_error,
            "sources": source_urls,
        }

    # --------------------------------------------------------
    # Success.
    # --------------------------------------------------------

    log("")
    log(
        f"      Geography: {geography}"
    )

    log(
        f"      Sector: {sector}"
    )

    log(
        "      Status: success"
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
    """
    Confirm that the local model can produce valid JSON before
    processing the fund universe.
    """

    log("")
    log("=" * 70)
    log("AI HEALTH CHECK")
    log("=" * 70)

    prompt = build_prompt(
        fund_name="HEALTH CHECK",
        source_text=(
            "This is a model health check. "
            "The document contains no actual fund exposure. "
            "Return empty arrays if no exposure is supported."
        ),
        geography_categories=geography_categories,
        sector_categories=sector_categories,
    )

    try:
        response = generate_ai_response(
            prompt
        )

        log(
            f"Health-check response: {response}"
        )

        parsed = extract_json_object(
            response
        )

        if parsed is None:
            log(
                "ERROR: AI health check returned invalid JSON."
            )
            return False

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
            log(
                "ERROR: AI health check geography is invalid."
            )
            return False

        if not isinstance(
            sector,
            list,
        ):
            log(
                "ERROR: AI health check sector is invalid."
            )
            return False

        log(
            "AI health check PASSED."
        )

        return True

    except Exception as exc:
        log(
            f"ERROR: AI health check failed: {exc}"
        )
        return False


# ============================================================
# COMPLETE RESULT VALIDATION
# ============================================================

def validate_complete_result(
    output: Dict[str, Any],
    total_expected: int,
    geography_categories: List[str],
    sector_categories: List[str],
) -> Tuple[bool, List[str]]:
    """
    Validate the entire run.

    A run is successful ONLY when every expected fund has
    status == success.
    """

    errors: List[str] = []

    if not isinstance(
        output,
        dict,
    ):
        errors.append(
            "Output root is not an object."
        )
        return False, errors

    funds = output.get(
        "funds"
    )

    if not isinstance(
        funds,
        list,
    ):
        errors.append(
            "Output does not contain a funds list."
        )
        return False, errors

    if len(funds) != total_expected:
        errors.append(
            f"Fund count mismatch: expected "
            f"{total_expected}, got {len(funds)}."
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

        fund_name = fund.get(
            "fundName",
            f"Fund #{index}",
        )

        status = fund.get(
            "status"
        )

        # ----------------------------------------------------
        # THE KEY RULE:
        # ANYTHING OTHER THAN SUCCESS FAILS THE ENTIRE RUN.
        # ----------------------------------------------------

        if status != "success":
            errors.append(
                f"{fund_name}: status={status!r}"
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
                f"{fund_name}: geography is not a list."
            )
            continue

        if not isinstance(
            sector,
            list,
        ):
            errors.append(
                f"{fund_name}: sector is not a list."
            )
            continue

        if len(geography) > 3:
            errors.append(
                f"{fund_name}: more than 3 geography categories."
            )

        if len(sector) > 3:
            errors.append(
                f"{fund_name}: more than 3 sector categories."
            )

        for category in geography:
            if not isinstance(
                category,
                str,
            ):
                errors.append(
                    f"{fund_name}: non-string geography."
                )
                continue

            if "%" in category:
                errors.append(
                    f"{fund_name}: percentage in geography."
                )

            if category not in geography_set:
                errors.append(
                    f"{fund_name}: invalid geography "
                    f"category {category!r}."
                )

        for category in sector:
            if not isinstance(
                category,
                str,
            ):
                errors.append(
                    f"{fund_name}: non-string sector."
                )
                continue

            if "%" in category:
                errors.append(
                    f"{fund_name}: percentage in sector."
                )

            if category not in sector_set:
                errors.append(
                    f"{fund_name}: invalid sector "
                    f"category {category!r}."
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
    """
    Safely replace the existing output.

    The existing file is not touched until the new file has been:

        1. Written completely.
        2. Flushed.
        3. Synced.
        4. Read back.
        5. Parsed successfully.

    Then os.replace() atomically swaps it in.
    """

    output_path.parent.mkdir(
        parents=True,
        exist_ok=True,
    )

    temp_path = output_path.with_suffix(
        output_path.suffix + ".tmp"
    )

    try:
        log("")
        log(
            f"Preparing temporary output: {temp_path}"
        )

        with temp_path.open(
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

        # ----------------------------------------------------
        # Verify temporary file.
        # ----------------------------------------------------

        with temp_path.open(
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
                "Temporary output validation failed: "
                "root is not an object."
            )

        if not isinstance(
            verified.get("funds"),
            list,
        ):
            raise RuntimeError(
                "Temporary output validation failed: "
                "'funds' is not a list."
            )

        # ----------------------------------------------------
        # ONLY HERE do we replace the existing file.
        # ----------------------------------------------------

        os.replace(
            temp_path,
            output_path,
        )

        log("")
        log(
            f"Successfully replaced: {output_path}"
        )

    except Exception:
        # ----------------------------------------------------
        # Most important safety rule:
        # delete only the temporary file.
        # Never delete or modify the existing output.
        # ----------------------------------------------------

        try:
            if temp_path.exists():
                temp_path.unlink()
        except Exception:
            pass

        raise


# ============================================================
# ARGUMENTS
# ============================================================

def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Research fund geography and sector exposure "
            "using local Qwen AI."
        )
    )

    parser.add_argument(
        "--limit",
        type=int,
        default=None,
        help=(
            "Process only the first N funds. "
            "Useful for testing."
        ),
    )

    return parser.parse_args()


# ============================================================
# MAIN
# ============================================================

def main() -> int:

    args = parse_args()

    log("")
    log("=" * 70)
    log("VGRAT FMS - AI FUND EXPOSURE RESEARCH")
    log("=" * 70)
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
    log("=" * 70)

    # --------------------------------------------------------
    # Load categories.
    # --------------------------------------------------------

    try:
        geography_categories, sector_categories = (
            load_categories(
                CATEGORIES_FILE
            )
        )

    except Exception as exc:
        log("")
        log(
            f"ERROR: Failed to load categories: {exc}"
        )
        log(
            "Existing fund_exposure.json has NOT been modified."
        )
        return 1

    log("")
    log(
        f"Geography categories: {len(geography_categories)}"
    )
    log(
        f"Sector categories: {len(sector_categories)}"
    )

    # --------------------------------------------------------
    # Load workbook.
    # --------------------------------------------------------

    try:
        workbook_path = find_funds_excel_file(
            REPOSITORY_ROOT
        )

        funds = load_funds(
            workbook_path
        )

    except Exception as exc:
        log("")
        log(
            f"ERROR: Failed to load fund universe: {exc}"
        )
        log(
            "Existing fund_exposure.json has NOT been modified."
        )
        return 1

    # --------------------------------------------------------
    # Apply optional test limit.
    # --------------------------------------------------------

    original_total = len(funds)

    if args.limit is not None:

        if args.limit <= 0:
            log(
                "ERROR: --limit must be greater than zero."
            )
            return 1

        funds = funds[
            :args.limit
        ]

        log("")
        log(
            f"TEST LIMIT ENABLED: "
            f"{len(funds)} of {original_total} funds"
        )

    total = len(funds)

    if total == 0:
        log(
            "ERROR: No funds selected."
        )
        return 1

    # --------------------------------------------------------
    # Load model.
    # --------------------------------------------------------

    try:
        load_ai_model()

    except Exception as exc:
        log("")
        log(
            f"ERROR: Failed to load AI model: {exc}"
        )
        log(
            "Existing fund_exposure.json has NOT been modified."
        )
        return 1

    # --------------------------------------------------------
    # Health check.
    # --------------------------------------------------------

    if not ai_health_check(
        geography_categories,
        sector_categories,
    ):
        log("")
        log(
            "AI HEALTH CHECK FAILED."
        )
        log(
            "Existing fund_exposure.json has NOT been modified."
        )
        return 1

    # --------------------------------------------------------
    # Research all selected funds.
    # --------------------------------------------------------

    results: List[Dict[str, Any]] = []

    success_count = 0
    failure_count = 0

    for index, fund in enumerate(
        funds,
        start=1,
    ):

        result = process_fund(
            fund=fund,
            geography_categories=geography_categories,
            sector_categories=sector_categories,
            index=index,
            total=total,
        )

        results.append(
            result
        )

        if result.get("status") == "success":
            success_count += 1
        else:
            failure_count += 1

    # --------------------------------------------------------
    # Build complete result in MEMORY ONLY.
    #
    # Nothing has been written to fund_exposure.json.
    # --------------------------------------------------------

    output: Dict[str, Any] = {
        "generatedAt": utc_now(),
        "model": MODEL_NAME,
        "fundCount": total,
        "funds": results,
        "summary": {
            "totalFunds": total,
            "successful": success_count,
            "failed": failure_count,
        },
    }

    # --------------------------------------------------------
    # Complete validation.
    #
    # Any single failed fund prevents publishing.
    # --------------------------------------------------------

    valid, errors = validate_complete_result(
        output=output,
        total_expected=total,
        geography_categories=geography_categories,
        sector_categories=sector_categories,
    )

    log("")
    log("=" * 70)
    log("FINAL VALIDATION")
    log("=" * 70)

    log(
        f"Total funds: {total}"
    )

    log(
        f"Successful: {success_count}"
    )

    log(
        f"Failed: {failure_count}"
    )

    if not valid:

        log("")
        log(
            "FINAL VALIDATION FAILED."
        )

        log(
            "The existing fund_exposure.json will NOT be replaced."
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
    # Additional success condition.
    # --------------------------------------------------------

    if success_count != total:

        log("")
        log(
            "ERROR: Not all funds succeeded."
        )

        log(
            "The existing fund_exposure.json will NOT be replaced."
        )

        return 1

    # --------------------------------------------------------
    # Publish atomically.
    # --------------------------------------------------------

    try:
        publish_output_atomically(
            output,
            OUTPUT_PATH,
        )

    except Exception as exc:

        log("")
        log(
            f"ERROR: Failed to publish output: {exc}"
        )

        log(
            "The existing fund_exposure.json "
            "was not intentionally replaced."
        )

        return 1

    # --------------------------------------------------------
    # SUCCESS
    # --------------------------------------------------------

    log("")
    log("=" * 70)
    log("RESEARCH SUCCESS")
    log("=" * 70)

    log(
        f"Funds successfully researched: "
        f"{success_count}/{total}"
    )

    log(
        f"Published: {OUTPUT_PATH}"
    )

    log(
        "Previous output replaced because the complete run succeeded."
    )

    log("=" * 70)

    return 0


if __name__ == "__main__":
    sys.exit(
        main()
    )
