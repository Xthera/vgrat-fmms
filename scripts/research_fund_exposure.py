#!/usr/bin/env python3

"""
VGrat FMS - AI FUND EXPOSURE RESEARCH

Adaptive AI-only geography and sector research.

Supports:
    1. Direct-investment funds
    2. Fund-of-funds / multi-manager funds

Architecture:
    Prudential source
        |
        +--> Qwen structure analysis
        |       |
        |       +--> direct investment
        |       |
        |       +--> fund of funds
        |
        +--> collect available supporting documents
                |
                +--> Qwen exposure analysis

Important:
    - No DuckDuckGo
    - No Google
    - No search-engine research
    - Qwen performs classification
    - Python only retrieves and prepares evidence
    - AI may only return canonical categories
    - Maximum 3 geography
    - Maximum 3 sector
    - No percentages in final output

Publishing:
    - Complete successful run -> replace fund_exposure.json
    - Any failure -> preserve existing fund_exposure.json
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
from typing import Any, Dict, List, Optional, Set, Tuple
from urllib.parse import urljoin, urlparse

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
    os.getenv("REQUEST_TIMEOUT", "60")
)

MAX_PDF_PAGES = int(
    os.getenv("MAX_PDF_PAGES", "60")
)

MAX_SOURCE_CHARS = int(
    os.getenv("MAX_SOURCE_CHARS", "90000")
)

MAX_AI_SOURCE_CHARS = int(
    os.getenv("MAX_AI_SOURCE_CHARS", "50000")
)

MAX_NEW_TOKENS = int(
    os.getenv("MAX_NEW_TOKENS", "900")
)

MAX_LINKED_DOCUMENTS = int(
    os.getenv("MAX_LINKED_DOCUMENTS", "20")
)

MAX_UNDERLYING_FUNDS = int(
    os.getenv("MAX_UNDERLYING_FUNDS", "12")
)

MAX_CRAWL_DEPTH = int(
    os.getenv("MAX_CRAWL_DEPTH", "1")
)

MAX_DOCUMENTS_PER_UNDERLYING = int(
    os.getenv("MAX_DOCUMENTS_PER_UNDERLYING", "5")
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
    print(message, flush=True)


def utc_now() -> str:
    return datetime.now(
        timezone.utc
    ).isoformat()


# ============================================================
# TEXT HELPERS
# ============================================================

def clean_text(text: str) -> str:

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


def normalize_name(
    value: str,
) -> str:

    value = value.lower()

    value = value.replace(
        "&",
        " and ",
    )

    value = re.sub(
        r"[^a-z0-9]+",
        " ",
        value,
    )

    value = re.sub(
        r"\s+",
        " ",
        value,
    )

    return value.strip()


def name_tokens(
    value: str,
) -> Set[str]:

    stopwords = {
        "fund",
        "funds",
        "class",
        "share",
        "shares",
        "unit",
        "units",
        "acc",
        "accumulation",
        "dist",
        "distribution",
        "sgd",
        "usd",
        "eur",
        "gbp",
        "hedged",
        "unhedged",
        "institutional",
        "retail",
        "portfolio",
        "prulink",
    }

    return {
        token
        for token in normalize_name(value).split()
        if token not in stopwords
        and len(token) > 2
    }


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
        "Could not find Funds Links.xlsx "
        "or Funds Links.xlsm."
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
                    f"({fund_name}) has no "
                    "Prudential URL."
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


def build_workbook_name_index(
    funds: List[Dict[str, str]],
) -> Dict[str, Dict[str, str]]:

    index: Dict[str, Dict[str, str]] = {}

    for fund in funds:

        name = fund["fundName"]

        index[
            normalize_name(name)
        ] = fund

    return index


def find_workbook_fund_match(
    underlying_name: str,
    workbook_index: Dict[str, Dict[str, str]],
) -> Optional[Dict[str, str]]:

    normalized = normalize_name(
        underlying_name
    )

    if normalized in workbook_index:
        return workbook_index[normalized]

    target_tokens = name_tokens(
        underlying_name
    )

    if not target_tokens:
        return None

    best_match = None
    best_score = 0.0

    for indexed_name, fund in workbook_index.items():

        candidate_tokens = name_tokens(
            indexed_name
        )

        if not candidate_tokens:
            continue

        intersection = (
            target_tokens
            & candidate_tokens
        )

        union = (
            target_tokens
            | candidate_tokens
        )

        score = (
            len(intersection)
            / max(len(union), 1)
        )

        if (
            len(intersection) >= 2
            and score > best_score
        ):

            best_score = score
            best_match = fund

    if best_score >= 0.50:
        return best_match

    return None


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

SESSION = requests.Session()
SESSION.headers.update(HEADERS)


def get_response(
    url: str,
) -> requests.Response:

    response = SESSION.get(
        url,
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
    base_url: str,
) -> Tuple[
    str,
    List[Dict[str, str]],
]:

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

    links: List[Dict[str, str]] = []

    for anchor in soup.find_all("a"):

        href = anchor.get("href")

        if not href:
            continue

        href = href.strip()

        absolute = urljoin(
            base_url,
            href,
        )

        link_text = clean_text(
            anchor.get_text(
                " ",
                strip=True,
            )
        )

        links.append(
            {
                "url": absolute,
                "text": link_text,
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
# DOCUMENT LINK SCORING
# ============================================================

DOCUMENT_TERMS = [
    "factsheet",
    "fact sheet",
    "fund factsheet",
    "fund-fact",
    "fund facts",
    "fund report",
    "fund report",
    "portfolio",
    "holdings",
    "investment",
    "annual report",
    "monthly report",
    "quarterly report",
    "semi annual",
    "semi-annual",
    "prospectus",
    "statement",
    "allocation",
]

UNDERLYING_TERMS = [
    "underlying fund",
    "underlying funds",
    "underlying investment",
    "underlying portfolio",
    "collective investment",
    "investment scheme",
    "manager",
    "fund manager",
]


def score_document_link(
    url: str,
    text: str,
) -> int:

    combined = (
        f"{url} {text}"
    ).lower()

    score = 0

    if ".pdf" in combined:
        score += 8

    for term in DOCUMENT_TERMS:

        if term in combined:
            score += 3

    for term in UNDERLYING_TERMS:

        if term in combined:
            score += 4

    return score


def same_domain(
    url_a: str,
    url_b: str,
) -> bool:

    return (
        urlparse(url_a).netloc.lower()
        == urlparse(url_b).netloc.lower()
    )


# ============================================================
# GENERIC SOURCE FETCH
# ============================================================

def fetch_document(
    url: str,
) -> Tuple[str, List[Dict[str, str]], str]:

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

    if (
        "application/pdf" in content_type
        or final_url.lower()
        .split("?")[0]
        .endswith(".pdf")
    ):

        text = extract_pdf_text(
            response.content
        )

        return (
            truncate_text(
                text,
                MAX_SOURCE_CHARS,
            ),
            [],
            final_url,
        )

    html_text, links = extract_html(
        response.text,
        final_url,
    )

    return (
        truncate_text(
            html_text,
            MAX_SOURCE_CHARS,
        ),
        links,
        final_url,
    )


# ============================================================
# SOURCE COLLECTION
# ============================================================

def fetch_source_bundle(
    url: str,
    label: str,
    max_documents: int = MAX_LINKED_DOCUMENTS,
) -> Dict[str, Any]:

    log(
        f"        Fetching {label}: {url}"
    )

    source_urls: List[str] = []
    documents: List[str] = []
    candidate_links: List[Dict[str, Any]] = []

    text, links, final_url = fetch_document(
        url
    )

    source_urls.append(
        final_url
    )

    if text:

        documents.append(
            "--- PRIMARY SOURCE ---\n"
            + text
        )

    # --------------------------------------------------------
    # Rank linked documents.
    # --------------------------------------------------------

    ranked_links = []

    for link in links:

        linked_url = link["url"]
        linked_text = link["text"]

        if not linked_url:
            continue

        score = score_document_link(
            linked_url,
            linked_text,
        )

        if score <= 0:
            continue

        ranked_links.append(
            {
                "url": linked_url,
                "text": linked_text,
                "score": score,
            }
        )

    ranked_links.sort(
        key=lambda item: (
            -item["score"],
            item["url"],
        )
    )

    seen_urls = {
        final_url
    }

    for candidate in ranked_links:

        linked_url = candidate["url"]

        if linked_url in seen_urls:
            continue

        seen_urls.add(
            linked_url
        )

        if len(documents) >= (
            max_documents + 1
        ):
            break

        try:

            log(
                "        Supporting document: "
                f"{linked_url}"
            )

            linked_text, _, linked_final_url = (
                fetch_document(
                    linked_url
                )
            )

            if not linked_text:
                continue

            source_urls.append(
                linked_final_url
            )

            documents.append(
                "--- SUPPORTING DOCUMENT ---\n"
                + linked_text
            )

        except Exception as exc:

            log(
                "        Supporting document failed: "
                f"{exc}"
            )

    combined = clean_text(
        "\n\n".join(documents)
    )

    if not combined:

        raise RuntimeError(
            f"No usable source text for {label}."
        )

    return {
        "label": label,
        "primaryUrl": final_url,
        "sourceUrls": list(
            dict.fromkeys(
                source_urls
            )
        ),
        "text": truncate_text(
            combined,
            MAX_SOURCE_CHARS,
        ),
        "candidateLinks": ranked_links,
    }


# ============================================================
# PRUDENTIAL PRIMARY SOURCE
# ============================================================

def fetch_prudential_document(
    url: str,
) -> Tuple[str, List[str]]:

    bundle = fetch_source_bundle(
        url=url,
        label="Prudential primary fund",
    )

    return (
        bundle["text"],
        bundle["sourceUrls"],
    )


# ============================================================
# EXPOSURE SECTION EXTRACTION
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
    r"collective investment",
    r"investment objective",
    r"investment strategy",
    r"investment approach",
    r"country",
    r"region",
    r"sector",
]


def extract_relevant_sections(
    source_text: str,
    maximum: int = MAX_AI_SOURCE_CHARS,
) -> str:

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

                start = max(
                    0,
                    index - 15,
                )

                end = min(
                    len(lines),
                    index + 55,
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
        "\n\n"
        "--- RELEVANT DOCUMENT SECTION ---\n\n"
        .join(chunks)
    )

    return truncate_text(
        relevant,
        maximum,
    )


def prepare_ai_source(
    source_text: str,
) -> str:

    relevant = extract_relevant_sections(
        source_text
    )

    if relevant:

        log(
            "        Using relevant financial "
            f"sections: {len(relevant)} chars"
        )

        return relevant

    log(
        "        No specific exposure sections "
        "detected; using full source."
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
# AI GENERATION
# ============================================================

def generate_ai_response(
    messages: List[Dict[str, str]],
) -> str:

    if TOKENIZER is None:
        raise RuntimeError(
            "Tokenizer is not loaded."
        )

    if MODEL is None:
        raise RuntimeError(
            "Model is not loaded."
        )

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
# STAGE 1 - FUND STRUCTURE ANALYSIS
# ============================================================

def build_structure_prompt(
    fund_name: str,
    source_text: str,
) -> str:

    return f"""
You are a financial fund research analyst.

Determine the structure of this fund using ONLY the supplied
documentation.

Fund:
{fund_name}

Possible structures:

1. "direct_investment"
   The fund directly invests in securities such as equities,
   bonds, cash or other securities.

2. "fund_of_funds"
   The fund primarily invests through one or more underlying
   investment funds, collective investment schemes, unit trusts,
   mutual funds, ETFs or similar pooled vehicles.

3. "mixed"
   The fund contains both direct investments and meaningful
   underlying investment funds.

4. "unknown"
   The documentation does not provide enough evidence.

For fund_of_funds or mixed:

Extract the names of the underlying funds ONLY when they are
explicitly supported by the document.

Do not invent names.

Return ONLY JSON.

Required format:

{{
  "structure": "direct_investment",
  "underlyingFunds": []
}}

Allowed structure values:

- direct_investment
- fund_of_funds
- mixed
- unknown

Maximum underlying funds:
{MAX_UNDERLYING_FUNDS}

DOCUMENTATION:

{source_text}
""".strip()


def analyze_fund_structure(
    fund_name: str,
    source_text: str,
) -> Tuple[
    bool,
    str,
    List[str],
    str,
]:

    prompt = build_structure_prompt(
        fund_name,
        source_text,
    )

    messages = [
        {
            "role": "system",
            "content": (
                "You are a precise financial "
                "document analyst. "
                "Return only valid JSON."
            ),
        },
        {
            "role": "user",
            "content": prompt,
        },
    ]

    try:

        response = generate_ai_response(
            messages
        )

    except Exception as exc:

        return (
            False,
            "ai_error",
            [],
            str(exc),
        )

    log("")
    log(
        "      Structure AI response:"
    )
    log(response)

    parsed = extract_json_object(
        response
    )

    if parsed is None:

        return (
            False,
            "ai_error",
            [],
            "Structure response "
            "could not be parsed as JSON.",
        )

    structure = parsed.get(
        "structure"
    )

    underlying = parsed.get(
        "underlyingFunds"
    )

    allowed = {
        "direct_investment",
        "fund_of_funds",
        "mixed",
        "unknown",
    }

    if structure not in allowed:

        return (
            False,
            "ai_error",
            [],
            f"Invalid fund structure: "
            f"{structure}",
        )

    if not isinstance(
        underlying,
        list,
    ):

        return (
            False,
            "ai_error",
            [],
            "underlyingFunds is not a list.",
        )

    clean_underlying: List[str] = []

    for item in underlying:

        if not isinstance(
            item,
            str,
        ):
            continue

        item = clean_text(
            item
        )

        if not item:
            continue

        if item not in clean_underlying:

            clean_underlying.append(
                item
            )

    clean_underlying = (
        clean_underlying[
            :MAX_UNDERLYING_FUNDS
        ]
    )

    if structure in {
        "fund_of_funds",
        "mixed",
    } and not clean_underlying:

        return (
            False,
            "insufficient_evidence",
            [],
            "Fund structure indicates "
            "underlying funds but no "
            "underlying fund names were "
            "identified.",
        )

    return (
        True,
        structure,
        clean_underlying,
        "",
    )


# ============================================================
# UNDERLYING FUND SOURCE DISCOVERY
# ============================================================

def discover_underlying_source(
    underlying_name: str,
    parent_source_bundle: Dict[str, Any],
    workbook_index: Dict[str, Dict[str, str]],
) -> List[Dict[str, Any]]:

    sources: List[Dict[str, Any]] = []

    # --------------------------------------------------------
    # 1. Match against Funds Links.xlsx.
    # --------------------------------------------------------

    workbook_match = find_workbook_fund_match(
        underlying_name,
        workbook_index,
    )

    if workbook_match:

        log(
            "        Workbook match for underlying "
            f"fund: {underlying_name}"
        )

        try:

            bundle = fetch_source_bundle(
                url=workbook_match[
                    "prudentialUrl"
                ],
                label=(
                    "underlying fund: "
                    + underlying_name
                ),
                max_documents=(
                    MAX_DOCUMENTS_PER_UNDERLYING
                ),
            )

            sources.append(
                bundle
            )

        except Exception as exc:

            log(
                "        Workbook-linked source "
                f"failed: {exc}"
            )

    # --------------------------------------------------------
    # 2. Search already discovered parent links.
    # --------------------------------------------------------

    target_tokens = name_tokens(
        underlying_name
    )

    candidate_links = (
        parent_source_bundle.get(
            "candidateLinks",
            [],
        )
    )

    scored_candidates = []

    for candidate in candidate_links:

        candidate_text = (
            f"{candidate.get('url', '')} "
            f"{candidate.get('text', '')}"
        )

        candidate_tokens = name_tokens(
            candidate_text
        )

        intersection = (
            target_tokens
            & candidate_tokens
        )

        if len(intersection) >= 2:

            score = (
                len(intersection)
                / max(
                    len(target_tokens),
                    1,
                )
            )

            scored_candidates.append(
                (
                    score,
                    candidate,
                )
            )

    scored_candidates.sort(
        key=lambda item: -item[0]
    )

    seen_urls = {
        item.get("primaryUrl")
        for item in sources
    }

    for _, candidate in scored_candidates[
        :MAX_DOCUMENTS_PER_UNDERLYING
    ]:

        candidate_url = candidate[
            "url"
        ]

        if candidate_url in seen_urls:
            continue

        try:

            log(
                "        Matching document for "
                f"{underlying_name}: "
                f"{candidate_url}"
            )

            bundle = fetch_source_bundle(
                url=candidate_url,
                label=(
                    "underlying fund: "
                    + underlying_name
                ),
                max_documents=2,
            )

            sources.append(
                bundle
            )

            seen_urls.add(
                candidate_url
            )

        except Exception as exc:

            log(
                "        Underlying candidate "
                f"failed: {exc}"
            )

    return sources


# ============================================================
# COMBINE EXPOSURE EVIDENCE
# ============================================================

def build_evidence_package(
    parent_source: str,
    underlying_sources: List[
        Dict[str, Any]
    ],
) -> str:

    sections: List[str] = []

    parent_relevant = (
        prepare_ai_source(
            parent_source
        )
    )

    sections.append(
        "=== PARENT FUND EVIDENCE ===\n"
        + parent_relevant
    )

    for index, bundle in enumerate(
        underlying_sources,
        start=1,
    ):

        underlying_text = bundle.get(
            "text",
            "",
        )

        if not underlying_text:
            continue

        relevant = (
            extract_relevant_sections(
                underlying_text,
                maximum=(
                    MAX_AI_SOURCE_CHARS
                    // max(
                        len(
                            underlying_sources
                        ),
                        1,
                    )
                ),
            )
        )

        if not relevant:
            relevant = truncate_text(
                underlying_text,
                max(
                    6000,
                    MAX_AI_SOURCE_CHARS
                    // max(
                        len(
                            underlying_sources
                        ),
                        1,
                    ),
                ),
            )

        sections.append(
            f"=== UNDERLYING SOURCE "
            f"{index} ===\n"
            + relevant
        )

    combined = clean_text(
        "\n\n".join(
            sections
        )
    )

    return truncate_text(
        combined,
        MAX_AI_SOURCE_CHARS,
    )


# ============================================================
# STAGE 2 - EXPOSURE ANALYSIS
# ============================================================

def build_exposure_prompt(
    fund_name: str,
    structure: str,
    underlying_names: List[str],
    evidence: str,
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

    underlying_text = (
        "\n".join(
            f"- {item}"
            for item in underlying_names
        )
        if underlying_names
        else "- None identified"
    )

    return f"""
You are a professional financial fund research analyst.

Analyze the supplied evidence for:

Fund:
{fund_name}

Fund structure:
{structure}

Identified underlying funds:
{underlying_text}

Your task is to determine the fund's supported:

1. Geographic exposure
2. Sector exposure

IMPORTANT:

The fund may be either a direct-investment fund or a fund-of-funds.

For a direct-investment fund:
    Analyze the fund's own documented portfolio evidence.

For a fund-of-funds:
    Analyze the documented evidence for the underlying funds
    and the parent fund.
    Look through to the underlying portfolio information when
    that information is supplied.

For a mixed fund:
    Analyze both direct and underlying-fund evidence.

STRICT EVIDENCE RULES:

- Use ONLY the supplied evidence.
- Do not use outside knowledge.
- Do not browse the internet.
- Do not use the fund name as evidence.
- Do not guess.
- Do not invent holdings.
- Do not invent countries.
- Do not invent sectors.
- Do not assume a company's geography or sector unless the
  supplied evidence supports it.
- Do not convert a fund name into an exposure without supporting
  portfolio evidence.
- Do not assume an underlying fund's portfolio merely from its
  name.
- Use country allocation, geographical allocation, regional
  allocation, sector allocation, industry allocation, holdings,
  asset allocation, investment strategy and other explicit
  portfolio evidence where available.
- If allocation weights are supplied, use them to understand
  which exposures are meaningful.
- Do NOT output percentages.
- Do NOT put percentages into category names.
- Maximum 3 geography categories.
- Maximum 3 sector categories.
- Fewer than 3 is allowed.
- Use ONLY canonical category names.
- Never create a new category.
- Return ONLY JSON.
- No explanation.
- No markdown.

CANONICAL GEOGRAPHY CATEGORIES:

{geography_text}

CANONICAL SECTOR CATEGORIES:

{sector_text}

OUTPUT:

{{
  "geography": [],
  "sector": []
}}

SUPPLIED EVIDENCE:

{evidence}
""".strip()


def analyze_exposure(
    fund_name: str,
    structure: str,
    underlying_names: List[str],
    evidence: str,
    geography_categories: List[str],
    sector_categories: List[str],
) -> Tuple[
    bool,
    List[str],
    List[str],
    str,
]:

    prompt = build_exposure_prompt(
        fund_name=fund_name,
        structure=structure,
        underlying_names=underlying_names,
        evidence=evidence,
        geography_categories=(
            geography_categories
        ),
        sector_categories=(
            sector_categories
        ),
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

        response = generate_ai_response(
            messages
        )

    except Exception as exc:

        return (
            False,
            [],
            [],
            str(exc),
        )

    log("")
    log(
        "      Exposure AI response:"
    )
    log(response)

    parsed = extract_json_object(
        response
    )

    if parsed is None:

        return (
            False,
            [],
            [],
            "Exposure response could "
            "not be parsed as JSON.",
        )

    valid, error, geography, sector = (
        validate_ai_result(
            parsed,
            geography_categories,
            sector_categories,
        )
    )

    if not valid:

        return (
            False,
            geography,
            sector,
            error,
        )

    return (
        True,
        geography,
        sector,
        "",
    )


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
                "Geography contains "
                "non-string value.",
                [],
                [],
            )

        category = category.strip()

        if not category:

            return (
                False,
                "Geography contains "
                "empty category.",
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
                "Sector contains "
                "non-string value.",
                [],
                [],
            )

        category = category.strip()

        if not category:

            return (
                False,
                "Sector contains "
                "empty category.",
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

    if (
        not clean_geography
        and not clean_sector
    ):

        return (
            False,
            "AI found no valid geography "
            "or sector exposure.",
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
    workbook_index: Dict[str, Dict[str, str]],
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
    log("-" * 70)
    log(
        f"[{index}/{total}] "
        f"{fund_name}"
    )
    log("-" * 70)

    # --------------------------------------------------------
    # Parent source
    # --------------------------------------------------------

    try:

        parent_bundle = fetch_source_bundle(
            url=prudential_url,
            label="parent fund",
        )

    except Exception as exc:

        log(
            f"        Source failed: {exc}"
        )

        return {
            "fundName": fund_name,
            "prudentialUrl": prudential_url,
            "geography": [],
            "sector": [],
            "status": "source_error",
            "error": str(exc),
            "structure": "unknown",
            "underlyingFunds": [],
            "sources": [
                prudential_url
            ],
        }

    parent_source = parent_bundle[
        "text"
    ]

    log(
        f"        Parent source characters: "
        f"{len(parent_source)}"
    )

    # --------------------------------------------------------
    # Diagnostic preview
    # --------------------------------------------------------

    parent_preview = prepare_ai_source(
        parent_source
    )

    log("")
    log(
        "        PARENT SOURCE PREVIEW:"
    )
    log(
        "        "
        + "-" * 60
    )

    for line in parent_preview[:4000].splitlines():

        log(
            f"        {line}"
        )

    log(
        "        "
        + "-" * 60
    )

    # --------------------------------------------------------
    # Stage 1: structure
    # --------------------------------------------------------

    log("")
    log(
        "      Stage 1 - detecting fund structure..."
    )

    structure_ok, structure, underlying_names, (
        structure_error
    ) = analyze_fund_structure(
        fund_name=fund_name,
        source_text=parent_preview,
    )

    if not structure_ok:

        log(
            f"        Structure analysis failed: "
            f"{structure_error}"
        )

        return {
            "fundName": fund_name,
            "prudentialUrl": prudential_url,
            "geography": [],
            "sector": [],
            "status": structure,
            "error": structure_error,
            "structure": structure,
            "underlyingFunds": underlying_names,
            "sources": parent_bundle[
                "sourceUrls"
            ],
        }

    log(
        f"        Detected structure: "
        f"{structure}"
    )

    if underlying_names:

        log(
            "        Underlying funds:"
        )

        for name in underlying_names:

            log(
                f"          - {name}"
            )

    # --------------------------------------------------------
    # Stage 2: underlying evidence
    # --------------------------------------------------------

    underlying_bundles: List[
        Dict[str, Any]
    ] = []

    unresolved_underlying: List[str] = []

    if structure in {
        "fund_of_funds",
        "mixed",
    }:

        for underlying_name in underlying_names:

            log("")
            log(
                "      Researching underlying fund: "
                f"{underlying_name}"
            )

            found = (
                discover_underlying_source(
                    underlying_name=(
                        underlying_name
                    ),
                    parent_source_bundle=(
                        parent_bundle
                    ),
                    workbook_index=(
                        workbook_index
                    ),
                )
            )

            if found:

                underlying_bundles.extend(
                    found
                )

            else:

                unresolved_underlying.append(
                    underlying_name
                )

        if unresolved_underlying:

            log("")
            log(
                "        Unresolved underlying "
                "fund evidence:"
            )

            for name in unresolved_underlying:

                log(
                    f"          - {name}"
                )

    # --------------------------------------------------------
    # Evidence package
    # --------------------------------------------------------

    evidence = build_evidence_package(
        parent_source=parent_source,
        underlying_sources=(
            underlying_bundles
        ),
    )

    log("")
    log(
        f"        Final AI evidence characters: "
        f"{len(evidence)}"
    )

    log("")
    log(
        "        FINAL AI EVIDENCE PREVIEW:"
    )
    log(
        "        "
        + "-" * 60
    )

    for line in evidence[:5000].splitlines():

        log(
            f"        {line}"
        )

    log(
        "        "
        + "-" * 60
    )

    # --------------------------------------------------------
    # Important evidence rule.
    #
    # For fund-of-funds:
    # If the parent source identifies underlying funds but
    # provides no useful portfolio evidence and none of the
    # underlying funds can be sourced, do not let Qwen guess.
    # --------------------------------------------------------

    if structure in {
        "fund_of_funds",
        "mixed",
    }:

        parent_relevant = (
            extract_relevant_sections(
                parent_source,
                maximum=20000,
            )
        )

        has_underlying_evidence = (
            len(underlying_bundles) > 0
        )

        if (
            not parent_relevant
            and not has_underlying_evidence
        ):

            return {
                "fundName": fund_name,
                "prudentialUrl": prudential_url,
                "geography": [],
                "sector": [],
                "status": "insufficient_evidence",
                "error": (
                    "Fund-of-funds identified, "
                    "but no usable underlying "
                    "portfolio evidence was "
                    "available."
                ),
                "structure": structure,
                "underlyingFunds": underlying_names,
                "unresolvedUnderlyingFunds": (
                    unresolved_underlying
                ),
                "sources": parent_bundle[
                    "sourceUrls"
                ],
            }

    # --------------------------------------------------------
    # Stage 2: exposure analysis
    # --------------------------------------------------------

    log("")
    log(
        "      Stage 2 - researching geography "
        "and sector..."
    )

    (
        exposure_ok,
        geography,
        sector,
        exposure_error,
    ) = analyze_exposure(
        fund_name=fund_name,
        structure=structure,
        underlying_names=underlying_names,
        evidence=evidence,
        geography_categories=(
            geography_categories
        ),
        sector_categories=(
            sector_categories
        ),
    )

    if not exposure_ok:

        log(
            f"        Exposure analysis failed: "
            f"{exposure_error}"
        )

        return {
            "fundName": fund_name,
            "prudentialUrl": prudential_url,
            "geography": geography,
            "sector": sector,
            "status": "ai_error",
            "error": exposure_error,
            "structure": structure,
            "underlyingFunds": underlying_names,
            "unresolvedUnderlyingFunds": (
                unresolved_underlying
            ),
            "sources": list(
                dict.fromkeys(
                    parent_bundle[
                        "sourceUrls"
                    ]
                    + [
                        url
                        for bundle in (
                            underlying_bundles
                        )
                        for url in bundle.get(
                            "sourceUrls",
                            [],
                        )
                    ]
                )
            ),
        }

    # --------------------------------------------------------
    # Success
    # --------------------------------------------------------

    log("")
    log(
        f"        Geography: {geography}"
    )

    log(
        f"        Sector: {sector}"
    )

    log(
        "        Status: success"
    )

    all_sources = (
        parent_bundle[
            "sourceUrls"
        ]
        + [
            url
            for bundle in underlying_bundles
            for url in bundle.get(
                "sourceUrls",
                [],
            )
        ]
    )

    return {
        "fundName": fund_name,
        "prudentialUrl": prudential_url,
        "geography": geography,
        "sector": sector,
        "status": "success",
        "structure": structure,
        "underlyingFunds": underlying_names,
        "unresolvedUnderlyingFunds": (
            unresolved_underlying
        ),
        "sources": list(
            dict.fromkeys(
                all_sources
            )
        ),
    }


# ============================================================
# AI HEALTH CHECK
# ============================================================

def ai_health_check() -> bool:

    log("")
    log("=" * 70)
    log("AI HEALTH CHECK")
    log("=" * 70)

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
            "content": (
                "Return exactly this JSON schema "
                "with empty arrays. "
                "Do not add explanation:\n"
                "{"
                "\"geography\": [], "
                "\"sector\": []"
                "}"
            ),
        },
    ]

    try:

        response = generate_ai_response(
            messages
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

        if (
            not geography
            and not sector
        ):

            errors.append(
                f"{name}: no geography "
                "or sector exposure."
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
            verified.get("funds"),
            list,
        ):

            raise RuntimeError(
                "Temporary output "
                "funds is not a list."
            )

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
    log("=" * 70)
    log(
        "VGRAT FMS - AI FUND EXPOSURE RESEARCH"
    )
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
        "Adaptive structure: ENABLED"
    )

    log(
        f"Output: {OUTPUT_PATH}"
    )

    log("=" * 70)

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
            f"ERROR: Fund loading failed: "
            f"{exc}"
        )

        log(
            "Existing fund_exposure.json "
            "has NOT been modified."
        )

        return 1

    workbook_index = (
        build_workbook_name_index(
            funds
        )
    )

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

    if not ai_health_check():

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
            workbook_index=(
                workbook_index
            ),
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
    # Final validation
    # --------------------------------------------------------

    log("")
    log("=" * 70)
    log("FINAL VALIDATION")
    log("=" * 70)

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
            "ERROR: Output publishing failed: "
            f"{exc}"
        )

        log(
            "Existing fund_exposure.json "
            "was preserved."
        )

        return 1

    # --------------------------------------------------------
    # Success
    # --------------------------------------------------------

    log("")
    log("=" * 70)
    log("RESEARCH SUCCESS")
    log("=" * 70)

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

    log("=" * 70)

    return 0


if __name__ == "__main__":
    sys.exit(
        main()
    )
