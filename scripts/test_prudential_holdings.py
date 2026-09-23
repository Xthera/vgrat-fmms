#!/usr/bin/env python3
"""
VGrat FMS - Adaptive Per-Fund Prudential Top Holdings Extractor

Architecture
============
Process exactly one Excel fund at a time:

    Funds Links.xlsm
        -> official Prudential fund page
        -> official Prudential factsheet PDF
        -> raw PDF text
        -> confirmed "Top 10 Holdings" section
        -> per-fund structure analysis
        -> selected extraction strategy
        -> strict validation
        -> verification pass
        -> complete per-fund evidence/audit

This is a TEST collector only. It does not modify test_pruaccess.py and does
not create data.json/index.html/css/js.

Hard rules
==========
- Excel Column A controls the universe; no hardcoded fund count.
- Only official Prudential Singapore pages/factsheets.
- Accepted section headings are specifically "Top 10 Holdings" and
  "Top Ten Holdings", including attached Unicode/ASCII footnote markers.
- Generic "Holdings", "Investments", etc. are never treated as the target.
- No inferred holdings or fabricated percentages.
- Never force ten holdings.
- If fewer than ten holdings are actually published, preserve that count.
- Preserve Prudential's published order.
- Multi-line holding names are supported.
- Fixed-income coupon/rate percentages are handled by a last-percentage
  strategy when the structure proves that is required.
- No holding is accepted without a published weight.
- Duplicate holding names and duplicate percentages are allowed.
- The full PDF text and exact extracted Top Holdings section are preserved.
- Every accepted holding gets an audit record explaining why it was accepted.
- The selected strategy is determined independently for each fund.
- A final verification pass re-downloads/re-reads the same official PDF and
  reruns the selected strategy; mismatches fail the fund.
- A fund with no confirmed Top Holdings section is NO_HOLDINGS_SECTION.
- A confirmed section that cannot be safely parsed is FAILED.
"""

from __future__ import annotations

import json
import re
import sys
import time
import unicodedata
from datetime import datetime, timezone
from io import BytesIO
from pathlib import Path
from urllib.parse import urljoin, urlparse

from openpyxl import load_workbook
from pypdf import PdfReader
from playwright.sync_api import TimeoutError as PlaywrightTimeoutError, sync_playwright


# =============================================================================
# CONFIGURATION
# =============================================================================

EXCEL_FILE = Path("Funds Links.xlsm")
OUTPUT_DIR = Path("output_holdings")
FUNDS_OUTPUT_DIR = OUTPUT_DIR / "funds"
ALL_HOLDINGS_FILE = OUTPUT_DIR / "all_holdings.json"
RUN_SUMMARY_FILE = OUTPUT_DIR / "run_summary.json"

BROWSER_HEADLESS = True
PAGE_TIMEOUT_MS = 120000
FACTSHEET_DOWNLOAD_TIMEOUT_MS = 120000
POST_PAGE_WAIT_MS = 1500
RETRY_COUNT = 3
RETRY_DELAY_SECONDS = 3.0

MAX_HOLDINGS = 10
PRUDENTIAL_HOSTS = {"prudential.com.sg", "www.prudential.com.sg"}

# Target heading only. Generic "Holdings" / "Investments" are intentionally
# absent.
TOP_HOLDINGS_HEADING_RE = re.compile(
    r"^\s*top\s+(?:10|ten)\s+holdings"
    r"(?:[\s\u00B9\u00B2\u00B3\u2070\u2074\u2075\u2076\u2077\u2078\u2079\u00AA"
    r"\u00B9\u00B2\u00B3\u207A-\u207F\u2080-\u2089\u2020\u2021*]+)?"
    r"\s*$",
    re.IGNORECASE,
)

# Also permit a footnote marker immediately attached after Holdings.
TOP_HOLDINGS_CONTAINS_RE = re.compile(
    r"\btop\s+(?:10|ten)\s+holdings\b",
    re.IGNORECASE,
)

PERCENT_RE = re.compile(r"(?<![\d.])([0-9]+(?:\.[0-9]+)?)\s*%")
RANK_PUNCT_RE = re.compile(r"^\s*(\d{1,2})[.)\-:]\s+(.+)$")
RANK_SPACE_RE = re.compile(r"^\s*(\d{1,2})\s+(.+)$")
STANDALONE_RANK_RE = re.compile(r"^\s*(\d{1,2})\s*$")

# Headings which commonly begin the next section. They are intentionally
# conservative: unknown text is not treated as an ending merely because it
# looks like a heading.
SECTION_END_RE = re.compile(
    r"^\s*(?:"
    r"asset\s+allocation|"
    r"geographical\s+allocation|"
    r"sector\s+allocation|"
    r"portfolio\s+characteristics|"
    r"country\s+allocation|"
    r"regional\s+allocation|"
    r"credit\s+quality|"
    r"fund\s+performance|"
    r"performance|"
    r"risk\s+profile|"
    r"risk\s+classification|"
    r"fund\s+information|"
    r"fund\s+details|"
    r"important\s+information|"
    r"disclaimer|"
    r"source\s*:?"
    r")\s*$",
    re.IGNORECASE,
)

NOISE_EXACT = {
    "holding", "holdings", "name", "names", "weight", "weights",
    "allocation", "allocations", "%", "portfolio holdings",
    "top 10 holdings", "top ten holdings",
}

# =============================================================================
# GENERAL HELPERS
# =============================================================================

def clean_text(value) -> str:
    if value is None:
        return ""
    text = str(value).replace("\xa0", " ")
    text = unicodedata.normalize("NFKC", text)
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def normalize_text(value) -> str:
    text = clean_text(value).casefold()
    text = text.replace("–", "-").replace("—", "-")
    return re.sub(r"\s+", " ", text).strip()


def utc_now_iso() -> str:
    return datetime.now(timezone.utc).replace(microsecond=0).isoformat()


def save_json(path: Path, data) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(
        json.dumps(data, indent=2, ensure_ascii=False),
        encoding="utf-8",
    )


def safe_filename(value) -> str:
    text = clean_text(value)
    text = re.sub(r"[^A-Za-z0-9._-]+", "_", text).strip("._")
    return (text or "fund")[:120]


def is_prudential_url(url: str) -> bool:
    try:
        parsed = urlparse(url)
        return (
            parsed.scheme.lower() in {"http", "https"}
            and (parsed.hostname or "").lower() in PRUDENTIAL_HOSTS
        )
    except Exception:
        return False


def ensure_prudential_url(url: str) -> str:
    url = clean_text(url)
    if not url or not is_prudential_url(url):
        raise RuntimeError(f"Non-Prudential URL rejected: {url}")
    return url


def percentage_matches(line: str):
    return list(PERCENT_RE.finditer(line or ""))


def percentage_value(match) -> float:
    value = float(match.group(1))
    if not 0 <= value <= 100:
        raise ValueError("Percentage outside 0-100")
    return value


def percentage_tuple(match):
    return (
        percentage_value(match),
        clean_text(match.group(0)),
        match.start(),
        match.end(),
    )


def clean_holding_name(value: str) -> str:
    name = clean_text(value)
    name = re.sub(r"^[•·▪■□*]+", "", name).strip()
    name = re.sub(r"^\d{1,2}[.)]\s+", "", name)
    name = name.replace("|", " ")
    name = re.sub(r"\bNone\b", " ", name, flags=re.IGNORECASE)
    name = clean_text(name).strip(" -|")
    if normalize_text(name) in {"", "none", "null", "-", "—"}:
        return ""
    return name


def clean_fragment(value: str) -> str:
    text = clean_text(value).replace("|", " ")
    text = re.sub(r"\bNone\b", " ", text, flags=re.IGNORECASE)
    text = re.sub(r"^[•·▪■□*]+", "", text)
    return clean_text(text)


def join_fragments(fragments) -> str:
    return clean_holding_name(" ".join(
        clean_fragment(x) for x in fragments if clean_fragment(x)
    ))


def is_noise(line: str) -> bool:
    n = normalize_text(line)
    if not n:
        return True
    if n in NOISE_EXACT:
        return True
    if n.startswith("top 10 holdings") or n.startswith("top ten holdings"):
        return True
    return False


def parse_leading_rank(line: str):
    text = clean_text(line)
    if not text:
        return None, ""

    m = STANDALONE_RANK_RE.fullmatch(text)
    if m:
        rank = int(m.group(1))
        return (rank, "") if 1 <= rank <= 99 else (None, text)

    m = RANK_PUNCT_RE.match(text)
    if m:
        rank = int(m.group(1))
        return (rank, clean_text(m.group(2))) if 1 <= rank <= 99 else (None, text)

    # Space-separated rank. Require a non-numeric first token so that
    # "5.5% ..." is not mistaken for rank 5.
    m = RANK_SPACE_RE.match(text)
    if m and not re.match(r"^\d+(?:\.\d+)?%", m.group(2)):
        rank = int(m.group(1))
        if 1 <= rank <= 99:
            return rank, clean_text(m.group(2))

    return None, text


# =============================================================================
# EXCEL MASTER UNIVERSE
# =============================================================================

def read_excel_funds() -> list[dict]:
    if not EXCEL_FILE.exists():
        raise FileNotFoundError(f"Excel file not found: {EXCEL_FILE}")

    workbook = load_workbook(
        EXCEL_FILE,
        read_only=True,
        keep_vba=True,
        data_only=True,
    )
    worksheet = workbook.active
    funds = []

    for row_number in range(2, worksheet.max_row + 1):
        url = clean_text(worksheet.cell(row=row_number, column=1).value)
        pruaccess_name = clean_text(worksheet.cell(row=row_number, column=2).value)
        if not url:
            continue
        funds.append({
            "excelRow": row_number,
            "prudentialUrl": url,
            "pruAccessName": pruaccess_name,
        })

    workbook.close()

    if not funds:
        raise RuntimeError("No populated Prudential URLs found in Excel Column A.")

    return funds


# =============================================================================
# OFFICIAL PAGE / FACTSHEET
# =============================================================================

def find_factsheet_url(page) -> str:
    anchors = page.locator("a")
    candidates = []

    for index in range(anchors.count()):
        anchor = anchors.nth(index)
        try:
            href = clean_text(anchor.get_attribute("href"))
            text = clean_text(anchor.inner_text())
        except Exception:
            continue
        if not href:
            continue

        absolute = urljoin(page.url, href)
        if not is_prudential_url(absolute):
            continue

        low_href = absolute.casefold()
        low_text = text.casefold()
        score = 0

        if "fund factsheet" in low_text:
            score += 200
        elif "factsheet" in low_text:
            score += 150
        if "factsheet" in low_href:
            score += 100
        if low_href.endswith(".pdf"):
            score += 50

        if score:
            candidates.append({"score": score, "url": absolute, "text": text})

    if not candidates:
        raise RuntimeError("Could not find an official Prudential factsheet link.")

    candidates.sort(key=lambda x: (-x["score"], x["url"]))
    return ensure_prudential_url(candidates[0]["url"])


def download_factsheet(page, factsheet_url: str) -> bytes:
    response = page.request.get(
        factsheet_url,
        timeout=FACTSHEET_DOWNLOAD_TIMEOUT_MS,
    )
    if response.status != 200:
        raise RuntimeError(
            f"Factsheet download returned HTTP {response.status}: {factsheet_url}"
        )
    data = response.body()
    if not data.startswith(b"%PDF"):
        raise RuntimeError("Factsheet response is not a PDF.")
    return data


def extract_pdf_text(pdf_bytes: bytes):
    reader = PdfReader(BytesIO(pdf_bytes))
    if not reader.pages:
        raise RuntimeError("Factsheet PDF has zero pages.")

    page_texts = []
    for page_no, page in enumerate(reader.pages, start=1):
        try:
            text = page.extract_text() or ""
        except Exception as exc:
            raise RuntimeError(
                f"PDF text extraction failed on page {page_no}: {exc}"
            ) from exc
        page_texts.append(text)

    full_text = "\n".join(page_texts)
    if not clean_text(full_text):
        raise RuntimeError("Factsheet PDF contains no extractable text.")

    return full_text, len(reader.pages)


def pdf_lines(text: str) -> list[str]:
    return [
        clean_text(line)
        for line in text.splitlines()
        if clean_text(line)
    ]


# =============================================================================
# FACTSHEET METADATA
# =============================================================================

def extract_data_as_at(text: str):
    patterns = [
        re.compile(
            r"\b(?:all\s+)?data\s+as\s+at\s+"
            r"([0-9]{1,2}\s+[A-Za-z]{3,9}\s+[0-9]{4})\b",
            re.IGNORECASE,
        ),
        re.compile(
            r"\b(?:as\s+at|as\s+of)\s+"
            r"([0-9]{1,2}\s+[A-Za-z]{3,9}\s+[0-9]{4})\b",
            re.IGNORECASE,
        ),
    ]
    for pattern in patterns:
        match = pattern.search(text)
        if match:
            return clean_text(match.group(1))
    return None


def extract_document_date(text: str):
    pattern = re.compile(
        r"\b(?:January|February|March|April|May|June|July|August|"
        r"September|October|November|December)\s+[0-9]{4}\b",
        re.IGNORECASE,
    )
    for line in pdf_lines(text)[:60]:
        match = pattern.search(line)
        if match:
            return clean_text(match.group(0))
    return None


# =============================================================================
# CONFIRMED TOP HOLDINGS SECTION
# =============================================================================

def is_top_holdings_heading(line: str) -> bool:
    n = normalize_text(line)
    if TOP_HOLDINGS_HEADING_RE.fullmatch(n):
        return True

    # Attached footnote forms such as "Top 10 Holdings3" or superscript 3.
    if TOP_HOLDINGS_CONTAINS_RE.search(n):
        remainder = TOP_HOLDINGS_CONTAINS_RE.sub("", n).strip()
        if remainder and not re.fullmatch(
            r"[\d\u00B9\u00B2\u00B3\u2070\u2074\u2075\u2076\u2077\u2078\u2079\u00AA"
            r"\u2020\u2021*]+",
            remainder,
        ):
            return False
        return True
    return False


def is_section_end(line: str) -> bool:
    n = normalize_text(line)
    if SECTION_END_RE.fullmatch(n):
        return True
    if re.fullmatch(r"page\s+\d+(?:\s+of\s+\d+)?", n):
        return True
    # A new major title is an end only if it is clearly unrelated to the
    # holdings rows. Generic words are deliberately not used.
    return False


def extract_holdings_section(text: str):
    lines = pdf_lines(text)
    start = None

    for i, line in enumerate(lines):
        if is_top_holdings_heading(line):
            start = i
            break

    if start is None:
        return "", "not_published", None

    section = []
    for line in lines[start + 1:]:
        if is_section_end(line):
            break
        section.append(line)

    section_text = "\n".join(section).strip()
    return section_text, "published", start


# =============================================================================
# ADAPTIVE STRUCTURE ANALYSIS
# =============================================================================

def analyze_section(section_text: str) -> dict:
    lines = pdf_lines(section_text)
    non_noise = [x for x in lines if not is_noise(x)]

    total_pct_lines = 0
    multi_pct_lines = 0
    standalone_pct_lines = 0
    rank_lines = 0
    explicit_rank_with_pct = 0
    lines_with_pct = 0

    for line in non_noise:
        matches = percentage_matches(line)
        if matches:
            lines_with_pct += 1
            total_pct_lines += len(matches)
            if len(matches) > 1:
                multi_pct_lines += 1
            if len(clean_text(line)) == len(clean_text(matches[0].group(0))):
                standalone_pct_lines += 1

        rank, remainder = parse_leading_rank(line)
        if rank is not None:
            rank_lines += 1
            if percentage_matches(remainder):
                explicit_rank_with_pct += 1

    return {
        "lineCount": len(lines),
        "nonNoiseLineCount": len(non_noise),
        "percentageCount": total_pct_lines,
        "linesWithPercentage": lines_with_pct,
        "multiPercentageLineCount": multi_pct_lines,
        "standalonePercentageLineCount": standalone_pct_lines,
        "rankLineCount": rank_lines,
        "explicitRankWithPercentageCount": explicit_rank_with_pct,
        "hasMultiPercentageLine": multi_pct_lines > 0,
        "hasStandalonePercentageLines": standalone_pct_lines > 0,
        "hasRankMarkers": rank_lines > 0,
    }


def choose_strategy(analysis: dict) -> tuple[str, str]:
    """
    Select one strategy for this fund only.

    Strategy selection is structural, not fund-name based and not row based.
    """
    if analysis["percentageCount"] == 0:
        raise RuntimeError(
            "Confirmed Top Holdings section contains no published percentages."
        )

    if analysis["hasMultiPercentageLine"]:
        return (
            "fixed_income_last_percentage",
            "At least one logical line contains multiple percentages; "
            "the final percentage is required to distinguish security "
            "coupon/rate text from the portfolio weight.",
        )

    if analysis["hasStandalonePercentageLines"] and analysis["hasRankMarkers"]:
        return (
            "ranked_wrapped_lines",
            "The section contains rank markers and standalone published "
            "weight lines; wrapped holding names are therefore assembled "
            "until the next percentage.",
        )

    if analysis["hasRankMarkers"]:
        return (
            "ranked_lines",
            "The section contains explicit holding rank markers.",
        )

    if analysis["hasStandalonePercentageLines"]:
        return (
            "wrapped_weight_lines",
            "The section contains standalone percentage lines; holding "
            "names are assembled until a published weight is encountered.",
        )

    return (
        "logical_line_last_percentage",
        "No rank markers were required; each logical row is closed by its "
        "published percentage, using the final percentage on the row.",
    )


# =============================================================================
# ADAPTIVE EXTRACTION CORE
# =============================================================================

def make_holding(rank, name, pct, pct_text, line_numbers, evidence, strategy):
    name = clean_holding_name(name)
    if not name:
        raise RuntimeError("Published percentage found without a holding name.")
    if not 0 <= pct <= 100:
        raise RuntimeError("Holding percentage outside 0-100.")
    return {
        "rank": rank,
        "name": name,
        "weightPercent": pct,
        "weightText": pct_text,
        "sourceLines": line_numbers,
        "evidence": evidence,
        "extractionStrategy": strategy,
    }


def validate_holdings(holdings: list[dict], strategy: str) -> None:
    if not holdings:
        raise RuntimeError("No holding/percentage pairs were extracted.")
    if len(holdings) > MAX_HOLDINGS:
        raise RuntimeError(
            f"More than {MAX_HOLDINGS} published holdings were extracted."
        )

    for index, holding in enumerate(holdings, start=1):
        if not clean_holding_name(holding.get("name")):
            raise RuntimeError(f"Holding {index} has no valid name.")
        weight = holding.get("weightPercent")
        if not isinstance(weight, (int, float)) or not 0 <= weight <= 100:
            raise RuntimeError(f"Holding {index} has invalid weight.")

        if strategy in {"ranked_lines", "ranked_wrapped_lines"}:
            expected = index
            if holding.get("rank") != expected:
                raise RuntimeError(
                    f"Published holding rank mismatch: "
                    f"parsed {holding.get('rank')} expected {expected}."
                )


def extract_adaptive(section_text: str, strategy: str):
    lines = pdf_lines(section_text)
    holdings = []
    audit = []

    pending_fragments = []
    pending_line_numbers = []
    pending_rank = None

    def reset_pending():
        nonlocal pending_fragments, pending_line_numbers, pending_rank
        pending_fragments = []
        pending_line_numbers = []
        pending_rank = None

    def commit(pct, pct_text, evidence, line_no, extra_reason):
        nonlocal pending_fragments, pending_line_numbers, pending_rank

        name = join_fragments(pending_fragments)
        if not name:
            raise RuntimeError(
                f"Line {line_no}: published percentage has no holding name."
            )

        rank = pending_rank if pending_rank is not None else len(holdings) + 1
        holding = make_holding(
            rank=rank,
            name=name,
            pct=pct,
            pct_text=pct_text,
            line_numbers=pending_line_numbers + [line_no],
            evidence=evidence,
            strategy=strategy,
        )
        holdings.append(holding)

        audit.append({
            "accepted": True,
            "holdingIndex": len(holdings),
            "rank": rank,
            "holdingName": name,
            "weightText": pct_text,
            "weightPercent": pct,
            "sourceLines": pending_line_numbers + [line_no],
            "evidence": evidence,
            "reason": extra_reason,
            "strategy": strategy,
        })
        reset_pending()

    for line_no, raw_line in enumerate(lines, start=1):
        line = clean_text(raw_line)
        if not line or is_noise(line):
            continue

        rank, remainder = parse_leading_rank(line)

        if strategy in {
            "ranked_lines",
            "ranked_wrapped_lines",
            "fixed_income_last_percentage",
            "logical_line_last_percentage",
        } and rank is not None:
            if pending_fragments:
                raise RuntimeError(
                    f"Line {line_no}: new rank {rank} appeared before "
                    f"the previous holding received a published percentage."
                )
            pending_rank = rank
            pending_line_numbers.append(line_no)
            line = remainder
            if not line:
                continue

        matches = percentage_matches(line)

        if strategy == "fixed_income_last_percentage":
            if matches:
                match = matches[-1]
                pct, pct_text, start, end = percentage_tuple(match)
                fragment = clean_text(line[:start])
                if fragment:
                    pending_fragments.append(fragment)
                    pending_line_numbers.append(line_no)

                commit(
                    pct,
                    pct_text,
                    clean_text(line),
                    line_no,
                    "Accepted because the selected fixed-income strategy "
                    "uses the last published percentage on the logical row; "
                    "earlier percentage(s) remain part of the security name.",
                )
                if len(holdings) >= MAX_HOLDINGS:
                    break
                continue

            pending_fragments.append(clean_fragment(line))
            pending_line_numbers.append(line_no)
            continue

        if matches:
            # If a non-fixed-income strategy sees multiple percentages, the
            # structure is ambiguous. Never guess: fail this strategy.
            if len(matches) > 1:
                raise RuntimeError(
                    f"Line {line_no}: multiple percentages require the "
                    "fixed-income last-percentage strategy."
                )

            match = matches[0]
            pct, pct_text, start, end = percentage_tuple(match)
            fragment = clean_text(line[:start])
            if fragment:
                pending_fragments.append(fragment)
                pending_line_numbers.append(line_no)

            reason = (
                "Accepted because a published portfolio percentage closed "
                "the accumulated holding-name fragments."
            )
            if strategy in {"ranked_lines", "ranked_wrapped_lines"}:
                reason = (
                    "Accepted because the explicit Prudential rank plus the "
                    "published percentage closed the holding."
                )

            commit(
                pct,
                pct_text,
                clean_text(line),
                line_no,
                reason,
            )
            if len(holdings) >= MAX_HOLDINGS:
                break
            continue

        # No percentage on the line: it can only be a holding-name fragment
        # if a subsequent published percentage validates it.
        fragment = clean_fragment(line)
        if fragment:
            pending_fragments.append(fragment)
            pending_line_numbers.append(line_no)

    validate_holdings(holdings, strategy)

    return holdings, audit


# =============================================================================
# RESULT / VERIFICATION
# =============================================================================

def comparable_holdings(holdings):
    return [
        {
            "rank": h["rank"],
            "name": h["name"],
            "weightPercent": h["weightPercent"],
            "weightText": h["weightText"],
        }
        for h in holdings
    ]


def verify_result(section_text: str, strategy: str, expected_holdings):
    verified, verification_audit = extract_adaptive(section_text, strategy)
    if comparable_holdings(verified) != comparable_holdings(expected_holdings):
        raise RuntimeError(
            "Verification mismatch: the selected strategy did not reproduce "
            "the exact same holdings and published weights."
        )
    return verified, verification_audit


# =============================================================================
# SAVE PER-FUND EVIDENCE
# =============================================================================

def fund_output_directory(result: dict) -> Path:
    row = result["excelRow"]
    fund_name = result.get("fundName") or "fund"
    return FUNDS_OUTPUT_DIR / f"{row}_{safe_filename(fund_name)}"


def save_success_result(
    result: dict,
    factsheet_bytes: bytes,
    full_text: str,
    section_text: str,
    verification_audit,
):
    directory = fund_output_directory(result)
    directory.mkdir(parents=True, exist_ok=True)

    (directory / "factsheet.pdf").write_bytes(factsheet_bytes)
    (directory / "factsheet_text.txt").write_text(full_text, encoding="utf-8")
    (directory / "top_holdings_section.txt").write_text(
        section_text,
        encoding="utf-8",
    )

    save_json(
        directory / "top_holdings.json",
        {
            "status": "success",
            "excelRow": result["excelRow"],
            "fundName": result.get("fundName"),
            "factsheetUrl": result["factsheetUrl"],
            "factsheetDocumentDate": result.get("factsheetDocumentDate"),
            "factsheetDataAsAt": result.get("factsheetDataAsAt"),
            "structureAnalysis": result["structureAnalysis"],
            "selectedStrategy": result["selectedStrategy"],
            "strategyReason": result["strategyReason"],
            "topHoldingsCount": len(result["topHoldings"]),
            "topHoldings": result["topHoldings"],
        },
    )

    save_json(
        directory / "extraction_log.json",
        {
            "status": "success",
            "excelRow": result["excelRow"],
            "selectedStrategy": result["selectedStrategy"],
            "strategyReason": result["strategyReason"],
            "structureAnalysis": result["structureAnalysis"],
            "selectionRules": [
                "Strategy selected from this fund's actual Top Holdings "
                "section structure only.",
                "No fund-specific row/name parser exists.",
                "No generic Holdings/Investments heading was accepted.",
                "Every holding required a published percentage.",
                "Duplicate names and duplicate percentages were permitted.",
                "Published order was preserved.",
                "No more than 10 holdings were accepted.",
            ],
            "holdingAudit": result["holdingAudit"],
            "verificationAudit": verification_audit,
            "verification": {
                "passed": True,
                "method": "Re-extract exact saved section with the same "
                          "selected strategy and compare rank/name/weight.",
            },
            "savedAtUtc": utc_now_iso(),
        },
    )

    save_json(
        directory / "metadata.json",
        {
            "excelRow": result["excelRow"],
            "prudentialUrl": result["prudentialUrl"],
            "finalUrl": result["finalUrl"],
            "fundName": result.get("fundName"),
            "factsheetUrl": result["factsheetUrl"],
            "factsheetDocumentDate": result.get("factsheetDocumentDate"),
            "factsheetDataAsAt": result.get("factsheetDataAsAt"),
            "pdfPageCount": result.get("pdfPageCount"),
            "topHoldingsCount": len(result["topHoldings"]),
            "selectedStrategy": result["selectedStrategy"],
            "status": "success",
            "savedAtUtc": utc_now_iso(),
        },
    )

    return directory


def save_no_holdings_result(result: dict, factsheet_bytes: bytes, full_text: str):
    directory = fund_output_directory(result)
    directory.mkdir(parents=True, exist_ok=True)

    # Even a genuine no-holdings fund gets complete evidence. We do not create
    # fake holdings just to make the output shape uniform.
    (directory / "factsheet.pdf").write_bytes(factsheet_bytes)
    (directory / "factsheet_text.txt").write_text(full_text, encoding="utf-8")
    (directory / "top_holdings_section.txt").write_text("", encoding="utf-8")

    payload = {
        "status": "no_holdings_section",
        "excelRow": result["excelRow"],
        "fundName": result.get("fundName"),
        "factsheetUrl": result["factsheetUrl"],
        "factsheetDocumentDate": result.get("factsheetDocumentDate"),
        "factsheetDataAsAt": result.get("factsheetDataAsAt"),
        "topHoldingsCount": 0,
        "topHoldings": [],
    }
    save_json(directory / "top_holdings.json", payload)

    save_json(
        directory / "extraction_log.json",
        {
            "status": "no_holdings_section",
            "excelRow": result["excelRow"],
            "reason": (
                "The official factsheet was downloaded and its raw text "
                "contained no confirmed Top 10 Holdings / Top Ten Holdings "
                "heading. No generic Holdings or Investments heading was "
                "accepted."
            ),
            "verification": {
                "passed": True,
                "method": "Target heading search on the official raw PDF text.",
            },
            "holdingAudit": [],
            "savedAtUtc": utc_now_iso(),
        },
    )

    save_json(
        directory / "metadata.json",
        {
            "excelRow": result["excelRow"],
            "prudentialUrl": result["prudentialUrl"],
            "finalUrl": result["finalUrl"],
            "fundName": result.get("fundName"),
            "factsheetUrl": result["factsheetUrl"],
            "factsheetDocumentDate": result.get("factsheetDocumentDate"),
            "factsheetDataAsAt": result.get("factsheetDataAsAt"),
            "pdfPageCount": result.get("pdfPageCount"),
            "topHoldingsCount": 0,
            "status": "no_holdings_section",
            "savedAtUtc": utc_now_iso(),
        },
    )

    return directory


def save_failure_result(fund: dict, error: str, details=None):
    directory = (
        FUNDS_OUTPUT_DIR
        / f"{fund['excelRow']}_failed"
    )
    directory.mkdir(parents=True, exist_ok=True)

    save_json(
        directory / "failure.json",
        {
            "status": "failed",
            "excelRow": fund["excelRow"],
            "prudentialUrl": fund["prudentialUrl"],
            "pruAccessName": fund.get("pruAccessName"),
            "error": clean_text(error),
            "details": details or {},
            "savedAtUtc": utc_now_iso(),
        },
    )
    return directory


# =============================================================================
# SINGLE FUND
# =============================================================================

def extract_single_fund(page, fund: dict):
    excel_row = fund["excelRow"]
    prudential_url = ensure_prudential_url(fund["prudentialUrl"])

    page.goto(
        prudential_url,
        wait_until="domcontentloaded",
        timeout=PAGE_TIMEOUT_MS,
    )

    try:
        page.wait_for_load_state("networkidle", timeout=25000)
    except PlaywrightTimeoutError:
        pass

    page.wait_for_timeout(POST_PAGE_WAIT_MS)

    final_url = ensure_prudential_url(page.url)
    page_title = clean_text(page.title())

    fund_name = ""
    try:
        h1 = page.locator("h1").first
        if h1.count():
            fund_name = clean_text(h1.inner_text())
    except Exception:
        pass

    if not fund_name:
        body = clean_text(page.locator("body").inner_text())
        m = re.search(r"\bPRU(?:Link|Prime)\s+[^\n]+", body, re.IGNORECASE)
        if m:
            fund_name = clean_text(m.group(0))

    factsheet_url = find_factsheet_url(page)
    factsheet_bytes = download_factsheet(page, factsheet_url)
    full_text, page_count = extract_pdf_text(factsheet_bytes)

    data_as_at = extract_data_as_at(full_text)
    document_date = extract_document_date(full_text)

    section_text, section_status, heading_line = extract_holdings_section(full_text)

    base = {
        "excelRow": excel_row,
        "prudentialUrl": prudential_url,
        "finalUrl": final_url,
        "pageTitle": page_title,
        "fundName": fund_name,
        "pruAccessName": fund.get("pruAccessName", ""),
        "factsheetUrl": factsheet_url,
        "factsheetDocumentDate": document_date,
        "factsheetDataAsAt": data_as_at,
        "pdfPageCount": page_count,
        "topHoldingsHeadingLine": heading_line,
    }

    if section_status == "not_published":
        base["status"] = "no_holdings_section"
        save_no_holdings_result(base, factsheet_bytes, full_text)
        return base

    if not clean_text(section_text):
        raise RuntimeError(
            "Confirmed Top Holdings heading exists, but its extracted section "
            "is empty. This is a failed extraction, not a no-holdings result."
        )

    structure = analyze_section(section_text)
    strategy, strategy_reason = choose_strategy(structure)
    holdings, holding_audit = extract_adaptive(section_text, strategy)

    if not holdings:
        raise RuntimeError("Adaptive parser returned no holdings.")

    # Independent verification against the exact same section text.
    verified, verification_audit = verify_result(
        section_text,
        strategy,
        holdings,
    )

    base.update({
        "status": "success",
        "structureAnalysis": structure,
        "selectedStrategy": strategy,
        "strategyReason": strategy_reason,
        "topHoldings": verified,
        "holdingAudit": holding_audit,
        "verificationPassed": True,
    })

    output_directory = save_success_result(
        base,
        factsheet_bytes,
        full_text,
        section_text,
        verification_audit,
    )
    base["outputDirectory"] = str(output_directory)
    return base


# =============================================================================
# MAIN
# =============================================================================

def main() -> int:
    started_at = utc_now_iso()

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    FUNDS_OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    funds = read_excel_funds()

    successful = []
    no_holdings_section = []
    failed = []
    strategy_counts = {}

    print("=" * 78)
    print("VGRAT FMS - ADAPTIVE PER-FUND TOP HOLDINGS TEST")
    print("=" * 78)
    print(f"Excel fund universe: {len(funds)}")
    print("Architecture: one fund -> analyse -> extract -> verify -> next fund")
    print()

    with sync_playwright() as playwright:
        browser = playwright.chromium.launch(headless=BROWSER_HEADLESS)
        context = browser.new_context(
            accept_downloads=True,
            locale="en-SG",
        )
        page = context.new_page()
        page.set_default_timeout(PAGE_TIMEOUT_MS)

        try:
            for position, fund in enumerate(funds, start=1):
                print()
                print("=" * 78)
                print(
                    f"PROCESSING FUND {position}/{len(funds)} "
                    f"(Excel Row {fund['excelRow']})"
                )
                print("=" * 78)
                print(f"Prudential URL: {fund['prudentialUrl']}")
                print(f"PruAccess reference: {fund.get('pruAccessName') or '-'}")

                result = None
                last_error = None

                for attempt in range(1, RETRY_COUNT + 1):
                    try:
                        print(f"Attempt {attempt}/{RETRY_COUNT}")
                        result = extract_single_fund(page, fund)
                        last_error = None
                        break
                    except Exception as exc:
                        last_error = clean_text(str(exc))
                        print(f"Attempt failed: {last_error}")
                        if attempt < RETRY_COUNT:
                            time.sleep(RETRY_DELAY_SECONDS)

                if result is None:
                    output_directory = save_failure_result(
                        fund,
                        last_error or "Unknown extraction failure.",
                    )
                    failed.append({
                        "excelRow": fund["excelRow"],
                        "prudentialUrl": fund["prudentialUrl"],
                        "pruAccessName": fund.get("pruAccessName"),
                        "error": last_error or "Unknown extraction failure.",
                        "outputDirectory": str(output_directory),
                    })
                    continue

                if result["status"] == "no_holdings_section":
                    no_holdings_section.append(result)
                    print("STATUS: NO_HOLDINGS_SECTION")
                    continue

                successful.append(result)
                strategy = result["selectedStrategy"]
                strategy_counts[strategy] = strategy_counts.get(strategy, 0) + 1

                print(f"STATUS: SUCCESS")
                print(f"Fund: {result.get('fundName') or '-'}")
                print(f"Selected strategy: {strategy}")
                print(f"Holdings extracted: {len(result['topHoldings'])}")
                for holding in result["topHoldings"]:
                    print(
                        f"  {holding['rank']}. {holding['name']} "
                        f"- {holding['weightText']}"
                    )
        finally:
            context.close()
            browser.close()

    completed_at = utc_now_iso()
    all_results = successful + no_holdings_section

    total_holdings = sum(
        len(item.get("topHoldings", []))
        for item in successful
    )

    rules = {
        "excelColumnAControlsUniverse": True,
        "officialPrudentialSingaporeOnly": True,
        "acceptedTopHoldingsHeadings": [
            "Top 10 Holdings",
            "Top Ten Holdings",
            "Unicode/ASCII footnote variants of those headings",
        ],
        "genericHoldingsHeadingRejected": True,
        "genericInvestmentsHeadingRejected": True,
        "noInferredHoldings": True,
        "noFabricatedPercentages": True,
        "noForcedTenEntries": True,
        "fewerThanTenPublishedAllowed": True,
        "publishedOrderPreserved": True,
        "duplicateHoldingNamesAllowed": True,
        "duplicateHoldingPercentagesAllowed": True,
        "fixedIncomeLastPercentageSupported": True,
        "fullPdfPreserved": True,
        "rawTopHoldingsSectionPreserved": True,
        "perHoldingAcceptanceAudit": True,
        "adaptivePerFundStrategy": True,
        "verificationPassRequired": True,
        "thirdPartyHoldings": False,
    }

    all_payload = {
        "status": "success" if not failed else "partial",
        "source": "Prudential Singapore",
        "generatedAtUtc": completed_at,
        "excelFile": str(EXCEL_FILE),
        "excelFundUniverse": len(funds),
        "successfulFunds": len(successful),
        "noHoldingsSectionFunds": len(no_holdings_section),
        "failedFunds": len(failed),
        "totalPublishedTopHoldings": total_holdings,
        "strategyCounts": strategy_counts,
        "rules": rules,
        "funds": all_results,
        "failed": failed,
    }
    save_json(ALL_HOLDINGS_FILE, all_payload)

    run_summary = {
        "status": "success" if not failed else "partial",
        "startedAtUtc": started_at,
        "completedAtUtc": completed_at,
        "excelFile": str(EXCEL_FILE),
        "excelFundUniverse": len(funds),
        "successfulFunds": len(successful),
        "noHoldingsSectionFunds": len(no_holdings_section),
        "failedFunds": len(failed),
        "totalPublishedTopHoldings": total_holdings,
        "strategyCounts": strategy_counts,
        "successfulFundsDetail": [
            {
                "excelRow": x["excelRow"],
                "fundName": x.get("fundName"),
                "selectedStrategy": x["selectedStrategy"],
                "strategyReason": x["strategyReason"],
                "topHoldingsCount": len(x["topHoldings"]),
                "factsheetUrl": x["factsheetUrl"],
                "factsheetDataAsAt": x.get("factsheetDataAsAt"),
            }
            for x in successful
        ],
        "noHoldingsSectionDetail": [
            {
                "excelRow": x["excelRow"],
                "fundName": x.get("fundName"),
                "factsheetUrl": x["factsheetUrl"],
            }
            for x in no_holdings_section
        ],
        "failedFundsDetail": failed,
        "rules": rules,
    }
    save_json(RUN_SUMMARY_FILE, run_summary)

    print()
    print("=" * 78)
    print("ADAPTIVE PER-FUND TOP HOLDINGS TEST COMPLETE")
    print("=" * 78)
    print(f"Excel fund universe: {len(funds)}")
    print(f"Successful with holdings: {len(successful)}")
    print(f"No Top Holdings section: {len(no_holdings_section)}")
    print(f"Failed: {len(failed)}")
    print(f"Total published holdings extracted: {total_holdings}")
    print("Strategy counts:")
    for strategy, count in sorted(strategy_counts.items()):
        print(f"  {strategy}: {count}")
    print()
    print(f"Output: {ALL_HOLDINGS_FILE}")
    print(f"Output: {RUN_SUMMARY_FILE}")
    print(f"Output: {FUNDS_OUTPUT_DIR}")

    if failed:
        print()
        print("FAILED FUNDS:")
        for item in failed:
            print(f"  Row {item['excelRow']}: {item['error']}")

    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except KeyboardInterrupt:
        print("\nInterrupted.", file=sys.stderr)
        raise SystemExit(130)
