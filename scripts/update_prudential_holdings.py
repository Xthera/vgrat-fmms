#!/usr/bin/env python3

"""
VGrat FMS - Prudential Top Holdings FAILED-FUND RECOVERY RUNNER

IMPORTANT
=========

This script does NOT replace or modify scripts/test_prudential_holdings.py.

The frozen baseline remains the extraction authority.
This recovery runner imports the frozen baseline module and calls its exact:

    read_excel_funds()
    extract_single_fund()
    extract_holdings_section()
    parse_holdings()
    parse_holdings_fallback()
    extract_pdf_text()
    download_factsheet()
    is_prudential_url()
    clean_text()
    safe_filename()
    utc_now_iso()

No alternate proximity parser is used.
No holding is inferred.
No percentage is fabricated, estimated, interpolated, or calculated.

WORKFLOW
========

1. Read the completed baseline output:
       output_holdings/run_summary.json

2. Read failedFundsDetail from that baseline run.

3. Read Funds Links.xlsm using the SAME frozen baseline reader.

4. Select ONLY the failed Excel rows from the baseline run.

5. Re-run those failed funds using the EXACT frozen baseline
   extract_single_fund() implementation.

6. Re-download the official Prudential factsheet and verify the result using
   the SAME parser selected by the baseline extraction.

7. Save recovery artifacts under:
       output_holdings_recovery/

8. Never modify:
       output_holdings/
       scripts/test_prudential_holdings.py
       scripts/test_pruaccess.py

This is deliberately a thin recovery runner. Parser behavior belongs to the
frozen baseline, not this file.
"""

from __future__ import annotations

import json
import sys
import time
from pathlib import Path
from urllib.parse import urlparse

from playwright.sync_api import sync_playwright


# =============================================================================
# PATHS
# =============================================================================

ROOT_DIR = Path(__file__).resolve().parent.parent
SCRIPT_DIR = Path(__file__).resolve().parent

if str(SCRIPT_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPT_DIR))

# Import the frozen baseline.  Do not copy or reimplement its parser logic.
import test_prudential_holdings as baseline


BASELINE_SUMMARY_FILE = ROOT_DIR / "output_holdings" / "run_summary.json"
RECOVERY_DIR = ROOT_DIR / "output_holdings_recovery"
RECOVERY_FUNDS_DIR = RECOVERY_DIR / "funds"
RECOVERY_SUMMARY_FILE = RECOVERY_DIR / "run_summary.json"
RECOVERY_ALL_FILE = RECOVERY_DIR / "all_holdings.json"


# Keep the same retry timing as the frozen baseline.
RETRY_COUNT = baseline.RETRY_COUNT
RETRY_DELAY_SECONDS = baseline.RETRY_DELAY_SECONDS
BROWSER_HEADLESS = baseline.BROWSER_HEADLESS
PAGE_TIMEOUT_MS = baseline.PAGE_TIMEOUT_MS
FACTSHEET_DOWNLOAD_TIMEOUT_MS = baseline.FACTSHEET_DOWNLOAD_TIMEOUT_MS


# =============================================================================
# HELPERS
# =============================================================================


def load_json(path: Path):
    if not path.exists():
        raise RuntimeError(f"Required file does not exist: {path}")

    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except Exception as error:
        raise RuntimeError(f"Could not read JSON file {path}: {error}") from error


def save_json(path: Path, value) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(
        json.dumps(value, indent=2, ensure_ascii=False),
        encoding="utf-8",
    )


def signature(holdings: list[dict]) -> list[tuple]:
    """Return the exact holding identity used for final recovery comparison."""
    return [
        (
            item.get("rank"),
            item.get("name"),
            item.get("weightPercent"),
        )
        for item in holdings
    ]


def recovery_identifier(result: dict, excel_row: int) -> str:
    parsed = urlparse(
        result.get("finalUrl")
        or result.get("prudentialUrl")
        or ""
    )

    identifier = baseline.safe_filename(
        parsed.path.rstrip("/").split("/")[-1]
        or result.get("fundName")
        or f"fund_{excel_row}"
    )

    return f"{excel_row}_{identifier}"


def save_recovery_success(
    result: dict,
    factsheet_bytes: bytes,
    full_text: str,
    section_text: str,
) -> Path:
    excel_row = int(result["excelRow"])
    directory = RECOVERY_FUNDS_DIR / recovery_identifier(result, excel_row)
    directory.mkdir(parents=True, exist_ok=True)

    (directory / "factsheet.pdf").write_bytes(factsheet_bytes)
    (directory / "factsheet_text.txt").write_text(
        full_text,
        encoding="utf-8",
    )
    (directory / "top_holdings_section.txt").write_text(
        section_text,
        encoding="utf-8",
    )
    save_json(directory / "top_holdings.json", result)

    metadata = {
        "excelRow": result.get("excelRow"),
        "fundName": result.get("fundName"),
        "excelPruAccessName": result.get("excelPruAccessName"),
        "prudentialUrl": result.get("prudentialUrl"),
        "finalUrl": result.get("finalUrl"),
        "factsheetUrl": result.get("factsheetUrl"),
        "factsheetDocumentDate": result.get("factsheetDocumentDate"),
        "factsheetDataAsAt": result.get("factsheetDataAsAt"),
        "factsheetPageCount": result.get("factsheetPageCount"),
        "topHoldingsCount": result.get("topHoldingsCount"),
        "holdingsParser": result.get("holdingsParser"),
        "savedAtUtc": baseline.utc_now_iso(),
        "source": "frozen_test_prudential_holdings.py",
    }
    save_json(directory / "metadata.json", metadata)

    return directory


def save_recovery_failure(excel_fund: dict, error: str) -> Path:
    excel_row = int(excel_fund["excelRow"])
    directory = RECOVERY_FUNDS_DIR / f"{excel_row}_failed"
    directory.mkdir(parents=True, exist_ok=True)

    failure = {
        "status": "failed",
        "excelRow": excel_row,
        "prudentialUrl": excel_fund.get("prudentialUrl"),
        "pruAccessName": excel_fund.get("pruAccessName"),
        "error": baseline.clean_text(error),
        "savedAtUtc": baseline.utc_now_iso(),
        "source": "frozen_test_prudential_holdings.py",
    }
    save_json(directory / "failure.json", failure)
    return directory


def save_no_holdings_result(result: dict) -> Path:
    excel_row = int(result["excelRow"])
    directory = RECOVERY_FUNDS_DIR / recovery_identifier(result, excel_row)
    directory.mkdir(parents=True, exist_ok=True)

    save_json(directory / "top_holdings.json", result)
    save_json(
        directory / "metadata.json",
        {
            "excelRow": result.get("excelRow"),
            "fundName": result.get("fundName"),
            "factsheetUrl": result.get("factsheetUrl"),
            "factsheetDocumentDate": result.get("factsheetDocumentDate"),
            "factsheetDataAsAt": result.get("factsheetDataAsAt"),
            "topHoldingsCount": 0,
            "status": result.get("status"),
            "savedAtUtc": baseline.utc_now_iso(),
            "source": "frozen_test_prudential_holdings.py",
        },
    )
    return directory


def load_failed_rows() -> tuple[dict, list[int]]:
    """
    Require the baseline run summary and use ONLY its failedFundsDetail rows.
    """
    summary = load_json(BASELINE_SUMMARY_FILE)

    if not isinstance(summary, dict):
        raise RuntimeError("Baseline run_summary.json is not a JSON object.")

    if summary.get("status") == "running":
        raise RuntimeError(
            "Baseline run_summary.json indicates the baseline run is still running. "
            "Wait for the baseline to complete before starting recovery."
        )

    details = summary.get("failedFundsDetail")

    if details is None:
        raise RuntimeError(
            "Baseline run_summary.json has no failedFundsDetail field. "
            "Recovery requires a completed baseline summary."
        )

    if not isinstance(details, list):
        raise RuntimeError("Baseline failedFundsDetail is not a list.")

    failed_rows: list[int] = []
    for item in details:
        if not isinstance(item, dict):
            continue
        if item.get("excelRow") is None:
            continue
        row = int(item["excelRow"])
        if row not in failed_rows:
            failed_rows.append(row)

    return summary, failed_rows


def select_failed_funds(funds: list[dict], failed_rows: list[int]) -> list[dict]:
    by_row = {int(item["excelRow"]): item for item in funds}

    missing = [row for row in failed_rows if row not in by_row]
    if missing:
        raise RuntimeError(
            "Baseline failed rows were not found in Funds Links.xlsm: "
            + ", ".join(str(row) for row in missing)
        )

    # Preserve the order recorded by the baseline failure summary.
    return [by_row[row] for row in failed_rows]


def verify_official_pdf_again(
    page,
    result: dict,
) -> tuple[bytes, str, str, list[dict]]:
    """
    Final verification using the same official PDF and the same frozen parser.

    The parser choice is retained from extract_single_fund().
    If the baseline used fallback, fallback is used here.
    If it used primary, primary is attempted first and the exact baseline
    fallback behavior is retained if primary fails.
    """
    response = page.request.get(
        result["factsheetUrl"],
        timeout=FACTSHEET_DOWNLOAD_TIMEOUT_MS,
    )

    if response.status != 200:
        raise RuntimeError(
            "Factsheet re-download returned "
            f"HTTP {response.status}."
        )

    factsheet_bytes = response.body()

    if not factsheet_bytes.startswith(b"%PDF"):
        raise RuntimeError(
            "Factsheet re-download did not return a valid PDF."
        )

    full_text, _page_count = baseline.extract_pdf_text(factsheet_bytes)

    section_text, section_status = baseline.extract_holdings_section(full_text)

    if section_status != "published":
        raise RuntimeError(
            "Factsheet holdings section disappeared during final verification."
        )

    parser_used = result.get("holdingsParser") or "primary"

    if parser_used == "fallback":
        verified_holdings = baseline.parse_holdings_fallback(section_text)
    else:
        try:
            verified_holdings = baseline.parse_holdings(section_text)
        except Exception as primary_error:
            verified_holdings = baseline.parse_holdings_fallback(section_text)
            result["holdingsParser"] = "fallback"
            result["primaryParserError"] = baseline.clean_text(
                str(primary_error)
            )

    original_holdings = result.get("topHoldings") or []

    if len(verified_holdings) != len(original_holdings):
        raise RuntimeError(
            "Holding count changed during final verification: "
            f"initial={len(original_holdings)}, "
            f"verified={len(verified_holdings)}."
        )

    if signature(verified_holdings) != signature(original_holdings):
        raise RuntimeError(
            "Holding signature changed during final verification. "
            "No recovery result was accepted."
        )

    result["topHoldings"] = verified_holdings
    result["topHoldingsCount"] = len(verified_holdings)

    return factsheet_bytes, full_text, section_text, verified_holdings


# =============================================================================
# MAIN RECOVERY RUN
# =============================================================================


def main() -> int:
    started_at = baseline.utc_now_iso()

    print("################################################################")
    print("VGrat FMS - PRUDENTIAL FAILED-FUND HOLDINGS RECOVERY")
    print("################################################################")
    print(f"Started UTC: {started_at}")
    print(f"Baseline summary: {BASELINE_SUMMARY_FILE}")
    print(f"Recovery output:  {RECOVERY_DIR}")
    print()
    print("Parser source: scripts/test_prudential_holdings.py")
    print("Parser mode:   EXACT FROZEN BASELINE")

    baseline_summary, failed_rows = load_failed_rows()

    if not failed_rows:
        print("\nBaseline has no failed funds. Nothing to recover.")
        save_json(
            RECOVERY_SUMMARY_FILE,
            {
                "status": "nothing_to_recover",
                "startedAtUtc": started_at,
                "completedAtUtc": baseline.utc_now_iso(),
                "baselineSummary": str(BASELINE_SUMMARY_FILE),
                "failedRows": [],
                "successfulFunds": [],
                "failedFunds": [],
                "rules": {
                    "usesFrozenBaselineParser": True,
                    "failedFundsOnly": True,
                    "noProximityMatching": True,
                    "noSyntheticData": True,
                },
            },
        )
        save_json(RECOVERY_ALL_FILE, [])
        return 0

    print(
        "\nBaseline failed rows: "
        + ", ".join(str(row) for row in failed_rows)
    )

    # -------------------------------------------------------------------------
    # Read the master Excel universe with the exact frozen reader.
    # -------------------------------------------------------------------------
    funds = baseline.read_excel_funds()
    failed_funds = select_failed_funds(funds, failed_rows)

    print(
        f"Selected {len(failed_funds)} failed fund(s) for recovery."
    )

    successful: list[dict] = []
    failed: list[dict] = []
    no_holdings_section: list[dict] = []

    RECOVERY_FUNDS_DIR.mkdir(parents=True, exist_ok=True)

    with sync_playwright() as playwright:
        browser = playwright.chromium.launch(
            headless=BROWSER_HEADLESS
        )

        context = browser.new_context(
            viewport={
                "width": 1440,
                "height": 1000,
            },
            user_agent=(
                "Mozilla/5.0 "
                "(Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 "
                "(KHTML, like Gecko) "
                "Chrome/153.0.0.0 Safari/537.36"
            ),
        )

        page = context.new_page()

        try:
            total = len(failed_funds)

            for index, excel_fund in enumerate(failed_funds, start=1):
                row = int(excel_fund["excelRow"])

                print("\n" + "=" * 72)
                print(f"RECOVERY FUND {index}/{total}")
                print(f"Excel row: {row}")
                print("=" * 72)
                print(f"URL: {excel_fund['prudentialUrl']}")

                result = None
                last_error = None

                # -----------------------------------------------------------------
                # Exact same whole-fund retry pattern as the frozen baseline.
                # -----------------------------------------------------------------
                for attempt in range(1, RETRY_COUNT + 1):
                    try:
                        print(f"\nAttempt {attempt}/{RETRY_COUNT}")

                        result = baseline.extract_single_fund(
                            page,
                            excel_fund,
                        )

                        last_error = None
                        break

                    except Exception as error:
                        last_error = baseline.clean_text(str(error))
                        print("Attempt failed:")
                        print(last_error)

                        if attempt < RETRY_COUNT:
                            print(
                                f"Retrying in {RETRY_DELAY_SECONDS} seconds..."
                            )
                            time.sleep(RETRY_DELAY_SECONDS)

                if result is None:
                    failure_dir = save_recovery_failure(
                        excel_fund,
                        last_error or "Unknown extraction failure.",
                    )

                    failed.append(
                        {
                            "excelRow": row,
                            "prudentialUrl": excel_fund.get("prudentialUrl"),
                            "pruAccessName": excel_fund.get("pruAccessName"),
                            "error": last_error or "Unknown extraction failure.",
                            "outputDirectory": str(failure_dir),
                        }
                    )
                    continue

                # -----------------------------------------------------------------
                # No published holdings section.
                # -----------------------------------------------------------------
                if result.get("status") == "no_holdings_section":
                    output_dir = save_no_holdings_result(result)
                    result["outputDirectory"] = str(output_dir)
                    no_holdings_section.append(result)
                    print("NO TOP HOLDINGS SECTION")
                    continue

                # -----------------------------------------------------------------
                # Final official-PDF verification.
                # -----------------------------------------------------------------
                try:
                    (
                        factsheet_bytes,
                        full_text,
                        section_text,
                        verified_holdings,
                    ) = verify_official_pdf_again(page, result)

                    output_dir = save_recovery_success(
                        result,
                        factsheet_bytes,
                        full_text,
                        section_text,
                    )

                    result["outputDirectory"] = str(output_dir)
                    result["verification"] = {
                        "status": "passed",
                        "verifiedTopHoldingsCount": len(verified_holdings),
                        "signatureMatch": True,
                    }

                    # Re-save top_holdings.json with the verification metadata.
                    save_json(output_dir / "top_holdings.json", result)

                    successful.append(result)

                    print(
                        f"RECOVERED: {len(verified_holdings)} holding(s) "
                        f"using {result.get('holdingsParser', 'unknown')} parser."
                    )

                except Exception as error:
                    error_text = baseline.clean_text(str(error))
                    failure_dir = save_recovery_failure(
                        excel_fund,
                        error_text,
                    )

                    failed.append(
                        {
                            "excelRow": row,
                            "prudentialUrl": excel_fund.get("prudentialUrl"),
                            "pruAccessName": excel_fund.get("pruAccessName"),
                            "error": error_text,
                            "outputDirectory": str(failure_dir),
                        }
                    )

                    print("FINAL VERIFICATION FAILED:")
                    print(error_text)

        finally:
            context.close()
            browser.close()

    # -------------------------------------------------------------------------
    # Consolidated recovery outputs.
    # -------------------------------------------------------------------------
    all_holdings = []
    for result in successful:
        all_holdings.append(
            {
                "excelRow": result.get("excelRow"),
                "fundName": result.get("fundName"),
                "excelPruAccessName": result.get("excelPruAccessName"),
                "prudentialUrl": result.get("prudentialUrl"),
                "finalUrl": result.get("finalUrl"),
                "factsheetUrl": result.get("factsheetUrl"),
                "factsheetDocumentDate": result.get("factsheetDocumentDate"),
                "factsheetDataAsAt": result.get("factsheetDataAsAt"),
                "topHoldingsCount": result.get("topHoldingsCount"),
                "holdingsParser": result.get("holdingsParser"),
                "topHoldings": result.get("topHoldings", []),
            }
        )

    save_json(RECOVERY_ALL_FILE, all_holdings)

    status = "success" if not failed else "partial"

    summary = {
        "status": status,
        "startedAtUtc": started_at,
        "completedAtUtc": baseline.utc_now_iso(),
        "baselineSummary": str(BASELINE_SUMMARY_FILE),
        "baselineStatus": baseline_summary.get("status"),
        "baselineFailedRows": failed_rows,
        "recoveryUniverse": len(failed_funds),
        "successfulFunds": [item.get("excelRow") for item in successful],
        "noHoldingsSectionFunds": [
            item.get("excelRow") for item in no_holdings_section
        ],
        "failedFunds": [item.get("excelRow") for item in failed],
        "successfulCount": len(successful),
        "noHoldingsSectionCount": len(no_holdings_section),
        "failedCount": len(failed),
        "failedFundsDetail": failed,
        "rules": {
            "usesFrozenBaselineParser": True,
            "failedFundsOnly": True,
            "officialPrudentialSourceOnly": True,
            "noProximityMatching": True,
            "noInferredHoldings": True,
            "noFabricatedPercentages": True,
            "noSyntheticData": True,
            "publishedCountUsedExactly": True,
            "finalOfficialPdfVerification": True,
            "exactHoldingSignatureVerification": True,
            "baselineScriptUnmodified": True,
            "testPruaccessUnmodified": True,
        },
    }

    save_json(RECOVERY_SUMMARY_FILE, summary)

    print("\n" + "#" * 72)
    print("RECOVERY COMPLETE")
    print("#" * 72)
    print(f"Baseline failed funds: {len(failed_funds)}")
    print(f"Recovered:             {len(successful)}")
    print(f"No holdings section:   {len(no_holdings_section)}")
    print(f"Still failed:          {len(failed)}")
    print(f"Recovery summary:      {RECOVERY_SUMMARY_FILE}")
    print(f"Recovery all holdings: {RECOVERY_ALL_FILE}")

    if failed:
        print("\nStill-failed rows:")
        for item in failed:
            print(
                f"  Row {item['excelRow']}: {item['error']}"
            )

    return 0 if not failed else 1


if __name__ == "__main__":
    raise SystemExit(main())
