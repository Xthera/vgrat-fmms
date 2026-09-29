#!/usr/bin/env python3

"""
VGrat FMS - BUILD FUNDS DATA

Merges the outputs of the extraction pipelines into two data files that
share the SAME layout (same envelope, same per-fund identity block):

    data/funds.json          fund info + Top Holdings
    data/bid_history.json    historical BID observations

MASTER UNIVERSE
===============

Funds Links.xlsm, Column A (Prudential URL) / Column B (PruAccess name).
Every populated row becomes exactly one record in BOTH files, ordered by
Excel row.

HOLDINGS RESOLUTION
===================

Stage order:

    baseline  ->  recovery1  ->  recovery2  ->  recovery3

- A fund that succeeded in the baseline keeps its baseline result.
- A fund the baseline marked "no_holdings_section" is kept AS IS
  (empty holdings, status "no_holdings_section"). It is never replaced.
- A fund that FAILED in the baseline is replaced by the first recovery
  stage that resolved it (holdings, or a genuine no_holdings_section).
- A fund still failed after Recovery 3 is "unresolved".

Unresolved funds (holdings or BID history) fail the build unless
ALLOW_UNRESOLVED=1, in which case they are written with status
"unresolved" and empty data. Nothing is ever inferred or fabricated.

BID HISTORY
===========

Read from output_pruaccess/funds/<row>_<id>/ (bid_history.json and
prudential_fund.json). Observations are copied exactly as extracted.

DIVIDEND NORMALIZATION
======================

Final funds.json rule:

    - If dividendRate is non-empty -> hasDividend = true
    - If dividendRate is empty/missing -> hasDividend = false

This rule is applied during the final build so that hasDividend is always
consistent with dividendRate in the published data.
"""

from __future__ import annotations

import json
import os
import re
import sys
from datetime import datetime, timezone
from pathlib import Path

from openpyxl import load_workbook


# =============================================================================
# CONFIGURATION
# =============================================================================

EXCEL_FILE = Path("Funds Links.xlsm")

BASELINE_FILE = Path("output_holdings/all_holdings.json")

RECOVERY_STAGES = [
    ("recovery1", Path("output_holdings_recovery")),
    ("recovery2", Path("output_holdings_recovery_2")),
    ("recovery3", Path("output_holdings_recovery_3")),
]

PRUACCESS_FUNDS_DIR = Path("output_pruaccess/funds")

DATA_DIR = Path("data")
FUNDS_OUT = DATA_DIR / "funds.json"
BID_OUT = DATA_DIR / "bid_history.json"

SCHEMA_VERSION = 1
MAX_HOLDINGS = 10

ALLOW_UNRESOLVED = os.environ.get(
    "ALLOW_UNRESOLVED", ""
).strip().lower() in {"1", "true", "yes"}


# =============================================================================
# HELPERS
# =============================================================================

def clean_text(value) -> str:
    if value is None:
        return ""
    return re.sub(r"\s+", " ", str(value).replace("\xa0", " ")).strip()


def utc_now_iso() -> str:
    return datetime.now(timezone.utc).replace(microsecond=0).isoformat()


def load_json(path: Path):
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except Exception as error:
        raise RuntimeError(f"Could not read {path}: {error}") from error


def write_json_atomic(path: Path, data, compact: bool = False) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp = path.with_suffix(path.suffix + ".tmp")

    if compact:
        text = json.dumps(
            data,
            ensure_ascii=False,
            separators=(",", ":"),
        )
    else:
        text = json.dumps(
            data,
            indent=2,
            ensure_ascii=False,
        )

    tmp.write_text(text + "\n", encoding="utf-8")
    tmp.replace(path)


# =============================================================================
# DIVIDEND NORMALIZATION
# =============================================================================

def normalize_dividend_fields(fund_info: dict) -> dict:
    """
    Ensure hasDividend is always consistent with dividendRate.

    Rule:
        non-empty dividendRate -> True
        empty/missing dividendRate -> False

    The existing dividendRate value itself is preserved unchanged.
    """

    dividend_rate = clean_text(
        fund_info.get("dividendRate")
    )

    fund_info["hasDividend"] = bool(dividend_rate)

    return fund_info


# =============================================================================
# EXCEL MASTER UNIVERSE
# =============================================================================

def read_excel_funds() -> dict[int, dict]:
    if not EXCEL_FILE.exists():
        raise FileNotFoundError(
            f"Excel file not found: {EXCEL_FILE}"
        )

    workbook = load_workbook(
        EXCEL_FILE,
        read_only=True,
        keep_vba=True,
        data_only=True,
    )

    funds = {}

    try:
        worksheet = workbook.active

        for row in range(2, worksheet.max_row + 1):
            url = clean_text(
                worksheet.cell(row=row, column=1).value
            )
            name = clean_text(
                worksheet.cell(row=row, column=2).value
            )

            if not url:
                continue

            funds[row] = {
                "excelRow": row,
                "prudentialUrl": url,
                "pruAccessName": name,
            }

    finally:
        workbook.close()

    if not funds:
        raise RuntimeError(
            "No populated URLs in Excel Column A."
        )

    return funds


# =============================================================================
# HOLDINGS
# =============================================================================

def clean_holdings(holdings, label: str) -> list[dict]:
    if not isinstance(holdings, list) or not holdings:
        raise ValueError(
            f"{label}: holdings list is empty or invalid."
        )

    if len(holdings) > MAX_HOLDINGS:
        raise ValueError(
            f"{label}: more than {MAX_HOLDINGS} holdings."
        )

    cleaned = []

    for position, item in enumerate(holdings, start=1):
        if not isinstance(item, dict):
            raise ValueError(
                f"{label}: holding {position} is not an object."
            )

        name = clean_text(item.get("name"))
        weight = item.get("weightPercent")

        if item.get("rank") != position:
            raise ValueError(
                f"{label}: rank {item.get('rank')} "
                f"at position {position}."
            )

        if not name:
            raise ValueError(
                f"{label}: holding {position} has no name."
            )

        if (
            isinstance(weight, bool)
            or not isinstance(weight, (int, float))
            or not 0 <= weight <= 100
        ):
            raise ValueError(
                f"{label}: holding {position} "
                f"has invalid weight {weight!r}."
            )

        cleaned.append(
            {
                "rank": position,
                "name": name,
                "weightPercent": weight,
                "weightText": clean_text(
                    item.get("weightText")
                ),
            }
        )

    return cleaned


def holdings_block(result: dict, stage: str, label: str):
    """Return the topHoldings block, or None when the result is a failure."""

    status = result.get("status")

    common = {
        "source": stage,
        "parser": result.get("holdingsParser"),
        "factsheetUrl": result.get("factsheetUrl"),
        "factsheetDocumentDate": result.get(
            "factsheetDocumentDate"
        ),
        "factsheetDataAsAt": result.get(
            "factsheetDataAsAt"
        ),
    }

    if status == "success":
        holdings = clean_holdings(
            result.get("topHoldings"),
            label,
        )

        return {
            "status": "published",
            **common,
            "count": len(holdings),
            "holdings": holdings,
        }

    if status == "no_holdings_section":
        return {
            "status": "no_holdings_section",
            **common,
            "count": 0,
            "holdings": [],
        }

    return None


def resolve_holdings(
    excel_funds: dict[int, dict],
) -> dict[int, dict]:

    resolved: dict[int, dict] = {}

    def accept(result: dict, stage: str) -> None:
        if result.get("excelRow") is None:
            raise RuntimeError(
                f"{stage}: result without excelRow."
            )

        row = int(result["excelRow"])

        if row not in excel_funds:
            raise RuntimeError(
                f"{stage}: row {row} is not in Excel."
            )

        # A resolved fund is never replaced by a later stage.
        if row in resolved:
            return

        result_url = clean_text(
            result.get("prudentialUrl")
        )

        if (
            result_url
            and result_url
            != excel_funds[row]["prudentialUrl"]
        ):
            raise RuntimeError(
                f"{stage}: URL mismatch for row {row}."
            )

        block = holdings_block(
            result,
            stage,
            f"{stage} row {row}",
        )

        if block is not None:
            block["fundName"] = (
                clean_text(result.get("fundName"))
                or None
            )
            resolved[row] = block

    # ---- baseline --------------------------------------------------------

    if not BASELINE_FILE.exists():
        raise FileNotFoundError(
            f"Baseline output not found: {BASELINE_FILE}"
        )

    baseline = load_json(BASELINE_FILE)

    for result in baseline.get("funds", []):
        accept(result, "baseline")

    # ---- recovery stages -------------------------------------------------

    for stage, stage_dir in RECOVERY_STAGES:
        funds_dir = stage_dir / "funds"

        if not funds_dir.exists():
            print(
                f"WARNING: {funds_dir} not found; "
                f"{stage} skipped."
            )
            continue

        for path in sorted(
            funds_dir.glob("*/top_holdings.json")
        ):
            accept(
                load_json(path),
                stage,
            )

    return resolved


# =============================================================================
# PRUACCESS
# =============================================================================

def validate_bid_history(
    history: dict,
    label: str,
) -> list[dict]:

    observations = history.get("observations")

    if not isinstance(observations, list) or not observations:
        raise ValueError(
            f"{label}: no BID observations."
        )

    if history.get("observationCount") != len(observations):
        raise ValueError(
            f"{label}: observationCount mismatch."
        )

    previous = None
    seen = set()

    for item in observations:
        date = item.get("date")
        price = item.get("bidPrice")

        if not isinstance(date, str) or not date:
            raise ValueError(
                f"{label}: invalid date {date!r}."
            )

        if (
            isinstance(price, bool)
            or not isinstance(price, (int, float))
        ):
            raise ValueError(
                f"{label}: invalid BID price on {date}."
            )

        if date in seen:
            raise ValueError(
                f"{label}: duplicate date {date}."
            )

        if previous is not None and date < previous:
            raise ValueError(
                f"{label}: observations not chronological."
            )

        seen.add(date)
        previous = date

    return [
        {
            "date": item["date"],
            "bidPrice": item["bidPrice"],
        }
        for item in observations
    ]


def load_pruaccess() -> dict[int, dict]:
    result = {}

    if not PRUACCESS_FUNDS_DIR.exists():
        print(
            f"WARNING: {PRUACCESS_FUNDS_DIR} not found."
        )
        return result

    for directory in sorted(
        PRUACCESS_FUNDS_DIR.iterdir()
    ):
        match = re.match(
            r"^(\d+)_",
            directory.name,
        )

        if not directory.is_dir() or not match:
            continue

        bid_file = directory / "bid_history.json"
        fund_file = directory / "prudential_fund.json"

        if not (
            bid_file.exists()
            and fund_file.exists()
        ):
            continue

        row = int(match.group(1))

        result[row] = {
            "prudential": load_json(fund_file),
            "bidHistory": load_json(bid_file),
        }

    return result


# =============================================================================
# BUILD
# =============================================================================

def main() -> int:
    print("=" * 72)
    print("VGRAT FMS - BUILD FUNDS DATA")
    print("=" * 72)
    print(
        f"Allow unresolved: {ALLOW_UNRESOLVED}"
    )

    excel_funds = read_excel_funds()
    holdings = resolve_holdings(excel_funds)
    pruaccess = load_pruaccess()

    generated_at = utc_now_iso()

    fund_records = []
    bid_records = []
    gaps = []

    stage_counts: dict[str, int] = {}
    published = 0
    no_section = 0
    unresolved_holdings = 0

    bid_ok = 0
    bid_missing = 0
    total_observations = 0

    dividend_true = 0
    dividend_false = 0

    for row in sorted(excel_funds):
        excel = excel_funds[row]
        pru = pruaccess.get(row)
        prudential = (
            (pru or {}).get("prudential")
            or {}
        )

        block = holdings.get(row)

        identity = {
            "excelRow": row,
            "fundIdentifier": (
                clean_text(
                    prudential.get("fundIdentifier")
                )
                or None
            ),
            "fundCode": (
                clean_text(
                    prudential.get("fundCode")
                )
                or None
            ),
            "fundName": (
                clean_text(
                    prudential.get("fundName")
                )
                or (block or {}).get("fundName")
                or None
            ),
            "pruAccessName": (
                excel["pruAccessName"]
                or None
            ),
            "prudentialUrl": excel["prudentialUrl"],
        }

        # ---- holdings ----------------------------------------------------

        if block is None:
            unresolved_holdings += 1
            gaps.append(
                f"Row {row}: holdings unresolved "
                f"after Recovery 3."
            )

            top_holdings = {
                "status": "unresolved",
                "source": None,
                "count": 0,
                "holdings": [],
            }

        else:
            top_holdings = {
                k: v
                for k, v in block.items()
                if k != "fundName"
            }

            stage_counts[block["source"]] = (
                stage_counts.get(
                    block["source"],
                    0,
                )
                + 1
            )

            if block["status"] == "published":
                published += 1
            else:
                no_section += 1

        # ---- fund info ---------------------------------------------------

        fund_info = None

        if prudential:
            fund_info = {
                k: v
                for k, v in prudential.items()
                if k not in {
                    "raw",
                    "fundIdentifier",
                    "fundCode",
                    "fundName",
                }
            }

            # -------------------------------------------------------------
            # DIVIDEND RULE
            #
            # If dividendRate contains a value, hasDividend is ALWAYS true.
            # If dividendRate is empty or missing, hasDividend is false.
            #
            # dividendRate itself is NOT modified.
            # -------------------------------------------------------------

            fund_info = normalize_dividend_fields(
                fund_info
            )

            if fund_info["hasDividend"]:
                dividend_true += 1
            else:
                dividend_false += 1

        fund_records.append(
            {
                **identity,
                "fund": fund_info,
                "topHoldings": top_holdings,
            }
        )

        # ---- bid history -------------------------------------------------

        if pru is None:
            bid_missing += 1
            gaps.append(
                f"Row {row}: BID history missing."
            )

            bid_history = {
                "status": "unresolved",
                "priceType": "BID",
                "observationCount": 0,
                "observations": [],
            }

        else:
            history = pru["bidHistory"]

            observations = validate_bid_history(
                history,
                f"row {row} BID",
            )

            bid_ok += 1
            total_observations += len(observations)

            bid_history = {
                "status": "success",
                "priceType": "BID",
                "pruAccessFundId": history.get(
                    "fundId"
                ),
                "currency": history.get(
                    "currency"
                ),
                "startDate": history.get(
                    "startDate"
                ),
                "endDate": history.get(
                    "endDate"
                ),
                "observationCount": len(
                    observations
                ),
                "observations": observations,
            }

        bid_records.append(
            {
                **identity,
                "bidHistory": bid_history,
            }
        )

    # ---- gate ------------------------------------------------------------

    if gaps:
        print("\nUNRESOLVED:")

        for gap in gaps:
            print(f" - {gap}")

        if not ALLOW_UNRESOLVED:
            print(
                "\nBUILD FAILED: unresolved funds present. "
                "Nothing was written. Set ALLOW_UNRESOLVED=1 "
                "to publish them as 'unresolved'."
            )
            return 1

    envelope = {
        "schemaVersion": SCHEMA_VERSION,
        "generatedAtUtc": generated_at,
        "source": str(EXCEL_FILE),
        "fundCount": len(fund_records),
    }

    write_json_atomic(
        FUNDS_OUT,
        {
            **envelope,
            "summary": {
                "holdingsPublished": published,
                "noHoldingsSection": no_section,
                "holdingsUnresolved": unresolved_holdings,
                "holdingsBySource": stage_counts,
                "fundsWithDividend": dividend_true,
                "fundsWithoutDividend": dividend_false,
            },
            "funds": fund_records,
        },
    )

    write_json_atomic(
        BID_OUT,
        {
            **envelope,
            "summary": {
                "fundsWithBidHistory": bid_ok,
                "bidHistoryUnresolved": bid_missing,
                "totalObservations": total_observations,
            },
            "funds": bid_records,
        },
        compact=True,
    )

    print("\nBUILD COMPLETE")
    print(
        f"Funds:                 {len(fund_records)}"
    )
    print(
        f"Holdings published:    {published}"
    )
    print(
        f"No holdings section:   {no_section}"
    )
    print(
        f"Holdings unresolved:   {unresolved_holdings}"
    )
    print(
        f"Holdings by source:    {stage_counts}"
    )
    print(
        f"BID history funds:     {bid_ok}"
    )
    print(
        f"BID observations:      {total_observations}"
    )
    print(
        f"Funds with dividend:   {dividend_true}"
    )
    print(
        f"Funds without dividend:{dividend_false}"
    )
    print(
        f"Wrote: {FUNDS_OUT}"
    )
    print(
        f"Wrote: {BID_OUT}"
    )

    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception as error:
        print(
            f"\nFATAL BUILD ERROR: {error}",
            file=sys.stderr,
        )
        raise SystemExit(1)
