#!/usr/bin/env python3

"""
VGrat FMS - Gemini API Diagnostic
=================================

Purpose
-------

Diagnose Gemini API key / project / model availability.

This script deliberately does NOT read or modify:

    data/market_news/current.json
    Research Funds.xlsx
    data/market_news/analysis/current.json

It performs two independent checks:

    1. LIST MODELS
       GET /v1beta/models

       This shows exactly which models are exposed to the API key.

    2. MINIMAL GENERATION TEST
       POST /v1beta/models/{model}:generateContent

       This sends an intentionally tiny prompt to one model that
       supports generateContent.

The script never prints the API key itself.

Environment variables
---------------------

Required:

    GEMINI_API_KEY

Optional:

    GEMINI_MODEL

If GEMINI_MODEL is not supplied, the script chooses the first
available generateContent model from the returned model list,
with preference for the following models:

    gemini-3.8-flash
    gemini-3.7-flash
    gemini-3.6-flash
    gemini-3.5-flash
    gemini-3.5-flash-lite
    gemini-2.5-flash-lite

Exit codes
----------

    0 = API key can list models and a minimal generation request
        succeeded.

    1 = diagnostic failure.

No files are written.
"""

from __future__ import annotations

import json
import os
import sys
import time
from datetime import datetime, timezone
from typing import Any

import requests


# ============================================================
# CONFIGURATION
# ============================================================

API_BASE = "https://generativelanguage.googleapis.com/v1beta"
API_KEY = os.getenv("GEMINI_API_KEY", "").strip()

REQUEST_TIMEOUT_SECONDS = int(
    os.getenv("GEMINI_DIAGNOSTIC_TIMEOUT_SECONDS", "60")
)

REQUESTED_MODEL = os.getenv("GEMINI_MODEL", "").strip()

PREFERRED_MODELS = [
    "gemini-3.8-flash",
    "gemini-3.7-flash",
    "gemini-3.6-flash",
    "gemini-3.5-flash",
    "gemini-3.5-flash-lite",
    "gemini-2.5-flash-lite",
    "gemini-2.5-flash",
]


# ============================================================
# LOGGING
# ============================================================

START_TIME = time.monotonic()


def elapsed() -> str:
    return f"{time.monotonic() - START_TIME:.2f}s"


def log(message: str = "") -> None:
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    print(f"[{timestamp}] [+{elapsed()}] {message}", flush=True)


def separator(char: str = "=", width: int = 72) -> None:
    print(char * width, flush=True)


# ============================================================
# HELPERS
# ============================================================

def model_short_name(name: str) -> str:
    if name.startswith("models/"):
        return name[len("models/"):]
    return name


def pretty_json(value: Any) -> str:
    return json.dumps(
        value,
        indent=2,
        ensure_ascii=False,
        sort_keys=False,
    )


def get_model_by_short_name(
    models: list[dict[str, Any]],
    requested: str,
) -> dict[str, Any] | None:

    requested_clean = requested.strip()

    for model in models:
        full_name = str(model.get("name", ""))
        short_name = model_short_name(full_name)

        if requested_clean in (full_name, short_name):
            return model

    return None


def supports_generate_content(model: dict[str, Any]) -> bool:
    actions = model.get("supportedGenerationMethods")

    if not isinstance(actions, list):
        actions = model.get("supportedActions")

    if not isinstance(actions, list):
        return False

    return "generateContent" in actions


def choose_generation_model(
    models: list[dict[str, Any]],
) -> dict[str, Any] | None:

    # --------------------------------------------------------
    # Explicit model requested by workflow/environment.
    # --------------------------------------------------------

    if REQUESTED_MODEL:
        selected = get_model_by_short_name(
            models,
            REQUESTED_MODEL,
        )

        if selected is None:
            log(
                f"Requested model '{REQUESTED_MODEL}' "
                "was NOT returned by models.list."
            )
            return None

        if not supports_generate_content(selected):
            log(
                f"Requested model '{REQUESTED_MODEL}' "
                "does not advertise generateContent."
            )
            return None

        return selected

    # --------------------------------------------------------
    # Preferred models.
    # --------------------------------------------------------

    for preferred in PREFERRED_MODELS:
        selected = get_model_by_short_name(
            models,
            preferred,
        )

        if selected and supports_generate_content(selected):
            return selected

    # --------------------------------------------------------
    # Fallback: first generateContent model.
    # --------------------------------------------------------

    for model in models:
        if supports_generate_content(model):
            return model

    return None


# ============================================================
# LIST MODELS
# ============================================================

def list_all_models() -> tuple[list[dict[str, Any]], bool]:
    """
    Returns:

        (models, success)
    """

    separator()
    log("STEP 1 - LIST MODELS")
    separator()

    url = f"{API_BASE}/models"

    all_models: list[dict[str, Any]] = []
    page_token: str | None = None
    page_number = 0

    while True:
        page_number += 1

        params = {
            "key": API_KEY,
            "pageSize": "1000",
        }

        if page_token:
            params["pageToken"] = page_token

        log(f"Requesting models page {page_number}...")

        started = time.monotonic()

        try:
            response = requests.get(
                url,
                params=params,
                timeout=REQUEST_TIMEOUT_SECONDS,
            )
        except requests.RequestException as exc:
            log(f"NETWORK ERROR: {exc}")
            return all_models, False

        duration = time.monotonic() - started

        log(
            f"HTTP status: {response.status_code} "
            f"({duration:.2f}s)"
        )

        if response.status_code != 200:
            log("MODELS LIST REQUEST FAILED")

            try:
                body = response.json()
                print(pretty_json(body), flush=True)
            except ValueError:
                print(response.text, flush=True)

            return all_models, False

        try:
            payload = response.json()
        except ValueError:
            log("ERROR: Google returned non-JSON response.")
            print(response.text, flush=True)
            return all_models, False

        page_models = payload.get("models", [])

        if not isinstance(page_models, list):
            log("ERROR: 'models' field is not a list.")
            print(pretty_json(payload), flush=True)
            return all_models, False

        all_models.extend(page_models)

        log(
            f"Models returned on page {page_number}: "
            f"{len(page_models)}"
        )

        page_token = payload.get("nextPageToken")

        if not page_token:
            break

    log(f"TOTAL MODELS RETURNED: {len(all_models)}")

    return all_models, True


# ============================================================
# PRINT MODEL INVENTORY
# ============================================================

def print_model_inventory(
    models: list[dict[str, Any]],
) -> None:

    separator()
    log("MODEL INVENTORY")
    separator()

    generate_models: list[dict[str, Any]] = []

    for model in models:
        if supports_generate_content(model):
            generate_models.append(model)

    print("", flush=True)

    log(
        f"Models supporting generateContent: "
        f"{len(generate_models)}"
    )

    print("", flush=True)

    for index, model in enumerate(generate_models, start=1):

        name = model.get("name", "")
        display_name = model.get("displayName", "")
        description = model.get("description", "")
        version = model.get("version", "")

        input_limit = model.get("inputTokenLimit")
        output_limit = model.get("outputTokenLimit")

        print(
            f"[{index}] {model_short_name(str(name))}",
            flush=True,
        )

        if display_name:
            print(
                f"    displayName: {display_name}",
                flush=True,
            )

        if version:
            print(
                f"    version: {version}",
                flush=True,
            )

        if input_limit is not None:
            print(
                f"    inputTokenLimit: {input_limit}",
                flush=True,
            )

        if output_limit is not None:
            print(
                f"    outputTokenLimit: {output_limit}",
                flush=True,
            )

        if description:
            one_line = " ".join(
                str(description).split()
            )

            if len(one_line) > 300:
                one_line = one_line[:300] + "..."

            print(
                f"    description: {one_line}",
                flush=True,
            )

        print(
            "    supportedGenerationMethods: "
            f"{model.get('supportedGenerationMethods', model.get('supportedActions', []))}",
            flush=True,
        )

        print("", flush=True)


# ============================================================
# CHECK SPECIFIC MODELS
# ============================================================

def print_requested_model_status(
    models: list[dict[str, Any]],
) -> None:

    separator()
    log("SPECIFIC MODEL CHECK")
    separator()

    targets = [
        "gemini-3.8-flash",
        "gemini-3.7-flash",
        "gemini-3.6-flash",
        "gemini-3.5-flash",
        "gemini-3.5-flash-lite",
        "gemini-2.5-flash",
        "gemini-2.5-flash-lite",
    ]

    for target in targets:

        model = get_model_by_short_name(
            models,
            target,
        )

        if model is None:
            print(
                f"{target}: NOT RETURNED BY models.list",
                flush=True,
            )
            continue

        actions = model.get(
            "supportedGenerationMethods",
            model.get("supportedActions", []),
        )

        print(
            f"{target}: AVAILABLE",
            flush=True,
        )

        print(
            f"    generateContent: "
            f"{'YES' if 'generateContent' in actions else 'NO'}",
            flush=True,
        )

        if model.get("inputTokenLimit") is not None:
            print(
                f"    inputTokenLimit: "
                f"{model['inputTokenLimit']}",
                flush=True,
            )

        if model.get("outputTokenLimit") is not None:
            print(
                f"    outputTokenLimit: "
                f"{model['outputTokenLimit']}",
                flush=True,
            )


# ============================================================
# MINIMAL GENERATION TEST
# ============================================================

def test_generation(
    model: dict[str, Any],
) -> bool:

    separator()
    log("STEP 2 - MINIMAL GENERATION TEST")
    separator()

    model_name = model_short_name(
        str(model.get("name", ""))
    )

    log(f"Selected model: {model_name}")

    prompt = "Reply with exactly: GEMINI_DIAGNOSTIC_OK"

    request_body = {
        "contents": [
            {
                "parts": [
                    {
                        "text": prompt
                    }
                ]
            }
        ],
        "generationConfig": {
            "temperature": 0,
            "maxOutputTokens": 32,
        },
    }

    url = (
        f"{API_BASE}/models/"
        f"{model_name}:generateContent"
    )

    log("Prompt: GEMINI_DIAGNOSTIC_OK")
    log("Prompt characters: 33")
    log("Sending minimal Gemini request...")

    started = time.monotonic()

    try:
        response = requests.post(
            url,
            params={"key": API_KEY},
            headers={
                "Content-Type": "application/json",
            },
            json=request_body,
            timeout=REQUEST_TIMEOUT_SECONDS,
        )
    except requests.RequestException as exc:
        log(f"NETWORK ERROR: {exc}")
        return False

    duration = time.monotonic() - started

    log(
        f"Gemini HTTP status: {response.status_code}"
    )

    log(
        f"Gemini response time: {duration:.2f}s"
    )

    print("", flush=True)

    if response.status_code != 200:

        log("MINIMAL GENERATION FAILED")
        log(f"HTTP {response.status_code}")

        print("Response:", flush=True)

        try:
            body = response.json()
            print(
                pretty_json(body),
                flush=True,
            )
        except ValueError:
            print(
                response.text,
                flush=True,
            )

        return False

    log("MINIMAL GENERATION SUCCEEDED")

    try:
        payload = response.json()
    except ValueError:
        log("ERROR: Successful response was not JSON.")
        print(response.text, flush=True)
        return False

    print("Raw successful response:", flush=True)
    print(
        pretty_json(payload),
        flush=True,
    )

    # --------------------------------------------------------
    # Extract text if available.
    # --------------------------------------------------------

    extracted_text = None

    try:
        candidates = payload.get("candidates", [])

        if candidates:
            parts = (
                candidates[0]
                .get("content", {})
                .get("parts", [])
            )

            for part in parts:
                if "text" in part:
                    extracted_text = part["text"]
                    break

    except Exception:
        extracted_text = None

    print("", flush=True)

    if extracted_text is not None:
        log(
            f"Generated text: {extracted_text!r}"
        )

    return True


# ============================================================
# MAIN
# ============================================================

def main() -> int:

    separator()
    log("VGrat FMS - GEMINI API DIAGNOSTIC")
    separator()

    log("Purpose: determine API-key model visibility and generation access.")
    log("No project files will be modified.")
    print("", flush=True)

    # --------------------------------------------------------
    # API key check
    # --------------------------------------------------------

    if not API_KEY:
        log("ERROR: GEMINI_API_KEY is not set.")
        log("Diagnostic cannot continue.")
        return 1

    log(
        "GEMINI_API_KEY detected."
    )

    log(
        f"API key length: {len(API_KEY)} characters"
    )

    log(
        "API key value will NOT be printed."
    )

    print("", flush=True)

    # --------------------------------------------------------
    # List models
    # --------------------------------------------------------

    models, list_success = list_all_models()

    if not list_success:
        separator()
        log("DIAGNOSTIC RESULT: FAILED")
        separator()
        log(
            "The API key could not successfully retrieve "
            "the Gemini model list."
        )
        log(
            "This is the first issue to investigate."
        )
        log("No files were modified.")
        return 1

    if not models:
        separator()
        log("DIAGNOSTIC RESULT: FAILED")
        separator()
        log(
            "The API returned an empty model list."
        )
        log(
            "The API key is accepted but no models were "
            "returned to this request."
        )
        log("No files were modified.")
        return 1

    # --------------------------------------------------------
    # Print inventory
    # --------------------------------------------------------

    print_model_inventory(models)

    print_requested_model_status(models)

    # --------------------------------------------------------
    # Select model
    # --------------------------------------------------------

    separator()
    log("MODEL SELECTION")
    separator()

    if REQUESTED_MODEL:
        log(
            f"GEMINI_MODEL explicitly supplied: "
            f"{REQUESTED_MODEL}"
        )
    else:
        log(
            "GEMINI_MODEL was not supplied."
        )
        log(
            "The diagnostic will select the first "
            "preferred generateContent model available."
        )

    selected_model = choose_generation_model(models)

    if selected_model is None:

        separator()
        log("DIAGNOSTIC RESULT: FAILED")
        separator()

        if REQUESTED_MODEL:
            log(
                f"The requested model '{REQUESTED_MODEL}' "
                "was not available for generateContent."
            )
        else:
            log(
                "No generateContent-capable model was returned."
            )

        log("No files were modified.")

        return 1

    selected_name = model_short_name(
        str(selected_model.get("name", ""))
    )

    log(
        f"Selected diagnostic model: {selected_name}"
    )

    # --------------------------------------------------------
    # Generation
    # --------------------------------------------------------

    generation_success = test_generation(
        selected_model
    )

    # --------------------------------------------------------
    # Final result
    # --------------------------------------------------------

    separator()

    if generation_success:

        log("DIAGNOSTIC RESULT: SUCCESS")
        separator()

        log(
            "The API key can list Gemini models and "
            "successfully generate content."
        )

        log(
            f"Working model: {selected_name}"
        )

        log(
            "The previous 503 problem is therefore likely "
            "specific to the tested model/capacity at that time "
            "or to the larger generation request."
        )

        log("No files were modified.")

        return 0

    else:

        log("DIAGNOSTIC RESULT: PARTIAL / FAILED")
        separator()

        log(
            "The API key successfully listed models, "
            "but the minimal generateContent request failed."
        )

        log(
            f"Model tested: {selected_name}"
        )

        log(
            "This strongly points toward generation availability, "
            "capacity, quota, project access, or service-side "
            "availability rather than model-list access."
        )

        log("No files were modified.")

        return 1


if __name__ == "__main__":
    sys.exit(main())
