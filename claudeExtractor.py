"""
claudeExtractor.py

Claude vision fallback for drawing pages the regex parser in readDrawings.py
doesn't recognise (unknown title block layout, garbled text layer, or scanned
pages with no text layer at all).

Sends a single page as a PDF document block — Claude reads both the text
layer and the rendered image — and forces the response into the same
title_block / pricing_factors shape readDrawings.py produces, via a
structured-output JSON schema.

Auth: ANTHROPIC_API_KEY in project/.env (loaded here so the module works
from the CLI as well as under app.py's load_dotenv()).

Dependencies:
    pip install anthropic pymupdf python-dotenv
"""

from __future__ import annotations

import base64
import json
import os
from pathlib import Path

import fitz  # PyMuPDF
from dotenv import load_dotenv

load_dotenv(Path(__file__).resolve().parent / ".env")

_MODEL = "claude-opus-4-8"
_MAX_TOKENS = 10000


class ClaudeExtractionError(RuntimeError):
    """Claude fallback failed for this page. `permanent` means the caller
    should stop trying for the rest of the run (missing/invalid API key)."""

    def __init__(self, message: str, permanent: bool = False):
        super().__init__(message)
        self.permanent = permanent


_PROMPT = """This is one page of an engineering drawing for sheet-metal fabrication pricing.
Extract the title block fields and pricing-relevant manufacturing features.

Title block guidance:
- part_number is usually the drawing number in the title block (bottom or bottom-right).
- revision is the revision letter/number for the current issue.
- material is the material specification (e.g. "ALUMINIUM ALLOY TO WMP-1096", "8mm Mild Steel").
- thickness_mm is the sheet thickness, often written like "3mm" or "3.000 mm" near the material.
- finish is the surface finish / paint / coating callout (e.g. "FA03N", "NO PAINT").
- status is the drawing status (Draft, Released, Approved, Obsolete, Preliminary).
- drawn_by / drawn_date and checked_by / checked_date come from the signature boxes.
- scale is the drawing scale — a number where possible, a ratio string like "1:2" otherwise.
- drawing_size is the sheet size (A0-A4, B, C, D).
- latest_change_description is the description text of the newest row in the revision table.
- Set any title block field you cannot find to null. Do not guess values.

Pricing factors guidance:
- bend_angles_deg: the unique bend/fold angles called out on the drawing
  (e.g. "FOLD UP 45" means a 45 degree bend). A single 90.00 callout is usually a
  perpendicularity tolerance, not a bend.
- bend_count: the total number of bend/fold callouts including repeats
  (a proxy for the number of bending operations).
- bending_required: true if there is at least one bend, otherwise false.
- slots: each slot formatted as "<width>MM SLOT", e.g. "7MM SLOT".
- holes_present: true if the drawing shows any drilled/punched through-holes
  (diameter symbol callouts), false if none. Do not list diameters or dimensions.
- radii_mm: radius values from R callouts (e.g. R5).
- process_notes: manufacturing/post-processing notes verbatim - deburring, welding,
  threading, coating/painting, compliance notes (e.g. "MUST COMPLY WITH WMP-2040"),
  "SEE DXF FOR CUT PROFILE", etc.
- Use empty lists for pricing factor lists with no callouts found."""

_TITLE_BLOCK_SCHEMA = {
    "type": "object",
    "properties": {
        "part_number":               {"type": ["string", "null"]},
        "revision":                  {"type": ["string", "null"]},
        "description":               {"type": ["string", "null"]},
        "material":                  {"type": ["string", "null"]},
        "thickness_mm":              {"type": ["number", "null"]},
        "finish":                    {"type": ["string", "null"]},
        "status":                    {"type": ["string", "null"]},
        "drawn_by":                  {"type": ["string", "null"]},
        "drawn_date":                {"type": ["string", "null"]},
        "checked_by":                {"type": ["string", "null"]},
        "checked_date":              {"type": ["string", "null"]},
        "scale":                     {"type": ["number", "string", "null"]},
        "drawing_size":              {"type": ["string", "null"]},
        "latest_change_description": {"type": ["string", "null"]},
    },
    "required": [
        "part_number", "revision", "description", "material", "thickness_mm",
        "finish", "status", "drawn_by", "drawn_date", "checked_by",
        "checked_date", "scale", "drawing_size", "latest_change_description",
    ],
    "additionalProperties": False,
}

_PRICING_FACTORS_SCHEMA = {
    "type": "object",
    "properties": {
        "bending_required": {"type": "boolean"},
        "bend_angles_deg":  {"type": "array", "items": {"type": "number"}},
        "bend_count":       {"type": "integer"},
        "slots":            {"type": "array", "items": {"type": "string"}},
        "holes_present":    {"type": "boolean"},
        "radii_mm":         {"type": "array", "items": {"type": "number"}},
        "process_notes":    {"type": "array", "items": {"type": "string"}},
    },
    "required": [
        "bending_required", "bend_angles_deg", "bend_count", "slots",
        "holes_present", "radii_mm", "process_notes",
    ],
    "additionalProperties": False,
}

_SCHEMA = {
    "type": "object",
    "properties": {
        "title_block":     _TITLE_BLOCK_SCHEMA,
        "pricing_factors": _PRICING_FACTORS_SCHEMA,
    },
    "required": ["title_block", "pricing_factors"],
    "additionalProperties": False,
}

# Set after an auth failure so a batch run doesn't retry a dead key per page
_disabled = False


def _single_page_pdf_b64(path: Path, page_index: int) -> str:
    with fitz.open(path) as src:
        out = fitz.open()
        out.insert_pdf(src, from_page=page_index, to_page=page_index)
        data = out.tobytes()
        out.close()
    return base64.standard_b64encode(data).decode()


def extract_page_with_claude(pdf_path: Path, page_index: int) -> dict:
    """
    Extract title block + pricing factors for one page (0-based index).

    Returns {"title_block": {...}, "pricing_factors": {...}} matching
    readDrawings.py conventions: unfound title block fields and empty
    pricing-factor lists are omitted rather than null/[].

    Raises:
        ClaudeExtractionError: on any API or parsing failure. Check
            .permanent to decide whether to stop retrying for the run.
    """
    global _disabled
    if _disabled:
        raise ClaudeExtractionError(
            "Claude fallback disabled after earlier authentication failure",
            permanent=True,
        )
    if not os.environ.get("ANTHROPIC_API_KEY"):
        raise ClaudeExtractionError(
            "ANTHROPIC_API_KEY not set — add it to project/.env",
            permanent=True,
        )

    import anthropic

    client = anthropic.Anthropic()
    try:
        message = client.messages.create(
            model=_MODEL,
            max_tokens=_MAX_TOKENS,
            output_config={"format": {"type": "json_schema", "schema": _SCHEMA}},
            messages=[{
                "role": "user",
                "content": [
                    {
                        "type": "document",
                        "source": {
                            "type": "base64",
                            "media_type": "application/pdf",
                            "data": _single_page_pdf_b64(pdf_path, page_index),
                        },
                    },
                    {"type": "text", "text": _PROMPT},
                ],
            }],
        )
    except anthropic.AuthenticationError as e:
        _disabled = True
        raise ClaudeExtractionError(
            f"invalid ANTHROPIC_API_KEY: {e.message}", permanent=True
        ) from e
    except anthropic.APIStatusError as e:
        raise ClaudeExtractionError(f"API error {e.status_code}: {e.message}") from e
    except anthropic.APIConnectionError as e:
        raise ClaudeExtractionError(f"connection error: {e}") from e

    if message.stop_reason == "max_tokens":
        raise ClaudeExtractionError("response truncated at max_tokens")

    raw = next((b.text for b in message.content if b.type == "text"), None)
    if raw is None:
        raise ClaudeExtractionError(
            f"no text in response (stop_reason={message.stop_reason})"
        )
    try:
        data = json.loads(raw)
    except json.JSONDecodeError as e:
        raise ClaudeExtractionError(f"unparseable response JSON: {e}") from e

    return {
        "title_block": {
            k: v for k, v in data["title_block"].items() if v is not None
        },
        "pricing_factors": {
            k: v for k, v in data["pricing_factors"].items() if v not in ([], None)
        },
    }


if __name__ == "__main__":
    import sys

    if len(sys.argv) < 2:
        print("Usage: python claudeExtractor.py <path-to-pdf> [page-number]")
        sys.exit(1)

    page_no = int(sys.argv[2]) if len(sys.argv) > 2 else 1
    print(json.dumps(extract_page_with_claude(Path(sys.argv[1]), page_no - 1), indent=2))
