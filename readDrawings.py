"""
drawing_parser.py

Parse engineering drawing PDFs and extract pricing-relevant data.

Works across different title block formats by using line-pattern matching
with multiple regex fallbacks. No external API calls required.

Public API:
    parse_drawing_pdf(path: str) -> dict

Dependencies:
    pip install pdfplumber
"""

from __future__ import annotations

import re
from pathlib import Path
from typing import Optional

import pdfplumber


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _clean(s: str) -> str:
    return " ".join(s.split()).strip()


# ---------------------------------------------------------------------------
# Title block extractor
# ---------------------------------------------------------------------------

def _extract_title_block_redrock(text: str) -> Optional[dict]:
    """
    Redrock Machinery title block. Layout (single line, columns interleaved
    with garbled boilerplate from the notes column):

        DRAWING NO. MASS QTY. / MC CHECKED BY UNLESS OTHERWISE ... MATERIAL
        300611-03 22.03 kg 1 <garble> COARSE W.MCC 12/11/2020 8mm Mild Steel

    Returns None if the sheet is not a Redrock drawing, so the caller falls
    through to the other formats.
    """
    if not re.search(r"\bREDROCK\b", text, re.I):
        return None

    tb: dict = {}

    # --- Main title row: part number, mass, qty, drawn by/date, thickness, material ---
    m = re.search(
        r"^([A-Z0-9][A-Z0-9\-]+)\s+([\d.]+)\s*kg\s+(\d+)\b.*?"
        r"(?:([A-Za-z][\w.']{0,15})\s+)?(\d{2}/\d{2}/\d{4})\s+"
        r"([\d.]+)\s*mm\s+(.+?)\s*$",
        text, re.I | re.M,
    )
    if m:
        tb["part_number"]     = m.group(1)
        tb["mass_kg"]         = float(m.group(2))
        tb["qty_per_machine"] = int(m.group(3))
        # "COARSE" is trailing tolerance-note garble, not a person
        if m.group(4) and m.group(4).upper() != "COARSE":
            tb["drawn_by"] = m.group(4)
        tb["drawn_date"]   = m.group(5)
        tb["thickness_mm"] = float(m.group(6))
        tb["material"]     = _clean(m.group(7))
    else:
        # Fallback: "Drawing No. Revision Sheet" box in the top-right corner
        m = re.search(
            r"Drawing\s+No\.?\s+Revision\s+Sheet\s*\n\s*([A-Z0-9][A-Z0-9\-]+)",
            text, re.I,
        )
        if not m:
            return None
        tb["part_number"] = m.group(1)

    # --- Revision: latest row of the REVISIONS table ---
    # e.g. "A Gearbox holes changed from 20mm to 18mm 05/07/2023 W McC"
    if re.search(r"^\s*REVISIONS\s*$", text, re.I | re.M):
        revs = re.findall(r"^\s*([A-Z])\s+(?:.*?\s)?\d{2}/\d{2}/\d{4}\b", text, re.M)
        if revs:
            tb["revision"] = max(revs)

    # --- Description: value line(s) below the "TITLE / DESCRIPTION" header,
    # with the paint column and tolerance-table garble appended ---
    lines = text.split("\n")
    for i, line in enumerate(lines):
        if not re.search(r"TITLE\s*/\s*DESCRIPTION", line, re.I):
            continue
        for cand_line in lines[i + 1:i + 5]:
            cand = re.split(
                r"\s+(?:NO\s+PAINT\b|PAINT\b|OVER\s+\d)", cand_line, maxsplit=1
            )[0].strip()
            if not cand or re.match(r"^(OVER|ALL\b|Reproduction|prohibited)", cand, re.I):
                continue
            if len(cand) >= 3 and re.search(r"[A-Za-z]", cand):
                tb["description"] = _clean(cand)
                break
        break

    # --- Paint / coating ---
    if re.search(r"NO\s+PAINT", text, re.I):
        tb["finish"] = "NO PAINT"

    # --- Sheet ---
    m = re.search(r"\b(\d+)\s+OF\s+(\d+)\b", text)
    if m:
        tb["sheet"] = f"{m.group(1)} OF {m.group(2)}"

    return tb


def _extract_title_block(page) -> dict:
    """
    Extract title block fields. Uses precise line-pattern matching for the
    Redrock and Wrightbus/Glovebox formats, with generic fallbacks for
    other formats.
    """
    text = page.extract_text() or ""

    redrock = _extract_title_block_redrock(text)
    if redrock is not None:
        return redrock

    tb: dict = {}

    # --- Part number + revision + status (Wrightbus/Glovebox) ---
    # Label line: "DRAWN SCALE (DO NOT SCALE OFF DRAWING)  STATUS  PART NUMBER  REVISION"
    # Value line: "2.000  PROJECTION  Draft  A11-00817  2"
    m = re.search(
        r"STATUS\s+PART\s+NUMBER\s+REVISION\s*\n"
        r".*?\b(Draft|Released|Approved|Obsolete|Preliminary)\b"
        r"\s+([A-Z][A-Z0-9\-]+)\s+(\d+)\s*$",
        text, re.I | re.M | re.S,
    )
    if m:
        tb["status"]      = m.group(1)
        tb["part_number"] = m.group(2)
        tb["revision"]    = m.group(3)
    else:
        for pat in (
            r"PART\s+NO\.?\s*[:\-]?\s*([A-Z][A-Z0-9\-]+)",
            r"DRAWING\s+NO\.?\s*[:\-]?\s*([A-Z][A-Z0-9\-]+)",
            r"DRG\.?\s+NO\.?\s*[:\-]?\s*([A-Z][A-Z0-9\-]+)",
        ):
            m = re.search(pat, text, re.I)
            if m:
                tb["part_number"] = m.group(1)
                break
        # Explicit "REVISION: x" label (SolidWorks/Groundsman title blocks)
        m = re.search(r"REVISION\s*:\s*([A-Z0-9]+)\b", text, re.I)
        if not m:
            # Generic revision: standalone value after a REV label. \b after the
            # group stops "REVISIONS" (table header) matching as REVISION + "S".
            m = re.search(r"\bREV(?:ISION)?\b\s*[:\-]?\s*([A-Z0-9]+)\b", text, re.I)
        if m and m.group(1).upper() not in ("BY", "DATE", "HISTORY", "ISED", "ISION"):
            tb["revision"] = m.group(1)

    # --- Description ---
    # Wrightbus: garbled copyright line ends with "Kg  <DESCRIPTION>  make_or_buy"
    m = re.search(r"\bKg\b\s+(.+?)\s+(?:make_or_buy|BUY|MAKE)\s*$", text, re.M)
    if m:
        tb["description"] = _clean(m.group(1))
    else:
        # Revision-table headers ("ZONE REV. DESCRIPTION DATE APPROVED") make the
        # generic DESCRIPTION pattern capture neighbouring labels — reject those.
        _label_junk = re.compile(
            r"(?:ZONE|REV\.?|DESCRIPTION|DATE|APPROVED|SOURCING)"
            r"(?:\s+(?:ZONE|REV\.?|DESCRIPTION|DATE|APPROVED|SOURCING))*",
            re.I,
        )
        for pat in (
            r"DESCRIPTION\s+SOURCING\s*\n\s*(.+?)(?:\n|make_or_buy)",
            r"DESCRIPTION\s*[:\-]?\s*(.+?)(?:\n|DRAWING|PART\s+NUMBER)",
            r"TITLE\s*[:\-]?\s*(.+?)(?:\n|$)",
        ):
            m = re.search(pat, text, re.I | re.S)
            if m:
                candidate = _clean(m.group(1))
                if len(candidate) > 3 and not _label_junk.fullmatch(candidate):
                    tb["description"] = candidate
                    break

    # --- Material / finish / thickness (Wrightbus) ---
    # Header: "MATERIAL SPECIFICATION  MATERIAL  FINISH  SHEETMETAL THICKNESS"
    # Value:  "ALUMINIUM ALLOY TO WMP-1096  FA03N  N/A  3.000 mm"
    m = re.search(
        r"MATERIAL\s+SPECIFICATION.+\n(.+?)\s+(FA\w+|N/A)\s+(?:N/A|\w+)\s+([\d.]+)\s*mm",
        text, re.I,
    )
    if m:
        tb["material"]     = _clean(m.group(1))
        tb["finish"]       = m.group(2)
        tb["thickness_mm"] = float(m.group(3))
    else:
        for pat in (
            r"MATERIAL\s*[:\-]?\s*(.+?)(?:\n|FINISH|THICKNESS)",
            r"MAT\.?\s*[:\-]?\s*(.+?)(?:\n|$)",
        ):
            m = re.search(pat, text, re.I)
            if m:
                tb["material"] = _clean(m.group(1))
                break
        for pat in (
            r"THICKNESS\s*[:\-]?\s*([\d.]+)\s*mm",
            r"([\d.]+)\s*mm\s+(?:THK|THICK)",
        ):
            m = re.search(pat, text, re.I)
            if m:
                try:
                    tb["thickness_mm"] = float(m.group(1))
                except ValueError:
                    pass
                break

    # --- Status (if not already set) ---
    if "status" not in tb:
        m = re.search(r"\b(Draft|Released|Approved|Obsolete|Preliminary)\b", text, re.I)
        if m:
            tb["status"] = _clean(m.group(1))

    # --- Drawn by / drawn date ---
    # Wrightbus: "ALL TOLERANCES TO COMPLY  MC  28-Mar-13  cad_type  1"
    m = re.search(
        r"ALL\s+TOLERANCES\s+TO\s+COMPLY\s+(\w+)\s+([\d]+-[A-Za-z]+-[\d]+)", text
    )
    if m:
        tb["drawn_by"]   = m.group(1)
        tb["drawn_date"] = m.group(2)
    else:
        m = re.search(r"DRAWN\s+BY\s*[:\-]?\s*(\w+)", text, re.I)
        if m and m.group(1) not in ("DRAWN", "DATE"):
            tb["drawn_by"] = m.group(1)
        m = re.search(r"DRAWN\s+DATE\s*[:\-]?\s*([\d\-A-Za-z]+)", text, re.I)
        if m:
            tb["drawn_date"] = m.group(1)

    # --- Checked by / checked date ---
    # Wrightbus: "MUST COMPLY WITH WMP-2040  CC  28-Mar-13  part_finish  2"
    m = re.search(
        r"MUST\s+COMPLY\s+WITH\s+\S+\s+(\w+)\s+([\d]+-[A-Za-z]+-[\d]+)", text
    )
    if m:
        tb["checked_by"]   = m.group(1)
        tb["checked_date"] = m.group(2)
    else:
        m = re.search(r"(?:DRAWING\s+)?CHECKED\s+BY\s*[:\-]?\s*(\w+)", text, re.I)
        if m and m.group(1) not in ("BY", "DATE", "DRAWING"):
            tb["checked_by"] = m.group(1)

    # --- Scale ---
    m = re.search(r"SCALE\s*\(DO\s+NOT[^)]+\)\s*\n?\s*([\d.]+)", text)
    if not m:
        m = re.search(r"\bSCALE\s+([\d.:]+)\b", text)
    if m:
        try:
            tb["scale"] = float(m.group(1))
        except ValueError:
            tb["scale"] = m.group(1)

    # --- Drawing size ---
    m = re.search(r"DRAWING\s+SIZE\s*[:\-]?\s*(A\d|B|C|D)\b", text, re.I)
    if m:
        tb["drawing_size"] = m.group(1).upper()

    # --- Latest revision change ---
    # Wrightbus: "2  Scott Logan  30-Apr-24  MANUFACTURING EFFICENCY CHANGE  ECR-33703"
    m = re.search(
        r"^\s*\d+\s+\w+\s+\w+\s+[\d]+-[A-Za-z]+-[\d]+\s+(.+?)\s+ECR-\w+\s*$",
        text, re.M,
    )
    if m:
        tb["latest_change_description"] = _clean(m.group(1))

    return tb


# ---------------------------------------------------------------------------
# Pricing factors extractor
# ---------------------------------------------------------------------------

_ANGLE_RE   = re.compile(r"(\d+(?:\.\d+)?)°")
# Redrock bend callouts: "FOLD UP 15o" / "FOLD DOWN 48°" (degree sign is
# sometimes extracted as a letter "o")
_FOLD_RE    = re.compile(r"FOLD\s+(?:UP|DOWN)\s+(\d+(?:\.\d+)?)", re.I)
_SLOT_RE    = re.compile(r"(\d+(?:\.\d+)?)\s*MM\s+SLOT", re.I)
_HOLE_RE    = re.compile(r"[ØΦ∅](\d+(?:\.\d+)?)")
_RADIUS_RE  = re.compile(r"R(\d+(?:\.\d+)?)\s*(?:TYP)?", re.I)
_DIM_RE     = re.compile(r"^(\d+(?:\.\d+)?)$")

_TB_NOISE_PHRASES = (
    "MATERIAL SPECIFICATION", "DRAWN BY", "CHECKED", "REVISION",
    "UNLESS STAMPED", "UNLESS OTHERWISE", "TOLERANCES", "WMP-",
    "WES-", "BALLYMENA", "ELECTRONIC DATA", "PROJECTION",
    "FORMAT NAME", "DRAWING SIZE", "DO NOT SCALE", "NORTHERN IRELAND",
    "ALL DIMENSIONS", "ESTIMATED MASS", "SOURCING", "COMPLY WITH",
    "BURRS", "SHARP EDGES", "ELECTRONIC", "Scott Logan",
    # Redrock boilerplate (tolerance table, address block, copyright)
    "REDROCK", "Redrock", "TOLERANCES (DIN", "UP TO", "PAINT / COATING",
    "DRAWING NO. MASS", "Reproduction in part", "prohibited",
    "www.redrockmachinery", "Collone", "Co.Armagh", "THIRD ANGLE",
)


def _extract_pricing_factors(page, tb_top_y: float) -> dict:
    """
    Extract bend angles, slots, holes, dimensions and process notes
    from the drawing area (above the title block).
    """
    text = page.extract_text() or ""
    words = page.extract_words()

    # Drawing area words only
    drawing_words = [w for w in words if w["bottom"] < tb_top_y]

    # Filter text lines to drawing area
    drawing_lines = [
        line for line in text.split("\n")
        if not any(phrase in line for phrase in _TB_NOISE_PHRASES)
    ]
    drawing_text = "\n".join(drawing_lines)

    factors: dict = {}

    # --- Bend angles ---
    # Explicit fold callouts first (Redrock), then strip them so the generic
    # degree-symbol pattern doesn't count the same callout twice.
    fold_angle_strs = _FOLD_RE.findall(drawing_text)
    all_angle_strs  = _ANGLE_RE.findall(_FOLD_RE.sub(" ", drawing_text))
    angle_counts: dict[float, int] = {}
    for a in all_angle_strs:
        angle_counts[float(a)] = angle_counts.get(float(a), 0) + 1

    # 90.00 is a flatness/perpendicularity callout; 90° appearing multiple
    # times means genuine bend angles are present
    bend_angles = sorted(set(
        float(a) for a in all_angle_strs
        if not (float(a) == 90.0 and "90.00" in drawing_text and angle_counts.get(90.0, 0) <= 1)
    ) | set(float(a) for a in fold_angle_strs))
    if bend_angles:
        factors["bend_angles_deg"] = bend_angles
        factors["bend_count"]      = len(all_angle_strs) + len(fold_angle_strs)
        factors["bending_required"] = True
    else:
        factors["bending_required"] = False

    # --- Slots ---
    slot_matches = _SLOT_RE.findall(drawing_text)
    if slot_matches:
        factors["slots"] = sorted(set(f"{s}MM SLOT" for s in slot_matches))

    # --- Holes / diameters ---
    holes = sorted(set(float(h) for h in _HOLE_RE.findall(drawing_text)))
    dia_matches = re.findall(r"(?:DIA\.?|DIAM\.?)\s*([\d.]+)", drawing_text, re.I)
    holes = sorted(set(holes) | {float(d) for d in dia_matches})
    if holes:
        factors["holes_dia_mm"] = holes

    # --- Radii ---
    radii = sorted(set(float(r) for r in _RADIUS_RE.findall(drawing_text)))
    if radii:
        factors["radii_mm"] = radii

    # --- Linear dimensions (best-effort) ---
    grid_labels = {str(i) for i in range(1, 9)}
    dim_values: set[float] = set()
    for w in drawing_words:
        t = w["text"]
        if _DIM_RE.match(t) and t not in grid_labels:
            val = float(t)
            if val > 1:
                dim_values.add(val)
        elif re.match(r"^\d+\.\d+$", t):
            val = float(t)
            if val > 1:
                dim_values.add(val)

    if dim_values:
        factors["dimensions_mm"] = sorted(dim_values)
        factors["dimensions_note"] = (
            "Best-effort extraction — rotated dimension text may be missing or garbled"
        )

    # --- Process / post-processing notes ---
    notes = []
    note_patterns = [
        r"ALL\s+BURRS?\s+AND\s+SHARP\s+EDGES?\s+TO\s+BE\s+REMOVED",
        r"REMOVE\s+ALL\s+BURRS?\s*&?\s*(?:AND\s+)?SHARP\s+EDGES?",
        r"DEBURR\s+ALL\s+EDGES?",
        r"BREAK\s+ALL\s+SHARP\s+EDGES?",
        r"COUNTERSINK",
        r"COUNTERBORE",
        r"THREAD(?:ED)?",
        r"WELD(?:ED)?",
        r"PRESS\s+FIT",
        r"POWDER\s+COAT",
        r"ANODIS(?:E|ED)",
        # not "NO PAINT" and not the "PAINT / COATING" column header
        r"(?<!NO )\bPAINT(?:ED)?\b(?!\s*/\s*COATING)",
        r"ZINC\s+PLAT(?:E|ED)",
        r"SEE\s+DXF\s+FOR\s+CUT\s+PROFILE",
        r"MUST\s+COMPLY\s+WITH\s+[\w-]+",
    ]
    for pat in note_patterns:
        m = re.search(pat, text, re.I)
        if m:
            notes.append(_clean(m.group(0)))
    if notes:
        factors["process_notes"] = notes

    return factors


# ---------------------------------------------------------------------------
# Public entry point
# ---------------------------------------------------------------------------

def parse_drawing_pdf(path: str) -> dict:
    """
    Parse an engineering drawing PDF and extract pricing-relevant data.

    Works across different title block formats using pattern matching —
    no external API calls required.

    Args:
        path: Absolute or relative path to the PDF file.

    Returns:
        {
            "source_file": str,
            "pages": [
                {
                    "page": int,
                    "title_block": {
                        "part_number": str,
                        "revision": str,
                        "description": str,
                        "material": str,
                        "thickness_mm": float,
                        "finish": str,
                        "status": str,
                        "drawn_by": str,
                        "drawn_date": str,
                        "checked_by": str,
                        "checked_date": str,
                        "scale": float,
                        "drawing_size": str,
                        "latest_change_description": str,
                        # fields not found are omitted rather than null
                    },
                    "pricing_factors": {
                        "bending_required": bool,
                        "bend_angles_deg": [float, ...],   # unique angles present
                        "bend_count": int,                  # total angle callouts (proxy for op count)
                        "slots": ["7MM SLOT", ...],
                        "holes_dia_mm": [float, ...],
                        "radii_mm": [float, ...],
                        "dimensions_mm": [float, ...],      # best-effort, may be incomplete
                        "dimensions_note": str,
                        "process_notes": [str, ...],
                    }
                }
            ]
        }

    Raises:
        FileNotFoundError: If the PDF path does not exist.
    """
    pdf_path = Path(path)
    if not pdf_path.exists():
        raise FileNotFoundError(f"PDF not found: {path}")

    result = {
        "source_file": pdf_path.name,
        "pages": [],
    }

    with pdfplumber.open(pdf_path) as pdf:
        for i, page in enumerate(pdf.pages):
            text = page.extract_text() or ""
            if not text.strip():
                continue

            # Title block starts roughly in the bottom 22% of the page
            tb_top_y = page.height * 0.78

            title_block     = _extract_title_block(page)
            pricing_factors = _extract_pricing_factors(page, tb_top_y)

            result["pages"].append({
                "page":            i + 1,
                "title_block":     title_block,
                "pricing_factors": pricing_factors,
            })

    return result


# ---------------------------------------------------------------------------
# CLI convenience
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    import sys
    import json

    if len(sys.argv) < 2:
        print("Usage: python drawing_parser.py <path-to-pdf>")
        sys.exit(1)

    result = parse_drawing_pdf(sys.argv[1])
    print(json.dumps(result, indent=2))