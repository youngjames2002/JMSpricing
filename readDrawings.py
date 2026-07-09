"""
readDrawings.py

Parse engineering drawing PDFs and extract pricing-relevant data.

Known title block layouts (Redrock, Wrightbus, Groundsman/SolidWorks) are
parsed with line-pattern regexes — fast and free. Pages in an unrecognised
format, pages where the regex extraction comes back too sparse to price
from, and scanned pages with no text layer fall back to the Claude API
(claudeExtractor.py, vision-based). If the fallback is unavailable or
fails, the parser degrades to best-effort generic regexes.

Each parsed page carries an "extraction_method" field so downstream
pricing can tell a precise format match from a model read or a generic
regex guess.

Public API:
    parse_drawing_pdf(path: str, use_claude: bool = True) -> dict

Dependencies:
    pip install pdfplumber          # regex path
    pip install anthropic pymupdf python-dotenv   # Claude fallback
    ANTHROPIC_API_KEY in project/.env             # Claude fallback
"""

from __future__ import annotations

import re
import sys
from pathlib import Path
from typing import Callable, Optional

import pdfplumber


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _clean(s: str) -> str:
    return " ".join(s.split()).strip()


# ---------------------------------------------------------------------------
# Title block extractors — one per known layout
# ---------------------------------------------------------------------------

def _extract_title_block_redrock(text: str, page=None) -> Optional[dict]:
    """
    Redrock Machinery title block. Layout (single line, columns interleaved
    with garbled boilerplate from the notes column):

        DRAWING NO. MASS QTY. / MC CHECKED BY UNLESS OTHERWISE ... MATERIAL
        300611-03 22.03 kg 1 <garble> COARSE W.MCC 12/11/2020 8mm Mild Steel

    Returns None if the expected patterns are absent, so the caller falls
    through to the Claude fallback.
    """
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


def _extract_title_block_groundsman(text: str, page=None) -> Optional[dict]:
    """
    Groundsman Industries (SolidWorks sheet format) title block. pdfplumber
    merges each row's label and value onto one line, and the TITLE box text
    lands on the same line as the MATERIAL row:

        DRAWN BY : B Warke DO NOT SCALE DRAWING REVISION: 1
        FRACTIONAL DATE CREATED : 29/07/2014
        TWO PLACE DECIMAL CHECKED: 26/05/2026
        MATERIAL: MSteel Dpth Stop Washer     <- material + TITLE box text
        THICKNESS: 4mm
        COMMENTS: DWG NO.
        LC_TMC26_022W A2
        SCALE:1:1 SHEET 1 OF 1

    The material/title split uses word x-positions (the TITLE: label marks
    the title column's left edge). Returns None if the part number line is
    absent, so the caller falls through to the Claude fallback.
    """
    tb: dict = {}

    # --- Part number + sheet size: value line directly above the SCALE row ---
    m = re.search(
        r"^\s*([A-Z][A-Z0-9_\-]{3,})\s+(A[0-4])\s*\n\s*SCALE\s*:", text, re.M
    )
    if not m:
        return None
    tb["part_number"]  = m.group(1)
    tb["drawing_size"] = m.group(2)

    m = re.search(r"REVISION\s*:\s*([A-Z0-9]+)\s*$", text, re.M)
    if m:
        tb["revision"] = m.group(1)

    m = re.search(r"DRAWN\s+BY\s*:\s*(.+?)\s+DO\s+NOT\s+SCALE", text, re.I)
    if m:
        tb["drawn_by"] = _clean(m.group(1))

    m = re.search(r"DATE\s+CREATED\s*:\s*(\d{2}/\d{2}/\d{4})", text, re.I)
    if m:
        tb["drawn_date"] = m.group(1)

    m = re.search(r"CHECKED\s*:\s*(\d{2}/\d{2}/\d{4})", text, re.I)
    if m:
        tb["checked_date"] = m.group(1)

    m = re.search(r"THICKNESS\s*:\s*([\d.]+)\s*mm", text, re.I)
    if m:
        tb["thickness_mm"] = float(m.group(1))

    # --- Material / description: the MATERIAL row also carries the TITLE
    # box text; split the two columns on the TITLE: label's x-position ---
    m = re.search(r"MATERIAL\s*:\s*(.+)$", text, re.M)
    if m:
        material, description = _clean(m.group(1)), None
        if page is not None:
            words = page.extract_words()
            mat_label = next(
                (w for w in words if w["text"].rstrip(":").upper() == "MATERIAL"), None
            )
            title_label = next(
                (w for w in words if w["text"].rstrip(":").upper() == "TITLE"), None
            )
            if mat_label and title_label:
                row = [
                    w for w in words
                    if abs(w["top"] - mat_label["top"]) < 3 and w["x0"] > mat_label["x1"]
                ]
                mat_words   = [w["text"] for w in row if w["x0"] <  title_label["x0"] - 2]
                title_words = [w["text"] for w in row if w["x0"] >= title_label["x0"] - 2]
                if mat_words:
                    material = _clean(" ".join(mat_words))
                if title_words:
                    description = _clean(" ".join(title_words))
        tb["material"] = material
        if description:
            tb["description"] = description

    m = re.search(r"SCALE\s*:\s*([\d.:]+)", text)
    if m:
        try:
            tb["scale"] = float(m.group(1))
        except ValueError:
            tb["scale"] = m.group(1)

    m = re.search(r"SHEET\s+(\d+)\s+OF\s+(\d+)", text, re.I)
    if m:
        tb["sheet"] = f"{m.group(1)} OF {m.group(2)}"

    # --- Latest revision change: "1 Pointer dimension 24mm to 29mm 200121 bw" ---
    m = re.search(r"^\s*[A-Z0-9]{1,2}\s+(.+?)\s+\d{6}\s+\w{1,4}\s*$", text, re.M)
    if m:
        tb["latest_change_description"] = _clean(m.group(1))

    return tb


def _extract_title_block_generic(text: str, page=None) -> dict:
    """
    Wrightbus/Glovebox title block extraction, with generic label-based
    fallbacks that also cover the Groundsman/SolidWorks sheets. Used for
    every known non-Redrock format and as the last-resort extractor when
    the Claude fallback is unavailable.
    """
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
# Known-format registry
# ---------------------------------------------------------------------------

def _detect_redrock(text: str) -> bool:
    return bool(re.search(r"\bREDROCK\b", text, re.I))


def _detect_wrightbus(text: str) -> bool:
    return bool(re.search(
        r"STATUS\s+PART\s+NUMBER\s+REVISION|make_or_buy|\bWMP-\d", text, re.I
    ))


def _detect_groundsman(text: str) -> bool:
    # Stock SolidWorks sheet format — detect the company name, or the label
    # pair the sheet always carries, so unbranded sheets in the same format
    # still route here.
    if re.search(r"\bGROUNDSMAN\b", text, re.I):
        return True
    return bool(
        re.search(r"DWG\s+NO\.", text, re.I)
        and re.search(r"DO\s+NOT\s+SCALE\s+DRAWING\s+REVISION\s*:", text, re.I)
    )


# (name, detector, title-block extractor) — first matching detector wins.
# Extractors take (text, page); page gives access to word coordinates for
# layouts where pdfplumber merges columns. To support a new customer
# format, add an entry here.
_KNOWN_FORMATS: list[tuple[str, Callable, Callable]] = [
    ("redrock",    _detect_redrock,    _extract_title_block_redrock),
    ("wrightbus",  _detect_wrightbus,  _extract_title_block_generic),
    ("groundsman", _detect_groundsman, _extract_title_block_groundsman),
]

# A regex extraction is good enough to skip the Claude fallback when it found
# the part number plus at least two of these. Anything sparser (e.g. a Redrock
# sheet with a mangled text layer) goes to Claude, because a half-read title
# block prices the part wrong silently.
_GATE_WANTED = ("material", "thickness_mm", "description", "revision")


def _extraction_acceptable(tb: Optional[dict]) -> bool:
    if not tb or "part_number" not in tb:
        return False
    return sum(f in tb for f in _GATE_WANTED) >= 2


# ---------------------------------------------------------------------------
# Pricing factors extractor
# ---------------------------------------------------------------------------

_ANGLE_RE   = re.compile(r"(\d+(?:\.\d+)?)°")
# Bend callouts: Redrock "FOLD UP 15o" / "FOLD DOWN 48°" (degree sign is
# sometimes extracted as a letter "o") and SolidWorks/Groundsman
# "DOWN 40° R 3" / "UP 90° R 2" bend notes
_FOLD_RE    = re.compile(r"(?:FOLD\s+)?\b(?:UP|DOWN)\s+(\d+(?:\.\d+)?)", re.I)
_SLOT_RE    = re.compile(r"(\d+(?:\.\d+)?)\s*MM\s+SLOT", re.I)
_HOLE_RE    = re.compile(r"[ØΦ∅]")
_RADIUS_RE  = re.compile(r"R(\d+(?:\.\d+)?)\s*(?:TYP)?", re.I)

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


def _page_has_circular_holes(page) -> bool:
    """
    Detect hole presence from vector geometry rather than text. SolidWorks/
    Groundsman exports draw the diameter symbol as vector artwork rather
    than an extractable text glyph, so Ø/DIA text matching alone misses
    every hole on those drawings. A real hole is drawn as a full closed
    circular curve; SolidWorks also draws a small (~8-9pt) center-mark dot
    at each hole's centre, excluded here by the width floor.

    Validated against 10 sample Groundsman drawings: every genuine hole
    produced a closed, square-bbox curve of 22-80pt; parts with no holes
    produced none (only the smaller center-mark dots, which this filters
    out). Presence-only — this does not attempt to size or count holes.
    """
    for c in page.curves:
        w, h = c.get("width", 0), c.get("height", 0)
        if w < 15 or h < 15:
            continue  # center-mark dot or other small artifact, not a hole
        if abs(w - h) / max(w, h) > 0.1:
            continue  # not circular
        pts = c.get("pts") or []
        if len(pts) < 8:
            continue
        if abs(pts[0][0] - pts[-1][0]) > 0.5 or abs(pts[0][1] - pts[-1][1]) > 0.5:
            continue  # not a closed path
        return True
    return False


def _extract_pricing_factors(text: str, page) -> dict:
    """
    Extract bend angles, slots, hole presence and process notes from the
    drawing area (title block boilerplate lines filtered out).
    """
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

    # --- Holes (presence only — not priced, exact sizes/dimensions not needed) ---
    # Text-based signal works on formats where the Ø glyph survives extraction
    # (Redrock/Wrightbus); geometry catches formats where it doesn't (Groundsman).
    has_hole_text = bool(_HOLE_RE.search(drawing_text)) or bool(
        re.search(r"(?:DIA\.?|DIAM\.?)\s*[\d.]+", drawing_text, re.I)
    )
    factors["holes_present"] = has_hole_text or _page_has_circular_holes(page)

    # --- Radii ---
    radii = sorted(set(float(r) for r in _RADIUS_RE.findall(drawing_text)))
    if radii:
        factors["radii_mm"] = radii

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

def parse_drawing_pdf(path: str, use_claude: bool = True) -> dict:
    """
    Parse an engineering drawing PDF and extract pricing-relevant data.

    Known formats are parsed with regexes; unrecognised or poorly-extracted
    pages fall back to the Claude API when use_claude is True and
    ANTHROPIC_API_KEY is configured.

    Args:
        path: Absolute or relative path to the PDF file.
        use_claude: Allow the Claude vision fallback for unrecognised pages.

    Returns:
        {
            "source_file": str,
            "pages": [
                {
                    "page": int,
                    "extraction_method": "redrock" | "wrightbus" | "groundsman"
                                         | "claude" | "generic-regex",
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
                        "scale": float | str,
                        "drawing_size": str,
                        "latest_change_description": str,
                        # fields not found are omitted rather than null
                    },
                    "pricing_factors": {
                        "bending_required": bool,
                        "bend_angles_deg": [float, ...],   # unique angles present
                        "bend_count": int,                  # total angle callouts (proxy for op count)
                        "slots": ["7MM SLOT", ...],
                        "holes_present": bool,               # for tube routing awareness; not priced
                        "radii_mm": [float, ...],
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

    claude_ok = use_claude
    with pdfplumber.open(pdf_path) as pdf:
        for i, page in enumerate(pdf.pages):
            text = page.extract_text() or ""

            page_result = None
            method = None

            # 1. Known layouts via regex, gated on extraction quality
            if text.strip():
                for fmt_name, detect, extract in _KNOWN_FORMATS:
                    if detect(text):
                        tb = extract(text, page)
                        if _extraction_acceptable(tb):
                            page_result = {
                                "title_block":     tb,
                                "pricing_factors": _extract_pricing_factors(text, page),
                            }
                            method = fmt_name
                        break  # detector matched; don't try other formats

            # 2. Unrecognised format, sparse extraction, or scanned page → Claude
            if page_result is None and claude_ok:
                try:
                    from claudeExtractor import (
                        ClaudeExtractionError,
                        extract_page_with_claude,
                    )
                except ImportError as e:
                    print(f"readDrawings: Claude fallback unavailable ({e})",
                          file=sys.stderr)
                    claude_ok = False
                else:
                    try:
                        page_result = extract_page_with_claude(pdf_path, i)
                        method = "claude"
                    except ClaudeExtractionError as e:
                        print(f"readDrawings: Claude fallback failed on page "
                              f"{i + 1} of {pdf_path.name}: {e}", file=sys.stderr)
                        if e.permanent:
                            claude_ok = False

            # 3. Last resort: best-effort generic regexes
            if page_result is None:
                if not text.strip():
                    continue  # scanned page and no Claude — nothing to extract
                page_result = {
                    "title_block":     _extract_title_block_generic(text),
                    "pricing_factors": _extract_pricing_factors(text, page),
                }
                method = "generic-regex"

            result["pages"].append({
                "page":              i + 1,
                "extraction_method": method,
                **page_result,
            })

    return result


# ---------------------------------------------------------------------------
# CLI convenience
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    import json

    args = [a for a in sys.argv[1:] if a != "--no-claude"]
    if not args:
        print("Usage: python readDrawings.py <path-to-pdf> [--no-claude]")
        sys.exit(1)

    result = parse_drawing_pdf(args[0], use_claude="--no-claude" not in sys.argv)
    print(json.dumps(result, indent=2))
