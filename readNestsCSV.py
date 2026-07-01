"""
nest_csv_parser.py

Parse a headerless nest CSV export and return the same JSON-serialisable
dict schema as schedule_parser.parse_schedule_pdf(), so pricer.py works
with either source without modification.

Expected CSV column order (no header row):
    0: Part Name          e.g. "IBP-62564-R1"
    1: Quantity           integer
    2: Size X (mm)        float
    3: Size Y (mm)        float
    4: Material Name      e.g. "ALU-5251-H22-3MM"
    5: Thickness (mm)     float
    6: Part Process Time  integer (seconds per part)

Parts appearing on multiple rows (same part number across sheets) are
aggregated into a single entry:
    - total_qty                  summed
    - total_weight_g             not available from CSV — omitted
    - total_process_time_seconds process_time_per_part_seconds * total_qty
    - materials                  deduplicated list
    - sheet_rows                 count of rows the part appeared on

Public API:
    parse_nest_csv(path: str) -> dict

CLI:
    python nest_csv_parser.py <path-to-csv>

Dependencies: none (stdlib only)
"""

from __future__ import annotations

import csv
import json
from pathlib import Path
from typing import Optional


# ---------------------------------------------------------------------------
# Column indices — update here if the export format changes
# ---------------------------------------------------------------------------

COL_PART_NUMBER   = 0
COL_QUANTITY      = 1
COL_SIZE_X        = 2
COL_SIZE_Y        = 3
COL_MATERIAL      = 4
COL_THICKNESS     = 5
COL_PROCESS_TIME  = 6

EXPECTED_COLUMNS  = 7


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _seconds_to_hhmmss(seconds: int) -> str:
    h  = seconds // 3600
    mn = (seconds % 3600) // 60
    s  = seconds % 60
    return f"{h:02d}:{mn:02d}:{s:02d}"


def _safe_float(value: str) -> Optional[float]:
    try:
        return float(value.strip())
    except (ValueError, AttributeError):
        return None


def _safe_int(value: str) -> Optional[int]:
    try:
        return int(value.strip())
    except (ValueError, AttributeError):
        return None


# ---------------------------------------------------------------------------
# Public entry point
# ---------------------------------------------------------------------------

def parse_nest_csv(path: str) -> dict:
    """
    Parse a headerless nest CSV and return a structured dict.

    Args:
        path: Path to the CSV file.

    Returns:
        {
            "source_file": str,
            "source_type": "csv",           # distinguishes from PDF output
            "schedule": {
                "total_process_qty": int,
                "sheets": []                # not available from CSV
            },
            "parts": [
                {
                    "part_number":                   str,
                    "size_x_mm":                     float,
                    "size_y_mm":                     float,
                    "material":                      str,
                    "thickness_mm":                  float,
                    "process_time_per_part":         str,    # HH:MM:SS
                    "process_time_per_part_seconds": int,
                    "total_qty":                     int,
                    "total_process_time_seconds":    int,
                    "total_process_time":            str,    # HH:MM:SS
                    "sheet_rows":                    int,    # rows in CSV (proxy for sheets)
                    "materials":                     [str],  # deduplicated
                    # Fields not available from CSV (present in PDF output) are omitted:
                    #   weight_per_part_g, total_weight_g, bending, comment,
                    #   order_qty, sheets (named sheet list)
                },
                ...
            ],
            "_warnings": [str]
        }

    Raises:
        FileNotFoundError: If the path does not exist.
        ValueError: If the file is empty or has no valid rows.
    """
    csv_path = Path(path)
    if not csv_path.exists():
        raise FileNotFoundError(f"CSV not found: {path}")

    warnings: list[str] = []
    aggregated: dict[str, dict] = {}
    skipped = 0

    with open(csv_path, newline="", encoding="utf-8-sig") as fh:
        # utf-8-sig strips BOM if present; handles Windows \r\n automatically
        reader = csv.reader(fh)
        for line_num, row in enumerate(reader, start=1):
            # Skip blank lines
            if not any(cell.strip() for cell in row):
                continue

            if len(row) < EXPECTED_COLUMNS:
                warnings.append(
                    f"Line {line_num}: expected {EXPECTED_COLUMNS} columns, "
                    f"got {len(row)} — skipped."
                )
                skipped += 1
                continue

            part_number  = row[COL_PART_NUMBER].strip()
            qty          = _safe_int(row[COL_QUANTITY])
            size_x       = _safe_float(row[COL_SIZE_X])
            size_y       = _safe_float(row[COL_SIZE_Y])
            material     = row[COL_MATERIAL].strip()
            thickness    = _safe_float(row[COL_THICKNESS])
            process_secs = _safe_int(row[COL_PROCESS_TIME])

            if not part_number:
                warnings.append(f"Line {line_num}: empty part number — skipped.")
                skipped += 1
                continue

            if qty is None or qty <= 0:
                warnings.append(f"Line {line_num}: invalid quantity '{row[COL_QUANTITY]}' — skipped.")
                skipped += 1
                continue

            # Aggregate by part number
            if part_number not in aggregated:
                aggregated[part_number] = {
                    "part_number":                   part_number,
                    "size_x_mm":                     size_x,
                    "size_y_mm":                     size_y,
                    "material":                      material,
                    "thickness_mm":                  thickness,
                    "process_time_per_part_seconds": process_secs,
                    "process_time_per_part":         _seconds_to_hhmmss(process_secs) if process_secs else None,
                    "total_qty":                     0,
                    "total_process_time_seconds":    0,
                    "sheet_rows":                    0,
                    "materials":                     [],
                    "_inconsistencies":              [],
                }

            agg = aggregated[part_number]
            agg["total_qty"]  += qty
            agg["sheet_rows"] += 1

            if process_secs is not None:
                agg["total_process_time_seconds"] += process_secs * qty

            if material and material not in agg["materials"]:
                agg["materials"].append(material)

            # Flag property inconsistencies across rows
            for field, new_val in (
                ("size_x_mm",    size_x),
                ("size_y_mm",    size_y),
                ("thickness_mm", thickness),
            ):
                if new_val is not None and agg[field] != new_val:
                    note = f"{field} differs on line {line_num}: {new_val} vs {agg[field]}"
                    if note not in agg["_inconsistencies"]:
                        agg["_inconsistencies"].append(note)

    if not aggregated and skipped == 0:
        raise ValueError(f"No valid rows found in {path}")

    if skipped:
        warnings.append(f"{skipped} row(s) skipped due to formatting issues.")

    # Finalise
    parts = []
    for agg in aggregated.values():
        agg["total_process_time"] = _seconds_to_hhmmss(agg["total_process_time_seconds"])
        if not agg["_inconsistencies"]:
            del agg["_inconsistencies"]
        parts.append(agg)

    total_qty = sum(p["total_qty"] for p in parts)

    return {
        "source_file": csv_path.name,
        "source_type": "csv",
        "schedule": {
            "total_process_qty": total_qty,
            "sheets":            [],  # not available from CSV
        },
        "parts":     parts,
        "_warnings": warnings,
    }


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    import sys

    if len(sys.argv) < 2:
        print("Usage: python nest_csv_parser.py <path-to-csv>")
        sys.exit(1)

    result = parse_nest_csv(sys.argv[1])
    print(json.dumps(result, indent=2))