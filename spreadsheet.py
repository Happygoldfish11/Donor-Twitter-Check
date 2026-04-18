"""
Spreadsheet I/O
---------------
Reads input spreadsheets (xlsx / csv) and writes results back
to a colour-coded output xlsx.

Expected input columns (case-insensitive, partial match):
  first name, last name, address, employer

Output adds:
  Twitter Handle | Profile URL | Confidence | Score | Score Breakdown | Bio | Location
"""

import io
import re
import pandas as pd
from openpyxl import load_workbook, Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from typing import IO, Union

from src.matcher import Person, MatchResult


# ---------------------------------------------------------------------------
# Colour palette (openpyxl uses ARGB hex without #)
# ---------------------------------------------------------------------------
FILLS = {
    "HIGH":     PatternFill("solid", start_color="DCFCE7"),   # green-100
    "MEDIUM":   PatternFill("solid", start_color="FEF9C3"),   # yellow-100
    "LOW":      PatternFill("solid", start_color="FFEDD5"),   # orange-100
    "NO MATCH": PatternFill("solid", start_color="FEE2E2"),   # red-100
    "HEADER":   PatternFill("solid", start_color="1E293B"),   # slate-800
}

CONFIDENCE_LABELS = {
    "HIGH":     "✓ HIGH",
    "MEDIUM":   "~ MEDIUM",
    "LOW":      "! LOW",
    "NO MATCH": "✗ NO MATCH",
}

thin = Side(style="thin", color="E2E8F0")
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)


# ---------------------------------------------------------------------------
# Column detection
# ---------------------------------------------------------------------------

def _find_col(columns: list[str], keywords: list[str]) -> str | None:
    for col in columns:
        col_l = col.lower()
        if any(kw in col_l for kw in keywords):
            return col
    return None


def load_persons(file: Union[str, IO]) -> tuple[pd.DataFrame, list[Person], list[str]]:
    """
    Load a spreadsheet and return:
      - original dataframe
      - list of Person objects
      - list of warnings for the UI
    """
    warnings = []

    if hasattr(file, "name") and file.name.endswith(".csv"):
        df = pd.read_csv(file)
    else:
        df = pd.read_excel(file)

    cols = list(df.columns)

    first_col   = _find_col(cols, ["first"])
    last_col    = _find_col(cols, ["last"])
    address_col = _find_col(cols, ["address", "addr", "street", "location"])
    employer_col = _find_col(cols, ["employer", "company", "organization", "org", "work"])

    if not first_col:
        raise ValueError("Could not find a 'First Name' column. Please ensure your spreadsheet has this column.")
    if not last_col:
        raise ValueError("Could not find a 'Last Name' column. Please ensure your spreadsheet has this column.")

    if not address_col:
        warnings.append("No address column detected — city/state matching disabled.")
    if not employer_col:
        warnings.append("No employer column detected — employer matching disabled.")

    persons = []
    for _, row in df.iterrows():
        persons.append(Person(
            first_name=str(row.get(first_col, "") or "").strip(),
            last_name=str(row.get(last_col, "") or "").strip(),
            address=str(row.get(address_col, "") or "").strip() if address_col else "",
            employer=str(row.get(employer_col, "") or "").strip() if employer_col else "",
        ))

    return df, persons, warnings


# ---------------------------------------------------------------------------
# Output writer
# ---------------------------------------------------------------------------

def write_results(original_file: Union[str, IO], results: list[MatchResult]) -> io.BytesIO:
    """
    Write results back into a copy of the original spreadsheet.
    Adds colour-coded result columns to the right.
    Returns a BytesIO buffer ready for download.
    """
    # Re-load original
    if hasattr(original_file, "seek"):
        original_file.seek(0)

    try:
        if hasattr(original_file, "name") and original_file.name.endswith(".csv"):
            df = pd.read_csv(original_file)
            wb = Workbook()
            ws = wb.active
            ws.title = "Results"
            ws.append(list(df.columns))
            for row in df.itertuples(index=False):
                ws.append(list(row))
        else:
            wb = load_workbook(original_file)
            ws = wb.active
    except Exception:
        wb = Workbook()
        ws = wb.active
        ws.title = "Results"

    orig_cols = ws.max_column

    # ── New column headers ───────────────────────────────────────────────────
    new_headers = [
        "Twitter Handle",
        "Profile URL",
        "Confidence",
        "Score (0–100)",
        "Display Name",
        "Bio",
        "Location",
        "Score Breakdown",
    ]

    header_font = Font(color="FFFFFF", bold=True, name="Calibri", size=10)
    for i, h in enumerate(new_headers, start=orig_cols + 1):
        cell = ws.cell(1, i, h)
        cell.fill = FILLS["HEADER"]
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = BORDER

    # Also style existing header row
    for c in range(1, orig_cols + 1):
        cell = ws.cell(1, c)
        cell.fill = FILLS["HEADER"]
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = BORDER

    ws.row_dimensions[1].height = 28

    # ── Data rows ────────────────────────────────────────────────────────────
    for row_idx, result in enumerate(results, start=2):
        fill = FILLS.get(result.confidence, FILLS["NO MATCH"])
        data_font = Font(name="Calibri", size=10)

        # Colour existing cells in this row
        for c in range(1, orig_cols + 1):
            cell = ws.cell(row_idx, c)
            cell.fill = fill
            cell.font = data_font
            cell.border = BORDER
            cell.alignment = Alignment(vertical="center")

        # Write new columns
        p = result.profile
        breakdown_str = " | ".join(f"{k}: +{v}" for k, v in result.score_breakdown.items())

        new_values = [
            p.handle if p else "",
            p.profile_url if p else "",
            CONFIDENCE_LABELS.get(result.confidence, result.confidence),
            result.score,
            p.display_name if p else "",
            (p.bio[:200] + "…") if p and len(p.bio) > 200 else (p.bio if p else ""),
            p.location if p else "",
            breakdown_str,
        ]

        for i, val in enumerate(new_values, start=orig_cols + 1):
            cell = ws.cell(row_idx, i, val)
            cell.fill = fill
            cell.font = data_font
            cell.border = BORDER
            cell.alignment = Alignment(vertical="center", wrap_text=(i > orig_cols + 2))

        ws.row_dimensions[row_idx].height = 18

    # ── Column widths ────────────────────────────────────────────────────────
    col_widths = {
        orig_cols + 1: 18,   # handle
        orig_cols + 2: 35,   # url
        orig_cols + 3: 14,   # confidence
        orig_cols + 4: 12,   # score
        orig_cols + 5: 20,   # display name
        orig_cols + 6: 45,   # bio
        orig_cols + 7: 20,   # location
        orig_cols + 8: 55,   # breakdown
    }
    for col_num, width in col_widths.items():
        ws.column_dimensions[get_column_letter(col_num)].width = width

    # ── Summary sheet ────────────────────────────────────────────────────────
    if "Summary" in wb.sheetnames:
        del wb["Summary"]
    ws_sum = wb.create_sheet("Summary")

    counts = {"HIGH": 0, "MEDIUM": 0, "LOW": 0, "NO MATCH": 0}
    for r in results:
        counts[r.confidence] = counts.get(r.confidence, 0) + 1

    ws_sum.column_dimensions["A"].width = 20
    ws_sum.column_dimensions["B"].width = 12
    ws_sum.column_dimensions["C"].width = 12

    ws_sum.append(["Confidence Level", "Count", "% of Total"])
    for c in range(1, 4):
        cell = ws_sum.cell(1, c)
        cell.fill = FILLS["HEADER"]
        cell.font = Font(color="FFFFFF", bold=True, name="Calibri", size=10)
        cell.alignment = Alignment(horizontal="center")
        cell.border = BORDER

    total = len(results)
    for level in ["HIGH", "MEDIUM", "LOW", "NO MATCH"]:
        n = counts[level]
        pct = f"{100 * n / total:.1f}%" if total else "0%"
        ws_sum.append([CONFIDENCE_LABELS[level], n, pct])
        for c in range(1, 4):
            cell = ws_sum.cell(ws_sum.max_row, c)
            cell.fill = FILLS[level]
            cell.font = Font(name="Calibri", size=10)
            cell.border = BORDER
            cell.alignment = Alignment(horizontal="center")

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf
