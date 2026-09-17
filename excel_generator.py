"""
excel_generator.py
===================
Layer D of the architecture (EXCEL GENERATION ENGINE).

Opens the Changan Excel file as the master template, never builds a new
layout from scratch. Preserves fonts, fills, borders, number formats,
column widths and the merged-cell structure while inserting/removing
item rows to match how many media lines Sales selected, then fills in
values and formulas.

Template layout (fixed, read from the master file — not re-invented):
    Row 3            : "Product: <client>"
    Row 4            : "Date : <date>"
    Row 7-9          : header (3 rows, some cells merged)
    Row 10 .. 10+n-1 : one row per selected media line
    Row 10+n         : "Total" row (SUM formulas)
    Row 10+n+1 ..     : Remarks / service-condition footer (verbatim from template)
"""

from __future__ import annotations

import copy
import datetime as dt
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

from pricing_engine import LineItem
from data_loader import NO_DATA

SHEET_NAME = "Media Package"
TEMPLATE_FIRST_ITEM_ROW = 10
TEMPLATE_SAMPLE_ITEM_COUNT = 3          # how many sample rows ship in the template
TEMPLATE_TOTAL_ROW = 13                  # in the untouched template
TEMPLATE_LAST_FOOTER_ROW = 29

# Column letters, from the confirmed Field Mapping Table
COL = {
    "type_of_media": "B",
    "media": "C",
    "quantity": "D",
    "detail": "E",
    "period": "F",
    "duration": "G",
    "normal_rate": "H",
    "agency_rate_secondary": "I",   # mirrors J, per confirmed decision
    "agency_rate": "J",
    "total_media_cost": "K",
    "production": "L",
    "total_media_investment": "M",
    "remark": "N",
    "tax3_q1": "P",
    "tax3_q2": "Q",
    "tax3_q3": "R",
    "tax3_q4": "S",
    "media_proposal": "U",
    "material_guideline": "V",
}
NUMERIC_FIELDS = ["normal_rate", "agency_rate_secondary", "agency_rate",
                   "total_media_cost", "production", "total_media_investment",
                   "tax3_q1", "tax3_q2", "tax3_q3", "tax3_q4"]


def _clone_row_style(ws: Worksheet, src_row: int, dst_row: int, min_col=2, max_col=22):
    for c in range(min_col, max_col + 1):
        src = ws.cell(row=src_row, column=c)
        dst = ws.cell(row=dst_row, column=c)
        dst.font = copy.copy(src.font)
        dst.border = copy.copy(src.border)
        dst.fill = copy.copy(src.fill)
        dst.number_format = src.number_format
        dst.alignment = copy.copy(src.alignment)
    ws.row_dimensions[dst_row].height = ws.row_dimensions[src_row].height


def _resize_item_block(ws: Worksheet, n_items: int) -> int:
    """Grow or shrink the item block so there are exactly n_items rows
    between the header and the Total row. Returns the new Total row number.

    openpyxl's insert_rows/delete_rows do NOT shift merged-cell ranges, so
    every merge is captured and torn down first, the row operation is
    performed on a merge-free sheet (safe), and then every merge is
    re-applied at its correct (possibly shifted) location.
    """
    delta = n_items - TEMPLATE_SAMPLE_ITEM_COUNT
    if delta == 0:
        return TEMPLATE_FIRST_ITEM_ROW + n_items

    above_row = (TEMPLATE_FIRST_ITEM_ROW + TEMPLATE_SAMPLE_ITEM_COUNT - 1) if delta > 0 \
        else (TEMPLATE_FIRST_ITEM_ROW + n_items - 1)

    saved_ranges = [(mr.min_row, mr.min_col, mr.max_row, mr.max_col) for mr in list(ws.merged_cells.ranges)]
    for (r1, c1, r2, c2) in saved_ranges:
        ws.unmerge_cells(start_row=r1, start_column=c1, end_row=r2, end_column=c2)

    if delta > 0:
        insert_at = TEMPLATE_FIRST_ITEM_ROW + TEMPLATE_SAMPLE_ITEM_COUNT  # before old Total row
        ws.insert_rows(insert_at, delta)
        style_src = TEMPLATE_FIRST_ITEM_ROW + TEMPLATE_SAMPLE_ITEM_COUNT - 1  # last sample row
        for i in range(delta):
            _clone_row_style(ws, style_src, insert_at + i)
    else:
        remove_from = TEMPLATE_FIRST_ITEM_ROW + n_items
        ws.delete_rows(remove_from, -delta)

    for (r1, c1, r2, c2) in saved_ranges:
        nr1 = r1 + delta if r1 > above_row else r1
        nr2 = r2 + delta if r2 > above_row else r2
        ws.merge_cells(start_row=nr1, start_column=c1, end_row=nr2, end_column=c2)

    total_row = TEMPLATE_FIRST_ITEM_ROW + n_items
    return total_row


def _write_cell(ws, row, col_letter, value):
    ws[f"{col_letter}{row}"] = value


def generate_media_package_excel(
    template_path: str,
    output_path: str,
    line_items: list[LineItem],
    client_name: str,
    package_date: dt.date | None = None,
):
    if not line_items:
        raise ValueError("At least one media line item is required.")

    wb = load_workbook(template_path)
    ws = wb[SHEET_NAME]

    package_date = package_date or dt.date.today()
    ws["C3"] = f"Product: {client_name}"
    ws["C4"] = f"Date : {package_date.strftime('%d %b %Y')}"

    n = len(line_items)
    total_row = _resize_item_block(ws, n)

    for i, li in enumerate(line_items):
        row = TEMPLATE_FIRST_ITEM_ROW + i
        d = li.to_row_dict()

        _write_cell(ws, row, COL["type_of_media"], d["type_of_media"])
        _write_cell(ws, row, COL["media"], d["media"])
        _write_cell(ws, row, COL["quantity"], d["quantity"])
        _write_cell(ws, row, COL["detail"], d["detail"])
        _write_cell(ws, row, COL["period"], d["period"])
        _write_cell(ws, row, COL["duration"], d["duration"])

        _write_cell(ws, row, COL["normal_rate"], d["normal_rate"])
        _write_cell(ws, row, COL["agency_rate_secondary"], d["agency_rate_secondary"])
        _write_cell(ws, row, COL["agency_rate"], d["agency_rate"])

        if d["agency_rate"] != NO_DATA:
            ws[f"{COL['total_media_cost']}{row}"] = f"={COL['agency_rate']}{row}*{COL['duration']}{row}"
        else:
            _write_cell(ws, row, COL["total_media_cost"], NO_DATA)

        _write_cell(ws, row, COL["production"], d["production"])

        if d["agency_rate"] != NO_DATA:
            ws[f"{COL['total_media_investment']}{row}"] = (
                f"={COL['total_media_cost']}{row}+{COL['production']}{row}"
                if d["production"] != NO_DATA
                else f"={COL['total_media_cost']}{row}"
            )
        else:
            _write_cell(ws, row, COL["total_media_investment"], NO_DATA)

        _write_cell(ws, row, COL["remark"], d["remark"] if d["remark"] != NO_DATA else "")

        for q in ["tax3_q1", "tax3_q2", "tax3_q3", "tax3_q4"]:
            val = d[q]
            _write_cell(ws, row, COL[q], val)

        # Media Proposal / Material Guideline hyperlinks
        for field, col_letter in [("media_proposal_url", COL["media_proposal"]),
                                   ("material_guideline_url", COL["material_guideline"])]:
            url = getattr(li.record, field)
            cell = ws[f"{col_letter}{row}"]
            if url and url != NO_DATA:
                cell.value = "Link"
                cell.hyperlink = url
                cell.style = "Hyperlink"
            else:
                cell.value = NO_DATA

    # Total row: label + SUM formulas over the actual item range
    first, last = TEMPLATE_FIRST_ITEM_ROW, TEMPLATE_FIRST_ITEM_ROW + n - 1
    ws[f"{COL['type_of_media']}{total_row}"] = "Total"
    for field in NUMERIC_FIELDS + ["total_media_cost", "total_media_investment"]:
        col_letter = COL[field]
        ws[f"{col_letter}{total_row}"] = f"=SUM({col_letter}{first}:{col_letter}{last})"
    # The original Changan template has a stray '=SUM(#REF!)' in the Total
    # row's Remark cell (a leftover broken reference from their own file,
    # confirmed against the source). Remarks are text, not something to
    # sum — clear it instead of carrying the broken formula forward.
    ws[f"{COL['remark']}{total_row}"] = ""

    wb.save(output_path)
    return output_path


if __name__ == "__main__":
    # Smoke test using real database rows.
    from data_loader import load_database
    db = load_database("Database.xlsx")

    items = [
        LineItem(record=db["Horizon"][0], period="1 Oct - 31 Dec'26", duration_months=3),
        LineItem(record=db["BTS MAX"][0], period="1 Oct - 31 Dec'26", duration_months=3),
        LineItem(record=db["Walkway Balustrade"][0], period="1 Oct - 31 Dec'26", duration_months=6),
        LineItem(record=db["BTS Station (Ground)"][0], period="1 Oct - 31 Dec'26", duration_months=3),
        LineItem(record=db["Move"][0], period="1 Oct - 31 Dec'26", duration_months=2),
    ]
    generate_media_package_excel(
        "ต_วอย_าง_Media_Package_Changan_27_07_2026.xlsx",
        "test_output_5items.xlsx",
        items,
        client_name="Test Client",
    )
    print("Generated test_output_5items.xlsx")

    items2 = items[:2]
    generate_media_package_excel(
        "ต_วอย_าง_Media_Package_Changan_27_07_2026.xlsx",
        "test_output_2items.xlsx",
        items2,
        client_name="Test Client (2 items)",
    )
    print("Generated test_output_2items.xlsx")
