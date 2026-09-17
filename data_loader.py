"""
data_loader.py
================
Layer A of the architecture (DATABASE LAYER).

Reads the Media Database workbook (every sheet = a Type of Media, discovered
dynamically — never hardcoded) and normalizes each into a flat list of
`MediaRecord` dicts.

Hard rules enforced here (per project spec):
  - Sheet/tab names are read directly from the workbook, never hardcoded.
  - No value is invented. Anything not present in the source becomes "-".
  - Static media with repeated names across multiple locations get their
    Location appended to the Media name (rule 16).
  - TAX3 values are looked up per-quarter, never computed.
  - Production is used exactly as given, never multiplied by duration.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field, asdict
from typing import Optional

import openpyxl


NO_DATA = "-"

# ---------------------------------------------------------------------------
# Header alias matching (structural mapping logic - not data - allowed to be
# defined in code per the project's own rules; the VALUES always come from
# the workbook, never from here).
# ---------------------------------------------------------------------------

def _norm(s) -> str:
    if s is None:
        return ""
    return re.sub(r"\s+", " ", str(s)).strip().lower()


ALIASES = {
    "media": ["media"],
    "station_or_media": ["station"],           # used as media name only if no "media" col
    "package_or_media": ["package", "pack"],    # used as media name only if no "media"/"station" col
    "location": ["location"],
    "size": ["size"],
    "quantity": ["quantity"],
    "min_loop": ["min / loop", "min/loop", "min / loop "],
    "rate_card": ["rate card"],
    "agency_rate": ["agency rate"],
    "production": ["production"],
    "tax3_q1": ["tax3 q1"],
    "tax3_q2": ["tax3 q2"],
    "tax3_q3": ["tax3 q3"],
    "tax3_q4": ["tax3 q4"],
    "remark": ["remark"],
    "available_pack": ["available pack"],
    "advertising_space": ["advertising space"],
}

# Fields whose value should be carried down (forward-filled) across the
# blank rows that belong to the same static-media group (e.g. one Media /
# PACKAGE spanning several Location / STATION rows). Row-specific fields
# (location, size, quantity) are intentionally excluded.
FILL_FIELDS = [
    "media", "station_or_media", "package_or_media",
    "rate_card", "agency_rate", "production",
    "tax3_q1", "tax3_q2", "tax3_q3", "tax3_q4",
    "remark", "available_pack",
]


@dataclass
class MediaRecord:
    media_id: str
    source_sheet: str
    source_row: int

    type_of_media: str
    media: str               # display name, Location already appended if needed
    media_base: str          # name before Location suffix (for lookups)
    location: str            # "-" if not applicable
    size: str
    quantity: str
    min_loop: str

    rate_card: str
    agency_rate: str
    production: str

    tax3_q1: str
    tax3_q2: str
    tax3_q3: str
    tax3_q4: str

    remark: str
    available_pack: str

    media_proposal_url: str = NO_DATA
    material_guideline_url: str = NO_DATA


def _find_header_row(ws, max_scan: int = 6) -> int:
    """Header is normally row 3, but scan defensively. A real header row has
    several populated cells (many column labels); a single-cell title row
    (e.g. 'BTS Station') must not be mistaken for one just because a keyword
    like 'station' happens to appear in the title text."""
    for r in range(1, max_scan + 1):
        vals = [ws.cell(row=r, column=c).value for c in range(1, ws.max_column + 1)]
        non_empty = [v for v in vals if v not in (None, "")]
        if len(non_empty) < 3:
            continue
        text = " ".join(_norm(v) for v in non_empty)
        if "media" in text or "station" in text or "package" in text or "tier" in text or "quantity" in text:
            return r
    return 3


def _map_columns(ws, header_row: int) -> dict:
    """Return {logical_field: column_index} for a sheet, by alias matching."""
    headers = {}
    for c in range(1, ws.max_column + 1):
        h = _norm(ws.cell(row=header_row, column=c).value)
        if h:
            headers[c] = h

    colmap = {}
    for field_name, aliases in ALIASES.items():
        for c, h in headers.items():
            if any(a in h for a in aliases):
                colmap[field_name] = c
                break

    # Resolve which column is the "media name" column, in priority order.
    if "media" in colmap:
        colmap["_media_name_col"] = colmap["media"]
    elif "station_or_media" in colmap:
        colmap["_media_name_col"] = colmap["station_or_media"]
    elif "package_or_media" in colmap:
        colmap["_media_name_col"] = colmap["package_or_media"]

    # Location: prefer an explicit "location" column; if the media-name
    # column ended up being something else and a "station" column exists
    # separately, use STATION as location.
    if "location" in colmap:
        colmap["_location_col"] = colmap["location"]
    elif "station_or_media" in colmap and colmap.get("_media_name_col") != colmap.get("station_or_media"):
        colmap["_location_col"] = colmap["station_or_media"]
    elif "advertising_space" in colmap:
        colmap["_location_col"] = colmap["advertising_space"]

    # Fallback: some sheets (e.g. BTS Cityline (Light Box)) have no
    # Media/Station/Package column at all — the leftmost labeled column
    # (e.g. a panel-count tier) is the only row identifier available.
    if "_media_name_col" not in colmap and headers:
        colmap["_media_name_col"] = min(headers.keys())
        colmap["_fallback_identifier"] = True

    return colmap


def _cell(ws, row, col) -> Optional[str]:
    if col is None:
        return None
    v = ws.cell(row=row, column=col).value
    if v is None:
        return None
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    return str(v).strip()


def _parse_generic_sheet(ws, sheet_name: str) -> list[MediaRecord]:
    header_row = _find_header_row(ws)
    colmap = _map_columns(ws, header_row)

    records = []
    last_values: dict[str, str] = {}
    idx = 0

    r = header_row + 1
    while r <= ws.max_row:
        row_values = [ws.cell(row=r, column=c).value for c in range(1, ws.max_column + 1)]
        if all(v is None for v in row_values):
            # Blank spacer row *within* a static-media group (common between
            # station/location lines). NOT a group boundary — new groups are
            # detected because they carry their own fresh values. Just skip.
            r += 1
            continue

        raw = {}
        for field_name in ALIASES:
            col = colmap.get(field_name)
            raw[field_name] = _cell(ws, r, col)

        fallback_val = None
        if colmap.get("_fallback_identifier"):
            fallback_val = _cell(ws, r, colmap.get("_media_name_col"))
            if fallback_val and fallback_val.strip().upper() in ("APPENDIX", "NOTE", "NOTES"):
                # A trailing appendix/notes sub-table with a different
                # structure — not part of the priced media list. Stop here.
                break
            if fallback_val is not None:
                last_values["_fallback"] = fallback_val
            elif "_fallback" in last_values:
                fallback_val = last_values["_fallback"]

        location_val = _cell(ws, r, colmap.get("_location_col"))

        # Forward-fill group-level fields
        for f in FILL_FIELDS:
            if raw.get(f) is None and f in last_values:
                raw[f] = last_values[f]
            elif raw.get(f) is not None:
                last_values[f] = raw[f]

        if colmap.get("_fallback_identifier"):
            media_base = fallback_val
        else:
            media_base = None
            for key in ("media", "station_or_media", "package_or_media"):
                if colmap.get("_media_name_col") == colmap.get(key):
                    media_base = raw.get(key)
                    break
        if not media_base:
            r += 1
            continue

        display_name = media_base
        loc_display = location_val if location_val else NO_DATA
        if location_val:
            display_name = f"{media_base} – {location_val}"

        rec = MediaRecord(
            media_id=f"{sheet_name}::{r}",
            source_sheet=sheet_name,
            source_row=r,
            type_of_media=sheet_name,
            media=display_name,
            media_base=media_base,
            location=loc_display,
            size=raw.get("size") or NO_DATA,
            quantity=raw.get("quantity") or NO_DATA,
            min_loop=raw.get("min_loop") or NO_DATA,
            rate_card=raw.get("rate_card") or NO_DATA,
            agency_rate=raw.get("agency_rate") or NO_DATA,
            production=raw.get("production") or NO_DATA,
            tax3_q1=raw.get("tax3_q1") or NO_DATA,
            tax3_q2=raw.get("tax3_q2") or NO_DATA,
            tax3_q3=raw.get("tax3_q3") or NO_DATA,
            tax3_q4=raw.get("tax3_q4") or NO_DATA,
            remark=raw.get("remark") or NO_DATA,
            available_pack=raw.get("available_pack") or NO_DATA,
        )
        records.append(rec)
        idx += 1
        r += 1

    return records


def colmap_field_name(colmap, key, raw):
    """Helper: given colmap['_media_name_col'], find which alias field name
    that column index corresponds to, so we can pull it from `raw`."""
    target_col = colmap.get(key)
    for field_name in ("media", "station_or_media", "package_or_media"):
        if colmap.get(field_name) == target_col:
            return field_name
    return "media"


def _parse_bts_max(ws, sheet_name: str) -> list[MediaRecord]:
    """BTS MAX has a unique layout: an informational spec table at the top,
    then three separate pricing blocks, each titled 'BTS MAX : <Pack Name>'
    in column B, with its own header row + one priced row + several
    unpriced 'included component' rows (per confirmed rule: components of a
    bundled pack are never sold/listed separately).
    """
    records = []
    r = 1
    while r <= ws.max_row:
        title = _cell(ws, r, 2)
        if title and title.strip().upper().startswith("BTS MAX :"):
            pack_name = title.split(":", 1)[1].strip()
            header_row = r + 1
            colmap = _map_columns(ws, header_row)
            price_row = header_row + 1

            rate_card = _cell(ws, price_row, colmap.get("rate_card")) or NO_DATA
            agency_rate = _cell(ws, price_row, colmap.get("agency_rate")) or NO_DATA
            available_pack = _cell(ws, price_row, colmap.get("available_pack")) or NO_DATA

            # Gather included components (this row + following rows until a
            # blank row or the next block title) for the Detail field.
            components = []
            cr = price_row
            while cr <= ws.max_row:
                media_c = _cell(ws, cr, colmap.get("_media_name_col"))
                qty_c = _cell(ws, cr, colmap.get("quantity"))
                loop_c = _cell(ws, cr, colmap.get("min_loop"))
                if media_c is None:
                    break
                next_title = _cell(ws, cr, 2)
                if cr != price_row and next_title and next_title.strip().upper().startswith("BTS MAX :"):
                    break
                parts = [media_c]
                if qty_c:
                    parts.append(qty_c)
                if loop_c:
                    parts.append(loop_c)
                components.append(" / ".join(parts))
                cr += 1

            rec = MediaRecord(
                media_id=f"{sheet_name}::{r}",
                source_sheet=sheet_name,
                source_row=r,
                type_of_media=sheet_name,
                media=pack_name,
                media_base=pack_name,
                location=NO_DATA,
                size=NO_DATA,
                quantity=available_pack,
                min_loop=NO_DATA,
                rate_card=rate_card,
                agency_rate=agency_rate,
                production=NO_DATA,
                tax3_q1=NO_DATA, tax3_q2=NO_DATA, tax3_q3=NO_DATA, tax3_q4=NO_DATA,
                remark="Included: " + "; ".join(components) if components else NO_DATA,
                available_pack=available_pack,
            )
            records.append(rec)
            r = cr
        else:
            r += 1
    return records


SPECIAL_PARSERS = {
    "BTS MAX": _parse_bts_max,
}


def load_database(path: str) -> dict[str, list[MediaRecord]]:
    """Returns {sheet_name: [MediaRecord, ...]} for every sheet in the
    workbook, discovered dynamically."""
    wb = openpyxl.load_workbook(path, data_only=True)
    result = {}
    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        parser = SPECIAL_PARSERS.get(sheet_name.strip(), _parse_generic_sheet)
        result[sheet_name] = parser(ws, sheet_name)
    return result


def flatten(db: dict[str, list[MediaRecord]]) -> list[MediaRecord]:
    out = []
    for recs in db.values():
        out.extend(recs)
    return out


if __name__ == "__main__":
    import sys
    import json

    path = sys.argv[1] if len(sys.argv) > 1 else "Database.xlsx"
    db = load_database(path)
    total = 0
    for sheet, recs in db.items():
        print(f"{sheet:30s} -> {len(recs):4d} records")
        total += len(recs)
    print(f"\nTOTAL: {total} records across {len(db)} sheets")
