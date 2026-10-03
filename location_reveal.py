"""Reads the data behind a "location reveal" from data.xlsx.

A reveal is identified by a Location ID. Given that ID:
  * the location's name comes from the `Locations` tab (Location ID -> Location Name), and
  * everything else comes from the tab named `Location_<ID>` (e.g. `Location_1`), whose
    columns are read by header name:
        Reveal_Text         text shown when the player enters the area
        Leaving_Text        text shown when the player leaves (loaded, not drawn yet)
        Sub-Locations             one bullet per non-empty cell in the column
        Sub-Location-Desc         description text for the sub-location on that same row
        Mandatory_Sub_Location    TRUE = the sub-location always appears; FALSE (or
                                   blank) = it's a 50/50 chance each time, decided by
                                   pdf_builder.draw_location_reveal when it draws it
        Sub_Location_ID           this sub-location's number, matched against the
                                   Sub_Location column on the Quests tab (see main_quest.py)

Use `get_location_reveal(location_id)` from anywhere; `pdf_builder.draw_location_reveal`
turns the result into page content, and can be called repeatedly for more locations.
"""
import os
from dataclasses import dataclass, field

import openpyxl

DATA_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data.xlsx")

REVEAL_TEXT_HEADER = "Reveal_Text"
LEAVING_TEXT_HEADER = "Leaving_Text"
SUB_LOCATIONS_HEADER = "Sub-Locations"
SUB_LOCATION_DESC_HEADER = "Sub-Location-Desc"
MANDATORY_SUB_LOCATION_HEADER = "Mandatory_Sub_Location"
SUB_LOCATION_ID_HEADER = "Sub_Location_ID"


@dataclass
class SubLocation:
    name: str
    description: str = ""
    mandatory: bool = True
    sub_location_id: int = None


@dataclass
class LocationReveal:
    location_id: int
    name: str
    reveal_text: str = ""
    leaving_text: str = ""
    sub_locations: list = field(default_factory=list)  # list[SubLocation]


def _clean(value):
    return "" if value is None else str(value).strip()


def _to_bool(value, default=True):
    """Excel booleans come back as real True/False; a typed "TRUE"/"FALSE" comes
    back as text. A blank cell keeps `default` (mandatory, if the column is missing
    or a row leaves it blank, so existing sub-locations without this column keep
    always appearing)."""
    if isinstance(value, bool):
        return value
    text = _clean(value).lower()
    if text in ("true", "yes", "1"):
        return True
    if text in ("false", "no", "0"):
        return False
    return default


def _to_int(value):
    """Excel numbers come back as float (e.g. 2.0); a blank or non-numeric cell
    becomes None rather than raising."""
    if value is None or value == "":
        return None
    try:
        return int(float(value))
    except (TypeError, ValueError):
        return None


def get_location_name(location_id, workbook=None):
    """Name of a location, from the `Locations` tab."""
    wb = workbook or openpyxl.load_workbook(DATA_FILE, read_only=True, data_only=True)
    for row in wb["Locations"].iter_rows(min_row=2, values_only=True):
        if row[0] is not None and int(row[0]) == int(location_id):
            return _clean(row[1])
    raise ValueError(f"Location ID {location_id} is not on the Locations tab")


def get_location_reveal(location_id):
    """Everything needed to reveal one location. Raises ValueError with a clear
    message if the location or its `Location_<ID>` tab is missing."""
    wb = openpyxl.load_workbook(DATA_FILE, read_only=True, data_only=True)
    name = get_location_name(location_id, wb)

    tab = f"Location_{int(location_id)}"
    if tab not in wb.sheetnames:
        raise ValueError(f"data.xlsx has no '{tab}' tab for location {location_id} ({name})")

    rows = list(wb[tab].iter_rows(values_only=True))
    headers = [_clean(h) for h in rows[0]] if rows else []

    def column(header):
        if header not in headers:
            return []
        idx = headers.index(header)
        return [_clean(r[idx]) for r in rows[1:] if idx < len(r) and _clean(r[idx])]

    reveal = column(REVEAL_TEXT_HEADER)
    leaving = column(LEAVING_TEXT_HEADER)

    sub_locations = []
    if SUB_LOCATIONS_HEADER in headers:
        name_idx = headers.index(SUB_LOCATIONS_HEADER)
        desc_idx = headers.index(SUB_LOCATION_DESC_HEADER) if SUB_LOCATION_DESC_HEADER in headers else None
        mand_idx = headers.index(MANDATORY_SUB_LOCATION_HEADER) if MANDATORY_SUB_LOCATION_HEADER in headers else None
        id_idx = headers.index(SUB_LOCATION_ID_HEADER) if SUB_LOCATION_ID_HEADER in headers else None
        for r in rows[1:]:
            sub_name = _clean(r[name_idx]) if name_idx < len(r) else ""
            if not sub_name:
                continue
            desc = _clean(r[desc_idx]) if desc_idx is not None and desc_idx < len(r) else ""
            mandatory = _to_bool(r[mand_idx]) if mand_idx is not None and mand_idx < len(r) else True
            sub_id = _to_int(r[id_idx]) if id_idx is not None and id_idx < len(r) else None
            sub_locations.append(SubLocation(sub_name, desc, mandatory, sub_id))

    return LocationReveal(
        location_id=int(location_id),
        name=name,
        reveal_text=reveal[0] if reveal else "",
        leaving_text=leaving[0] if leaving else "",
        sub_locations=sub_locations,
    )
