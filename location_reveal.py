"""Reads the data behind a "location reveal" from data.xlsx.

A reveal is identified by a Location ID. Given that ID:
  * the location's name comes from the `Locations` tab (Location ID -> Location Name), and
  * everything else comes from the tab named `Location_<ID>` (e.g. `Location_1`), whose
    columns are read by header name:
        Reveal_Text    text shown when the player enters the area
        Leaving_Text   text shown when the player leaves (loaded, not drawn yet)
        Sub-Locations  one bullet per non-empty cell in the column

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


@dataclass
class LocationReveal:
    location_id: int
    name: str
    reveal_text: str = ""
    leaving_text: str = ""
    sub_locations: list = field(default_factory=list)


def _clean(value):
    return "" if value is None else str(value).strip()


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
    return LocationReveal(
        location_id=int(location_id),
        name=name,
        reveal_text=reveal[0] if reveal else "",
        leaving_text=leaving[0] if leaving else "",
        sub_locations=column(SUB_LOCATIONS_HEADER),
    )
