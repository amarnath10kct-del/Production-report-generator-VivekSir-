"""Header-row detection for the master sheet (shared by the reader and the builder).

Matching ignores case and runs of whitespace (including non-breaking spaces) **for detection
only**; master data is never altered.
"""

from __future__ import annotations

import re
from collections.abc import Iterable, Sequence
from typing import Any

#: Expected master headers, in the column order the approved template relies on (A, B, C).
EXPECTED_HEADERS: tuple[str, str, str] = ("Item Name", "Work Center", "Week Plan Qty")
#: The header must appear within this many rows from the top of the sheet.
HEADER_SEARCH_ROWS = 10

_WS = re.compile(r"\s+", re.UNICODE)


def normalise_header(value: Any) -> str:
    """Return a comparison key for a header cell: whitespace collapsed, case-folded."""
    if value is None:
        return ""
    return _WS.sub(" ", str(value)).strip().casefold()


_KEYS = tuple(normalise_header(h) for h in EXPECTED_HEADERS)


def find_header_row(rows: Iterable[Sequence[Any]]) -> int | None:
    """Return the 1-based row whose first three cells are the expected headers, else ``None``.

    Args:
        rows: Row value tuples starting at sheet row 1. Only the first
            :data:`HEADER_SEARCH_ROWS` rows are inspected.

    """
    for idx, row in enumerate(rows, start=1):
        if idx > HEADER_SEARCH_ROWS:
            break
        vals = tuple(normalise_header(row[c] if c < len(row) else None) for c in range(3))
        if vals == _KEYS:
            return idx
    return None


def describe_header_problem(rows: Sequence[Sequence[Any]]) -> str:
    """Explain why no header row was found: which headers are missing or misplaced.

    Args:
        rows: Row value tuples for the first :data:`HEADER_SEARCH_ROWS` rows.

    """
    best_row: int | None = None
    best_found: dict[str, int] = {}
    for idx, row in enumerate(rows[:HEADER_SEARCH_ROWS], start=1):
        found = {}
        for col, value in enumerate(row):
            key = normalise_header(value)
            for want, wkey in zip(EXPECTED_HEADERS, _KEYS, strict=True):
                if key == wkey and want not in found:
                    found[want] = col
        if len(found) > len(best_found):
            best_row, best_found = idx, found
    if not best_found:
        return (
            "Could not find the header row. The first sheet must have the headers "
            f"'{EXPECTED_HEADERS[0]}', '{EXPECTED_HEADERS[1]}' and '{EXPECTED_HEADERS[2]}' "
            f"in columns A, B and C within the first {HEADER_SEARCH_ROWS} rows."
        )
    missing = [h for h in EXPECTED_HEADERS if h not in best_found]
    if missing:
        return (
            f"Header row {best_row} is missing: "
            + ", ".join(f"'{m}'" for m in missing)
            + ". Expected 'Item Name' (A), 'Work Center' (B), 'Week Plan Qty' (C)."
        )
    return (
        f"Headers found on row {best_row} but not in the expected columns. "
        "They must be 'Item Name' in column A, 'Work Center' in B and 'Week Plan Qty' in C."
    )
