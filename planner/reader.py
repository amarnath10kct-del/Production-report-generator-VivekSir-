"""Read and check the master sheet (streaming, read-only) and collect non-blocking warnings.

The reader never changes data. Its row logic mirrors :func:`planner.builder.read_master` exactly
so warnings describe what the builder will do.
"""

from __future__ import annotations

import io
import numbers
import warnings
from dataclasses import dataclass, field
from typing import Any

from openpyxl import load_workbook

from planner.builder import RESERVED_SHEETS
from planner.errors import InvalidUpload, MasterFormatError, UploadTooLarge
from planner.headers import HEADER_SEARCH_ROWS, describe_header_problem, find_header_row
from planner.validate import UploadLimits

MAX_ROW_WARNINGS = 50
#: Rows that exist only as formatting are allowed up to this multiple of ``max_rows``.
HARD_ROW_FACTOR = 5
#: The approved day-tab layout uses rows 1-2 (title, merged instructions) above the header.
MIN_HEADER_ROW = 3


@dataclass
class MasterReport:
    """Facts about the master needed for the summary.

    Attributes:
        sheet_title: First sheet's name.
        header_row: Row holding the headers.
        week_title: Master A1.
        location: Master A2.
        date_line: Master C2.
        data_rows: Non-blank rows below the header.
        planning_rows: Distinct (work centre, item) pairs = rows per day sheet.
        work_centres: Distinct work centres that have at least one item.
        total_qty: Sum of numeric Week Plan Qty over data rows.
        warnings: Non-blocking issues, each naming the master row.

    """

    sheet_title: str = ""
    header_row: int = 0
    week_title: Any = None
    location: Any = None
    date_line: Any = None
    data_rows: int = 0
    planning_rows: int = 0
    work_centres: int = 0
    total_qty: float = 0
    warnings: list[str] = field(default_factory=list)


def _is_number(v: Any) -> bool:
    return isinstance(v, numbers.Number) and not isinstance(v, bool)


def read_master(data: bytes, limits: UploadLimits) -> MasterReport:
    """Stream the first sheet of ``data`` and validate it.

    Raises:
        InvalidUpload: the workbook cannot be parsed.
        UploadTooLarge: more rows/columns than allowed.
        MasterFormatError: headers missing or no usable rows.

    """
    try:
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")
            wb = load_workbook(io.BytesIO(data), read_only=True, data_only=False)
    except Exception as exc:  # openpyxl raises many types for corrupt input
        raise InvalidUpload("The workbook could not be read. It may be damaged.") from exc
    try:
        return _scan(wb, limits)
    finally:
        wb.close()


def _scan(wb: Any, limits: UploadLimits) -> MasterReport:
    if not wb.worksheets:
        raise MasterFormatError("The workbook has no sheets.")
    ws = wb.worksheets[0]
    rep = MasterReport(sheet_title=ws.title)
    if ws.title.lower() in RESERVED_SHEETS:
        raise MasterFormatError(
            f"The first sheet is named '{ws.title}', which is reserved for generated tabs. "
            "Put the master data on the first sheet."
        )
    replaced = [n for n in wb.sheetnames if n.lower() in RESERVED_SHEETS]
    if replaced:
        rep.warnings.append(
            "Existing tabs were replaced by freshly generated ones: " + ", ".join(replaced) + "."
        )

    top: list[tuple[Any, ...]] = []
    hdr: int | None = None
    seen: dict[tuple[Any, Any], int] = {}
    wcs: set[Any] = set()
    row_warnings: list[str] = []
    try:
        for r, row in enumerate(ws.iter_rows(values_only=True), start=1):
            used = [i for i, v in enumerate(row) if v is not None]
            # Formatted-but-empty rows are tolerated up to a hard cap; data rows are not.
            if (used and r > limits.max_rows) or r > HARD_ROW_FACTOR * limits.max_rows:
                raise UploadTooLarge(f"The master sheet has more than {limits.max_rows:,} rows.")
            if used and used[-1] >= limits.max_cols:
                raise UploadTooLarge(f"The master sheet uses more than {limits.max_cols} columns.")
            vals = tuple(row[:3]) + (None,) * (3 - min(len(row), 3))
            if r == 1:
                rep.week_title = vals[0]
            if r == 2:
                rep.location, rep.date_line = vals[0], vals[2]
            if hdr is None:
                top.append(vals)
                if r <= HEADER_SEARCH_ROWS:
                    found = find_header_row(top)
                    if found is not None:
                        hdr = found
                    continue
                raise MasterFormatError(describe_header_problem(top))
            item, wc, qty = vals
            if item is None and wc is None and qty is None:
                continue
            rep.data_rows += 1
            if qty is not None and not _is_number(qty):
                row_warnings.append(
                    f"Row {r}: Week Plan Qty is not a number; it counts as 0 in totals, and the "
                    "item's Weekly plan will show #VALUE! when picked until the master is fixed."
                )
            elif _is_number(qty):
                rep.total_qty += qty
            if wc is None and item is None:
                row_warnings.append(
                    f"Row {r}: quantity without Item Name or Work Center; ignored on day tabs."
                )
                continue
            if wc is None:
                row_warnings.append(
                    f"Row {r}: item has no Work Center; it is excluded from the day tabs "
                    "(still listed in Consolidated)."
                )
                continue
            if item is None:
                row_warnings.append(
                    f"Row {r}: Work Center has no Item Name; it is excluded from the day tabs "
                    "(still listed in Consolidated)."
                )
                continue
            key = (wc, item)
            if key in seen:
                row_warnings.append(
                    f"Row {r}: duplicate of row {seen[key]} (same Work Center and Item); they share "
                    "one row on the day tabs and the Weekly plan is their sum."
                )
            else:
                seen[key] = r
                wcs.add(wc)
    except (UploadTooLarge, MasterFormatError):
        raise
    except Exception as exc:
        raise InvalidUpload("The workbook could not be read. It may be damaged.") from exc

    if hdr is None:
        raise MasterFormatError(describe_header_problem(top))
    if hdr < MIN_HEADER_ROW:
        raise MasterFormatError(
            f"The header row is row {hdr}; it must be on row {MIN_HEADER_ROW} or lower, because rows 1-2 "
            "of the generated day tabs hold the title and instructions. Insert rows above the header "
            "(the week title normally goes in A1, location/date in row 2)."
        )
    rep.header_row = hdr
    rep.planning_rows = len(seen)
    rep.work_centres = len(wcs)
    if not seen:
        raise MasterFormatError(
            "No rows with both an Item Name and a Work Center were found below the header."
        )
    if len(row_warnings) > MAX_ROW_WARNINGS:
        extra = len(row_warnings) - MAX_ROW_WARNINGS
        row_warnings = [*row_warnings[:MAX_ROW_WARNINGS], f"...and {extra} more row warning(s)."]
    rep.warnings.extend(row_warnings)
    return rep
