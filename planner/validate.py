"""Upload checks done before any parsing, and checks on the generated workbook.

Nothing here evaluates uploaded content. The only XML parsed here (``[Content_Types].xml``) goes
through :mod:`defusedxml`; openpyxl itself uses defusedxml automatically when it is installed
(asserted at import time below).
"""

from __future__ import annotations

import io
import zipfile
from dataclasses import dataclass
from typing import Any

import openpyxl.xml
from defusedxml import DefusedXmlException
from defusedxml import ElementTree as defused_et  # noqa: N813

from planner.errors import InvalidUpload, UploadTooLarge

if not openpyxl.xml.DEFUSEDXML:  # pragma: no cover - depends on the environment
    raise RuntimeError("defusedxml must be installed so openpyxl parses XML safely")

_OLE_MAGIC = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"
_ZIP_MAGIC = b"PK\x03\x04"
_MB = 1024 * 1024
#: An entry whose uncompressed/compressed ratio exceeds this is treated as a zip bomb.
MAX_COMPRESSION_RATIO = 200
ERROR_VALUES = frozenset({"#REF!", "#VALUE!", "#NAME?", "#DIV/0!", "#N/A", "#NUM!", "#NULL!"})


@dataclass(frozen=True)
class UploadLimits:
    """Size limits for an uploaded workbook."""

    max_upload_bytes: int = 10 * _MB
    max_uncompressed_bytes: int = 100 * _MB
    max_zip_entries: int = 1000
    max_rows: int = 20_000
    max_cols: int = 50


def check_filename(filename: str) -> None:
    """Reject anything that is not named ``*.xlsx`` (content is checked separately)."""
    lower = (filename or "").lower().strip()
    if lower.endswith((".xlsm", ".xltm", ".xlam")):
        raise InvalidUpload("Macro-enabled workbooks (.xlsm) are not accepted. Save the master as .xlsx.")
    if lower.endswith(".xlsb"):
        raise InvalidUpload("Binary workbooks (.xlsb) are not accepted. Save the master as .xlsx.")
    if lower.endswith(".xls"):
        raise InvalidUpload("Old Excel files (.xls) are not accepted. Save the master as .xlsx.")
    if not lower.endswith(".xlsx"):
        raise InvalidUpload("Only Excel .xlsx files are accepted.")


def check_package(data: bytes, limits: UploadLimits) -> None:
    """Verify ``data`` is a plain, unencrypted, macro-free .xlsx within the size limits.

    Raises:
        InvalidUpload: not an .xlsx, encrypted, macro content, external links, corrupt.
        UploadTooLarge: file, entry count or uncompressed size over the limit.

    """
    if not data:
        raise InvalidUpload("The file is empty.")
    if len(data) > limits.max_upload_bytes:
        raise UploadTooLarge(f"The file is larger than the {limits.max_upload_bytes // _MB} MB limit.")
    if data.startswith(_OLE_MAGIC):
        raise InvalidUpload(
            "This file is password-protected/encrypted or in the old .xls format. "
            "Remove the password and save it as a normal .xlsx, then try again."
        )
    if not data.startswith(_ZIP_MAGIC):
        raise InvalidUpload("This is not a valid .xlsx file (its content is not an Excel workbook).")
    try:
        zf = zipfile.ZipFile(io.BytesIO(data))
    except (zipfile.BadZipFile, OSError, ValueError) as exc:
        raise InvalidUpload("This .xlsx file is damaged and cannot be opened.") from exc
    with zf:
        infos = zf.infolist()
        if len(infos) > limits.max_zip_entries:
            raise UploadTooLarge("The workbook contains too many internal parts.")
        total = 0
        for info in infos:
            total += info.file_size
            if info.flag_bits & 0x1:
                raise InvalidUpload("This file is encrypted. Remove the password and try again.")
            if info.file_size > _MB and info.file_size > MAX_COMPRESSION_RATIO * max(info.compress_size, 1):
                raise UploadTooLarge("The workbook's content is abnormally compressed and was rejected.")
        if total > limits.max_uncompressed_bytes:
            raise UploadTooLarge(
                f"The workbook expands to more than {limits.max_uncompressed_bytes // _MB} MB."
            )
        names = {i.filename for i in infos}
        lowered = {n.lower() for n in names}
        if "[Content_Types].xml" not in names or "xl/workbook.xml" not in names:
            if "xl/workbook.bin" in lowered:
                raise InvalidUpload("Binary workbooks (.xlsb) are not accepted. Save the master as .xlsx.")
            raise InvalidUpload("This is not a valid .xlsx file (workbook parts are missing).")
        if any(n.endswith("vbaproject.bin") or n.endswith("vbaprojectsignature.bin") for n in lowered):
            raise InvalidUpload("The workbook contains macros (VBA). Save it as a macro-free .xlsx.")
        if any(n.startswith("xl/externallinks/") for n in lowered):
            raise InvalidUpload(
                "The workbook has links to other files (external links). "
                "Break the links in Excel (Data > Edit Links) and upload again."
            )
        try:
            ct = defused_et.fromstring(zf.read("[Content_Types].xml"))
        except (DefusedXmlException, defused_et.ParseError, zipfile.BadZipFile, OSError) as exc:
            raise InvalidUpload("This .xlsx file is damaged and cannot be opened.") from exc
        for el in ct:
            ctype = (el.get("ContentType") or "").lower()
            if "macroenabled" in ctype or "vbaproject" in ctype:
                raise InvalidUpload("The workbook contains macros. Save it as a macro-free .xlsx.")
            if "externallink" in ctype:
                raise InvalidUpload(
                    "The workbook has links to other files (external links). "
                    "Break the links in Excel and upload again."
                )
        bad = zf.testzip()
        if bad is not None:
            raise InvalidUpload("This .xlsx file is damaged and cannot be opened.")


@dataclass(frozen=True)
class Expected:
    """What a generated workbook must contain (derived from the build layout)."""

    sheet_order: list[str]
    first: int
    lastd: int
    hdr: int
    last: int


def verify_output(formulas_wb: Any, values_wb: Any, exp: Expected) -> list[str]:
    """Check a (recalculated) workbook kept every template feature and has no formula errors.

    Args:
        formulas_wb: The workbook loaded normally (formulas, rules, layout).
        values_wb: The same workbook loaded with ``data_only=True`` (cached values).
        exp: Expected layout.

    Returns:
        A list of problems (empty when the workbook is good). Problems never contain cell data.

    """
    from planner.builder import DAYS, ENTRIES  # local import keeps module import order simple

    problems: list[str] = []
    if formulas_wb.sheetnames != exp.sheet_order:
        problems.append("sheet order changed")
        return problems
    if formulas_wb["_Lists"].sheet_state != "hidden":
        problems.append("_Lists is not hidden")
    for day in DAYS:
        ws = formulas_wb[day]
        sq = sorted(str(dv.sqref) for dv in ws.data_validations.dataValidation)
        want = sorted([f"B{exp.first}:B{exp.lastd}", f"E{exp.first}:E{exp.lastd}"])
        if sq != want:
            problems.append(f"{day}: data validations missing")
        cf = sorted(str(r.sqref) for r in ws.conditional_formatting)
        if cf != sorted([f"B{exp.first}:B{exp.lastd}", f"E{exp.first}:E{exp.lastd}"]):
            problems.append(f"{day}: conditional formats missing")
        if ws.freeze_panes != f"A{exp.first}":
            problems.append(f"{day}: freeze panes lost")
        if ws.auto_filter.ref != f"A{exp.hdr}:G{exp.lastd}":
            problems.append(f"{day}: autofilter lost")
    cons = formulas_wb["Consolidated"]
    if [str(r.sqref) for r in cons.conditional_formatting] != [f"K{exp.hdr + 1}:K{exp.last}"]:
        problems.append("Consolidated: conditional format missing")
    if cons.freeze_panes != f"A{exp.hdr + 1}":
        problems.append("Consolidated: freeze panes lost")
    if cons.auto_filter.ref != f"A{exp.hdr}:K{exp.last}":
        problems.append("Consolidated: autofilter lost")
    entries = formulas_wb[ENTRIES]
    if entries.freeze_panes != f"A{exp.hdr + 1}":
        problems.append(f"{ENTRIES}: freeze panes lost")
    if entries.auto_filter.ref != f"A{exp.hdr}:D{exp.hdr + len(DAYS) * (exp.lastd - exp.first + 1)}":
        problems.append(f"{ENTRIES}: autofilter lost")
    errors = 0
    for ws in values_wb.worksheets[1:]:
        for row in ws.iter_rows(values_only=True):
            errors += sum(1 for v in row if isinstance(v, str) and v in ERROR_VALUES)
    if errors:
        problems.append(f"{errors} formula error value(s)")
    return problems
