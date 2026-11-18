"""One call that turns an uploaded master into the planning workbook: validate, build, recalc."""

from __future__ import annotations

import io
import logging
import re
import time
import unicodedata
import warnings
from dataclasses import dataclass, field

from openpyxl import load_workbook

from planner import builder
from planner.reader import read_master
from planner.recalc import RecalcError, RecalcOptions, recalculate
from planner.summary import Summary
from planner.validate import Expected, UploadLimits, check_filename, check_package, verify_output

log = logging.getLogger(__name__)

XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
FALLBACK_WARNING = (
    "The workbook could not be pre-calculated on the server ({reason}). It is still complete: "
    "Excel calculates all values when you open it, but previews (phone, Outlook, web viewers) "
    "may show empty cells until it is opened in Excel."
)


@dataclass(frozen=True)
class GenerateOptions:
    """Everything :func:`generate` needs; built from the app config."""

    limits: UploadLimits = field(default_factory=UploadLimits)
    recalc_enabled: bool = True
    recalc: RecalcOptions = field(default_factory=RecalcOptions)


@dataclass
class GenerateResult:
    """Output of :func:`generate`."""

    content: bytes
    filename: str
    summary: Summary
    mimetype: str = XLSX_MIME


_SAFE = re.compile(r"[^A-Za-z0-9._ ()-]+")


def output_filename(original: str) -> str:
    """Derive ``<stem>_Planning.xlsx`` from the uploaded name, safely.

    Path parts, control characters and anything outside ``A-Z a-z 0-9 . _ - ( ) space`` are
    replaced with ``_``.
    """
    name = (original or "").replace("\\", "/").rsplit("/", 1)[-1]
    stem = name[:-5] if name.lower().endswith(".xlsx") else name.rsplit(".", 1)[0]
    stem = unicodedata.normalize("NFKD", stem).encode("ascii", "ignore").decode("ascii")
    stem = _SAFE.sub("_", stem).strip(" ._-")
    stem = re.sub(r"_+", "_", stem)[:100].strip(" ._-")
    return f"{stem or 'Master'}_Planning.xlsx"


def _ms(t0: float) -> int:
    return int((time.perf_counter() - t0) * 1000)


def generate(data: bytes, filename: str, opts: GenerateOptions | None = None) -> GenerateResult:
    """Validate ``data``, build the planning workbook and (if possible) recalculate it.

    Raises:
        PlannerError: for invalid uploads (message is safe to show). Recalculation problems never
            raise; they fall back to an Excel-calculates-on-open workbook with a warning.

    """
    opts = opts or GenerateOptions()
    timings: dict[str, int] = {}
    t0 = time.perf_counter()
    check_filename(filename)
    check_package(data, opts.limits)
    report = read_master(data, opts.limits)
    timings["validate"] = _ms(t0)

    t1 = time.perf_counter()
    out = io.BytesIO()
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        info = builder.build(io.BytesIO(data), out)
    built = out.getvalue()
    timings["build"] = _ms(t1)

    summary = Summary.from_report(report)
    expected = Expected(
        sheet_order=[info.master_title, *builder.DAYS, "Consolidated", builder.ENTRIES, "_Lists"],
        first=info.first,
        lastd=info.lastd,
        hdr=info.hdr,
        last=info.last,
    )
    content = built
    if opts.recalc_enabled:
        t2 = time.perf_counter()
        try:
            recalced = recalculate(built, opts.recalc)
            problems = _verify(recalced, expected)
            if problems:
                log.warning("recalc.verify_failed", extra={"problems": len(problems)})
                raise RecalcError("verification failed")
            content = recalced
            summary.recalculated = True
        except RecalcError as exc:
            log.warning("recalc.fallback", extra={"reason": str(exc)})
            summary.warnings.append(FALLBACK_WARNING.format(reason=str(exc)))
        timings["recalc"] = _ms(t2)
    else:
        summary.warnings.append(FALLBACK_WARNING.format(reason="recalculation is turned off"))
    timings["total"] = _ms(t0)
    summary.timings_ms = timings
    return GenerateResult(content=content, filename=output_filename(filename), summary=summary)


def _verify(xlsx: bytes, expected: Expected) -> list[str]:
    try:
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")
            fwb = load_workbook(io.BytesIO(xlsx))
            vwb = load_workbook(io.BytesIO(xlsx), data_only=True)
    except Exception as exc:  # corrupted LibreOffice output
        raise RecalcError("recalculated file unreadable") from exc
    return verify_output(fwb, vwb, expected)
