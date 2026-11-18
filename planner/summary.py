"""The summary shown to the user after a generation."""

from __future__ import annotations

import base64
import datetime as dt
import json
from dataclasses import asdict, dataclass, field
from typing import Any

from planner.reader import MasterReport


def _text(value: Any) -> str:
    """Render a master header cell (A1/A2/C2) as display text."""
    if value is None:
        return ""
    if isinstance(value, dt.datetime | dt.date):
        return value.strftime("%d-%m-%Y")
    return str(value)


@dataclass
class Summary:
    """Result summary (also serialisable as base64 JSON for an HTTP header if ever needed)."""

    week_title: str
    location: str
    date_line: str
    work_centres: int
    items: int
    total_weekly_qty: float
    warnings: list[str] = field(default_factory=list)
    recalculated: bool = False
    timings_ms: dict[str, int] = field(default_factory=dict)

    @classmethod
    def from_report(cls, rep: MasterReport) -> Summary:
        """Build the summary from the reader's report."""
        total = rep.total_qty
        return cls(
            week_title=_text(rep.week_title),
            location=_text(rep.location),
            date_line=_text(rep.date_line),
            work_centres=rep.work_centres,
            items=rep.planning_rows,
            total_weekly_qty=int(total) if float(total).is_integer() else total,
            warnings=list(rep.warnings),
        )

    def to_dict(self) -> dict[str, Any]:
        """Plain dict representation."""
        return asdict(self)

    def to_header(self) -> str:
        """Base64-encoded JSON (the ``X-Planning-Summary`` format from the original spec)."""
        raw = json.dumps(self.to_dict(), ensure_ascii=False).encode("utf-8")
        return base64.b64encode(raw).decode("ascii")
