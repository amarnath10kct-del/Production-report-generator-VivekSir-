"""Structured (JSON-lines) logging to stdout and, optionally, a rotating log file.

Log records must never contain uploaded content: callers pass only request ids, outcomes,
durations, counts, sizes and error class names. The formatter emits only whitelisted extras.
"""

from __future__ import annotations

import json
import logging
import logging.handlers
import sys
import time
from typing import Any

#: Extra fields allowed into log lines. Anything else passed via ``extra=`` is dropped.
SAFE_FIELDS = (
    "rid",
    "outcome",
    "error_class",
    "status",
    "size_bytes",
    "out_bytes",
    "work_centres",
    "items",
    "warnings",
    "recalculated",
    "ms_total",
    "ms_validate",
    "ms_build",
    "ms_recalc",
    "reason",
    "problems",
    "soffice",
)

_HANDLER_TAG = "_mdk_planner_handler"


class JsonFormatter(logging.Formatter):
    """Format records as one JSON object per line."""

    def format(self, record: logging.LogRecord) -> str:
        """Serialise ``record`` with only whitelisted extra fields."""
        out: dict[str, Any] = {
            "ts": time.strftime("%Y-%m-%dT%H:%M:%S", time.gmtime(record.created))
            + f".{int(record.msecs):03d}Z",
            "level": record.levelname,
            "logger": record.name,
            "event": record.getMessage(),
        }
        for key in SAFE_FIELDS:
            if key in record.__dict__:
                out[key] = record.__dict__[key]
        if record.exc_info and record.exc_info[0] is not None:
            # class name only: tracebacks can echo uploaded values
            out["exc"] = record.exc_info[0].__name__
        return json.dumps(out, ensure_ascii=False, default=str)


def configure_logging(level: str = "INFO", log_file: str = "", max_bytes: int = 0, backups: int = 5) -> None:
    """Install JSON handlers on the ``app`` and ``planner`` loggers (idempotent)."""
    fmt = JsonFormatter()
    handlers: list[logging.Handler] = [logging.StreamHandler(sys.stdout)]
    if log_file:
        handlers.append(
            logging.handlers.RotatingFileHandler(
                log_file, maxBytes=max_bytes or 5_000_000, backupCount=backups, encoding="utf-8"
            )
        )
    for name in ("app", "planner"):
        logger = logging.getLogger(name)
        for h in list(logger.handlers):
            if getattr(h, _HANDLER_TAG, False):
                logger.removeHandler(h)
                h.close()
        for h in handlers:
            h.setFormatter(fmt)
            setattr(h, _HANDLER_TAG, True)
            logger.addHandler(h)
        logger.setLevel(level)
        logger.propagate = False
