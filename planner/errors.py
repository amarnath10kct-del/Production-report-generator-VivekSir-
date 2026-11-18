"""Exceptions raised by the planner. Messages are safe to show to the user (no data, no paths)."""

from __future__ import annotations


class PlannerError(Exception):
    """Base class for user-facing errors.

    Attributes:
        message: Plain-language explanation shown to the user.
        status: The HTTP status this error would map to (kept for logging/metrics).

    """

    status: int = 400

    def __init__(self, message: str) -> None:
        super().__init__(message)
        self.message = message


class InvalidUpload(PlannerError):
    """The upload is not an acceptable .xlsx workbook (wrong type, macros, encrypted, corrupt...)."""

    status = 400


class UploadTooLarge(PlannerError):
    """The upload, its uncompressed content, or its sheet exceeds a configured size limit."""

    status = 413


class MasterFormatError(PlannerError):
    """The workbook is valid but the master sheet does not have the expected layout."""

    status = 422
