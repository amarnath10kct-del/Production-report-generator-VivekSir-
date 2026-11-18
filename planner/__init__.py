"""Planner core: read a weekly master, build the approved planning workbook, recalculate it.

This package has no web-framework dependency; the Streamlit UI (``streamlit_app.py``) and the
tests call :func:`planner.service.generate`.
"""

from planner.errors import PlannerError
from planner.service import GenerateOptions, GenerateResult, generate

__all__ = ["GenerateOptions", "GenerateResult", "PlannerError", "generate"]
