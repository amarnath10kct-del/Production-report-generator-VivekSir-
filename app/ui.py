"""Streamlit page: Upload master -> Generate -> Download.

The page keeps no data on disk. The generated workbook lives only in this browser session's
memory until a new file is uploaded or the tab is closed.
"""

from __future__ import annotations

import logging
import uuid
from pathlib import Path
from typing import Any

import streamlit as st

from app.access import client_ip, is_allowed
from app.config import ConfigError, Settings, load_dotenv_if_present
from app.logging_setup import configure_logging
from app.ratelimit import RateLimiter
from planner.errors import PlannerError
from planner.recalc import find_soffice, soffice_version
from planner.service import XLSX_MIME, GenerateResult, generate

log = logging.getLogger("app.ui")

APP_TITLE = "MDK Weekly Planning Generator"
STATIC = Path(__file__).resolve().parent.parent / "static"
GENERIC_ERROR = "Something went wrong while generating the workbook. Please try again."


@st.cache_resource(show_spinner=False)
def get_settings() -> Settings:
    """Load and validate settings once per server process."""
    load_dotenv_if_present()
    settings = Settings.from_env()
    configure_logging(settings.log_level, settings.log_file, settings.log_max_bytes, settings.log_backups)
    path = find_soffice(settings.soffice_path)
    log.info("app.start", extra={"soffice": bool(path), "status": "ready" if path else "degraded"})
    return settings


@st.cache_resource(show_spinner=False)
def get_limiter(limit: int) -> RateLimiter:
    """Process-wide rate limiter shared by all sessions."""
    return RateLimiter(limit)


@st.cache_data(ttl=300, show_spinner=False)
def readiness(soffice_path: str, enabled: bool) -> tuple[bool, str]:
    """Return ``(ready, text)`` describing LibreOffice availability (cached 5 minutes)."""
    if not enabled:
        return False, "Recalculation is turned off (RECALC_ENABLED=false)."
    path = find_soffice(soffice_path)
    if path is None:
        return False, "LibreOffice not found: files are generated without pre-calculated values."
    version = soffice_version(path)
    if version is None:
        return False, "LibreOffice is installed but cannot be started."
    return True, f"Ready: {version}"


def request_context() -> tuple[dict[str, str], str | None]:
    """Return the current request's headers and socket IP (monkeypatched in tests)."""
    try:
        headers = {str(k): str(v) for k, v in st.context.headers.items()}
        ip = st.context.ip_address
    except Exception:  # pragma: no cover - outside a browser session
        return {}, None
    return headers, ip if isinstance(ip, str) else None


def _css() -> None:
    css = (STATIC / "style.css").read_text(encoding="utf-8")
    st.markdown(f"<style>{css}</style>", unsafe_allow_html=True)


def _steps(current: int) -> None:
    labels = ["Upload master", "Generate", "Download"]
    items = []
    for i, label in enumerate(labels, start=1):
        state = "done" if i < current else ("active" if i == current else "todo")
        aria = ' aria-current="step"' if i == current else ""
        items.append(f'<li class="step {state}"{aria}><span class="num">{i}</span>{label}</li>')
    st.markdown(
        '<nav aria-label="Progress"><ol class="steps">' + "".join(items) + "</ol></nav>",
        unsafe_allow_html=True,
    )


def _clear_result() -> None:
    st.session_state.pop("result", None)
    st.session_state.pop("error", None)


def _fmt_qty(v: float) -> str:
    return f"{v:,.0f}" if float(v).is_integer() else f"{v:,.2f}"


def _result_panel(res: GenerateResult) -> None:
    s = res.summary
    st.success(f"Planning workbook ready: {s.week_title or 'week'}", icon=":material/check_circle:")
    sub = "  ·  ".join(x for x in (s.location, s.date_line) if x)
    if sub:
        st.caption(sub)
    c1, c2, c3 = st.columns(3)
    c1.metric("Work centres", f"{s.work_centres:,}")
    c2.metric("Items", f"{s.items:,}")
    c3.metric("Total weekly plan", _fmt_qty(s.total_weekly_qty))
    if s.warnings:
        with st.expander(f"{len(s.warnings)} warning(s) - the file was still generated", expanded=True):
            for w in s.warnings:
                st.warning(w, icon=":material/warning:")
    st.download_button(
        f"Download {res.filename}",
        data=res.content,
        file_name=res.filename,
        mime=XLSX_MIME,
        type="primary",
        icon=":material/download:",
        use_container_width=True,
        key="download",
    )


def _run_generation(upload: Any, settings: Settings, ip: str) -> None:
    rid = uuid.uuid4().hex[:12]
    allowed, retry = get_limiter(settings.rate_limit_per_hour).hit(ip)
    if not allowed:
        log.warning("generate.rejected", extra={"rid": rid, "status": 429, "outcome": "rate_limited"})
        st.session_state["error"] = (
            f"Too many generations in the last hour. Please try again in about {retry // 60 + 1} minute(s).",
            rid,
        )
        return
    data: bytes = upload.getvalue()
    try:
        with st.spinner("Generating the planning workbook... this can take up to a minute."):
            res = generate(data, upload.name, settings.generate_options())
    except PlannerError as exc:
        log.info(
            "generate.rejected",
            extra={
                "rid": rid,
                "status": exc.status,
                "error_class": type(exc).__name__,
                "size_bytes": len(data),
            },
        )
        st.session_state["error"] = (exc.message, rid)
        return
    except Exception as exc:
        log.error(
            "generate.failed",
            extra={"rid": rid, "status": 500, "error_class": type(exc).__name__, "size_bytes": len(data)},
        )
        st.session_state["error"] = (GENERIC_ERROR, rid)
        return
    t = res.summary.timings_ms
    log.info(
        "generate.ok",
        extra={
            "rid": rid,
            "status": 200,
            "size_bytes": len(data),
            "out_bytes": len(res.content),
            "work_centres": res.summary.work_centres,
            "items": res.summary.items,
            "warnings": len(res.summary.warnings),
            "recalculated": res.summary.recalculated,
            "ms_total": t.get("total"),
            "ms_validate": t.get("validate"),
            "ms_build": t.get("build"),
            "ms_recalc": t.get("recalc"),
        },
    )
    st.session_state["result"] = res


def main() -> None:
    """Render the page."""
    st.set_page_config(page_title=APP_TITLE, page_icon=":material/calendar_month:", layout="centered")
    _css()
    try:
        settings = get_settings()
    except ConfigError as exc:
        st.error(f"The app is misconfigured: {exc}. Fix the setting and restart.")
        st.stop()
        return

    headers, remote = request_context()
    ip = client_ip(headers, remote, settings.trusted_proxy_hops)
    if not is_allowed(ip, settings.allowed_ips):
        log.warning("access.denied", extra={"status": 403})
        st.error("Access denied (403). This app can only be used from the company network.")
        st.stop()
        return

    st.title(APP_TITLE)
    st.write("Upload the weekly production master (.xlsx) and download the day-wise planning workbook.")
    result: GenerateResult | None = st.session_state.get("result")
    upload = st.file_uploader(
        "Weekly production master (.xlsx)",
        type=["xlsx"],
        accept_multiple_files=False,
        help=f"Excel .xlsx only, up to {settings.max_upload_mb} MB. Nothing you upload is stored.",
        on_change=_clear_result,
        key="master",
    )
    _steps(3 if result else (2 if upload else 1))

    clicked = st.button(
        "Generate planning workbook",
        type="primary",
        disabled=upload is None,
        use_container_width=True,
        icon=":material/play_arrow:",
        key="generate",
    )
    if clicked and upload is not None:
        _clear_result()
        _run_generation(upload, settings, ip)
        result = st.session_state.get("result")

    err = st.session_state.get("error")
    if err:
        message, rid = err
        st.error(f"{message}\n\nReference: {rid}", icon=":material/error:")
    if result is not None:
        _result_panel(result)

    # Hidden by default; set SHOW_SERVER_STATUS=true to show it for troubleshooting.
    if settings.show_server_status:
        ready, text = readiness(settings.soffice_path, settings.recalc_enabled)
        with st.expander("Server status"):
            st.write(("✅ " if ready else "⚠️ ") + text)
            st.caption(
                f"Limits: {settings.max_upload_mb} MB per file, {settings.max_rows:,} rows, "
                f"{settings.rate_limit_per_hour or 'unlimited'} generations per hour."
            )
