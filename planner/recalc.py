"""Recalculate a workbook with headless LibreOffice so cached values are embedded.

Each call runs ``soffice`` in its own private temp directory with its own user profile (so calls
never share state or lock files), under a hard timeout and a process-wide concurrency limit.
Everything is deleted in ``finally``, also after a timeout or crash.
"""

from __future__ import annotations

import contextlib
import logging
import os
import shutil
import signal
import subprocess
import sys
import tempfile
import threading
import time
from dataclasses import dataclass
from functools import lru_cache
from pathlib import Path

log = logging.getLogger(__name__)

#: Forces "always recalculate on load" for OOXML/ODF in the throw-away profile.
_REGISTRY = """<?xml version="1.0" encoding="UTF-8"?>
<oor:items xmlns:oor="http://openoffice.org/2001/registry" \
xmlns:xs="http://www.w3.org/2001/XMLSchema" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
<item oor:path="/org.openoffice.Office.Calc/Formula/Load">\
<prop oor:name="OOXMLRecalcMode" oor:op="fuse"><value>0</value></prop></item>
<item oor:path="/org.openoffice.Office.Calc/Formula/Load">\
<prop oor:name="ODFRecalcMode" oor:op="fuse"><value>0</value></prop></item>
<item oor:path="/org.openoffice.Office.Common/Misc">\
<prop oor:name="FirstRun" oor:op="fuse"><value>false</value></prop></item>
</oor:items>
"""

_WINDOWS_CANDIDATES = (
    r"C:\Program Files\LibreOffice\program\soffice.exe",
    r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
    # Per-user copy (no admin): msiexec /a LibreOffice.msi TARGETDIR=%LOCALAPPDATA%\LibreOffice
    os.path.expandvars(r"%LOCALAPPDATA%\LibreOffice\program\soffice.exe"),
)
_POSIX_CANDIDATES = (
    "/usr/bin/soffice",
    "/usr/lib/libreoffice/program/soffice",
    "/opt/libreoffice/program/soffice",
    "/Applications/LibreOffice.app/Contents/MacOS/soffice",
)
TEMP_PREFIX = "mdkplan-"


class RecalcError(Exception):
    """LibreOffice is unavailable, failed, timed out, or the system is busy."""


@dataclass(frozen=True)
class RecalcOptions:
    """Settings for :func:`recalculate`."""

    soffice_path: str = ""
    timeout_s: float = 60.0
    concurrency: int = 1
    temp_root: str | None = None


def find_soffice(configured: str = "") -> str | None:
    """Return the path of the LibreOffice binary, or ``None`` if it is not installed."""
    if configured:
        return configured if Path(configured).is_file() else None
    for name in ("soffice", "libreoffice"):
        found = shutil.which(name)
        if found:
            return found
    for cand in _WINDOWS_CANDIDATES if os.name == "nt" else _POSIX_CANDIDATES:
        if Path(cand).is_file():
            return cand
    return None


@lru_cache(maxsize=4)
def soffice_version(path: str) -> str | None:
    """Return LibreOffice's version string (cached), or ``None`` if it cannot be run."""
    try:
        out = subprocess.run(  # noqa: S603 - fixed argv, no shell
            [path, "--version"], capture_output=True, timeout=60, check=False
        )
    except (OSError, subprocess.SubprocessError):
        return None
    text = out.stdout.decode("utf-8", "replace").strip()
    return text or ("LibreOffice" if out.returncode == 0 else None)


_sem_lock = threading.Lock()
_semaphores: dict[int, threading.BoundedSemaphore] = {}


def _semaphore(n: int) -> threading.BoundedSemaphore:
    with _sem_lock:
        if n not in _semaphores:
            _semaphores[n] = threading.BoundedSemaphore(n)
        return _semaphores[n]


def _kill_tree(proc: subprocess.Popen[bytes]) -> None:
    """Kill ``proc`` and everything it started."""
    with contextlib.suppress(OSError, subprocess.SubprocessError):
        if sys.platform == "win32":
            subprocess.run(  # noqa: S603
                ["taskkill", "/F", "/T", "/PID", str(proc.pid)],  # noqa: S607
                capture_output=True,
                timeout=30,
                check=False,
            )
        else:
            os.killpg(proc.pid, signal.SIGKILL)
    with contextlib.suppress(OSError):
        proc.kill()
    with contextlib.suppress(subprocess.TimeoutExpired):
        proc.wait(timeout=10)


def _rmtree(path: Path) -> None:
    """Delete ``path``; retry briefly because Windows may hold files for a moment after exit."""
    # The profile nests files deeper than MAX_PATH (260); the \\?\ prefix lifts that limit.
    target = "\\\\?\\" + str(path.resolve()) if os.name == "nt" else str(path)
    for _ in range(10):
        shutil.rmtree(target, ignore_errors=True)
        if not path.exists():
            return
        time.sleep(0.3)
    log.warning("recalc.cleanup_incomplete")


def recalculate(xlsx: bytes, opts: RecalcOptions) -> bytes:
    """Return ``xlsx`` re-saved by LibreOffice with all formulas recalculated.

    Raises:
        RecalcError: on any failure (missing binary, busy, timeout, crash, no output).

    """
    binary = find_soffice(opts.soffice_path)
    if binary is None:
        raise RecalcError("LibreOffice not found")
    sem = _semaphore(max(1, opts.concurrency))
    if not sem.acquire(timeout=opts.timeout_s):
        raise RecalcError("recalculation busy")
    try:
        return _run(binary, xlsx, opts)
    finally:
        sem.release()


def _run(binary: str, xlsx: bytes, opts: RecalcOptions) -> bytes:
    work = Path(tempfile.mkdtemp(prefix=TEMP_PREFIX, dir=opts.temp_root))
    proc: subprocess.Popen[bytes] | None = None
    try:
        os.chmod(work, 0o700)
        src_dir, out_dir, profile = work / "in", work / "out", work / "profile"
        for d in (src_dir, out_dir, profile / "user"):
            d.mkdir(parents=True)
        (profile / "user" / "registrymodifications.xcu").write_text(_REGISTRY, encoding="utf-8")
        src = src_dir / "workbook.xlsx"
        src.write_bytes(xlsx)
        argv = [
            binary,
            f"-env:UserInstallation={profile.resolve().as_uri()}",
            "--headless",
            "--invisible",
            "--norestore",
            "--nolockcheck",
            "--nodefault",
            "--nologo",
            "--convert-to",
            "xlsx:Calc MS Excel 2007 XML",
            "--outdir",
            str(out_dir),
            str(src),
        ]
        kwargs: dict[str, object] = {}
        if os.name == "nt":
            kwargs["creationflags"] = subprocess.CREATE_NEW_PROCESS_GROUP | getattr(
                subprocess, "CREATE_NO_WINDOW", 0
            )
        else:
            kwargs["start_new_session"] = True
        env = dict(os.environ)
        env["HOME"] = str(work)  # keep LibreOffice away from the service user's home
        proc = subprocess.Popen(  # noqa: S603 - fixed argv, shell=False
            argv,
            stdin=subprocess.DEVNULL,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            env=env,
            **kwargs,  # type: ignore[call-overload]
        )
        try:
            proc.communicate(timeout=opts.timeout_s)
        except subprocess.TimeoutExpired as exc:
            _kill_tree(proc)
            raise RecalcError("recalculation timed out") from exc
        out = out_dir / "workbook.xlsx"
        if proc.returncode != 0 or not out.is_file() or out.stat().st_size == 0:
            raise RecalcError(f"LibreOffice failed (exit {proc.returncode})")
        return out.read_bytes()
    except OSError as exc:
        raise RecalcError(f"LibreOffice could not run ({type(exc).__name__})") from exc
    finally:
        if proc is not None and proc.poll() is None:
            _kill_tree(proc)
        _rmtree(work)
