"""Configuration from environment variables (optionally a ``.env`` file), validated at startup.

On Streamlit Community Cloud, root-level keys in the app's *Secrets* are exposed as environment
variables, so the same names work there.
"""

from __future__ import annotations

import ipaddress
import os
from collections.abc import Mapping
from dataclasses import dataclass

from planner.recalc import RecalcOptions
from planner.service import GenerateOptions
from planner.validate import UploadLimits

_MB = 1024 * 1024
_LEVELS = {"DEBUG", "INFO", "WARNING", "ERROR", "CRITICAL"}


class ConfigError(ValueError):
    """An environment variable has an invalid value."""


@dataclass(frozen=True)
class Settings:
    """Validated application settings (see README for meaning and defaults)."""

    max_upload_mb: int = 10
    max_uncompressed_mb: int = 100
    max_zip_entries: int = 1000
    max_rows: int = 20_000
    max_cols: int = 50
    recalc_enabled: bool = True
    soffice_path: str = ""
    recalc_timeout_s: int = 60
    recalc_concurrency: int = 1
    rate_limit_per_hour: int = 60
    allowed_ips: tuple[ipaddress.IPv4Network | ipaddress.IPv6Network, ...] = ()
    trusted_proxy_hops: int = 1
    log_level: str = "INFO"
    log_file: str = ""
    log_max_bytes: int = 5 * _MB
    log_backups: int = 5

    @classmethod
    def from_env(cls, env: Mapping[str, str] | None = None) -> Settings:
        """Read and validate settings; raise :class:`ConfigError` naming the bad variable."""
        e = os.environ if env is None else env

        def integer(name: str, default: int, lo: int, hi: int) -> int:
            raw = e.get(name, "").strip()
            if not raw:
                return default
            try:
                val = int(raw)
            except ValueError as exc:
                raise ConfigError(f"{name} must be a whole number, got {raw!r}") from exc
            if not lo <= val <= hi:
                raise ConfigError(f"{name} must be between {lo} and {hi}, got {val}")
            return val

        def boolean(name: str, default: bool) -> bool:
            raw = e.get(name, "").strip().lower()
            if not raw:
                return default
            if raw in {"1", "true", "yes", "on"}:
                return True
            if raw in {"0", "false", "no", "off"}:
                return False
            raise ConfigError(f"{name} must be true or false, got {raw!r}")

        level = (e.get("LOG_LEVEL", "") or "INFO").strip().upper()
        if level not in _LEVELS:
            raise ConfigError(f"LOG_LEVEL must be one of {sorted(_LEVELS)}, got {level!r}")
        soffice = e.get("SOFFICE_PATH", "").strip()
        if soffice and not os.path.isfile(soffice):
            raise ConfigError(f"SOFFICE_PATH does not point to a file: {soffice!r}")
        return cls(
            max_upload_mb=integer("MAX_UPLOAD_MB", 10, 1, 200),
            max_uncompressed_mb=integer("MAX_UNCOMPRESSED_MB", 100, 1, 2000),
            max_zip_entries=integer("MAX_ZIP_ENTRIES", 1000, 10, 100_000),
            max_rows=integer("MAX_ROWS", 20_000, 10, 1_000_000),
            max_cols=integer("MAX_COLS", 50, 3, 16_384),
            recalc_enabled=boolean("RECALC_ENABLED", True),
            soffice_path=soffice,
            recalc_timeout_s=integer("RECALC_TIMEOUT_S", 60, 5, 600),
            recalc_concurrency=integer("RECALC_CONCURRENCY", 1, 1, 16),
            rate_limit_per_hour=integer("RATE_LIMIT_PER_HOUR", 60, 0, 100_000),
            allowed_ips=parse_allowed_ips(e.get("ALLOWED_IPS", "")),
            trusted_proxy_hops=integer("TRUSTED_PROXY_HOPS", 1, 0, 10),
            log_level=level,
            log_file=e.get("LOG_FILE", "").strip(),
            log_max_bytes=integer("LOG_MAX_BYTES", 5 * _MB, 10_000, 1_000 * _MB),
            log_backups=integer("LOG_BACKUPS", 5, 0, 100),
        )

    def generate_options(self) -> GenerateOptions:
        """Options for :func:`planner.service.generate`."""
        return GenerateOptions(
            limits=UploadLimits(
                max_upload_bytes=self.max_upload_mb * _MB,
                max_uncompressed_bytes=self.max_uncompressed_mb * _MB,
                max_zip_entries=self.max_zip_entries,
                max_rows=self.max_rows,
                max_cols=self.max_cols,
            ),
            recalc_enabled=self.recalc_enabled,
            recalc=RecalcOptions(
                soffice_path=self.soffice_path,
                timeout_s=float(self.recalc_timeout_s),
                concurrency=self.recalc_concurrency,
            ),
        )


def parse_allowed_ips(raw: str) -> tuple[ipaddress.IPv4Network | ipaddress.IPv6Network, ...]:
    """Parse ``ALLOWED_IPS`` (comma-separated IPs/CIDRs). Empty means allow everyone."""
    nets = []
    for part in (p.strip() for p in raw.split(",")):
        if not part:
            continue
        try:
            nets.append(ipaddress.ip_network(part, strict=False))
        except ValueError as exc:
            raise ConfigError(f"ALLOWED_IPS has an invalid entry: {part!r}") from exc
    return tuple(nets)


def load_dotenv_if_present(path: str = ".env") -> None:
    """Load ``.env`` (never overriding real environment variables)."""
    from dotenv import load_dotenv

    if os.path.isfile(path):
        load_dotenv(path, override=False)
