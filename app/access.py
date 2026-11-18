"""Optional IP allow-list (``ALLOWED_IPS``). This is a safety net, not authentication."""

from __future__ import annotations

import ipaddress
from collections.abc import Mapping, Sequence

Network = ipaddress.IPv4Network | ipaddress.IPv6Network


def client_ip(headers: Mapping[str, str], remote_ip: str | None, trusted_hops: int) -> str:
    """Return the client's IP, honouring ``X-Forwarded-For`` from ``trusted_hops`` proxies.

    With N trusted proxies the client is the N-th address from the right of X-Forwarded-For
    (the proxies append to the right). Without the header, the socket address is used; a local
    connection (``None``) counts as ``127.0.0.1``.
    """
    xff = ""
    for k, v in headers.items():
        if k.lower() == "x-forwarded-for":
            xff = v
            break
    if trusted_hops > 0 and xff:
        parts = [p.strip() for p in xff.split(",") if p.strip()]
        if parts:
            return parts[-trusted_hops] if len(parts) >= trusted_hops else parts[0]
    return remote_ip or "127.0.0.1"


def is_allowed(ip: str, allowed: Sequence[Network]) -> bool:
    """Return True when ``allowed`` is empty or ``ip`` is inside one of its networks."""
    if not allowed:
        return True
    try:
        addr = ipaddress.ip_address(ip.split("%", 1)[0])
    except ValueError:
        return False
    if isinstance(addr, ipaddress.IPv6Address) and addr.ipv4_mapped is not None:
        addr = addr.ipv4_mapped
    return any(addr.version == net.version and addr in net for net in allowed)
