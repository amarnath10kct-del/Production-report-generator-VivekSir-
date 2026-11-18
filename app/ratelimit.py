"""A small thread-safe sliding-window rate limiter (in process memory).

Streamlit runs one server process, so one shared instance covers every browser session.
"""

from __future__ import annotations

import threading
import time
from collections import deque
from collections.abc import Callable


class RateLimiter:
    """Allow at most ``limit`` events per ``window_s`` seconds per key (``limit=0`` disables)."""

    def __init__(
        self, limit: int, window_s: float = 3600.0, clock: Callable[[], float] = time.monotonic
    ) -> None:
        self.limit = limit
        self.window_s = window_s
        self._clock = clock
        self._events: dict[str, deque[float]] = {}
        self._lock = threading.Lock()

    def hit(self, key: str) -> tuple[bool, int]:
        """Record an attempt for ``key``.

        Returns:
            ``(allowed, retry_after_seconds)``. A refused attempt is not recorded.

        """
        if self.limit <= 0:
            return True, 0
        now = self._clock()
        with self._lock:
            q = self._events.setdefault(key, deque())
            while q and now - q[0] >= self.window_s:
                q.popleft()
            if len(q) >= self.limit:
                return False, max(1, int(self.window_s - (now - q[0])))
            q.append(now)
            # drop idle keys so memory stays bounded
            if len(self._events) > 10_000:
                for k in [k for k, v in self._events.items() if not v]:
                    del self._events[k]
            return True, 0
