"""In-memory cache for one-shot Excel downloads.

The matcher / compare routes generate an XLSX in memory, store it here
keyed by an unguessable token, and return that token in the JSON response.
The SPA then GETs /api/<service>/download/<token> to retrieve the file.

TTL: 30 minutes. The cache is per-process — a uvicorn restart clears it.
That's intentional: results are ephemeral. Single-process deploy means no
cross-worker coordination needed; if we ever scale to multiple workers,
swap this for Redis or write the bytes to UPLOAD_FOLDER.
"""
from __future__ import annotations

import secrets
import time
from threading import Lock

_TTL_SECONDS = 30 * 60

_lock = Lock()
_cache: dict[str, dict] = {}


def store(excel_bytes: bytes, filename: str) -> str:
    """Stash bytes under a fresh token; return the token."""
    with _lock:
        _evict_expired()
        token = secrets.token_urlsafe(24)
        _cache[token] = {
            "excel_bytes": excel_bytes,
            "filename": filename,
            "created_at": time.time(),
        }
        return token


def fetch(token: str) -> tuple[bytes, str] | None:
    """Return (bytes, filename) or None if missing/expired."""
    with _lock:
        _evict_expired()
        entry = _cache.get(token)
        if entry is None:
            return None
        return entry["excel_bytes"], entry["filename"]


def _evict_expired() -> None:
    """Caller must hold _lock."""
    now = time.time()
    for t in list(_cache.keys()):
        if now - _cache[t]["created_at"] > _TTL_SECONDS:
            del _cache[t]
