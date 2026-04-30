"""LRU cache of deserialized matcher DataFrames keyed by dataset_id.

The Dataset table stores `df_hub_pickle` / `df_rta_pickle` as encrypted BLOBs.
Deserializing them on every search would be wasteful, so we keep the most
recently-used datasets' DataFrames in process memory.

This is RAM-only — no disk fallback. Process restarts wipe the cache; the
next access deserializes from the (encrypted) DB again. Capped to keep
memory bounded if a user has many datasets.
"""
from __future__ import annotations

import os
from collections import OrderedDict
from threading import Lock

import pandas as pd

_MAX_CACHED = int(os.getenv("DF_CACHE_SIZE", "10"))

_lock = Lock()
_cache: "OrderedDict[int, dict]" = OrderedDict()


def put(dataset_id: int, df_hub: pd.DataFrame, df_rta: pd.DataFrame) -> None:
    with _lock:
        _cache.pop(dataset_id, None)
        _cache[dataset_id] = {"df_hub": df_hub, "df_rta": df_rta}
        while len(_cache) > _MAX_CACHED:
            _cache.popitem(last=False)


def get(dataset_id: int) -> dict | None:
    """Return {'df_hub': ..., 'df_rta': ...} or None if not cached."""
    with _lock:
        if dataset_id not in _cache:
            return None
        _cache.move_to_end(dataset_id)
        return _cache[dataset_id]


def evict(dataset_id: int) -> None:
    with _lock:
        _cache.pop(dataset_id, None)
