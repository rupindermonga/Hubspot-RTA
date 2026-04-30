"""Address search across saved Hubspot+RTA datasets.

Reuses the matcher's normalization helpers + 5-pass key system, applies them
to the user's free-text query, and finds matching rows in any of the user's
saved datasets. Returns the original row data + which side(s) the address
appears on + how it was matched (exact / direction-strip / alias / unit-strip).
"""
from __future__ import annotations

from typing import Iterable

import pandas as pd

from .matcher import (
    DEFAULT_ALIASES,
    apply_canonical,
    normalize,
    norm_pc,
    strip_direction,
    strip_unit,
)


def build_query_keys(q: str, postal: str = "", aliases: Iterable[tuple[str, str]] | None = None) -> dict:
    """Build the normalization-key set for a free-text query.

    Mirrors the 6 keys the matcher uses (_k_exact, _k_dir, _k_canon, _k_canon_dir,
    _k_unit, _k_unit_dir). When postal is empty, the keys still work — they just
    end with '|' which only matches rows with empty/missing postal codes. So we
    also build a postal-stripped variant in that case.
    """
    canonical_map = dict(aliases) if aliases is not None else dict(DEFAULT_ALIASES)
    street = normalize(q)
    street_canon = apply_canonical(street, canonical_map)
    pc = norm_pc(postal) if postal else ""

    keys = {
        "exact": f"{street}|{pc}",
        "dir": f"{strip_direction(street)}|{pc}",
        "canon": f"{street_canon}|{pc}",
        "canon_dir": f"{strip_direction(street_canon)}|{pc}",
        "unit": f"{strip_unit(street)}|{pc}",
        "unit_dir": f"{strip_direction(strip_unit(street))}|{pc}",
    }
    return {
        "street": street,
        "street_canon": street_canon,
        "pc": pc,
        "keys": keys,
    }


# Map of column-name → matcher's pass-name (what we report as match_type)
_KEY_TO_TYPE = {
    "_k_exact": "exact",
    "_k_dir": "direction_strip",
    "_k_canon": "fuzzy",
    "_k_canon_dir": "fuzzy",
    "_k_unit": "unit_strip",
    "_k_unit_dir": "unit_strip",
}


def _matched_rows(df: pd.DataFrame, query_keys: dict) -> list[dict]:
    """For a single side's keyed DataFrame, return rows where any key matches.

    Earliest-strict-pass wins for `match_type` reporting (so an exact match
    isn't downgraded to 'fuzzy' just because the row's canon key also matched).
    """
    if len(df) == 0:
        return []

    hits: dict[int, str] = {}  # row_index → match_type (assigned once, by strictest pass)
    for col, mtype in (
        ("_k_exact", "exact"),
        ("_k_dir", "direction_strip"),
        ("_k_canon", "fuzzy"),
        ("_k_canon_dir", "fuzzy"),
        ("_k_unit", "unit_strip"),
        ("_k_unit_dir", "unit_strip"),
    ):
        if col not in df.columns:
            continue
        # Compare against the matching variant of the query
        target = query_keys.get({
            "_k_exact": "exact",
            "_k_dir": "dir",
            "_k_canon": "canon",
            "_k_canon_dir": "canon_dir",
            "_k_unit": "unit",
            "_k_unit_dir": "unit_dir",
        }[col])
        if not target:
            continue
        mask = df[col] == target
        for idx in df[mask].index:
            if idx not in hits:
                hits[idx] = mtype

    if not hits:
        return []

    # Drop the underscore columns from the row payload (those are internal)
    public_cols = [c for c in df.columns if not c.startswith("_")]
    out = []
    for idx, mtype in hits.items():
        row = df.loc[idx, public_cols].to_dict()
        # Coerce NaN → '' for clean JSON
        for k, v in list(row.items()):
            if isinstance(v, float) and pd.isna(v):
                row[k] = ""
            elif v is None:
                row[k] = ""
        out.append({"match_type": mtype, "row": row})
    return out


def search_in_dataset(
    df_hub: pd.DataFrame,
    df_rta: pd.DataFrame,
    q: str,
    postal: str = "",
    aliases: Iterable[tuple[str, str]] | None = None,
) -> dict:
    """Search a single dataset; return {in_hubspot: [...], in_rta: [...], normalized_keys: [...]}."""
    qk = build_query_keys(q, postal, aliases)
    keys = qk["keys"]
    in_hub = _matched_rows(df_hub, keys)
    in_rta = _matched_rows(df_rta, keys)
    return {
        "in_hubspot": in_hub,
        "in_rta": in_rta,
        "normalized_keys": sorted({v for v in keys.values()}),
    }
