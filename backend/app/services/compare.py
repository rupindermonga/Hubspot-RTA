"""RTA Status Compare engine — Old RTA snapshot vs New RTA snapshot.

Pure-Python port of the Streamlit prototype's comparison logic. Reuses the
normalization helpers from matcher.py so format quirks ("Street" vs "ST",
trailing N/S, alias variants) don't falsely show as add/remove between
snapshots.

Output Excel sheets:
  1. Dashboard         summary counts
  2. Old RTA           raw upload
  3. New RTA           raw upload
  4. Removed           in Old, not in New (purple)
  5. Added             in New, not in Old (blue)
  6. Status Changed    same address, different status — green for In Construction → RTA, yellow for the reverse
  7. Conflicts in New  same address with multiple statuses (orange)
"""
from __future__ import annotations

import io
from typing import Iterable

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

from .matcher import (
    DEFAULT_ALIASES,
    apply_canonical,
    normalize,
    norm_pc,
    sanitize_dataframe,
    strip_direction,
    strip_unit,
)


def run_compare(
    df_old: pd.DataFrame,
    df_new: pd.DataFrame,
    *,
    addr_no_col: str,
    street_col: str,
    locality_col: str,
    pc_col: str,
    status_col: str,
    aliases: Iterable[tuple[str, str]] | None = None,
) -> dict:
    """Compare old vs new RTA snapshots.

    Both files are assumed to have identical column names — caller is responsible
    for surfacing a friendly error if not.

    Returns a dict with:
      - stats (dict)
      - removed_preview, added_preview, status_changed_preview, conflicts_preview
      - excel_bytes (bytes)  multi-sheet color-coded XLSX
    """
    canonical_map = dict(aliases) if aliases is not None else dict(DEFAULT_ALIASES)

    df_old = df_old.copy()
    df_new = df_new.copy()

    for df in (df_old, df_new):
        df['_full'] = (
            df[addr_no_col].fillna('').astype(str).str.strip() + ' ' +
            df[street_col].fillna('').astype(str).str.strip() + ' ' +
            df[locality_col].fillna('').astype(str).str.strip() + ' ' +
            df[pc_col].fillna('').astype(str).str.strip()
        ).str.strip()
        df['_street'] = (
            df[addr_no_col].fillna('').astype(str) + ' ' + df[street_col].fillna('')
        ).apply(normalize)
        df['_street_canon'] = df['_street'].apply(lambda s: apply_canonical(s, canonical_map))
        df['_pc'] = df[pc_col].fillna('').apply(norm_pc)
        df['_key'] = df['_street_canon'].apply(strip_direction).apply(strip_unit) + '|' + df['_pc']
        df['_status_norm'] = df[status_col].fillna('').astype(str).str.strip().str.upper()

    # Conflicts WITHIN new file — same key, multiple distinct statuses
    new_dup = df_new.groupby('_key')['_status_norm'].nunique()
    conflict_keys = set(new_dup[new_dup > 1].index)
    df_conflicts = df_new[df_new['_key'].isin(conflict_keys)][
        [addr_no_col, street_col, locality_col, pc_col, status_col]
    ].copy()
    if len(df_conflicts) > 0:
        df_conflicts = df_conflicts.sort_values(
            by=[street_col, addr_no_col, status_col]
        ).reset_index(drop=True)

    # Per-file key → first status maps
    old_key_status = df_old.drop_duplicates(subset='_key').set_index('_key')['_status_norm'].to_dict()
    new_key_status = df_new.drop_duplicates(subset='_key').set_index('_key')['_status_norm'].to_dict()

    old_keys = set(old_key_status.keys()) - {''}
    new_keys = set(new_key_status.keys()) - {''}

    removed_keys = old_keys - new_keys
    added_keys = new_keys - old_keys
    common_keys = old_keys & new_keys

    df_removed = df_old[df_old['_key'].isin(removed_keys)].drop_duplicates(subset='_key')[
        [addr_no_col, street_col, locality_col, pc_col, status_col]
    ].sort_values(by=[street_col, addr_no_col]).reset_index(drop=True)

    df_added = df_new[df_new['_key'].isin(added_keys)].drop_duplicates(subset='_key')[
        [addr_no_col, street_col, locality_col, pc_col, status_col]
    ].sort_values(by=[street_col, addr_no_col]).reset_index(drop=True)

    # Status changes among addresses present in both files
    new_first_row_by_key = df_new.drop_duplicates(subset='_key').set_index('_key')
    status_change_rows = []
    for key in common_keys:
        old_s = old_key_status.get(key, '')
        new_s = new_key_status.get(key, '')
        if old_s != new_s:
            row = new_first_row_by_key.loc[key]
            status_change_rows.append({
                'Address Number': str(row[addr_no_col]),
                'Street Name': str(row[street_col]),
                'Locality': str(row[locality_col]),
                'Postal Code': str(row[pc_col]),
                'Old Status': old_s.title() if old_s else '',
                'New Status': new_s.title() if new_s else '',
            })
    if status_change_rows:
        df_status_changed = pd.DataFrame(status_change_rows).sort_values(
            by=['Street Name', 'Address Number']
        ).reset_index(drop=True)
    else:
        df_status_changed = pd.DataFrame(columns=[
            'Address Number', 'Street Name', 'Locality', 'Postal Code', 'Old Status', 'New Status'
        ])

    # Status-change colour breakdown
    if len(df_status_changed) > 0:
        forward_count = int((
            (df_status_changed['Old Status'].str.upper() == 'IN CONSTRUCTION') &
            (df_status_changed['New Status'].str.upper() == 'RTA')
        ).sum())
        regress_count = int((
            (df_status_changed['Old Status'].str.upper() == 'RTA') &
            (df_status_changed['New Status'].str.upper() == 'IN CONSTRUCTION')
        ).sum())
        other_count = len(df_status_changed) - forward_count - regress_count
    else:
        forward_count = regress_count = other_count = 0

    # ── Build the colored XLSX ────────────────────────────────────────────
    df_dashboard = pd.DataFrame([
        ['Old RTA total rows', len(df_old)],
        ['New RTA total rows', len(df_new)],
        ['Removed (in Old, not in New)', len(df_removed)],
        ['Added (in New, not in Old)', len(df_added)],
        ['Status Changed (same address, different status)', len(df_status_changed)],
        ['  └─ In Construction → RTA', forward_count],
        ['  └─ RTA → In Construction', regress_count],
        ['Conflicts in New — unique addresses', len(conflict_keys)],
        ['Conflicts in New — total rows', len(df_conflicts)],
    ], columns=['Metric', 'Value'])

    df_old_out = df_old.drop(columns=[c for c in df_old.columns if c.startswith('_')])
    df_new_out = df_new.drop(columns=[c for c in df_new.columns if c.startswith('_')])

    df_dashboard_san = sanitize_dataframe(df_dashboard)
    df_old_san = sanitize_dataframe(df_old_out)
    df_new_san = sanitize_dataframe(df_new_out)
    df_removed_san = sanitize_dataframe(df_removed)
    df_added_san = sanitize_dataframe(df_added)
    df_status_changed_san = sanitize_dataframe(df_status_changed)
    df_conflicts_san = sanitize_dataframe(df_conflicts)

    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        df_dashboard_san.to_excel(writer, sheet_name='Dashboard', index=False)
        df_old_san.to_excel(writer, sheet_name='Old RTA', index=False)
        df_new_san.to_excel(writer, sheet_name='New RTA', index=False)
        df_removed_san.to_excel(writer, sheet_name='Removed', index=False)
        df_added_san.to_excel(writer, sheet_name='Added', index=False)
        df_status_changed_san.to_excel(writer, sheet_name='Status Changed', index=False)
        df_conflicts_san.to_excel(writer, sheet_name='Conflicts in New', index=False)
    buffer.seek(0)

    wb = load_workbook(buffer)

    purple_fill = PatternFill(start_color='D8B4FE', end_color='D8B4FE', fill_type='solid')
    blue_fill = PatternFill(start_color='BFDBFE', end_color='BFDBFE', fill_type='solid')
    green_fill = PatternFill(start_color='86EFAC', end_color='86EFAC', fill_type='solid')
    yellow_fill = PatternFill(start_color='FFFF00', end_color='FFFF00', fill_type='solid')
    orange_fill = PatternFill(start_color='FFA500', end_color='FFA500', fill_type='solid')

    # Removed: purple
    ws = wb['Removed']
    for row_idx in range(2, ws.max_row + 1):
        for col_idx in range(1, ws.max_column + 1):
            ws.cell(row=row_idx, column=col_idx).fill = purple_fill

    # Added: blue
    ws = wb['Added']
    for row_idx in range(2, ws.max_row + 1):
        for col_idx in range(1, ws.max_column + 1):
            ws.cell(row=row_idx, column=col_idx).fill = blue_fill

    # Status Changed: green for In Construction → RTA, yellow for RTA → In Construction
    ws = wb['Status Changed']
    old_status_idx = None
    new_status_idx = None
    for cell in ws[1]:
        if cell.value == 'Old Status':
            old_status_idx = cell.column
        elif cell.value == 'New Status':
            new_status_idx = cell.column
    if old_status_idx and new_status_idx:
        for row_idx in range(2, ws.max_row + 1):
            old_s = str(ws.cell(row=row_idx, column=old_status_idx).value or '').upper().strip()
            new_s = str(ws.cell(row=row_idx, column=new_status_idx).value or '').upper().strip()
            fill = None
            if old_s == 'IN CONSTRUCTION' and new_s == 'RTA':
                fill = green_fill
            elif old_s == 'RTA' and new_s == 'IN CONSTRUCTION':
                fill = yellow_fill
            if fill:
                for col_idx in range(1, ws.max_column + 1):
                    ws.cell(row=row_idx, column=col_idx).fill = fill

    # Conflicts: orange
    ws = wb['Conflicts in New']
    for row_idx in range(2, ws.max_row + 1):
        for col_idx in range(1, ws.max_column + 1):
            ws.cell(row=row_idx, column=col_idx).fill = orange_fill

    out_buffer = io.BytesIO()
    wb.save(out_buffer)
    excel_bytes = out_buffer.getvalue()

    # Previews for the SPA dashboard (cap at 50 rows)
    return {
        'stats': {
            'old_total': len(df_old),
            'new_total': len(df_new),
            'removed': len(df_removed),
            'added': len(df_added),
            'status_changed': len(df_status_changed),
            'in_construction_to_rta': forward_count,
            'rta_to_in_construction': regress_count,
            'other_status_changes': other_count,
            'conflict_addresses': len(conflict_keys),
            'conflict_rows': len(df_conflicts),
        },
        'removed_preview': df_removed.head(50).to_dict(orient='records'),
        'added_preview': df_added.head(50).to_dict(orient='records'),
        'status_changed_preview': df_status_changed.head(100).to_dict(orient='records'),
        'conflicts_preview': df_conflicts.head(100).to_dict(orient='records'),
        'excel_bytes': excel_bytes,
    }
