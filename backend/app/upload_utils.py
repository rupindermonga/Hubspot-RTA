"""UploadFile → DataFrame helpers, with size + row guards."""
from __future__ import annotations

import io
import os

import pandas as pd
from fastapi import HTTPException, UploadFile

MAX_UPLOAD_MB = int(os.getenv("MAX_UPLOAD_MB", "50"))
MAX_UPLOAD_ROWS = int(os.getenv("MAX_UPLOAD_ROWS", "50000"))


async def read_to_dataframe(
    upload: UploadFile,
    label: str,
    sheet_name: str | None = None,
) -> pd.DataFrame:
    """Load an uploaded xlsx/csv into a DataFrame, enforcing size + row guards.

    label is used in error messages (e.g. "Hubspot", "RTA", "Old", "New").
    sheet_name is honored only for xlsx; ignored for csv.
    """
    raw = await upload.read()
    size_mb = len(raw) / (1024 * 1024)
    if size_mb > MAX_UPLOAD_MB:
        raise HTTPException(
            status_code=413,
            detail=f"{label} file too large ({size_mb:.1f} MB). Maximum is {MAX_UPLOAD_MB} MB.",
        )

    name = (upload.filename or "").lower()
    try:
        if name.endswith(".csv"):
            df = pd.read_csv(io.BytesIO(raw))
        elif name.endswith(".xlsx") or name.endswith(".xls"):
            xl = pd.ExcelFile(io.BytesIO(raw))
            chosen = sheet_name if (sheet_name and sheet_name in xl.sheet_names) else xl.sheet_names[0]
            df = pd.read_excel(xl, sheet_name=chosen)
        else:
            raise HTTPException(
                status_code=415,
                detail=f"{label} file type not supported. Upload .xlsx or .csv.",
            )
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=400, detail=f"Could not read {label} file: {e}")

    if len(df) > MAX_UPLOAD_ROWS:
        raise HTTPException(
            status_code=413,
            detail=f"{label} file has {len(df):,} rows. Maximum is {MAX_UPLOAD_ROWS:,}.",
        )
    return df


async def peek_columns(upload: UploadFile) -> tuple[list[str], list[str] | None]:
    """Read just enough of an upload to return its column list (and sheet names if xlsx).

    Returns (columns, sheet_names_or_None). For csv, sheet_names_or_None is None.
    """
    raw = await upload.read()
    await upload.seek(0)  # rewind so the same UploadFile can be read again later
    name = (upload.filename or "").lower()
    try:
        if name.endswith(".csv"):
            df = pd.read_csv(io.BytesIO(raw), nrows=0)
            return list(df.columns), None
        elif name.endswith(".xlsx") or name.endswith(".xls"):
            xl = pd.ExcelFile(io.BytesIO(raw))
            df = pd.read_excel(xl, sheet_name=xl.sheet_names[0], nrows=0)
            return list(df.columns), xl.sheet_names
        else:
            raise HTTPException(status_code=415, detail="Unsupported file type")
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=400, detail=f"Could not read file: {e}")
