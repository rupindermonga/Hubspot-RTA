from fastapi import APIRouter, Depends, File, Form, HTTPException, UploadFile
from fastapi.responses import Response

from ..dependencies import get_current_user
from ..models import User
from ..services.matcher import run_match
from ..upload_utils import peek_columns, read_to_dataframe
from ..download_cache import fetch as cache_fetch
from ..download_cache import store as cache_store

router = APIRouter(prefix="/api/matcher", tags=["matcher"])


@router.post("/peek")
async def peek(
    file: UploadFile = File(...),
    current_user: User = Depends(get_current_user),
):
    """Return column names + sheet list for an uploaded xlsx/csv.

    Used by the SPA to populate column dropdowns before the user runs the match.
    """
    columns, sheet_names = await peek_columns(file)
    return {"columns": columns, "sheet_names": sheet_names}


@router.post("/run")
async def run(
    hub_file: UploadFile = File(...),
    rta_file: UploadFile = File(...),
    hub_street_col: str = Form(...),
    hub_pc_col: str = Form(...),
    rta_addr_no_col: str = Form(...),
    rta_street_col: str = Form(...),
    rta_locality_col: str = Form(...),
    rta_pc_col: str = Form(...),
    rta_status_col: str = Form(...),
    hub_sheet: str | None = Form(None),
    rta_sheet: str | None = Form(None),
    enable_no_pc: bool = Form(False),
    current_user: User = Depends(get_current_user),
):
    df_hub = await read_to_dataframe(hub_file, "Hubspot", sheet_name=hub_sheet)
    df_rta = await read_to_dataframe(rta_file, "RTA", sheet_name=rta_sheet)

    # Validate column choices exist in the uploaded files
    missing_hub = [c for c in (hub_street_col, hub_pc_col) if c not in df_hub.columns]
    missing_rta = [c for c in (rta_addr_no_col, rta_street_col, rta_locality_col, rta_pc_col, rta_status_col) if c not in df_rta.columns]
    if missing_hub or missing_rta:
        raise HTTPException(
            status_code=422,
            detail={
                "message": "One or more selected columns were not found.",
                "missing_in_hubspot": missing_hub,
                "missing_in_rta": missing_rta,
            },
        )

    result = run_match(
        df_hub,
        df_rta,
        hub_street_col=hub_street_col,
        hub_pc_col=hub_pc_col,
        rta_addr_no_col=rta_addr_no_col,
        rta_street_col=rta_street_col,
        rta_locality_col=rta_locality_col,
        rta_pc_col=rta_pc_col,
        rta_status_col=rta_status_col,
        enable_no_pc=enable_no_pc,
    )

    token = cache_store(result["excel_bytes"], "hubspot_rta_matched_output.xlsx")

    return {
        "download_token": token,
        "stats": result["stats"],
        "conflicts": result["conflicts"],
        "flagged": result["flagged"],
        "rta_not_in_hubspot_preview": result["rta_not_in_hubspot_preview"],
    }


@router.get("/download/{token}")
def download(token: str, current_user: User = Depends(get_current_user)):
    entry = cache_fetch(token)
    if entry is None:
        raise HTTPException(status_code=410, detail="Download expired or not found. Re-run the match.")
    excel_bytes, filename = entry
    return Response(
        content=excel_bytes,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f'attachment; filename="{filename}"'},
    )
