from fastapi import APIRouter, Depends, File, Form, HTTPException, UploadFile
from fastapi.responses import Response

from ..dependencies import get_current_user
from ..models import User
from ..services.compare import run_compare
from ..upload_utils import peek_columns, read_to_dataframe
from ..download_cache import fetch as cache_fetch
from ..download_cache import store as cache_store

router = APIRouter(prefix="/api/compare", tags=["compare"])


@router.post("/peek")
async def peek(
    file: UploadFile = File(...),
    current_user: User = Depends(get_current_user),
):
    columns, sheet_names = await peek_columns(file)
    return {"columns": columns, "sheet_names": sheet_names}


@router.post("/run")
async def run(
    old_file: UploadFile = File(...),
    new_file: UploadFile = File(...),
    addr_no_col: str = Form(...),
    street_col: str = Form(...),
    locality_col: str = Form(...),
    pc_col: str = Form(...),
    status_col: str = Form(...),
    old_sheet: str | None = Form(None),
    new_sheet: str | None = Form(None),
    current_user: User = Depends(get_current_user),
):
    df_old = await read_to_dataframe(old_file, "Old RTA", sheet_name=old_sheet)
    df_new = await read_to_dataframe(new_file, "New RTA", sheet_name=new_sheet)

    required = (addr_no_col, street_col, locality_col, pc_col, status_col)
    missing_in_old = [c for c in required if c not in df_old.columns]
    missing_in_new = [c for c in required if c not in df_new.columns]
    if missing_in_old or missing_in_new:
        raise HTTPException(
            status_code=422,
            detail={
                "message": "One or more selected columns were not found.",
                "missing_in_old": missing_in_old,
                "missing_in_new": missing_in_new,
            },
        )

    result = run_compare(
        df_old,
        df_new,
        addr_no_col=addr_no_col,
        street_col=street_col,
        locality_col=locality_col,
        pc_col=pc_col,
        status_col=status_col,
    )

    token = cache_store(result["excel_bytes"], "rta_status_compare_output.xlsx")

    return {
        "download_token": token,
        "stats": result["stats"],
        "removed_preview": result["removed_preview"],
        "added_preview": result["added_preview"],
        "status_changed_preview": result["status_changed_preview"],
        "conflicts_preview": result["conflicts_preview"],
    }


@router.get("/download/{token}")
def download(token: str, current_user: User = Depends(get_current_user)):
    entry = cache_fetch(token)
    if entry is None:
        raise HTTPException(status_code=410, detail="Download expired or not found. Re-run the comparison.")
    excel_bytes, filename = entry
    return Response(
        content=excel_bytes,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f'attachment; filename="{filename}"'},
    )
