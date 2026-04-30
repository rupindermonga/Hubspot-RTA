import pickle
from datetime import datetime

from fastapi import APIRouter, Depends, File, Form, HTTPException, Query, UploadFile
from fastapi.responses import Response
from sqlalchemy.orm import Session

from .. import df_cache
from ..database import get_db
from ..dependencies import get_current_user
from ..models import Dataset, User
from ..services.matcher import DEFAULT_ALIASES, run_match
from ..services.search import search_in_dataset
from ..upload_utils import peek_columns, read_to_dataframe

router = APIRouter(prefix="/api/matcher", tags=["matcher"])


# ── Aliases (read-only) ──────────────────────────────────────────────
@router.get("/aliases")
def aliases(current_user: User = Depends(get_current_user)):
    return {"aliases": [{"variant": v, "canonical": c} for v, c in DEFAULT_ALIASES]}


# ── Peek (column detection before run) ───────────────────────────────
@router.post("/peek")
async def peek(
    file: UploadFile = File(...),
    current_user: User = Depends(get_current_user),
):
    columns, sheet_names = await peek_columns(file)
    return {"columns": columns, "sheet_names": sheet_names}


# ── Run match: parse files, match, persist as Dataset ────────────────
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
    label: str | None = Form(None),
    db: Session = Depends(get_db),
    current_user: User = Depends(get_current_user),
):
    df_hub = await read_to_dataframe(hub_file, "Hubspot", sheet_name=hub_sheet)
    df_rta = await read_to_dataframe(rta_file, "RTA", sheet_name=rta_sheet)

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

    col_map = {
        "hub_street_col": hub_street_col,
        "hub_pc_col": hub_pc_col,
        "rta_addr_no_col": rta_addr_no_col,
        "rta_street_col": rta_street_col,
        "rta_locality_col": rta_locality_col,
        "rta_pc_col": rta_pc_col,
        "rta_status_col": rta_status_col,
        "enable_no_pc": enable_no_pc,
    }

    auto_label = label or datetime.utcnow().strftime("%Y-%m-%d %H:%M UTC")
    ds = Dataset(
        user_id=current_user.id,
        label=auto_label,
        col_map=col_map,
        stats=result["stats"],
        conflicts=result["conflicts"],
        flagged=result["flagged"],
        rta_not_in_hubspot_preview=result["rta_not_in_hubspot_preview"],
        hub_output_preview=result["hub_output_preview"],
        excel_bytes=result["excel_bytes"],
        df_hub_pickle=pickle.dumps(result["df_hub_keyed"]),
        df_rta_pickle=pickle.dumps(result["df_rta_keyed"]),
    )
    db.add(ds)
    db.commit()
    db.refresh(ds)

    # Cache deserialized DFs so the next /search hit doesn't pay the unpickle cost
    df_cache.put(ds.id, result["df_hub_keyed"], result["df_rta_keyed"])

    return {
        "dataset_id": ds.id,
        "created_at": ds.created_at.isoformat(),
        "label": ds.label,
        "stats": ds.stats,
        "conflicts": ds.conflicts,
        "flagged": ds.flagged,
        "rta_not_in_hubspot_preview": ds.rta_not_in_hubspot_preview,
        "hub_output_preview": ds.hub_output_preview,
    }


# ── List user's saved datasets ───────────────────────────────────────
@router.get("/datasets")
def list_datasets(
    db: Session = Depends(get_db),
    current_user: User = Depends(get_current_user),
):
    rows = (
        db.query(Dataset)
        .filter(Dataset.user_id == current_user.id)
        .order_by(Dataset.created_at.desc())
        .all()
    )
    return {
        "datasets": [
            {
                "id": d.id,
                "created_at": d.created_at.isoformat(),
                "label": d.label,
                "hubspot_total": (d.stats or {}).get("hubspot_total", 0),
                "hubspot_matched": (d.stats or {}).get("hubspot_matched", 0),
                "rta_total": (d.stats or {}).get("rta_total", 0),
                "rta_in_hubspot": (d.stats or {}).get("rta_in_hubspot", 0),
                "has_conflicts": bool(d.conflicts),
            }
            for d in rows
        ]
    }


# ── Fetch one dataset (full result for re-rendering the result panel) ─
@router.get("/datasets/{dataset_id}")
def get_dataset(
    dataset_id: int,
    db: Session = Depends(get_db),
    current_user: User = Depends(get_current_user),
):
    d = (
        db.query(Dataset)
        .filter(Dataset.id == dataset_id, Dataset.user_id == current_user.id)
        .first()
    )
    if not d:
        raise HTTPException(status_code=404, detail="Dataset not found")
    return {
        "dataset_id": d.id,
        "created_at": d.created_at.isoformat(),
        "label": d.label,
        "col_map": d.col_map,
        "stats": d.stats,
        "conflicts": d.conflicts,
        "flagged": d.flagged,
        "rta_not_in_hubspot_preview": d.rta_not_in_hubspot_preview,
        "hub_output_preview": d.hub_output_preview,
    }


@router.delete("/datasets/{dataset_id}")
def delete_dataset(
    dataset_id: int,
    db: Session = Depends(get_db),
    current_user: User = Depends(get_current_user),
):
    d = (
        db.query(Dataset)
        .filter(Dataset.id == dataset_id, Dataset.user_id == current_user.id)
        .first()
    )
    if not d:
        raise HTTPException(status_code=404, detail="Dataset not found")
    db.delete(d)
    db.commit()
    df_cache.evict(dataset_id)
    return {"deleted": dataset_id}


# ── Search across user's datasets ────────────────────────────────────
@router.get("/search")
def search(
    q: str = Query(..., min_length=1, max_length=200),
    postal: str = Query("", max_length=20),
    dataset_id: int | None = Query(None, description="If set, search only this dataset"),
    db: Session = Depends(get_db),
    current_user: User = Depends(get_current_user),
):
    base = db.query(Dataset).filter(Dataset.user_id == current_user.id)
    if dataset_id is not None:
        base = base.filter(Dataset.id == dataset_id)
    datasets = base.order_by(Dataset.created_at.desc()).all()

    results = []
    total_hub = 0
    total_rta = 0
    normalized_keys: list[str] = []
    for d in datasets:
        cached = df_cache.get(d.id)
        if cached is None:
            df_hub = pickle.loads(d.df_hub_pickle)
            df_rta = pickle.loads(d.df_rta_pickle)
            df_cache.put(d.id, df_hub, df_rta)
        else:
            df_hub = cached["df_hub"]
            df_rta = cached["df_rta"]
        hit = search_in_dataset(df_hub, df_rta, q, postal)
        if not normalized_keys:
            normalized_keys = hit["normalized_keys"]
        if hit["in_hubspot"] or hit["in_rta"]:
            results.append({
                "dataset_id": d.id,
                "created_at": d.created_at.isoformat(),
                "label": d.label,
                "in_hubspot": hit["in_hubspot"],
                "in_rta": hit["in_rta"],
            })
            total_hub += len(hit["in_hubspot"])
            total_rta += len(hit["in_rta"])

    if total_hub and total_rta:
        verdict = "both"
    elif total_hub:
        verdict = "hubspot_only"
    elif total_rta:
        verdict = "rta_only"
    else:
        verdict = "neither"

    return {
        "query": q,
        "postal": postal,
        "normalized_keys": normalized_keys,
        "verdict": verdict,
        "datasets": results,
        "summary": {
            "datasets_searched": len(datasets),
            "datasets_with_hits": len(results),
            "total_hubspot_hits": total_hub,
            "total_rta_hits": total_rta,
        },
    }


# ── Download Excel for a saved dataset ───────────────────────────────
@router.get("/download/{dataset_id}")
def download(
    dataset_id: int,
    db: Session = Depends(get_db),
    current_user: User = Depends(get_current_user),
):
    d = (
        db.query(Dataset)
        .filter(Dataset.id == dataset_id, Dataset.user_id == current_user.id)
        .first()
    )
    if not d:
        raise HTTPException(status_code=404, detail="Dataset not found")
    return Response(
        content=d.excel_bytes,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f'attachment; filename="hubspot_rta_dataset_{d.id}.xlsx"'},
    )
