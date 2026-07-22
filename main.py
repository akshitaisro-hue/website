import os
import shutil
import pandas as pd
from fastapi import FastAPI, HTTPException, UploadFile, File
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel

import db
from pipeline.tctm_wrapper import run_pipeline
from pipeline.gen_summary_wrapper import run_gen_summary

app = FastAPI()
app.add_middleware(
    CORSMiddleware,
    allow_origins=['*'],
    allow_credentials = True,
    allow_methods = ['*'],
    allow_headers = ['*'],
)
DATA_DIR = "data"

# in-memory cache of last sync/check result per p_id, used by sync/confirm
_last_check_cache: dict[int, dict] = {}


def project_paths(p_id):
    base = os.path.join(DATA_DIR, str(p_id))
    return {"raw": os.path.join(base, "raw"), "output": os.path.join(base, "output")}


def require_project(p_id):
    proj = db.get_project_by_id(p_id)
    if not proj:
        raise HTTPException(status_code=404, detail="Project not found")
    return proj


# ---------------- 1. PROJECT MANAGEMENT ----------------

class ProjectCreate(BaseModel):
    p_id: str | None = None  # optional custom id, else DB auto-assigns
    name: str


def _enrich(proj: dict) -> dict:
    """adds files list + created_at (from raw folder), matching frontend Project shape"""
    paths = project_paths(proj["p_id"])
    files = []
    created_at = None
    if os.path.isdir(paths["raw"]):
        files = [f for f in os.listdir(paths["raw"]) if f.lower().endswith((".xlsx", ".xls"))]
        created_at = pd.Timestamp(os.path.getctime(paths["raw"]), unit="s").isoformat()
    return {
        "p_id": str(proj["p_id"]),
        "name": proj["project_name"],
        "created_at": created_at or pd.Timestamp.now().isoformat(),
        "files": files,
    }


@app.get("/api/projects")
def list_projects():
    return [_enrich(p) for p in db.list_projects()]


@app.post("/api/projects", status_code=201)
def create_project(payload: ProjectCreate):
    result = db.create_project(payload.name)
    if result is None:
        raise HTTPException(status_code=409, detail=f"Project '{payload.name}' already exists")
    return _enrich(result)


# ---------------- 2. FOLDER UPLOAD ----------------

@app.post("/api/projects/{p_id}/upload")
def upload_files(p_id: int, files: list[UploadFile] = File(...)):
    require_project(p_id)
    paths = project_paths(p_id)
    os.makedirs(paths["raw"], exist_ok=True)

    for f in files:
        if not f.filename.lower().endswith((".xlsx", ".xls")):
            continue  # server-side filter, mirrors client-side folder filter
        safe_name = os.path.basename(f.filename.replace("\\", "/"))  # strip subfolder path (webkitdirectory sends "Folder/file.xlsx")
        dest = os.path.join(paths["raw"], safe_name)
        with open(dest, "wb") as out:
            shutil.copyfileobj(f.file, out)

    all_files = [f for f in os.listdir(paths["raw"]) if f.lower().endswith((".xlsx", ".xls"))]
    return {"files": all_files}


# ---------------- 3. CORE ENGINE OPERATIONS ----------------

@app.post("/api/projects/{p_id}/run")
def run(p_id: int):
    require_project(p_id)
    paths = project_paths(p_id)
    if not os.path.isdir(paths["raw"]) or not os.listdir(paths["raw"]):
        raise HTTPException(status_code=400, detail="No uploaded files — upload first")

    result = run_pipeline(paths["raw"], paths["output"])
    return {
        "message": "TCTM processing complete.",
        "master_excel": f"/api/projects/{p_id}/master-excel",
        "pdf_report": f"/api/projects/{p_id}/report.pdf",
        "tc_count": result.get("tc_count", 0),
        "tm_count": result.get("tm_count", 0),
        "card_baseline_note" : result.get(card_baseline_note"),
    }


@app.get("/api/projects/{p_id}/report.pdf")
def get_report(p_id: int):
    pdf_path = os.path.join(project_paths(p_id)["output"], "report.pdf")
    if not os.path.isfile(pdf_path):
        raise HTTPException(status_code=404, detail="Report not generated — run pipeline first")
    return FileResponse(pdf_path, media_type="application/pdf", filename="report.pdf")


@app.get("/api/projects/{p_id}/master-excel")
def get_master_excel(p_id: int):
    xlsx_path = os.path.join(project_paths(p_id)["output"], "master_summary_tctm.xlsx")
    if not os.path.isfile(xlsx_path):
        raise HTTPException(status_code=404, detail="master_summary_tctm.xlsx not found — run pipeline first")
    return FileResponse(
        xlsx_path,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename="master_summary_tctm.xlsx",
    )


# ---------------- 4. STATE PARITY & DB SYNC ----------------

@app.post("/api/projects/{p_id}/sync")
def sync(p_id: int):
    """runs gen_summary.py -> all_tc.xlsx, all_tm.xlsx"""
    paths = project_paths(p_id)
    master_path = os.path.join(paths["output"], "master_summary_tctm.xlsx")
    if not os.path.isfile(master_path):
        raise HTTPException(status_code=400, detail="master_summary_tctm.xlsx not found — run pipeline first")
    result = run_gen_summary(master_path, paths["output"], p_id)
    return {"status": "done", **result}


@app.post("/api/projects/{p_id}/sync/check")
def sync_check(p_id: int):
    paths = project_paths(p_id)
    tc_path = os.path.join(paths["output"], "all_tc.xlsx")
    tm_path = os.path.join(paths["output"], "all_tm.xlsx")
    if not os.path.isfile(tc_path) or not os.path.isfile(tm_path):
        raise HTTPException(status_code=400, detail="all_tc.xlsx/all_tm.xlsx not found — run /sync first")

    tc_df = pd.read_excel(tc_path)
    tm_df = pd.read_excel(tm_path)
    tc_rows = tc_df.to_dict(orient="records")
    tm_rows = tm_df.to_dict(orient="records")
    for r in tc_rows:
        r["match_key"] = db.build_match_key(r, "all_tc")
    for r in tm_rows:
        r["match_key"] = db.build_match_key(r, "all_tm")

    result = db.check_sync(p_id, tc_rows, tm_rows)
    _last_check_cache[p_id] = result
    return result


class SyncConfirm(BaseModel):
    action: str  # "insert_all" | "insert_new_only"


@app.post("/api/projects/{p_id}/sync/confirm")
def sync_confirm(p_id: int, payload: SyncConfirm):
    cached = _last_check_cache.get(p_id)
    if not cached:
        raise HTTPException(status_code=400, detail="Run sync/check first")

    tc_insert = cached["new_rows"]["tc"]
    tm_insert = cached["new_rows"]["tm"]
    tc_update = cached.get("updated_rows",{}).get("tc",[])
    tm_update = cached.get("updated_rows",{}).get("tm",[])

    inserted = {"tc": 0, "tm": 0}
    
    if tc_insert:
        cols = [c for c in tc_insert[0].keys() if c not in ("match_key", "p_id","tc_id")]
        inserted["tc"] = db.insert_records(p_id, "all_tc", tc_insert, cols)
    if tm_insert:
        cols = [c for c in tm_insert[0].keys() if c not in ("match_key", "p_id","tm_id)]
        inserted["tm"] = db.insert_records(p_id, "all_tm", tm_insert, cols)
    if tc_update:
        cols = [c for c in tc_update[0].keys() if c not in ("match_key", "p_id","tc_id")]
        updated["tc"] = db.update_records(p_id, "all_tc", tc_update, cols)
    if tm_update:
        cols = [c for c in tm_update[0].keys() if c not in ("match_key", "p_id","tm_id)]
        updated["tm"] = db.update_records(p_id, "all_tm", tm_update, cols)

    _last_check_cache.pop(p_id, None)
    msg = f"Inserted {inserted['tc']+inserted['tm']} new rows, updated and {updated['tc']+updated['tm']} changed rows "
    return {"status": "inserted", "message": msg,"inserted":inserted,"updated" : updated}


app.mount("/", StaticFiles(directory="static", html=True), name="static")


if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000, log_config = None)
