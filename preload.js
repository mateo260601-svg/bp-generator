import os, uuid, sys, webbrowser, threading
from pathlib import Path
from datetime import datetime
from fastapi import FastAPI, HTTPException, Header, Request
from fastapi.responses import FileResponse, HTMLResponse
from fastapi.staticfiles import StaticFiles
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
from typing import Optional, List, Dict, Any

sys.path.insert(0, str(Path(__file__).parent))
from build import build_workbook

import os
PORT = int(os.environ.get('PORT', 8765))

app = FastAPI(title="BP Generator", version="1.0.0")
app.add_middleware(CORSMiddleware, allow_origins=["*"], allow_methods=["*"], allow_headers=["*"])

BASE     = Path(__file__).parent
STATIC   = BASE.parent / "static"
OUTPUTS  = BASE / "outputs"
OUTPUTS.mkdir(exist_ok=True)

VALID_LICENSES = {
    "JRC-MATEO-2025":  {"user": "Mateo Girard — JRC",   "expires": "2026-12-31"},
    "JRC-JAUFRE-2025": {"user": "Jaufré Rouanet — JRC", "expires": "2026-12-31"},
    "JRC-OZAN-2025":   {"user": "Ozan OK — JRC",        "expires": "2026-12-31"},
    "JRC-DEV-LOCAL":   {"user": "Developer",             "expires": "2099-12-31"},
    # Ajouter clients ici :
    # "JRC-CLIENT-001": {"user": "Nom Client", "expires": "2026-06-30"},
}

def check_license(key: str) -> dict:
    if key not in VALID_LICENSES:
        raise HTTPException(status_code=401, detail="Code d'accès invalide.")
    lic = VALID_LICENSES[key]
    if datetime.now().strftime("%Y-%m-%d") > lic["expires"]:
        raise HTTPException(status_code=403, detail="Licence expirée.")
    return lic

class LoginRequest(BaseModel):
    license_key: str

class BPConfig(BaseModel):
    company_name:    str        = "Business Plan"
    business_type:   str        = "industrial"
    currency:        str        = "USD"
    n_years:         int        = 7
    fy_start_month:  int        = 1
    modules:         List[str]  = []
    debt:            Dict[str, Any] = {}
    mechanics:       List[str]  = []
    base_rate_index: str        = "SOFR"
    output_format:   str        = "annual"

@app.post("/api/login")
def login(req: LoginRequest):
    lic = check_license(req.license_key)
    return {"ok": True, "user": lic["user"], "expires": lic["expires"]}

@app.post("/api/generate")
def generate(cfg: BPConfig, x_license: str = Header(...)):
    check_license(x_license)
    n = cfg.n_years

    DEBT_DEFAULTS = {
        "tla": {"tenor_yrs": 7,  "margin_pct": 0.0325, "amort_type": "quarterly"},
        "tlb": {"tenor_yrs": 7,  "margin_pct": 0.0400, "amort_type": "bullet"},
        "ss":  {"tenor_yrs": 3,  "margin_pct": 0.0450, "amort_type": "quarterly"},
        "rcf": {"tenor_yrs": 5,  "margin_pct": 0.0325, "amort_type": "revolving"},
        "mur": {"tenor_yrs": 7,  "margin_pct": 0.0340, "amort_type": "quarterly"},
        "mez": {"tenor_yrs": 7,  "margin_pct": 0.0700, "amort_type": "bullet"},
        "pik": {"tenor_yrs": 7,  "margin_pct": 0.0000, "amort_type": "bullet", "pik_margin_pct": 0.08},
        "shl": {"tenor_yrs": 7,  "margin_pct": 0.0000, "amort_type": "pik_only", "pik_margin_pct": 0.08},
        "uni": {"tenor_yrs": 7,  "margin_pct": 0.0600, "amort_type": "amortising"},
        "hyb": {"tenor_yrs": 7,  "margin_pct": 0.0750, "amort_type": "bullet"},
        "dip": {"tenor_yrs": 2,  "margin_pct": 0.0500, "amort_type": "quarterly"},
        "d2e": {"tenor_yrs": 0,  "margin_pct": 0.0000, "amort_type": "n/a"},
    }
    debt_map = {}
    for key, amt in cfg.debt.items():
        d = DEBT_DEFAULTS.get(key, {"tenor_yrs": 5, "margin_pct": 0.04, "amort_type": "bullet"})
        debt_map[key] = {
            "active": 1, "amount": int(amt), "currency": cfg.currency,
            "tenor_yrs": d["tenor_yrs"], "margin_pct": d["margin_pct"],
            "pik_margin_pct": d.get("pik_margin_pct", 0.0),
        }

    vol = [180000 + i*5000 for i in range(n)]
    pmt = [1050   + i*15   for i in range(n)]

    config = {
        "company_name":     cfg.company_name,
        "business_type":    cfg.business_type,
        "currency":         cfg.currency,
        "units":            "k",
        "fy_start_month":   cfg.fy_start_month,
        "proj_start_year":  datetime.now().year,
        "n_years":          n,
        "actuals_end_year": datetime.now().year - 1,
        "actuals_end_month":cfg.fy_start_month - 1 or 12,
        "freq":             "Annual",
        "modules": {k: (1 if k in cfg.modules else 0) for k in
                    ["debt","tax","scenarios","returns","valuation","consol"]},
        "revenue": {
            "capacity_mt":   [v*1.2 for v in vol],
            "base_volume":   vol,
            "volume_growth": [0.0]+[round((vol[i]-vol[i-1])/vol[i-1],3) for i in range(1,n)],
            "price_per_mt":  pmt, "freight_mt": [35]*n, "commission_pct": [0.01]*n,
        },
        "costs": {
            "direct_mat_mt": [520]*n, "direct_mat_gr": [0.02]*n,
            "utilities_mt":  [45]*n,  "packing_mt":    [18]*n,
            "var_opex_other_mt": [12]*n,
            "staff_prod":  [8500+i*170 for i in range(n)],
            "staff_sga":   [3200+i*65  for i in range(n)],
            "headcount_gr":[0.02]*n, "maintenance":[2400]*n,
            "insurance":[800]*n,"rent":[600]*n,"it":[400]*n,
            "prof_fees":[350]*n,"other_sga":[500]*n,"restr":[0]*n,
        },
        "nwc":   {"dso":[50]*n,"dio":[35]*n,"dpo":[55]*n,"oca_pct":[0.01]*n,"ocl_pct":[0.02]*n},
        "capex": {"maint":[3500]*n,"expan":([8000,5000,3000]+[1500]*(n-3) if n>3 else [5000]*n),
                  "opening_ppe":95000,"useful_life":20,"accum_dep_open":28000},
        "tax":   {"rate":[0.25]*n,"loss_cf_open":0,"dt_rate":[0.25]*n},
        "macro": {"cpi":[0.025]*n,"fx_usd_eur":[1.09]*n,"fx_usd_gbp":[1.27]*n,"fx_usd_aed":[3.67]*n},
        "debt":  debt_map,
        "debt_mechanics": {
            "base_rate_index":   cfg.base_rate_index,
            "base_rate_pct":     0.045, "base_rate_floor": 0.0,
            "upfront_fee_pct":   0.015, "commit_fee_pct":  0.005, "oid_pct": 0.0,
            "cash_sweep_pct":    0.75,
            "cash_sweep_active": 1 if "sweep"   in cfg.mechanics else 0,
            "margin_ratchet":    1 if "ratchet" in cfg.mechanics else 0,
            "covenant_tracking": 1 if "cov"     in cfg.mechanics else 0,
            "lev_covenant": 5.0, "icr_covenant": 2.0, "dscr_covenant": 0.0,
            "fx_hedge":   1 if "fx"  in cfg.mechanics else 0,
            "irs_active": 1 if "irs" in cfg.mechanics else 0,
        },
    }

    safe = cfg.company_name[:20].replace(" ","_").replace("/","_")
    fname = f"BP_{safe}_{uuid.uuid4().hex[:6]}.xlsx"
    outpath = str(OUTPUTS / fname)

    try:
        build_workbook(config=config, output_path=outpath)
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

    return {"ok": True, "filename": fname}

@app.get("/api/download/{filename}")
def download(filename: str, x_license: str = Header(...)):
    check_license(x_license)
    path = OUTPUTS / filename
    if not path.exists():
        raise HTTPException(status_code=404, detail="Fichier introuvable.")
    return FileResponse(
        path=str(path), filename=filename,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

@app.get("/api/health")
def health():
    return {"status": "ok"}

# Serve static files if directory exists
if STATIC.exists():
    app.mount("/", StaticFiles(directory=str(STATIC), html=True), name="static")


# ── Import module ──────────────────────────────────────────────────────────────
import tempfile, shutil
from fastapi import UploadFile, File
from extractor import extract_financials, map_to_bp_actuals, build_projections_from_actuals

ALLOWED_EXTENSIONS = {".xlsx", ".xls", ".xlsm", ".csv", ".txt", ".tsv", ".pdf"}

@app.post("/api/import")
async def import_financials(
    file: UploadFile = File(...),
    x_license: str = Header(...)
):
    check_license(x_license)

    ext = Path(file.filename).suffix.lower()
    if ext not in ALLOWED_EXTENSIONS:
        raise HTTPException(400, f"Format non supporté : {ext}. Acceptés : Excel, CSV, PDF")

    # Save upload to temp file
    with tempfile.NamedTemporaryFile(suffix=ext, delete=False) as tmp:
        shutil.copyfileobj(file.file, tmp)
        tmp_path = tmp.name

    try:
        raw     = extract_financials(tmp_path)
        mapped  = map_to_bp_actuals(raw)
        proj    = build_projections_from_actuals(mapped)
    except Exception as e:
        raise HTTPException(500, f"Erreur extraction : {str(e)}")
    finally:
        Path(tmp_path).unlink(missing_ok=True)

    if not raw:
        raise HTTPException(422, "Aucune donnée financière reconnue dans ce fichier. Vérifie que les lignes contiennent des labels standards (Revenue, EBITDA, Net Income, etc.)")

    return {
        "ok": True,
        "data_quality":     proj["data_quality"],
        "hist_years":       proj["hist_years"],
        "last_actuals":     proj["last_actuals"],
        "proj_assumptions": proj["proj_assumptions"],
        "actuals_overlay":  proj["actuals_overlay"],
    }

# ── Scenarios endpoint ─────────────────────────────────────────────────────────
from scenarios import build_scenario_assumptions, SCENARIO_DEFS

class ScenariosConfig(BaseModel):
    base_config:  Dict[str, Any]
    scenarios:    List[str] = ["low", "base", "best"]
    generate_separate: bool = True  # one Excel per scenario

@app.post("/api/generate-scenarios")
def generate_scenarios(req: ScenariosConfig, x_license: str = Header(...)):
    check_license(x_license)

    results = []
    for sk in req.scenarios:
        if sk not in SCENARIO_DEFS:
            continue

        sc_cfg = build_scenario_assumptions(req.base_config, sk)
        sdef   = SCENARIO_DEFS[sk]
        safe   = sc_cfg.get("company_name","BP")[:16].replace(" ","_")
        fname  = f"BP_{safe}_{sdef['label'].replace(' ','_')}_{uuid.uuid4().hex[:5]}.xlsx"
        outpath = str(OUTPUTS / fname)

        try:
            build_workbook(config=sc_cfg, output_path=outpath)
        except Exception as e:
            raise HTTPException(500, f"Erreur {sk}: {str(e)}")

        results.append({
            "scenario":    sk,
            "label":       sdef["label"],
            "filename":    fname,
            "description": sdef["description"],
        })

    return {"ok": True, "files": results}

# ── Preview endpoint ───────────────────────────────────────────────────────────
from preview import build_preview

@app.get("/api/preview/{filename}")
def preview_file(filename: str, x_license: str = Header(...)):
    check_license(x_license)
    path = OUTPUTS / filename
    if not path.exists():
        raise HTTPException(404, "Fichier introuvable.")
    try:
        data = build_preview(str(path))
    except Exception as e:
        raise HTTPException(500, f"Erreur aperçu : {str(e)}")
    return data
