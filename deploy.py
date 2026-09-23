#!/usr/bin/env python3
"""
Build MethodTech Excel add-in for Amplify (production .env), zip dist/,
upload the zip to Amplify, start the deployment, then delete local zips.

Usage:
  python deploy.py

Requires:
  boto3 + AWS credentials with amplify:CreateDeployment / StartDeployment
"""

from __future__ import annotations

import json
import os
import subprocess
import sys
import zipfile
from datetime import datetime
from pathlib import Path
from urllib.error import HTTPError, URLError

import boto3
import urllib.request

ROOT = Path(__file__).resolve().parent
ENV_FILE = ROOT / ".env"
ENV_LOCAL = ROOT / ".env.local"
DIST = ROOT / "dist"
CATALOG_META = ROOT / ".methodtech" / "catalog-functions.json"
REQUIRED_DIST_FILES = (
    "taskpane.html",
    "functions.js",
    "functions.json",
    "polyfill.js",
)

# Amplify (manual zip deploy)
AMPLIFY_REGION = "ap-south-1"
AMPLIFY_APP_ID = "d3jl8gkkjv76iz"
AMPLIFY_BRANCH = "staging"
CATALOG_PATH = "/api/microsoft/excel/catalog/"


def load_dotenv(path: Path) -> dict[str, str]:
    env: dict[str, str] = {}
    if not path.exists():
        return env
    for raw in path.read_text(encoding="utf-8").splitlines():
        line = raw.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, value = line.split("=", 1)
        key = key.strip()
        value = value.strip().strip('"').strip("'")
        env[key] = value
    return env


def confirm_production_env() -> dict[str, str]:
    print("=== 1) Confirm production .env ===")
    if not ENV_FILE.exists():
        raise SystemExit(f"ERROR: Missing {ENV_FILE}. Create it for production builds.")

    if ENV_LOCAL.exists():
        print(f"NOTE: {ENV_LOCAL.name} exists but is ignored for `npm run build` (production).")

    env = load_dotenv(ENV_FILE)
    django = env.get("DJANGO_BASE_URL", "").strip()
    scope = env.get("MICROSOFT_API_SCOPE", "").strip()
    app_id = env.get("MICROSOFT_APP_ID", "").strip()

    if not django:
        raise SystemExit("ERROR: DJANGO_BASE_URL is empty in .env")
    if not app_id:
        raise SystemExit("ERROR: MICROSOFT_APP_ID is empty in .env")
    if not scope:
        raise SystemExit("ERROR: MICROSOFT_API_SCOPE is empty in .env")
    if "localhost" in scope.lower():
        raise SystemExit(
            "ERROR: MICROSOFT_API_SCOPE in .env points at localhost. "
            "Use Amplify Application ID URI for production."
        )

    print(f"  DJANGO_BASE_URL     = {django}")
    print(f"  MICROSOFT_APP_ID    = {app_id}")
    print(f"  MICROSOFT_API_SCOPE = {scope}")
    print("  OK: using .env (production), not .env.local")
    return env


def catalog_param_to_excel(param: dict) -> dict:
    """Match src/catalog/registerFunctions.js catalogParamToExcel."""
    excel_param: dict = {
        "name": param.get("name"),
        "description": param.get("description") or param.get("name"),
        "optional": param.get("required") is False,
    }
    ptype = param.get("type")
    if ptype == "number":
        excel_param["type"] = "number"
    elif ptype == "boolean":
        excel_param["type"] = "boolean"
    elif ptype == "string":
        excel_param["type"] = "string"
    else:
        # array / any / default — matrix so Excel can pass ranges
        excel_param["type"] = "any"
        excel_param["dimensionality"] = "matrix"
    return excel_param


def catalog_item_to_excel_metadata(item: dict) -> dict:
    """Match src/catalog/registerFunctions.js catalogItemToExcelMetadata."""
    id_ = str(item.get("id") or item.get("name") or "").upper()
    name = str(item.get("name") or item.get("id") or "").upper()
    result = item.get("result") or {}
    rtype = result.get("type")
    if rtype == "array":
        result_type = "any"
    elif rtype == "number":
        result_type = "number"
    elif rtype == "string":
        result_type = "string"
    else:
        result_type = "any"

    meta: dict = {
        "id": id_,
        "name": name,
        "description": item.get("description") or name,
        "parameters": [catalog_param_to_excel(p) for p in (item.get("parameters") or [])],
        "result": {"type": result_type},
    }
    if rtype in ("array", "any") or result_type == "any":
        meta["result"] = {"type": "any", "dimensionality": "matrix"}
    return meta


def fetch_catalog_list(django_base_url: str) -> list[dict]:
    base = django_base_url.rstrip("/")
    url = f"{base}{CATALOG_PATH}"
    req = urllib.request.Request(url, headers={"Accept": "application/json"})
    try:
        with urllib.request.urlopen(req, timeout=60) as resp:
            data = json.loads(resp.read().decode("utf-8"))
    except HTTPError as error:
        raise SystemExit(f"ERROR: catalog fetch failed HTTP {error.code}: {url}") from error
    except URLError as error:
        raise SystemExit(f"ERROR: catalog fetch failed: {error.reason} ({url})") from error
    except json.JSONDecodeError as error:
        raise SystemExit(f"ERROR: catalog response was not JSON: {error}") from error

    if isinstance(data, list):
        return data
    if isinstance(data, dict):
        for key in ("functions", "results"):
            if isinstance(data.get(key), list):
                return data[key]
    raise SystemExit("ERROR: catalog response was not a function list.")


def ensure_catalog_metadata(django_base_url: str) -> int:
    """
    Fetch catalog from Django and write .methodtech/catalog-functions.json so
    webpack merges Excel IntelliSense metadata into dist/functions.json.
    """
    print("\n=== 1b) Seed catalog metadata for functions.json ===")
    catalog = fetch_catalog_list(django_base_url)
    excel_fns = [
        catalog_item_to_excel_metadata(item)
        for item in catalog
        if item and (item.get("id") or item.get("name"))
    ]
    if not excel_fns:
        raise SystemExit(
            "ERROR: catalog API returned 0 functions — cannot build IntelliSense metadata.\n"
            f"  Check {django_base_url.rstrip('/')}{CATALOG_PATH}"
        )

    CATALOG_META.parent.mkdir(parents=True, exist_ok=True)
    CATALOG_META.write_text(json.dumps(excel_fns, indent=2) + "\n", encoding="utf-8")
    print(f"  Fetched {len(catalog)} catalog function(s)")
    print(f"  Wrote {CATALOG_META.relative_to(ROOT)} ({len(excel_fns)} Excel metadata entries)")
    return len(excel_fns)


def run_production_build() -> None:
    print("\n=== 2) npm run build (production) ===")
    cmd = ["npm", "run", "build"]
    if os.name == "nt":
        cmd = ["npm.cmd", "run", "build"]

    proc = subprocess.run(
        cmd,
        cwd=str(ROOT),
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
    )
    combined = (proc.stdout or "") + "\n" + (proc.stderr or "")
    print(combined)

    if proc.returncode != 0:
        raise SystemExit(f"ERROR: npm run build failed with exit code {proc.returncode}")

    if "Loaded env from .env (production build)" not in combined:
        print(
            "WARNING: Did not see 'Loaded env from .env (production build)' in build log. "
            "Confirm webpack is using production mode."
        )
    else:
        print("  OK: build used .env (production)")


def validate_dist() -> None:
    print("\n=== 3) Validate dist/ ===")
    if not DIST.is_dir():
        raise SystemExit(f"ERROR: Missing dist folder at {DIST}")

    for name in REQUIRED_DIST_FILES:
        path = DIST / name
        if not path.exists():
            raise SystemExit(f"ERROR: Missing required file dist/{name}")
        if name == "functions.json" and path.stat().st_size == 0:
            raise SystemExit("ERROR: dist/functions.json is empty (0 bytes)")

    functions_path = DIST / "functions.json"
    try:
        meta = json.loads(functions_path.read_text(encoding="utf-8"))
        count = len(meta.get("functions") or [])
    except Exception as error:
        raise SystemExit(f"ERROR: dist/functions.json is invalid JSON: {error}") from error

    if count == 0:
        raise SystemExit(
            "ERROR: dist/functions.json has 0 functions — Excel will show no =MTECH. suggestions.\n"
            "  deploy.py should have seeded .methodtech/catalog-functions.json from the catalog API.\n"
            "  Check DJANGO_BASE_URL and /api/microsoft/excel/catalog/."
        )

    print(f"  OK: dist has required files ({DIST})")
    print(f"  OK: functions.json has {count} function(s) for IntelliSense")


def zip_dist() -> Path:
    print("\n=== 4) Zip dist contents (forward-slash paths) ===")
    stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    out = ROOT / f"Excel-Methodtech-amplify-{stamp}.zip"

    files = [p for p in DIST.rglob("*") if p.is_file()]
    if not files:
        raise SystemExit("ERROR: dist/ has no files to zip")

    with zipfile.ZipFile(out, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        for path in files:
            arcname = path.relative_to(DIST).as_posix()
            zf.write(path, arcname)

    print(f"  Created: {out.name}")
    print(f"  Entries: {len(files)}")
    return out


def amplify_upload_and_deploy(zip_path: Path) -> dict:
    """
    Manual Amplify zip deploy:
      create_deployment -> PUT zip to zipUploadUrl -> start_deployment
    """
    print("\n=== 5) Upload zip to Amplify + start deployment ===")
    amplify = boto3.client("amplify", region_name=AMPLIFY_REGION)

    created = amplify.create_deployment(
        appId=AMPLIFY_APP_ID,
        branchName=AMPLIFY_BRANCH,
    )
    job_id = created["jobId"]
    upload_url = created["zipUploadUrl"]
    print(f"  create_deployment jobId = {job_id}")

    zip_bytes = zip_path.read_bytes()
    req = urllib.request.Request(
        upload_url,
        data=zip_bytes,
        method="PUT",
        headers={"Content-Type": "application/zip"},
    )
    with urllib.request.urlopen(req, timeout=300) as resp:
        status = getattr(resp, "status", None) or resp.getcode()
        if status not in (200, 201):
            raise SystemExit(f"ERROR: zip upload failed HTTP {status}")
    print(f"  Uploaded {zip_path.name} ({len(zip_bytes)} bytes)")

    started = amplify.start_deployment(
        appId=AMPLIFY_APP_ID,
        branchName=AMPLIFY_BRANCH,
        jobId=job_id,
    )
    summary = started.get("jobSummary") or started
    print(f"  start_deployment: {summary}")
    return summary if isinstance(summary, dict) else {"jobSummary": summary}


def delete_local_zips(zip_path: Path) -> None:
    print("\n=== 6) Delete local zip(s) ===")
    candidates = {zip_path, ROOT / "Excel-Methodtech-amplify.zip"}
    for path in sorted(candidates):
        if path.exists():
            path.unlink()
            print(f"  Deleted: {path.name}")
        else:
            print(f"  Skip (missing): {path.name}")


def main() -> int:
    os.chdir(ROOT)
    print(f"MethodTech deploy — {ROOT}\n")
    env = confirm_production_env()
    ensure_catalog_metadata(env["DJANGO_BASE_URL"].strip())
    run_production_build()
    validate_dist()
    zip_path = zip_dist()
    try:
        amplify_upload_and_deploy(zip_path)
    finally:
        delete_local_zips(zip_path)

    print("\n=== Done ===")
    print(f"Amplify app {AMPLIFY_APP_ID} branch {AMPLIFY_BRANCH} deploy started.")
    print("Verify:")
    print("  https://staging.d3jl8gkkjv76iz.amplifyapp.com/taskpane.html")
    print("  https://staging.d3jl8gkkjv76iz.amplifyapp.com/functions.json")
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except KeyboardInterrupt:
        print("\nCancelled.")
        raise SystemExit(130)
