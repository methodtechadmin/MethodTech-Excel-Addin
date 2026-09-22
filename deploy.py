#!/usr/bin/env python3
"""
Build MethodTech Excel add-in for Amplify (production .env) and zip dist/.

Usage:
  python deploy.py

Output:
  Excel-Methodtech-amplify.zip  (dist contents at zip root, forward-slash paths)
"""

from __future__ import annotations

import os
import subprocess
import sys
import zipfile
from datetime import datetime
from pathlib import Path

ROOT = Path(__file__).resolve().parent
ENV_FILE = ROOT / ".env"
ENV_LOCAL = ROOT / ".env.local"
DIST = ROOT / "dist"
REQUIRED_DIST_FILES = (
    "taskpane.html",
    "functions.js",
    "functions.json",
    "polyfill.js",
)


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


def confirm_production_env() -> None:
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


def run_production_build() -> None:
    print("\n=== 2) npm run build (production) ===")
    cmd = ["npm", "run", "build"]
    # On Windows, npm is often npm.cmd
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

    print(f"  OK: dist has required files ({DIST})")


def zip_dist() -> Path:
    print("\n=== 4) Zip dist contents (forward-slash paths) ===")
    stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    out = ROOT / f"Excel-Methodtech-amplify-{stamp}.zip"
    latest = ROOT / "Excel-Methodtech-amplify.zip"

    files = [p for p in DIST.rglob("*") if p.is_file()]
    if not files:
        raise SystemExit("ERROR: dist/ has no files to zip")

    with zipfile.ZipFile(out, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        for path in files:
            # Zip root = dist contents (not a nested dist/ folder)
            arcname = path.relative_to(DIST).as_posix()
            zf.write(path, arcname)

    # Also write/replace a stable name for easy Amplify upload
    if latest.exists():
        latest.unlink()
    latest.write_bytes(out.read_bytes())

    print(f"  Created: {out.name}")
    print(f"  Also:    {latest.name}")
    print(f"  Entries: {len(files)}")
    return latest


def main() -> int:
    os.chdir(ROOT)
    print(f"MethodTech deploy — {ROOT}\n")
    confirm_production_env()
    run_production_build()
    validate_dist()
    zip_path = zip_dist()
    print("\n=== Done ===")
    print(f"Upload this zip to Amplify: {zip_path}")
    print("Then verify:")
    print("  …/taskpane.html")
    print("  …/functions.json  (must not be empty)")
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except KeyboardInterrupt:
        print("\nCancelled.")
        raise SystemExit(130)
