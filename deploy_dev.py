#!/usr/bin/env python3
"""
Restart local MethodTech Excel add-in dev server.

1) npm stop  (office-addin-debugging stop)
2) Kill whatever is holding the dev port (default 3001)
3) npm start (sideload + webpack serve via manifest.localhost.xml)

Usage:
  python deploy_dev.py
"""

from __future__ import annotations

import json
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
DEFAULT_PORT = 3001
NPM = "npm.cmd" if sys.platform == "win32" else "npm"


def run(cmd: list[str], *, check: bool = False, capture: bool = False) -> subprocess.CompletedProcess:
    print(f"\n> {' '.join(cmd)}")
    return subprocess.run(
        cmd,
        cwd=ROOT,
        check=check,
        text=True,
        capture_output=capture,
        shell=False,
    )


def read_dev_port() -> int:
    pkg = ROOT / "package.json"
    try:
        data = json.loads(pkg.read_text(encoding="utf-8"))
        port = int(data.get("config", {}).get("dev_server_port", DEFAULT_PORT))
        return port if port > 0 else DEFAULT_PORT
    except Exception:
        return DEFAULT_PORT


def npm_stop() -> None:
    print("=== 1) npm stop ===")
    result = run([NPM, "run", "stop"])
    if result.returncode != 0:
        print("NOTE: npm stop returned non-zero (often OK if nothing was running).")


def pids_on_port(port: int) -> list[int]:
    """Find PIDs listening on TCP port (Windows)."""
    if sys.platform != "win32":
        result = run(["lsof", "-ti", f":{port}"], capture=True)
        if result.returncode != 0 or not result.stdout.strip():
            return []
        return [int(x) for x in result.stdout.split() if x.strip().isdigit()]

    result = run(["netstat", "-ano"], capture=True)
    if result.returncode != 0 or not result.stdout:
        return []

    pids: set[int] = set()
    needle = f":{port}"
    for line in result.stdout.splitlines():
        if needle not in line:
            continue
        # Prefer LISTENING rows
        if "LISTENING" not in line.upper() and "LISTEN" not in line.upper():
            continue
        parts = line.split()
        if not parts:
            continue
        pid_str = parts[-1]
        if pid_str.isdigit() and int(pid_str) > 0:
            pids.add(int(pid_str))
    return sorted(pids)


def kill_port(port: int) -> None:
    print(f"=== 2) Kill port {port} ===")
    pids = pids_on_port(port)
    if not pids:
        print(f"Port {port} is free.")
        return

    for pid in pids:
        print(f"Killing PID {pid} on port {port}...")
        if sys.platform == "win32":
            run(["taskkill", "/PID", str(pid), "/F", "/T"])
        else:
            run(["kill", "-9", str(pid)])

    leftover = pids_on_port(port)
    if leftover:
        print(f"WARNING: still in use by PIDs: {leftover}")
    else:
        print(f"Port {port} cleared.")


def npm_start() -> None:
    print("=== 3) npm start ===")
    print("Starting local add-in (Ctrl+C to stop)...")
    # Don't capture — stream logs; this is the long-running process.
    result = subprocess.run([NPM, "start"], cwd=ROOT)
    raise SystemExit(result.returncode)


def main() -> None:
    if not (ROOT / "package.json").exists():
        raise SystemExit(f"ERROR: package.json not found in {ROOT}")

    port = read_dev_port()
    print(f"MethodTech local restart (port {port})")
    npm_stop()
    kill_port(port)
    npm_start()


if __name__ == "__main__":
    main()
