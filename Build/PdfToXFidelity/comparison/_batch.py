"""Shared batch runner: converts each job in a child process so one hang or crash cannot stop the batch."""
from __future__ import annotations

import json
import os
import subprocess
import sys
import time


def run_batch(script: str, jobs_path: str, out_dir: str, ext: str, timeout: int = 300) -> None:
    os.makedirs(out_dir, exist_ok=True)
    with open(jobs_path, encoding="utf-8") as f:
        jobs = json.load(f)
    log = []
    write_log(out_dir, log)
    for job in jobs:
        target = os.path.join(out_dir, job["id"] + "." + ext)
        if os.path.exists(target):
            os.remove(target)
        start = time.perf_counter()
        try:
            proc = subprocess.run([sys.executable, "-I", script, "--one", job["pdf"], target], capture_output=True, timeout=timeout)
            ok = proc.returncode == 0 and os.path.exists(target)
            status = "ok" if ok else "error: " + proc.stderr.decode("utf-8", "replace")[-300:]
        except subprocess.TimeoutExpired:
            status = "timeout"
        log.append({"id": job["id"], "status": status, "seconds": round(time.perf_counter() - start, 2)})
        print(job["id"], status[:80], flush=True)
        write_log(out_dir, log)


def write_log(out_dir: str, log: list) -> None:
    """Rewrites _log.json after every job, so an interrupted batch still records each finished outcome."""
    path = os.path.join(out_dir, "_log.json")
    with open(path + ".tmp", "w", encoding="utf-8") as f:
        json.dump(log, f, indent=1)
    os.replace(path + ".tmp", path)
