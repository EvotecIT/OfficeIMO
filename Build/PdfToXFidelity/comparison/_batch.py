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
    for job in jobs:
        target = os.path.join(out_dir, job["id"] + "." + ext)
        start = time.perf_counter()
        try:
            proc = subprocess.run([sys.executable, "-I", script, "--one", job["pdf"], target], capture_output=True, timeout=timeout)
            ok = proc.returncode == 0 and os.path.exists(target)
            status = "ok" if ok else "error: " + proc.stderr.decode("utf-8", "replace")[-300:]
        except subprocess.TimeoutExpired:
            status = "timeout"
        log.append({"id": job["id"], "status": status, "seconds": round(time.perf_counter() - start, 2)})
        print(job["id"], status[:80], flush=True)
    with open(os.path.join(out_dir, "_log.json"), "w", encoding="utf-8") as f:
        json.dump(log, f, indent=1)
