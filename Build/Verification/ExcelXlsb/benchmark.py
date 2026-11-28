#!/usr/bin/env python3
"""Run one paired BenchmarkDotNet XLSB comparison over the same master-generated workbook."""

import argparse
import hashlib
import json
import os
from pathlib import Path
import shutil
import subprocess
import sys

from compare import HERE, run


def driver(output, label):
    directory = output / f"driver-{label}"
    executable = directory / "bin/Release/net10.0/BenchmarkRevisions.dll"
    env = {**os.environ, "ExcelSourceRoot": str(output / f"source-{label}")}
    return executable, env


def measure(output):
    fixture = output / "benchmark.xlsb"
    executable, env = driver(output, "baseline")
    run(["dotnet", str(executable), "seed", str(fixture)], executable.parent, env=env)
    inputs_path = output / "inputs.json"
    inputs = json.loads(inputs_path.read_text())
    inputs["fixtureSha256"] = hashlib.sha256(fixture.read_bytes()).hexdigest()
    inputs["initialLoad"] = run(["uptime"], output, stdout=subprocess.PIPE).stdout.strip()
    inputs_path.write_text(json.dumps(inputs, indent=2) + "\n")

    # Alternate the starting revision on the second round. Each report keeps the raw
    # samples, allocations and runtime environment.
    for round_number, order in enumerate((("baseline", "candidate"), ("candidate", "baseline")), 1):
        for label in order:
            executable, env = driver(output, label)
            report = output / f"{round_number}-{label}"
            run(["dotnet", str(executable), "--inProcess", "--filter", "*XlsbMigrationBenchmarks*",
                 "--job", "short", "--artifacts", str(report), "--exporters", "json"], executable.parent,
                env={**env, "OFFICEIMO_XLSB_BENCH_INPUT": str(fixture)})
            reports = list((report / "results").glob("*-report*.json"))
            if len(reports) != 1:
                raise RuntimeError(f"Expected one benchmark report in {report}")
            results = json.loads(reports[0].read_text())["Benchmarks"]
            if len(results) != 3 or any(not result.get("Statistics") for result in results):
                raise RuntimeError(f"One or more workloads did not finish in {report}")


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("baseline", help="Baseline Git revision")
    parser.add_argument("candidate", help="Candidate Git revision")
    parser.add_argument("output", type=Path, help="New scratch directory for the measurements")
    parser.add_argument("--lock", type=Path, required=True, help="Shared benchmark lock path to acquire")
    parser.add_argument("--repository", type=Path, default=HERE.parents[2], help="Repository containing both revisions")
    locked = len(sys.argv) > 1 and sys.argv[1] == "--under-benchmark-lock"
    args = parser.parse_args(sys.argv[2:] if locked else sys.argv[1:])
    if not locked:
        run(["lockf", "-k", str(args.lock.expanduser().resolve()), sys.executable,
             str(Path(__file__).resolve()), "--under-benchmark-lock", *sys.argv[1:]], Path.cwd())
        return

    output = args.output.resolve()
    output.mkdir(parents=True)
    root = args.repository.resolve()
    checkouts = []
    revisions = {}

    try:
        for label, ref in (("baseline", args.baseline), ("candidate", args.candidate)):
            checkout = output / f"source-{label}"
            run(["git", "worktree", "add", "--detach", "--quiet", str(checkout), ref], root)
            checkouts.append(checkout)
            revision = run(["git", "rev-parse", "HEAD"], checkout, stdout=subprocess.PIPE).stdout.strip()
            directory = output / f"driver-{label}"
            directory.mkdir()
            for file in ("BenchmarkRevisions.cs", "BenchmarkRevisions.csproj"):
                shutil.copyfile(HERE / file, directory / file)
            _, env = driver(output, label)
            run(["dotnet", "build", "BenchmarkRevisions.csproj", "-c", "Release", "-f", "net10.0", "-m:2"],
                directory, env=env)
            revisions[label] = revision

        (output / "inputs.json").write_text(json.dumps({"revisions": revisions}, indent=2) + "\n")
        measure(output)
    finally:
        for checkout in reversed(checkouts):
            run(["git", "worktree", "remove", str(checkout)], root)


if __name__ == "__main__":
    main()
