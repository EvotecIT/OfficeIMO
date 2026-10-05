#!/usr/bin/env python3
"""Opt-in EPUBCheck and optional Ace evidence; does not install tools or certify accessibility."""
import argparse
import hashlib
import json
from pathlib import Path
import subprocess
import sys


def digest(path):
    value = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            value.update(chunk)
    return value.hexdigest()


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("epubs", nargs="+", type=Path)
    parser.add_argument("--epubcheck-jar", required=True, type=Path)
    parser.add_argument("--output", required=True, type=Path, help="New task-owned evidence directory")
    parser.add_argument("--java", default="java")
    parser.add_argument("--ace", type=Path, help="Explicit installed Ace CLI executable (optional)")
    parser.add_argument("--timeout", type=int, default=120, help="Seconds per validator invocation")
    args = parser.parse_args()
    jar = args.epubcheck_jar.resolve(strict=True)
    ace = args.ace.resolve(strict=True) if args.ace else None
    inputs = [path.resolve(strict=True) for path in args.epubs]
    if args.timeout < 1 or len(inputs) > 256:
        parser.error("Use a positive timeout and at most 256 publications per run")
    if not jar.is_file() or any(not path.is_file() or path.stat().st_size > 128 * 1024 * 1024 for path in inputs):
        parser.error("Inputs must be regular files; publications must not exceed 128 MiB")
    output = args.output.resolve()
    output.mkdir(parents=True, exist_ok=False)
    command = [args.java, "-jar", str(jar)]
    summary = {"epubcheckJarSha256": digest(jar), "publications": [],
               "accessibilityAssessment": "not-checked", "readingSystemPresentation": "not-checked"}
    failed = False
    try:
        version = subprocess.run(command + ["--version"], capture_output=True, text=True,
                                 timeout=args.timeout, check=True)
        summary["epubcheckVersion"] = (version.stdout + version.stderr).strip()
        if ace:
            version = subprocess.run([str(ace), "--version"], capture_output=True, text=True,
                                     timeout=args.timeout, check=True)
            summary["aceVersion"] = (version.stdout + version.stderr).strip()
        for index, source in enumerate(inputs):
            folder = output / f"publication-{index + 1:04d}"
            folder.mkdir()
            snapshot = folder / "publication.epub"
            with source.open("rb") as incoming, snapshot.open("xb") as captured:
                remaining = 128 * 1024 * 1024
                while True:
                    chunk = incoming.read(min(1024 * 1024, remaining + 1))
                    if not chunk:
                        break
                    if len(chunk) > remaining:
                        raise ValueError("Publication grew beyond the input bound during capture")
                    captured.write(chunk)
                    remaining -= len(chunk)
            record = {"source": str(source), "sha256": digest(snapshot), "status": "failed"}
            summary["publications"].append(record)
            with (folder / "epubcheck.log").open("w", encoding="utf-8") as log:
                result = subprocess.run(command + [str(snapshot), "--json", str(folder / "epubcheck.json")],
                                        stdout=log, stderr=subprocess.STDOUT, timeout=args.timeout, check=False)
            record["exitCode"] = result.returncode
            record["report"] = str((folder / "epubcheck.json").relative_to(output))
            if result.returncode == 0 and (folder / "epubcheck.json").is_file():
                # Parse the actual report so a successful process without usable evidence cannot pass.
                report = json.loads((folder / "epubcheck.json").read_text(encoding="utf-8"))
                checker = report.get("checker") if isinstance(report, dict) else None
                if not isinstance(checker, dict) or any(type(checker.get(key)) is not int or checker[key] < 0
                                                        for key in ("nError", "nFatal", "nWarning")):
                    raise ValueError("EPUBCheck report is missing valid outcome counts")
                record["errors"] = checker["nError"]
                record["fatalErrors"] = checker["nFatal"]
                record["warnings"] = checker["nWarning"]
                if record["errors"] == 0 and record["fatalErrors"] == 0:
                    record["status"] = "passed"
            failed |= record["status"] != "passed"
            record["automatedAccessibility"] = {"status": "not-checked"}
            if ace:
                automated = record["automatedAccessibility"] = {"status": "failed"}
                with (folder / "ace.log").open("w", encoding="utf-8") as log:
                    result = subprocess.run([str(ace), "--exiterror2", "--timeout", str(args.timeout * 1000),
                                             "--outdir", str(folder / "ace"), "--tempdir", str(folder / "ace-temp"), str(snapshot)],
                                            stdout=log, stderr=subprocess.STDOUT, timeout=args.timeout, check=False)
                automated["exitCode"] = result.returncode
                report_path = folder / "ace" / "report.json"
                automated["report"] = str(report_path.relative_to(output))
                if report_path.is_file():
                    report = json.loads(report_path.read_text(encoding="utf-8"))
                    result_node = report.get("earl:result") if isinstance(report, dict) else None
                    outcome = result_node.get("earl:outcome") if isinstance(result_node, dict) else None
                    automated["outcome"] = outcome
                    if result.returncode == 0 and outcome == "pass":
                        automated["status"] = "passed"
                failed |= automated["status"] != "passed"
    except (OSError, ValueError, subprocess.SubprocessError) as error:
        summary["failure"] = str(error)
        failed = True
    finally:
        (output / "summary.json").write_text(json.dumps(summary, indent=2) + "\n", encoding="utf-8")
    print(output / "summary.json")
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main())
