"""Verify matched-runner net472 evidence, including test identities and outcomes."""
import argparse
from collections import Counter
import json
from pathlib import Path
import statistics
import subprocess
import xml.etree.ElementTree as ET


NS = {"t": "http://microsoft.com/schemas/VisualStudio/TeamTest/2010"}


def read_layout(directory, expected):
    if len(expected) != len(set(expected)):
        raise ValueError(f"Duplicated project executions in {directory}")
    records = [row for path in directory.rglob("execution.json") for row in json.loads(path.read_text(encoding="utf-8-sig"))]
    if Counter(r["project"] for r in records) != Counter(expected):
        raise ValueError(f"Missing or duplicated project executions in {directory}")
    for row in records:
        names = [step["name"] for step in row["steps"]]
        obligations = [f'{row["name"]} / restore', f'{row["name"]} / net472']
        if row["name"] == "Drawing-CodeGlyphX":
            obligations += [f'Drawing-CodeGlyphX / pack {p}' for p in ("Drawing", "Word", "Excel", "PowerPoint", "Pdf")]
        if names != obligations or any(s["exit_code"] != 0 for s in row["steps"]):
            raise ValueError(f'Failed or missing execution obligations: {row["project"]}')
    results = {Path(p).stem: Counter() for p in expected}
    files = Counter()
    for path in directory.rglob("*.trx"):
        tree = ET.parse(path)
        definitions = {}
        assemblies = set()
        for definition in tree.findall(".//t:TestDefinitions/t:UnitTest", NS):
            method = definition.find("t:TestMethod", NS)
            assembly = Path(method.attrib["codeBase"].replace("\\", "/")).stem
            assemblies.add(assembly)
            definitions[definition.attrib["id"]] = (assembly, method.attrib["className"], method.attrib["name"])
        if len(assemblies) != 1 or not assemblies <= results.keys():
            raise ValueError(f"Unexpected or empty test assembly in {path}: {assemblies}")
        assembly = next(iter(assemblies))
        entries = tree.findall(".//t:Results/t:UnitTestResult", NS)
        counters = tree.find(".//t:ResultSummary/t:Counters", NS)
        if not entries or counters is None or int(counters.attrib["total"]) != len(entries):
            raise ValueError(f"Missing or incomplete test results in {path}")
        for result in entries:
            outcome = result.attrib["outcome"]
            if outcome not in ("Passed", "NotExecuted"):
                raise ValueError(f"Unsuccessful test in {path}: {result.attrib['testName']} ({outcome})")
            identity = (*definitions[result.attrib["testId"]], result.attrib["testName"], outcome)
            results[assembly][identity] += 1
        files[assembly] += 1
    if files != Counter({name: 1 for name in results}):
        raise ValueError(f"Missing or duplicate per-project TRX in {directory}: {files}")
    return results, sum(s["seconds"] for r in records for s in r["steps"])


def compare(directory, expected, rounds=3):
    report = []
    reference = None
    environment = None
    for sample in range(1, rounds + 1):
        layouts = {}
        for layout in ("baseline", "sharded"):
            artifacts = sorted(directory.glob(f"net472-evidence-{sample}-{layout}-*"))
            if not artifacts:
                raise ValueError(f"No artifacts for sample {sample} / {layout}")
            # Artifacts remain separate so overwritten TRX or manifests cannot hide duplicates.
            records = []
            seconds = 0
            for artifact in artifacts:
                metadata = list(artifact.rglob("environment.json"))
                if len(metadata) != 1:
                    raise ValueError(f"Missing or duplicated measurement environment in {artifact}")
                current = json.loads(metadata[0].read_text(encoding="utf-8-sig"))
                if not all(current.get(k) for k in ("source", "sdk", "image", "arch")):
                    raise ValueError(f"Incomplete measurement environment in {artifact}")
                if environment is not None and current != environment:
                    raise ValueError(f"Source, SDK or runner environment differs in {artifact}: {current}")
                environment = current
                manifests = list(artifact.rglob("execution.json"))
                projects = [r["project"] for p in manifests for r in json.loads(p.read_text(encoding="utf-8-sig"))]
                results, elapsed = read_layout(artifact, projects)
                records.extend(results.items())
                seconds += elapsed
            if Counter(name for name, _ in records) != Counter(Path(p).stem for p in expected):
                raise ValueError(f"Missing or duplicated selected projects: {sample} / {layout}")
            layouts[layout] = dict(records)
            report.append({"sample": sample, "layout": layout, "jobs": len(artifacts), "command_seconds": seconds,
                           "test_results": sum(sum(c.values()) for _, c in records), "environment": environment})
        if layouts["baseline"] != layouts["sharded"]:
            different = [p for p in layouts["baseline"] if layouts["baseline"][p] != layouts["sharded"][p]]
            raise ValueError(f"Test identity/outcome parity failed in sample {sample}: {different}")
        if reference is not None and layouts["baseline"] != reference:
            raise ValueError(f"Baseline test identities/outcomes changed between samples: {sample}")
        reference = layouts["baseline"]
    return report


def add_job_timings(report, repo, run_id):
    pages = json.loads(subprocess.check_output([
        "gh", "api", "--paginate", "--slurp", f"repos/{repo}/actions/runs/{run_id}/jobs?per_page=100"
    ]))
    from datetime import datetime
    jobs = [job for page in pages for job in page["jobs"]]
    for row in report:
        prefix = f'net472 evidence / {row["sample"]} / {row["layout"]} / '
        selected = [job for job in jobs if job["name"].startswith(prefix)]
        if len(selected) != row["jobs"] or any(j["conclusion"] != "success" for j in selected):
            raise ValueError(f"Incomplete or failed hosted job evidence: {prefix}")
        elapsed = lambda start, end: (datetime.fromisoformat(end) - datetime.fromisoformat(start)).total_seconds()
        row["runner_seconds"] = sum(elapsed(j["started_at"], j["completed_at"]) for j in selected)
        row["span_seconds"] = elapsed(min(j["started_at"] for j in selected), max(j["completed_at"] for j in selected))
        row["job_urls"] = [j["html_url"] for j in selected]


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("directory", type=Path)
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument("--repo")
    parser.add_argument("--run-id")
    args = parser.parse_args()
    config = json.loads(Path(__file__).with_name("pr-suites.json").read_text())
    report = compare(args.directory, [r["project"] for r in config["windows"]])
    if args.repo and args.run_id:
        add_job_timings(report, args.repo, args.run_id)
    summary = {"samples": report, "parity": "identical test identities, outcomes and pack obligations in all samples"}
    if "runner_seconds" in report[0]:
        medians = {layout: statistics.median(r["runner_seconds"] for r in report if r["layout"] == layout)
                   for layout in ("baseline", "sharded")}
        summary["median_runner_seconds"] = medians
        summary["median_runner_seconds_saved"] = medians["baseline"] - medians["sharded"]
        summary["runner_seconds_range"] = {
            layout: [min(r["runner_seconds"] for r in report if r["layout"] == layout),
                     max(r["runner_seconds"] for r in report if r["layout"] == layout)]
            for layout in ("baseline", "sharded")}
    args.output.write_text(json.dumps(summary, indent=2) + "\n")
    print(json.dumps(summary, indent=2))


if __name__ == "__main__":
    main()
