"""Observe selector differences without producing CI selection outputs or building projects."""
import argparse
from datetime import datetime, timezone
import hashlib
import importlib.util
import json
import os
from pathlib import Path
import platform
import re
import shutil
import statistics
import subprocess
import tempfile
import time
import xml.etree.ElementTree as ET


TOOL_VERSION = "7.0.1"


def git(root, *args):
    return subprocess.check_output(["git", *args], cwd=root).decode("utf-8")


def load_selection(root):
    spec = importlib.util.spec_from_file_location("officeimo_selection", root / "Build/Ci/select.py")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def read_projects(path, root):
    """Normalize the tool's per-framework output to repository-relative project identities."""
    rows = json.loads(path.read_text(encoding="utf-8-sig"))
    if not isinstance(rows, list):
        raise ValueError("Expected a JSON project array from dotnet-affected")
    projects = set()
    for row in rows:
        if not isinstance(row, dict) or not isinstance(row.get("FilePath"), str):
            raise ValueError("Expected a FilePath for every dotnet-affected project")
        project = Path(row["FilePath"])
        if not project.is_absolute():
            raise ValueError("Expected absolute project paths from dotnet-affected")
        project = project.resolve()
        relative = project.relative_to(root).as_posix()
        if not relative.endswith(".csproj") or not project.is_file():
            raise ValueError(f"Unexpected project in dotnet-affected output: {relative}")
        projects.add(relative)
    return projects


def differences(current, affected, universe, stable=True):
    """Separate discovery omissions from differences inside the evaluated project graph."""
    if not affected <= universe:
        raise ValueError("Affected projects were absent from the discovery inventory")
    comparable = current & universe
    current_only = comparable - affected
    affected_only = affected - current
    return {
        "common": sorted(current & affected),
        "currentOnlyReviewCandidates": sorted(current_only),
        "affectedOnlyDependencyCandidates": sorted(affected_only),
        "currentOutsideDiscovery": sorted(current - universe),
        "counts": {
            "currentComparableProjects": len(comparable),
            "affectedProjects": len(affected),
            "candidateRemovals": len(current_only),
            "candidateAdditions": len(affected_only),
        },
        "potentialSavings": {
            "netProjectReduction": len(comparable) - len(affected),
            "candidateRemovalPercent": round(100 * len(current_only) / len(comparable), 2) if comparable else None,
            "buildMinutes": None,
        } if stable else None,
    }


def assess_observations(current, observations, universe, graph):
    """Consistency alone cannot excuse omitting a declared consumer of an affected project."""
    stable = all(p == observations[0] for p in observations)
    affected = set.union(*observations)
    missing = [{"consumer": consumer, "dependency": dependency}
               for consumer in sorted(universe - affected)
               for dependency in sorted(graph.get(consumer, set()) & affected)]
    return {
        "status": "unstable" if not stable else "dependency-gaps" if missing else "complete",
        "affectedProjects": sorted(affected),
        "missingDeclaredReferences": missing,
        "differences": differences(current, affected, universe, stable and not missing),
    }


def run_tool(tool, root, base, filter_path, output, name, timeout, inventory=False, exclude=None):
    """Each invocation owns a fresh output directory; errors never become empty selections."""
    command = [tool, "--repository-path", str(root), "--from", base,
               "--uncommitted", "none", "--format", "json", "--output-name", name]
    if filter_path:
        command += ["--filter-file-path", str(filter_path)]
    if exclude:
        command += ["--exclude-discovery", exclude]
    if inventory:
        roots = ([p.get("Path") for p in ET.parse(filter_path).iter("Project")]
                 if filter_path else git(root, "ls-files", "*.csproj").splitlines())
        names = {Path(p.replace("\\", "/")).stem for p in roots
                 if not exclude or not re.search(exclude, str((root / p).resolve()))}
        for project in sorted(names):
            command += ["--assume-changes", project]
    started = time.perf_counter()
    with tempfile.TemporaryDirectory(prefix="affected-", dir=output) as temporary:
        command += ["--output-dir", temporary]
        log = output / f"{name}.log"
        # The apphost is invoked directly so the timeout owns the evaluating process.
        with log.open("w", encoding="utf-8") as stream:
            result = subprocess.run(command, cwd=root, stdout=stream, stderr=subprocess.STDOUT,
                                    timeout=timeout, check=False)
        elapsed = round(time.perf_counter() - started, 3)
        if result.returncode == 166 and not inventory:
            return set(), elapsed
        if result.returncode != 0:
            raise RuntimeError(f"dotnet-affected {name} exited {result.returncode}; see {log.name}")
        projects = read_projects(Path(temporary) / f"{name}.json", root)
        if inventory and not projects:
            raise ValueError("dotnet-affected discovered no projects")
        return projects, elapsed


def markdown(report):
    lines = ["# Affected-selection comparison", "",
             "Observation only. Existing CI selection and coverage remain authoritative.", "",
             f"- Status: {report['status']}",
             f"- Source: `{report['head']}`; baseline: `{report['base']}`",
             f"- Platform: {report['platform']}; SDK: `{report['sdk']}`",
             f"- Discovery: `{report['filter'] or 'all projects'}`; dotnet-affected: `{TOOL_VERSION}`",
             f"- Explicit discovery exclusion: `{report.get('excludeDiscovery') or 'none'}`",
             f"- Changed files: {len(report['changedFiles'])}", ""]
    if report["status"] not in ("complete", "unstable", "dependency-gaps"):
        lines += [f"Comparison unavailable: {report['error']}", "",
                  "No missed-dependency or savings conclusion can be drawn from this run."]
        return "\n".join(lines) + "\n"
    delta = report["differences"]
    counts = delta["counts"]
    lines += ["| Measure | Projects |", "| --- | ---: |",
              f"| Current selector, within discovery | {counts['currentComparableProjects']} |",
              f"| dotnet-affected, union of observations | {counts['affectedProjects']} |",
              f"| Current-only projects to investigate | {counts['candidateRemovals']} |",
              f"| Tool-only dependency candidates | {counts['candidateAdditions']} |",
              f"| Current projects outside discovery | {len(delta['currentOutsideDiscovery'])} |", "",
              "These are project counts, not saved build minutes or permission to skip tests. "
              "Conditional references, runtime-loaded fixtures and non-.NET contracts need separate review. "
              "Projects outside discovery are excluded from savings calculations.", "",
              f"Selection seconds: {json.dumps(report['seconds'], sort_keys=True)}. "
              "No builds or product tests were timed.", ""]
    lines += [f"Tool observation sizes: {[len(p) for p in report['observations']]}", ""]
    if report["status"] == "unstable":
        lines += ["**Identical inputs produced different project sets. Savings are withheld.** "
                  "The union above is diagnostic; it is not a complete or safe selection.", ""]
    if report["missingDeclaredReferences"]:
        lines += ["**Declared consumers are missing from the observations. Savings are withheld.** "
                  "These references include conditional declarations; investigate them with the "
                  "appropriate MSBuild configuration before deciding which are active.", "",
                  "## Missing declared references", ""]
        lines += [f"- `{r['consumer']}` references `{r['dependency']}`"
                  for r in report["missingDeclaredReferences"]]
        lines.append("")
    for label, key in [("Tool-only dependency candidates", "affectedOnlyDependencyCandidates"),
                       ("Current-only review candidates", "currentOnlyReviewCandidates"),
                       ("Current projects outside discovery", "currentOutsideDiscovery")]:
        lines += [f"## {label}", ""]
        lines += [f"- `{p}`" for p in delta[key]] or ["None."]
        lines.append("")
    return "\n".join(lines)


def compare(root, base, tool, filter_path, output, timeout, exclude=None):
    root = root.resolve()
    base = git(root, "rev-parse", "--verify", "--end-of-options", f"{base}^{{commit}}").strip()
    head = git(root, "rev-parse", "HEAD").strip()
    # Both selectors must see the same committed source. Untracked project files also
    # change MSBuild discovery even with --uncommitted none.
    if git(root, "status", "--porcelain", "--untracked-files=no").strip():
        raise ValueError("Use a checkout without tracked changes for a comparable graph")
    untracked = git(root, "ls-files", "--others", "--exclude-standard", "-z").split("\0")
    if any(untracked):
        raise ValueError("Untracked files can change MSBuild inputs; use a clean checkout")
    changed = git(root, "diff", "--no-renames", "--name-only", "-z", base, head).split("\0")[:-1]
    report = {
        "schemaVersion": 1, "status": "incomplete", "coverageAuthority": "Build/Ci/select.py",
        "selectionMode": "PR project selection; existing full-CI overrides remain unchanged",
        "createdUtc": datetime.now(timezone.utc).isoformat(), "head": head, "base": base,
        "comparisonScriptSha256": hashlib.sha256(Path(__file__).read_bytes()).hexdigest(),
        "platform": platform.platform(),
        "evaluationEnvironment": {key: os.environ.get(key) for key in (
            "DOTNET_PROCESSOR_COUNT", "Configuration", "TargetFramework", "TargetFrameworks",
            "OfficeIMOCodeQLBuild", "OfficeIMOCompatibilityTarget")},
        "sdk": subprocess.check_output(["dotnet", "--version"], cwd=root, text=True).strip(),
        "filter": filter_path.relative_to(root).as_posix() if filter_path else None,
        "excludeDiscovery": exclude,
        "changedFiles": changed, "seconds": {},
    }
    try:
        selector = load_selection(root)
        selector_times = []
        for _ in range(3):
            started = time.perf_counter()
            current = selector.select(root, changed)
            selector_times.append(round(time.perf_counter() - started, 3))
        report["seconds"]["currentSelector"] = selector_times
        report["currentSelection"] = current
        graph, _ = selector.project_graph(root, selector.tracked(root))
        version = subprocess.check_output([tool, "--version"], cwd=root, text=True, timeout=30).strip()
        report["toolVersion"] = version
        if version.split("+", 1)[0] != TOOL_VERSION:
            raise ValueError(f"Expected dotnet-affected {TOOL_VERSION}, received {version}")
        universe, seconds = run_tool(tool, root, base, filter_path, output, "inventory", timeout, True, exclude)
        report["seconds"]["discoveryInventory"] = seconds
        report["discoveredProjects"] = sorted(universe)
        observations = []
        tool_times = []
        report["observations"] = []
        report["seconds"]["dotnetAffected"] = tool_times
        for sample in range(3):
            projects, seconds = run_tool(tool, root, base, filter_path, output,
                                         f"affected-{sample + 1}", timeout, exclude=exclude)
            observations.append(projects)
            report["observations"].append(sorted(projects))
            tool_times.append(seconds)
        report["seconds"]["dotnetAffectedMedian"] = statistics.median(tool_times)
        report["seconds"]["currentSelectorMedian"] = statistics.median(selector_times)
        report.update(assess_observations(set(current["projects"]), observations, universe, graph))
    except (OSError, ValueError, RuntimeError, ET.ParseError, re.error, subprocess.SubprocessError) as error:
        report["error"] = str(error)
    return report


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--repository", type=Path, default=Path(__file__).resolve().parents[2])
    parser.add_argument("--base", required=True, help="Commit to compare with this checkout's HEAD")
    parser.add_argument("--tool", default="dotnet-affected", help="Path to the pinned tool's apphost")
    parser.add_argument("--filter", default="OfficeIMO.slnx", help="Repository-relative .slnx, or 'all'")
    parser.add_argument("--exclude-discovery", help="Explicit dotnet-affected exclusion regex, recorded in the report")
    parser.add_argument("--output", type=Path, required=True, help="New report directory outside the checkout")
    parser.add_argument("--timeout", type=int, default=180, help="Seconds allowed per tool evaluation")
    args = parser.parse_args()
    root = args.repository.resolve()
    output = args.output.resolve()
    if output.is_relative_to(root):
        parser.error("--output must be outside the evaluated checkout")
    if args.timeout < 1:
        parser.error("--timeout must be positive")
    tool = shutil.which(args.tool)
    if tool is None:
        parser.error(f"Tool not found: {args.tool}; install dotnet-affected {TOOL_VERSION} into an isolated tool path")
    filter_path = None if args.filter == "all" else (root / args.filter).resolve()
    if filter_path and (not filter_path.is_relative_to(root) or not filter_path.is_file() or filter_path.suffix != '.slnx'):
        parser.error("--filter must name an existing .slnx file inside the evaluated checkout")
    output.mkdir(parents=True, exist_ok=False)
    try:
        report = compare(root, args.base, tool, filter_path, output, args.timeout, args.exclude_discovery)
    except (OSError, ValueError, subprocess.SubprocessError) as error:
        parser.exit(1, f"Comparison setup failed: {error}\n")
    (output / "comparison.json").write_text(json.dumps(report, indent=2) + "\n", encoding="utf-8")
    (output / "comparison.md").write_text(markdown(report), encoding="utf-8")
    print(json.dumps({"status": report["status"], "report": str(output / "comparison.md")}))
    return 0 if report["status"] == "complete" else 1


if __name__ == "__main__":
    raise SystemExit(main())
