"""Select PR contracts from tracked project references, conservatively on unknown inputs.

No packages are restored or MSBuild code executed during selection. Conditional
references are included so a platform-specific consumer cannot be silently lost.
"""
import argparse
import fnmatch
import json
import os
from pathlib import Path
import subprocess
import xml.etree.ElementTree as ET


def tracked(root):
    return subprocess.check_output(["git", "ls-files", "-z"], cwd=root).decode().split("\0")[:-1]


def project_graph(root, files):
    projects = {p: set() for p in files if p.endswith(".csproj")}
    linked_inputs = {}
    for path, dependencies in projects.items():
        tree = ET.parse(root / path)
        for reference in tree.iter("ProjectReference"):
            include = reference.get("Include", "").replace("\\", "/")
            if include.startswith("$(OfficeIMORoot)"):
                include = str(root) + "/" + include.removeprefix("$(OfficeIMORoot)").lstrip("/")
            if "$" in include:  # External source overrides are not enabled in ordinary CI.
                continue
            resolved = ((root / path).parent / include).resolve().relative_to(root).as_posix()
            if resolved in projects:
                dependencies.add(resolved)
        # Linked source and fixtures are real inputs even without ProjectReference.
        for item in tree.iter():
            if item.tag not in ("Compile", "None", "Content", "EmbeddedResource", "AdditionalFiles", "Import"):
                continue
            for include in item.get("Include", item.get("Project", "")).split(";"):
                include = include.replace("\\", "/")
                if not include or "$" in include:
                    continue
                candidate = ((root / path).parent / include).resolve()
                if not candidate.is_relative_to(root):
                    continue
                pattern = candidate.relative_to(root).as_posix()
                linked_inputs.setdefault(pattern, set()).add(path)
    return projects, linked_inputs


def select(root, changed, full=False):
    config = json.loads((root / "Build/Ci/pr-suites.json").read_text())
    files = tracked(root)
    graph, linked_inputs = project_graph(root, files)
    affected = set()
    directly_changed = set()
    wide = full
    wide_paths = []
    for path in changed:
        linked_consumers = {project for pattern, consumers in linked_inputs.items()
                            if fnmatch.fnmatchcase(path, pattern) for project in consumers}
        affected.update(linked_consumers)
        # Executable Markdown and other fixtures are inputs, regardless of suffix.
        if path.startswith(("OfficeIMO.TestAssets/", "OfficeIMO.SharedSource/")) or "/Fixtures/" in path:
            wide = True
            wide_paths.append(path)
            continue
        # Exclude known prose only after accounting for linked build inputs.
        prose = path.endswith(".md") and (path.startswith("Docs/") or Path(path).name.lower() in
                  ("readme.md", "support.md", "migration.md", "agents.md", "license.md"))
        if prose or path == ".gitignore":
            continue
        if path.startswith(("Website/", "OfficeIMO.JavaScript/")):
            continue
        owners = [p for p in graph if path.startswith(str(Path(p).parent).replace("\\", "/") + "/")]
        if owners:
            affected.update(owners)
            directly_changed.update(owners)
        elif path.startswith("PowerShell/"):
            # Unowned module assets run through the PowerShell suite below.
            # Linked consumers were collected before this routing decision.
            continue
        else:
            # Shared source, fixtures, build scripts, workflow and graph changes fail open.
            wide = True
            wide_paths.append(path)
    routed_tests = set(sum(config["suites"].values(), []) + sum(config["partitions"].values(), []))
    routed_tests.update(Path(row["project"]).stem for row in config["windows"])
    unmapped_tests = {Path(p).stem for p in directly_changed if Path(p).stem.endswith("Tests")} - routed_tests
    if unmapped_tests:
        raise ValueError("Test projects need an executable CI route in pr-suites.json: " + ", ".join(sorted(unmapped_tests)))
    def with_consumers(projects):
        projects = set(projects)
        while True:
            consumers = {p for p, refs in graph.items() if refs & projects}
            if consumers <= projects:
                return projects
            projects |= consumers
    matches = lambda patterns, paths=changed: any(fnmatch.fnmatchcase(path, pattern) for path in paths for pattern in patterns)
    # Workflow and CI edits fail open for hosted Linux and Windows lanes, but
    # scarce macOS lanes only follow the projects those edits actually touch.
    ci_only = wide and not full and all(matches(config.get("ci_only_inputs", []), [p]) for p in wide_paths)
    narrow = with_consumers(affected)
    affected = set(graph) if wide else narrow
    names = {Path(p).stem for p in affected}
    selected = lambda roots: bool(names.intersection(roots))
    suites = {key: selected(roots) for key, roots in config["suites"].items()}
    suites["build-matrix"] = bool(affected)
    # JavaScript asset generation has its own executable browser and consumer proof.
    suites["browser-exports"] |= any(p.startswith(("OfficeIMO.JavaScript/", "Build/BrowserExports/")) for p in changed)
    # Full website generation has its own inputs. Library correctness, browser
    # exports and package contracts run in their owning lanes; full site/API
    # regeneration still runs on publication and explicit full-ci requests.
    suites["website"] = full or any(fnmatch.fnmatchcase(path, pattern)
                                    for path in changed for pattern in config["website_inputs"])
    suites["powershell"] |= any(p.startswith("PowerShell/") for p in changed)
    # Non-.NET lanes declare their own inputs. The selector, SDK and NuGet
    # configuration still select every one of them; full runs always do.
    shared = full or matches(config.get("shared_inputs", []))
    for key, patterns in config.get("path_suites", {}).items():
        suites[key] = shared or matches(patterns)
    if ci_only:
        narrow_names = {Path(p).stem for p in narrow}
        for key in config.get("macos_suites", []):
            suites[key] = shared or bool(narrow_names.intersection(config["suites"][key]))
    partitions = [p for p, roots in config["partitions"].items() if selected(roots)]
    frameworks = ["net8.0", "net10.0"] if full else ["net10.0"]
    ubuntu = [{"framework": f, "partition": p} for f in frameworks for p in partitions]
    windows = [row for row in config["windows"] if Path(row["project"]).stem in names]
    return {
        "full": full,
        "suites": suites,
        "ubuntu": {"include": ubuntu or [{"framework": "net10.0", "partition": "none"}]},
        "ubuntu_enabled": bool(ubuntu),
        "windows": {"include": windows or [{"name": "none", "project": "none"}]},
        "windows_enabled": bool(windows),
        "managed": bool(affected),
        "projects": sorted(affected),
    }


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--base")
    parser.add_argument("--files", nargs="*")
    parser.add_argument("--full", action="store_true")
    parser.add_argument("--output")
    args = parser.parse_args()
    root = Path(__file__).resolve().parents[2]
    event = json.loads(Path(os.environ["GITHUB_EVENT_PATH"]).read_text()) if os.environ.get("GITHUB_EVENT_PATH") else {}
    full = args.full or (bool(event) and "pull_request" not in event) or any(
        label["name"] == "full-ci" for label in event.get("pull_request", {}).get("labels", []))
    if args.files is not None:
        changed = args.files
    elif args.base:
        changed = subprocess.check_output(["git", "diff", "--no-renames", "--name-only", "-z", args.base, "HEAD"], cwd=root).decode().split("\0")[:-1]
    elif full:
        changed = []
    else:
        parser.error("--base or --files is required for a pull request")
    result = select(root, changed, full)
    if args.output:
        Path(args.output).write_text(json.dumps(result, indent=2) + "\n")
    print(json.dumps(result, indent=2))
    if os.environ.get("GITHUB_OUTPUT"):
        with open(os.environ["GITHUB_OUTPUT"], "a") as output:
            for key, value in result.items():
                output.write(f"{key}={json.dumps(value, separators=(',', ':'))}\n")


if __name__ == "__main__":
    main()
