#!/usr/bin/env python3
"""Check that third-party material in OfficeIMO packages is declared and shipped with its notices.

Build/third-party-material.json lists every piece of third-party code or data that a NuGet
package compiles in or embeds. This script checks that list against the repository:

  source   (default) the manifest, the sources, the project files and the license expressions
  package  built .nupkg files: license expression, notices file and license texts

Rules:
  * Each entry's files and license texts exist, and its package packs them, unless the entry
    records why its notice does not ship (noticeOmitted).
  * Files recorded under firstParty were written by OfficeIMO's author in another of their
    projects and need no notice.
  * A package with entries packs a notices file that names every entry by its title.
  * A package's PackageLicenseExpression is MIT plus the license of every third-party *code*
    entry it compiles. Data licenses go in the notices, not in the expression: nuget.org only
    accepts OSI- or FSF-approved identifiers, so data terms such as APAFML or CC-BY-SA-3.0
    cannot be expressed there, and listing only some of them would be misleading.
  * Source files that say they were adapted, ported or vendored from elsewhere, carry a
    third-party copyright, or cite Unicode data, belong to an entry. So does any license or
    notice text file inside a packaged project.
"""

from __future__ import annotations

import argparse
import fnmatch
import json
import re
import subprocess
import sys
import xml.etree.ElementTree as ET
import zipfile
from dataclasses import dataclass, field
from pathlib import Path, PurePosixPath

MANIFEST = "Build/third-party-material.json"
SHARED_SOURCE = "OfficeIMO.SharedSource"
ALL_PACKAGES = "*"

MARKER = re.compile(
    r"\b(adapted|vendored|ported|translated) from\b"
    r"|\bport of\b"
    r"|SPDX-License-Identifier"
    r"|\bcopyright\s*(\(c\)|©|\d{4})"
    r"|\blicen[sc]ed under\b"
    r"|unicode\.org/(Public|license)",
    re.IGNORECASE,
)
OWN_TEXT = re.compile(r"Evotec|Przemys|OfficeIMO's|repository's MIT", re.IGNORECASE)
COMMENT = re.compile(r"^\s*(//(?!/)|/\*|\*)")
LICENSE_FILE = re.compile(r"(^|[-_.])(LICEN[CS]ES?|NOTICES?|PATENTS?|COPYING)([-_.]|$)|ThirdPartyNotices", re.IGNORECASE)
EXCLUDED_PARTS = {"bin", "obj", "node_modules", ".git"}


@dataclass
class Report:
    errors: list[str] = field(default_factory=list)

    def error(self, message: str) -> None:
        self.errors.append(message)


@dataclass
class Project:
    name: str
    path: PurePosixPath  # repository-relative csproj path
    expression: str | None
    package_id: str

    @property
    def directory(self) -> PurePosixPath:
        return self.path.parent


def repository_files(root: Path) -> list[str]:
    if (root / ".git").exists():
        result = subprocess.run(
            ["git", "-C", str(root), "ls-files", "-z", "--cached", "--others", "--exclude-standard"],
            check=True,
            capture_output=True,
        )
        files = [name for name in result.stdout.decode("utf-8").split("\0") if name]
        return [name for name in files if (root / name).exists()]
    files = []
    for path in root.rglob("*"):
        if path.is_file() and not EXCLUDED_PARTS.intersection(path.relative_to(root).parts):
            files.append(path.relative_to(root).as_posix())
    return files


def read_text(path: Path) -> str:
    return path.read_text(encoding="utf-8-sig", errors="replace")


def split_expression(expression: str) -> set[str] | None:
    """Returns the terms of an AND-only SPDX expression, or None for anything else."""
    terms = [term.strip() for term in re.split(r"\s+AND\s+", expression.strip())]
    if any(not term or re.search(r"\s|\(|\)", term) for term in terms):
        return None
    return set(terms)


def load_projects(root: Path, files: list[str]) -> dict[str, Project]:
    projects: dict[str, Project] = {}
    for name in files:
        if not name.endswith(".csproj"):
            continue
        text = read_text(root / name)
        match = re.search(r"<PackageLicenseExpression>(.*?)</PackageLicenseExpression>", text, re.DOTALL)
        if not match:
            continue
        path = PurePosixPath(name)
        package_id = re.search(r"<PackageId>(.*?)</PackageId>", text)
        projects[path.stem] = Project(
            name=path.stem,
            path=path,
            expression=match.group(1).strip(),
            package_id=package_id.group(1).strip() if package_id else path.stem,
        )
    return projects


def expand(files: list[str], pattern: str) -> list[str]:
    return [name for name in files if fnmatch.fnmatchcase(name, pattern)]


def packed_items(root: Path, project_file: PurePosixPath, properties: dict[str, str]) -> dict[str, str]:
    """Maps repository-relative source paths to package paths for items packed by an MSBuild file."""
    try:
        document = ET.parse(root / project_file)
    except ET.ParseError:
        return {}
    directory = (root / project_file).parent
    packed: dict[str, str] = {}
    for element in document.iter():
        tag = element.tag.split("}")[-1]
        if tag not in {"None", "Content", "EmbeddedResource", "Compile"}:
            continue
        include = element.get("Include")
        if not include:
            continue
        values = {child.tag.split("}")[-1]: (child.text or "").strip() for child in element}
        pack = element.get("Pack", values.get("Pack", ""))
        if pack.lower() != "true":
            continue
        package_path = element.get("PackagePath", values.get("PackagePath", ""))
        for name, value in properties.items():
            include = include.replace(f"$({name})", value)
        for part in include.split(";"):
            part = part.strip().replace("\\", "/")
            if not part:
                continue
            for source in sorted(directory.glob(part)) if any(c in part for c in "*?") else [directory / part]:
                try:
                    relative = source.resolve().relative_to(root.resolve()).as_posix()
                except ValueError:
                    continue
                target = package_path.replace("\\", "/").strip()
                if target in {"", "/"}:
                    target = source.name
                elif target.endswith("/") or not PurePosixPath(target).suffix:
                    target = target.rstrip("/") + "/" + source.name
                packed[relative] = target.lstrip("/")
    return packed


def entry_projects(entry: dict, projects: dict[str, Project]) -> list[Project]:
    if entry.get("project") == ALL_PACKAGES:
        return list(projects.values())
    project = projects.get(entry.get("project", ""))
    return [project] if project else []


def expected_expression(project: Project, entries: list[dict]) -> set[str]:
    terms = {"MIT"}
    for entry in entries:
        if entry["kind"] == "code" and entry["project"] in {project.name, ALL_PACKAGES}:
            terms |= split_expression(entry["license"]) or {entry["license"]}
    return terms


def check_source(root: Path) -> Report:
    report = Report()
    manifest = json.loads(read_text(root / MANIFEST))
    entries: list[dict] = manifest.get("entries", [])
    notices: dict[str, dict] = manifest.get("notices", {})
    files = repository_files(root)
    projects = load_projects(root, files)
    covered: set[str] = set()
    license_sources: set[str] = set()
    ids: set[str] = set()

    for entry in entries:
        label = entry.get("id", "<missing id>")
        for key in ("id", "title", "kind", "license", "source", "project", "files", "licenseFiles"):
            if key not in entry:
                report.error(f"{label}: missing '{key}'.")
        if any(key not in entry for key in ("kind", "project", "files", "licenseFiles", "license", "title")):
            continue
        if label in ids:
            report.error(f"{label}: duplicate id.")
        ids.add(label)
        if entry["kind"] not in {"code", "data"}:
            report.error(f"{label}: kind must be 'code' or 'data'.")
        if entry["kind"] == "code" and split_expression(entry["license"]) is None:
            report.error(f"{label}: a code license must be an SPDX identifier or an AND expression.")
        if entry["project"] != ALL_PACKAGES and entry["project"] not in projects:
            report.error(f"{label}: '{entry['project']}' is not a packaged project.")
            continue
        for pattern in entry["files"]:
            matched = expand(files, pattern)
            if not matched:
                report.error(f"{label}: '{pattern}' matches no file.")
            if entry["project"] != ALL_PACKAGES:
                directory = projects[entry["project"]].directory.as_posix() + "/"
                for name in matched:
                    if not name.startswith(directory):
                        report.error(f"{label}: '{name}' is outside {entry['project']}.")
            covered.update(matched)
        for source in entry["licenseFiles"]:
            license_sources.add(source)
            if not (root / source).is_file():
                report.error(f"{label}: license file '{source}' does not exist.")
        if "noticeOmitted" in entry and not str(entry["noticeOmitted"]).strip():
            report.error(f"{label}: 'noticeOmitted' must say why the notice does not ship.")
        if entry["project"] == ALL_PACKAGES and not entry["licenseFiles"] and not entry.get("noticeOmitted"):
            report.error(f"{label}: material in every package needs a license file to pack.")

    for item in manifest.get("firstParty", []):
        for key in ("id", "files", "reason"):
            if key not in item:
                report.error(f"{item.get('id', '<missing id>')}: first-party record is missing '{key}'.")
        for pattern in item.get("files", []):
            matched = expand(files, pattern)
            if not matched:
                report.error(f"{item.get('id')}: '{pattern}' matches no file.")
            covered.update(matched)

    packed_by_props = packed_items(
        root,
        PurePosixPath("Directory.Build.props"),
        {"OfficeIMOSharedSourceDir": str((root / SHARED_SOURCE).resolve()), "MSBuildThisFileDirectory": str(root.resolve()) + "/"},
    ) if (root / "Directory.Build.props").exists() else {}

    for project in projects.values():
        mine = [e for e in entries if e.get("project") == project.name and "licenseFiles" in e]
        shared = [e for e in entries if e.get("project") == ALL_PACKAGES and "licenseFiles" in e and not e.get("noticeOmitted")]
        packed = packed_items(root, project.path, {})
        for entry in shared:
            for source, target in entry["licenseFiles"].items():
                if packed_by_props.get(source, packed.get(source)) != target:
                    report.error(f"{project.name}: does not pack '{source}' as '{target}' for {entry['id']}.")
        for entry in mine:
            if entry.get("noticeOmitted"):
                continue
            for source, target in entry["licenseFiles"].items():
                if packed.get(source) != target:
                    report.error(f"{project.name}: does not pack '{source}' as '{target}' for {entry['id']}.")
        notice = notices.get(project.name)
        if mine and not notice:
            report.error(f"{project.name}: has third-party material but no notices file in the manifest.")
        if notice:
            if not mine:
                report.error(f"{project.name}: has a notices file but no third-party material in the manifest.")
            path = root / notice["path"]
            if not path.is_file():
                report.error(f"{project.name}: notices file '{notice['path']}' does not exist.")
            else:
                text = read_text(path)
                for entry in mine:
                    if not entry.get("noticeOmitted") and entry["title"] not in text:
                        report.error(f"{project.name}: '{notice['path']}' does not mention '{entry['title']}'.")
            if packed.get(notice["path"]) != notice["packagePath"]:
                report.error(f"{project.name}: does not pack '{notice['path']}' as '{notice['packagePath']}'.")

        declared = split_expression(project.expression or "")
        expected = expected_expression(project, [e for e in entries if "kind" in e and "license" in e])
        if declared is None:
            report.error(f"{project.name}: PackageLicenseExpression '{project.expression}' must be MIT joined with AND terms.")
        elif declared != expected:
            missing = ", ".join(sorted(expected - declared)) or "none"
            extra = ", ".join(sorted(declared - expected)) or "none"
            report.error(
                f"{project.name}: PackageLicenseExpression is '{project.expression}', expected "
                f"'{' AND '.join(sorted(expected, key=lambda t: (t != 'MIT', t)))}' "
                f"(missing: {missing}; without third-party code behind it: {extra}). "
                "Data licenses belong in the notices, not in the expression."
            )

    scanned_roots = [project.directory.as_posix() + "/" for project in projects.values()] + [SHARED_SOURCE + "/"]
    notices_paths = {notice["path"] for notice in notices.values()}
    dependency_folders = tuple(item["path"] for item in manifest.get("dependencyLicenses", []))
    for name in files:
        if not any(name.startswith(prefix) for prefix in scanned_roots):
            continue
        if EXCLUDED_PARTS.intersection(PurePosixPath(name).parts):
            continue
        filename = PurePosixPath(name).name
        if LICENSE_FILE.search(PurePosixPath(filename).stem) and not name.endswith((".cs", ".csproj")):
            if name in license_sources or name in notices_paths or name.startswith(dependency_folders):
                continue
            if OWN_TEXT.search(read_text(root / name)[:2000]):
                continue
            report.error(f"{name}: license or notice file that no manifest entry declares.")
            continue
        if not name.endswith(".cs") or name in covered:
            continue
        for number, line in enumerate(read_text(root / name).splitlines(), start=1):
            if COMMENT.match(line) and MARKER.search(line) and not OWN_TEXT.search(line):
                report.error(f"{name}:{number}: third-party marker outside any manifest entry: {line.strip()[:120]}")
                break
    return report


def normalized_text(data: bytes) -> str:
    """License text compared without a byte order mark, line-ending differences or trailing blank space."""
    text = data.decode("utf-8-sig", errors="replace").replace("\r\n", "\n").replace("\r", "\n")
    return "\n".join(line.rstrip() for line in text.split("\n")).strip()


def nuspec_metadata(archive: zipfile.ZipFile) -> tuple[str, str | None]:
    nuspec = [n for n in archive.namelist() if n.endswith(".nuspec") and "/" not in n]
    document = ET.fromstring(archive.read(nuspec[0]))
    metadata = next(child for child in document if child.tag.split("}")[-1] == "metadata")
    values = {child.tag.split("}")[-1]: child for child in metadata}
    license_element = values.get("license")
    expression = license_element.text.strip() if license_element is not None and license_element.get("type") == "expression" else None
    return values["id"].text.strip(), expression


def check_packages(root: Path, packages: list[Path]) -> Report:
    report = Report()
    manifest = json.loads(read_text(root / MANIFEST))
    entries: list[dict] = manifest.get("entries", [])
    notices: dict[str, dict] = manifest.get("notices", {})
    files = repository_files(root)
    by_package_id = {project.package_id.lower(): project for project in load_projects(root, files).values()}
    found = [p for path in packages for p in ([path] if path.is_file() else sorted(path.rglob("*.nupkg")))]
    found = [p for p in found if not p.name.endswith(".snupkg")]
    if not found:
        report.error("no .nupkg files found.")
    for package in found:
        with zipfile.ZipFile(package) as archive:
            package_id, expression = nuspec_metadata(archive)
            project = by_package_id.get(package_id.lower())
            if project is None:
                continue
            names = {name.lower() for name in archive.namelist()}
            mine = [e for e in entries if e["project"] in {project.name, ALL_PACKAGES}]
            expected = expected_expression(project, entries)
            if expression is None or split_expression(expression) != expected:
                report.error(f"{package.name}: license expression is '{expression}', expected '{' AND '.join(sorted(expected, key=lambda t: (t != 'MIT', t)))}'.")
            for entry in mine:
                if entry.get("noticeOmitted"):
                    continue
                for source, target in entry["licenseFiles"].items():
                    if target.lower() not in names:
                        report.error(f"{package.name}: missing '{target}' for {entry['id']}.")
                        continue
                    actual = next(n for n in archive.namelist() if n.lower() == target.lower())
                    packed = normalized_text(archive.read(actual))
                    if not packed:
                        report.error(f"{package.name}: '{target}' for {entry['id']} is empty.")
                    elif (root / source).is_file() and packed != normalized_text((root / source).read_bytes()):
                        report.error(f"{package.name}: '{target}' for {entry['id']} differs from '{source}'.")
            notice = notices.get(project.name)
            if notice:
                if notice["packagePath"].lower() not in names:
                    report.error(f"{package.name}: missing '{notice['packagePath']}'.")
                else:
                    actual = next(n for n in archive.namelist() if n.lower() == notice["packagePath"].lower())
                    text = archive.read(actual).decode("utf-8-sig", errors="replace")
                    for entry in mine:
                        if entry["project"] != ALL_PACKAGES and not entry.get("noticeOmitted") and entry["title"] not in text:
                            report.error(f"{package.name}: '{notice['packagePath']}' does not mention '{entry['title']}'.")
    return report


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("mode", nargs="?", choices=["source", "package"], default="source")
    parser.add_argument("paths", nargs="*", type=Path, help="package files or folders (package mode)")
    parser.add_argument("--root", type=Path, default=Path(__file__).resolve().parent.parent)
    args = parser.parse_args(argv)
    if args.mode == "package":
        report = check_packages(args.root, args.paths or [args.root / "Artefacts" / "ProjectBuild"])
    else:
        report = check_source(args.root)
    for message in report.errors:
        print(f"error: {message}", file=sys.stderr)
    if report.errors:
        print(f"{len(report.errors)} third-party notice problem(s). See Docs/building-and-releasing.md, 'Third-party material and package licenses'.", file=sys.stderr)
        return 1
    print("Third-party material is declared and shipped with its notices.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
