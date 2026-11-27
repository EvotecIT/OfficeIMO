#!/usr/bin/env python3
"""Compare XLSX projections, unchanged XLSB saves and edited XLSB saves on both revisions."""

import argparse
import json
from io import BytesIO
import os
from pathlib import Path
import shutil
import subprocess
import xml.etree.ElementTree as ET
import zipfile


HERE = Path(__file__).resolve().parent
READER = Path("OfficeIMO.Excel/Xlsb/XlsbWorkbookReader.cs")
CAPTURE = """            string? captureDirectory = Environment.GetEnvironmentVariable("OFFICEIMO_XLSB_CORPUS");
            if (!string.IsNullOrEmpty(captureDirectory)) {
                using var hash = System.Security.Cryptography.SHA256.Create();
                string name = BitConverter.ToString(hash.ComputeHash(packageBytes)).Replace("-", string.Empty);
                string capturePath = Path.Combine(captureDirectory, name + ".xlsb");
                lock (CaptureCorpusLock) {
                    if (!File.Exists(capturePath)) File.WriteAllBytes(capturePath, packageBytes);
                }
            }
"""
CAPTURE_INSERTION = "            ArgumentNullException.ThrowIfNull(packageBytes);"
READER_SIGNATURE = "        internal static XlsbWorkbook Load(byte[] packageBytes, XlsbImportOptions? options = null) {"
CORE_MODIFIED = "{http://purl.org/dc/terms/}modified"
MC = "{http://schemas.openxmlformats.org/markup-compatibility/2006}"
XSI_TYPE = "{http://www.w3.org/2001/XMLSchema-instance}type"


def run(command, cwd, **kwargs):
    return subprocess.run(command, cwd=cwd, check=True, text=True, **kwargs)


def run_locked(command, cwd, build_lock, **kwargs):
    if build_lock is not None:
        command = ["lockf", "-k", str(build_lock), *command]
    return run(command, cwd, **kwargs)


def capture_test_inputs(checkout, destination, build_lock):
    destination.mkdir(parents=True)
    reader = checkout / READER
    source = reader.read_text()
    anchor = READER_SIGNATURE + "\n" + CAPTURE_INSERTION
    if source.count(anchor) != 1:
        raise ValueError(f"Cannot locate the XLSB reader entry in {reader}")
    field = "        private static readonly object CaptureCorpusLock = new();\n\n"
    reader.write_text(source.replace(anchor, field + anchor + "\n" + CAPTURE, 1))
    try:
        project = "OfficeIMO.Excel.Tests/OfficeIMO.Excel.Tests.csproj"
        run_locked(["dotnet", "build", project, "-c", "Release", "-f", "net10.0", "-m:2"], checkout, build_lock)
        run_locked(["dotnet", "test", project, "-c", "Release", "-f", "net10.0", "--no-build", "-m:2",
                    "--filter", "FullyQualifiedName~Xlsb&Category!=Performance"], checkout, build_lock,
                   env={**os.environ, "OFFICEIMO_XLSB_CORPUS": str(destination)})
    finally:
        reader.write_text(source)
    if not any(destination.iterdir()):
        raise ValueError("The Excel tests did not capture any native inputs")


def copy_fixtures(checkout, corpus):
    tracked = run(["git", "ls-files", "-z", "*.xlsb"], checkout, stdout=subprocess.PIPE).stdout
    if not tracked:
        raise ValueError("The baseline contains no tracked XLSB fixtures")
    for name in tracked.split("\0"):
        if not name:
            continue
        path = corpus / "fixtures" / name
        path.parent.mkdir(parents=True, exist_ok=True)
        shutil.copyfile(checkout / name, path)


def project(checkout, corpus, output, driver, build_lock):
    driver.mkdir()
    for file in ("Program.cs", "ProjectCorpus.csproj"):
        shutil.copyfile(HERE / file, driver / file)
    run_locked(["dotnet", "build", str(driver / "ProjectCorpus.csproj"), "-c", "Release",
                "-f", "net10.0", "-m:2", f"-p:ExcelSourceRoot={checkout}"], checkout, build_lock)
    executable = driver / "bin/Release/net10.0/ProjectCorpus.dll"
    run_locked(["dotnet", str(executable), str(corpus), str(output)], checkout, build_lock)


def normalize_qualified_values(xml):
    # ElementTree expands element and attribute names. Preserve the meaning of
    # schema attributes whose values also name prefixes, including MC's prefix lists.
    scopes = [{"xml": "http://www.w3.org/XML/1998/namespace"}]
    declared = {}
    for event, node in xml:
        if event == "start-ns":
            prefix, uri = node
            declared[prefix] = uri
        elif event == "start":
            scope = {**scopes[-1], **declared}
            scopes.append(scope)
            declared.clear()
            attributes = [(MC + "Ignorable", True), (MC + "MustUnderstand", True),
                          (MC + "ProcessContent", False), (MC + "PreserveElements", False),
                          (MC + "PreserveAttributes", False), (XSI_TYPE, False)]
            if node.tag == MC + "Choice":
                attributes.append(("Requires", True))
            for name, prefix_list in attributes:
                if name not in node.attrib:
                    continue
                values = []
                for token in node.get(name).split():
                    if prefix_list:
                        values.append("{" + scope.get(token, "unbound:" + token) + "}")
                    else:
                        prefix, has_prefix, local = token.partition(":")
                        if not has_prefix:
                            prefix, local = "", prefix
                        values.append("{" + scope.get(prefix, "unbound:" + prefix if prefix else "") + "}" + local)
                node.set(name, " ".join(values))
        else:
            scopes.pop()
    return xml.root


def normalized_xml(data, part):
    # ZIP timestamps and compression are ignored by comparing entries. Excel's
    # conversion save can refresh the modified timestamp to the current clock.
    root = normalize_qualified_values(ET.iterparse(BytesIO(data), events=("start-ns", "start", "end")))
    if part == "docProps/core.xml":
        modified = root.find(CORE_MODIFIED)
        if modified is not None:
            modified.text = ""
    if part.endswith(".rels") or part == "[Content_Types].xml":
        root[:] = sorted(root, key=lambda child: (child.tag, sorted(child.attrib.items())))
    return ET.canonicalize(ET.tostring(root, encoding="unicode"), rewrite_prefixes=True)


def compare_packages(left, right, allow_markup_equivalence):
    if left.read_bytes() == right.read_bytes():
        return "identical", []
    if not allow_markup_equivalence:
        return "different", ["XLSB bytes differ"]
    with zipfile.ZipFile(left) as baseline, zipfile.ZipFile(right) as candidate:
        left_names = sorted(baseline.namelist())
        right_names = sorted(candidate.namelist())
        if left_names != right_names:
            return "different", [
                f"Baseline only: {sorted(set(left_names) - set(right_names))}",
                f"Candidate only: {sorted(set(right_names) - set(left_names))}",
            ]
        changed = []
        for part in left_names:
            before = baseline.read(part)
            after = candidate.read(part)
            if before == after:
                continue
            if part.endswith((".xml", ".rels")):
                try:
                    if normalized_xml(before, part) == normalized_xml(after, part):
                        continue
                except ET.ParseError:
                    pass
            changed.append(part)
        return ("different", changed) if changed else ("equivalent", [])


def compare(output):
    baseline_dir, candidate_dir = output / "baseline", output / "candidate"
    baseline = json.loads((baseline_dir / "results.json").read_text())
    candidate = json.loads((candidate_dir / "results.json").read_text())
    if baseline.keys() != candidate.keys():
        raise ValueError("The revisions did not process the same inputs")
    results = {}
    counts = {}
    for name, before in baseline.items():
        results[name] = {}
        for key, baseline_error in before.items():
            candidate_error = candidate[name][key]
            if baseline_error or candidate_error:
                result = ("refused", []) if baseline_error == candidate_error else (
                    "different", [baseline_error, candidate_error])
            else:
                result = compare_packages(baseline_dir / f"{name}.{key.lower()}",
                                          candidate_dir / f"{name}.{key.lower()}",
                                          allow_markup_equivalence=key == "Xlsx")
            status, details = result
            counts[status] = counts.get(status, 0) + 1
            results[name][key] = {"status": status, "differences": details, "baselineError": baseline_error}
    (output / "comparison.json").write_text(json.dumps({"counts": counts, "results": results}, indent=2) + "\n")
    print(json.dumps({"inputs": len(results), "outputs": counts, "report": str(output / "comparison.json")}))
    return 1 if counts.get("different") else 0


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("baseline", help="Baseline Git revision, usually origin/master")
    parser.add_argument("candidate", help="Candidate Git revision")
    parser.add_argument("output", type=Path, help="New scratch directory for inputs, outputs and the difference report")
    parser.add_argument("--corpus", type=Path, help="Reuse fixtures and test outputs from an earlier run")
    parser.add_argument("--build-lock", type=lambda path: Path(path).expanduser().resolve(),
                        help="Serialize builds and tests through the machine's shared lockf path")
    args = parser.parse_args()
    root = HERE.parents[2]
    output = args.output.resolve()
    output.mkdir(parents=True)
    corpus = args.corpus.resolve() if args.corpus else output / "corpus"
    checkouts = []
    try:
        for label, ref in (("baseline", args.baseline), ("candidate", args.candidate)):
            checkout = output / f"source-{label}"
            run(["git", "worktree", "add", "--detach", "--quiet", str(checkout), ref], root)
            checkouts.append(checkout)
        if args.corpus:
            if any(not any((corpus / kind).rglob("*.xlsb")) for kind in ("fixtures", "test-outputs")):
                raise ValueError("The reused corpus must contain both fixtures and test-produced XLSB inputs")
        else:
            copy_fixtures(checkouts[0], corpus)
            capture_test_inputs(checkouts[0], corpus / "test-outputs", args.build_lock)
        for label, checkout in zip(("baseline", "candidate"), checkouts):
            project(checkout, corpus, output / label, output / f"driver-{label}", args.build_lock)
        return compare(output)
    finally:
        for checkout in reversed(checkouts):
            run(["git", "worktree", "remove", "--force", str(checkout)], root)


if __name__ == "__main__":
    raise SystemExit(main())
