#!/usr/bin/env python3
"""Exercise the maintained corpus workflow's provenance and deadline contracts."""

from __future__ import annotations

import hashlib
import json
import math
import os
from pathlib import Path
import re
import subprocess
import sys
import tempfile
import textwrap
import zipfile


ROOT = Path(__file__).resolve().parents[1]
WORKFLOW_PATH = ROOT / ".github" / "workflows" / "real-world-corpus-evidence.yml"
WORKFLOW = WORKFLOW_PATH.read_text(encoding="utf-8")
GUARD_WORKFLOW_PATH = ROOT / ".github" / "workflows" / "real-world-corpus-contract-guard.yml"
GUARD_WORKFLOW = GUARD_WORKFLOW_PATH.read_text(encoding="utf-8")
PINNED_SHA256 = "7fb4673ac1905a3ed16a9c5780a70c729c3abc575b06c155b7042b91fa5946aa"
SOURCE_MANIFEST_PATH = ROOT / "Build" / "RealWorldCorpus" / "openxml-corpus-sources.json"


def workflow_script() -> str:
    match = re.search(
        r"(?ms)^      - name: Download and validate the pinned corpus chunk\r?$.*?"
        r"^        run: \|\r?\n(?P<script>.*?)"
        r"(?=^      - name: Measure the bounded content-detected sample\r?$)",
        WORKFLOW,
    )
    if match is None:
        raise AssertionError("The maintained download step could not be located.")
    script = textwrap.dedent(match.group("script"))
    if script.index('archive_sha256="$(sha256sum') > script.index('if ! python3'):
        raise AssertionError("The archive is extracted before its checksum is verified.")
    return script


def workflow_integer(name: str) -> int:
    match = re.search(rf"^      {re.escape(name)}: '([0-9]+)'\r?$", WORKFLOW, re.MULTILINE)
    if match is None:
        raise AssertionError(f"Workflow integer {name} is missing or not fixed.")
    return int(match.group(1))


def corpus_job_timeout_minutes() -> int:
    match = re.search(
        r"(?ms)^  corpus-evidence:\r?$.*?^    timeout-minutes: ([0-9]+)\r?$",
        WORKFLOW,
    )
    if match is None:
        raise AssertionError("The corpus evidence job timeout is missing.")
    return int(match.group(1))


def job_block(name: str) -> str:
    match = re.search(rf"(?ms)^  {re.escape(name)}:\r?$(?P<body>.*?)(?=^  [A-Za-z0-9_-]+:\r?$|\Z)", WORKFLOW)
    if match is None:
        raise AssertionError(f"The workflow job {name} is missing.")
    return match.group("body")


def job_integer(block: str, name: str) -> int:
    match = re.search(rf"^      {re.escape(name)}: '([0-9]+)'\r?$", block, re.MULTILINE)
    if match is None:
        raise AssertionError(f"Job integer {name} is missing or not fixed.")
    return int(match.group(1))


def assert_openxml_differential_contract() -> None:
    block = job_block("openxml-differential")
    timeout = re.search(r"^    timeout-minutes: ([0-9]+)\r?$", block, re.MULTILINE)
    if timeout is None:
        raise AssertionError("The Open XML differential job has no timeout.")
    fetch_seconds = job_integer(block, "FETCH_TIMEOUT_SECONDS") + 30
    deadline_seconds = job_integer(block, "RUN_DEADLINE_SECONDS")
    file_timeout_seconds = job_integer(block, "FILE_TIMEOUT_SECONDS")
    job_integer(block, "MAX_TRAVERSAL_ENTRIES")
    job_integer(block, "MAX_TOTAL")
    if not 1 <= job_integer(block, "CORPUS_PARALLELISM") <= 8:
        raise AssertionError("The Open XML differential job must bound its parallelism.")
    for binding in (
        '--deadline-seconds "$RUN_DEADLINE_SECONDS"',
        '--timeout-seconds "$FILE_TIMEOUT_SECONDS"',
        '--parallelism "$CORPUS_PARALLELISM"',
        '--max-total "$MAX_TOTAL"',
        '--max-traversal-entries "$MAX_TRAVERSAL_ENTRIES"',
        'timeout --signal=TERM --kill-after=30s "$FETCH_TIMEOUT_SECONDS"',
        '--manifest "$SOURCE_MANIFEST"',
    ):
        if binding not in block:
            raise AssertionError(f"The Open XML differential job does not bind {binding}.")
    # The deadline stops new work; running workers end within one per-file timeout. Thirty minutes covers checkout,
    # SDK setup, the build, reporting and upload.
    reserve_seconds = int(timeout.group(1)) * 60 - fetch_seconds - deadline_seconds - file_timeout_seconds
    if reserve_seconds < 30 * 60:
        raise AssertionError(f"The Open XML differential job leaves only {reserve_seconds} seconds of reserve.")

    manifest = json.loads(SOURCE_MANIFEST_PATH.read_text(encoding="utf-8"))
    if not manifest.get("sources"):
        raise AssertionError("The Open XML source manifest lists no sources.")
    for source in manifest["sources"]:
        if not re.fullmatch(r"[0-9a-f]{40}", source.get("commit", "")):
            raise AssertionError(f"Source {source.get('id')} does not pin a full commit.")
        if not re.fullmatch(r"[0-9a-f]{64}", source.get("contentSha256", "")) or source["contentSha256"] == "0" * 64:
            raise AssertionError(f"Source {source.get('id')} does not pin its content SHA-256.")
        if source.get("fileCount", 0) <= 0 or not source.get("license") or not source.get("repository", "").startswith("https://"):
            raise AssertionError(f"Source {source.get('id')} lacks its file count, license or https repository.")


def assert_unrestricted_pull_request_trigger(workflow: str, name: str) -> None:
    lines = workflow.splitlines()
    try:
        trigger_index = lines.index("  pull_request:")
    except ValueError as error:
        raise AssertionError(f"{name} does not have a plain pull-request trigger.") from error
    for line in lines[trigger_index + 1:]:
        if not line.strip():
            continue
        if line.startswith("    "):
            raise AssertionError(f"{name} restricts its pull-request trigger.")
        break


def assert_contract_triggers_are_cross_guarded() -> None:
    invocation = "run: python3 ./Build/Test-RealWorldCorpusWorkflowContract.py"
    assert_unrestricted_pull_request_trigger(WORKFLOW, "Corpus evidence workflow")
    assert_unrestricted_pull_request_trigger(GUARD_WORKFLOW, "Corpus contract guard workflow")
    if invocation not in WORKFLOW or invocation not in GUARD_WORKFLOW:
        raise AssertionError("Both unrestricted workflows must execute the corpus workflow contract.")


def bash_path(path: Path) -> str:
    if os.name != "nt":
        return str(path)
    resolved = path.resolve()
    drive = resolved.drive.rstrip(":").lower()
    relative = resolved.relative_to(resolved.anchor).as_posix()
    return f"/mnt/{drive}/{relative}"


def run_case(
    script: str,
    root: Path,
    name: str,
    archive: Path,
    chunk: str,
    requested_sha256: str,
    timeout_seconds: int = 30,
) -> tuple[int, dict[str, str] | None, dict[str, str]]:
    case_root = root / name
    case_root.mkdir()
    fixture = bash_path(archive)
    runner_temp = bash_path(case_root)
    test_script = script.replace(
        'archive="${RUNNER_TEMP}/govdocs1-${CORPUS_CHUNK}.zip"',
        'archive="$FIXTURE_ARCHIVE"',
    )
    test_script = re.sub(
        r"if ! timeout --signal=TERM --kill-after=30s .*? curl .*?; then",
        'if ! test -f "$archive"; then',
        test_script,
    )
    test_script = test_script.replace("exit 0", "return 0").replace("exit 2", "return 2")
    prefix = f"""\
export CORPUS_CHUNK='{chunk}'
export REQUESTED_ARCHIVE_SHA256='{requested_sha256}'
export FILE_TIMEOUT_SECONDS='{timeout_seconds}'
export MAX_FILE_TIMEOUT_SECONDS='{workflow_integer("MAX_FILE_TIMEOUT_SECONDS")}'
export DOWNLOAD_TIMEOUT_SECONDS='1800'
export FIXTURE_ARCHIVE='{fixture}'
export RUNNER_TEMP='{runner_temp}'
export GITHUB_OUTPUT="$RUNNER_TEMP/github-output"
run_contract() {{
"""
    suffix = "\n}\nrun_contract\n"
    executable = ["wsl.exe", "bash", "-s"] if os.name == "nt" else ["bash", "-s"]
    completed = subprocess.run(executable, input=(prefix + test_script + suffix).encode(), capture_output=True)
    status_path = case_root / "real-world-corpus-reports" / "download-status.json"
    status = json.loads(status_path.read_text(encoding="utf-8")) if status_path.exists() else None
    output_path = case_root / "github-output"
    outputs: dict[str, str] = {}
    if output_path.exists():
        for line in output_path.read_text(encoding="utf-8").splitlines():
            key, value = line.split("=", 1)
            outputs[key] = value
    return completed.returncode, status, outputs


def main() -> None:
    assert_contract_triggers_are_cross_guarded()
    assert_openxml_differential_contract()
    script = workflow_script()
    if 'timeout --signal=TERM --kill-after=30s "$DOWNLOAD_TIMEOUT_SECONDS" curl' not in script:
        raise AssertionError("The complete download is not enclosed by its wall-clock deadline.")

    traversal = workflow_integer("MAX_TRAVERSAL_ENTRIES")
    selected = workflow_integer("MAX_TOTAL")
    parallelism = workflow_integer("CORPUS_PARALLELISM")
    download_seconds = workflow_integer("DOWNLOAD_TIMEOUT_SECONDS") + 30
    file_timeout_seconds = workflow_integer("MAX_FILE_TIMEOUT_SECONDS")
    required_runner_bindings = (
        '--max-total "$MAX_TOTAL"',
        '--max-traversal-entries "$MAX_TRAVERSAL_ENTRIES"',
        '--timeout-seconds "$FILE_TIMEOUT_SECONDS"',
        '--parallelism "$CORPUS_PARALLELISM"',
    )
    if any(binding not in WORKFLOW for binding in required_runner_bindings):
        raise AssertionError("A proven workflow limit is not bound to its runner argument.")
    if "FILE_TIMEOUT_SECONDS > MAX_FILE_TIMEOUT_SECONDS" not in script:
        raise AssertionError("The manual per-file timeout is not capped by the proven maximum.")
    worker_seconds = (
        math.ceil(traversal / parallelism) + math.ceil(selected / parallelism)
    ) * file_timeout_seconds
    job_seconds = corpus_job_timeout_minutes() * 60
    reserve_seconds = job_seconds - download_seconds - worker_seconds
    if reserve_seconds < 2 * 60 * 60:
        raise AssertionError(f"Workflow deadline reserve is only {reserve_seconds} seconds.")
    if traversal < 982:
        raise AssertionError("The pinned 000 baseline no longer fits the traversal budget.")

    with tempfile.TemporaryDirectory(prefix="officeimo-corpus-workflow-contract-") as temporary:
        root = Path(temporary)
        valid_archive = root / "valid.zip"
        with zipfile.ZipFile(valid_archive, "w") as package:
            package.writestr("document.html", "<!doctype html><p>contract</p>")
        valid_sha256 = hashlib.sha256(valid_archive.read_bytes()).hexdigest()
        invalid_archive = root / "invalid.zip"
        invalid_archive.write_bytes(b"not a zip archive")
        invalid_sha256 = hashlib.sha256(invalid_archive.read_bytes()).hexdigest()

        code, status, outputs = run_case(script, root, "alternate-missing", valid_archive, "001", "")
        if code != 0 or status is None or status.get("reason") != "expected-archive-sha256-required" or outputs.get("available") != "false":
            raise AssertionError("An alternate chunk without an expected checksum was not recorded as not measured.")

        wrong_sha256 = "0" * 64
        code, status, outputs = run_case(script, root, "checksum-mismatch", valid_archive, "001", wrong_sha256)
        if (
            code != 0
            or status is None
            or status.get("reason") != "archive-sha256-mismatch"
            or status.get("expectedArchiveSha256") != wrong_sha256
            or status.get("observedArchiveSha256") != valid_sha256
            or outputs.get("available") != "false"
        ):
            raise AssertionError("A checksum mismatch did not retain expected and observed provenance.")

        code, status, outputs = run_case(script, root, "matching-checksum", valid_archive, "001", valid_sha256.upper())
        if code != 0 or status is not None or outputs.get("available") != "true" or outputs.get("archive_sha256") != valid_sha256:
            raise AssertionError("A verified alternate archive did not reach extraction and measurement.")

        code, status, outputs = run_case(script, root, "invalid-archive", invalid_archive, "001", invalid_sha256)
        if (
            code != 0
            or status is None
            or status.get("reason") != "archive-validation-failed"
            or status.get("archiveSha256") != invalid_sha256
            or outputs.get("available") != "false"
        ):
            raise AssertionError("Archive validation failure did not retain verified provenance.")

        code, _, _ = run_case(script, root, "timeout-too-large", valid_archive, "001", valid_sha256, 31)
        if code == 0:
            raise AssertionError("The maintained workflow accepted a timeout beyond its proven runtime budget.")

        code, status, outputs = run_case(script, root, "pinned-conflict", valid_archive, "000", valid_sha256)
        if (
            code != 0
            or status is None
            or status.get("reason") != "pinned-archive-sha256-conflict"
            or status.get("expectedArchiveSha256") != PINNED_SHA256
            or status.get("requestedArchiveSha256") != valid_sha256
            or outputs.get("available") != "false"
        ):
            raise AssertionError("The pinned baseline accepted a conflicting manual checksum.")

        code, status, outputs = run_case(script, root, "invalid-checksum", valid_archive, "001", "not-a-sha256")
        if code != 0 or status is None or status.get("reason") != "expected-archive-sha256-invalid" or outputs.get("available") != "false":
            raise AssertionError("An invalid expected checksum was not recorded as not measured.")

    print(
        "Real-world corpus workflow contract passed: checksum ordering and outcomes, "
        f"rejected-archive provenance, and {reserve_seconds // 60} minutes of runtime reserve."
    )


if __name__ == "__main__":
    main()
