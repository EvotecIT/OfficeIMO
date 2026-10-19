#!/usr/bin/env python3
"""Independent, exact-artifact checks for the existing PDF compliance proof lane.

pyHanko proves CMS integrity, byte coverage and RFC 3161 signature timestamps
against ephemeral fixture roots. This is not qualified-signature or full PAdES
certification, and does not consult system trust, a live TSA or revocation feeds.
"""

import argparse
import hashlib
import importlib.metadata
import json
import re
import subprocess
import sys
import xml.etree.ElementTree as ET
from datetime import datetime, timezone
from pathlib import Path


PROFILES = {
    "pdfa2a": ("2a",),
    "pdfa2u": ("2u",),
    "pdfa3a": ("3a",),
    "pdfa3u": ("3u",),
    "pdfa4": ("4",),
    "pdfa4e": ("4e",),
    "pdfa4f": ("4f",),
    "pdfua2": ("ua2",),
    "pdfa2b-ua1": ("2b", "ua1"),
    "pdfa3b-ua1": ("3b", "ua1"),
    "invoice-minimum-pdfa3b": ("3b",),
    "invoice-basic-wl-pdfa3b": ("3b",),
    "invoice-basic-pdfa3b": ("3b",),
    "invoice-en16931-pdfa3a": ("3a",),
    "invoice-extended-pdfa3u": ("3u",),
    "invoice-pdfa3b-ua1": ("3b", "ua1"),
}
PROFILE_NAMES = {
    "2a": "PDF/A-2a validation profile",
    "2u": "PDF/A-2u validation profile",
    "3a": "PDF/A-3a validation profile",
    "3u": "PDF/A-3u validation profile",
    "4": "PDF/A-4 validation profile",
    "4e": "PDF/A-4e validation profile",
    "4f": "PDF/A-4f validation profile",
    "ua2": "PDF/UA-2 + Tagged PDF validation profile",
    "2b": "PDF/A-2b validation profile",
    "3b": "PDF/A-3b validation profile",
    "ua1": "PDF/UA-1 validation profile",
}


class ValidatorFailure(ValueError):
    """Retains raw validator streams when its process or semantic checks fail."""

    def __init__(self, message, diagnostic):
        super().__init__(message)
        self.diagnostic = f"{message}\n{diagnostic}"


def require(condition, message):
    if not condition:
        raise ValueError(message)


def fingerprint(path):
    return {"file": path.name, "sha256": hashlib.sha256(path.read_bytes()).hexdigest(),
            "sizeBytes": path.stat().st_size}


def execute(command):
    try:
        result = subprocess.run(command, capture_output=True, text=True, encoding="utf-8",
                                errors="replace", timeout=120, check=False)
    except subprocess.TimeoutExpired as exc:
        def decode(value):
            return value.decode("utf-8", errors="replace") if isinstance(value, bytes) else value or ""
        raise ValidatorFailure("Validator timed out after 120 seconds",
                               json.dumps({"command": command, "exitCode": None}) + "\nSTDOUT:\n" +
                               decode(exc.stdout) + "\nSTDERR:\n" + decode(exc.stderr)) from exc
    return result.returncode, result.stdout, result.stderr


def check_verapdf(output, flavor):
    root = ET.fromstring(output)
    reports = [e for e in root.iter() if e.tag.rsplit("}", 1)[-1] == "validationReport"]
    require(len(reports) == 1, "Expected exactly one veraPDF validation report")
    report = reports[0]
    require(report.get("profileName") == PROFILE_NAMES[flavor],
            f"Wrong veraPDF profile: {report.get('profileName')}")
    require(report.get("jobEndStatus") == "normal" and report.get("isCompliant") == "true",
            f"veraPDF rejected {flavor}")
    summaries = [e for e in root.iter() if e.tag.rsplit("}", 1)[-1] == "batchSummary"]
    require(len(summaries) == 1 and summaries[0].get("totalJobs") == "1", "Invalid veraPDF job count")
    for key in ("failedToParse", "encrypted", "outOfMemory", "veraExceptions"):
        require(summaries[0].get(key) == "0", f"veraPDF batch error: {key}")
    builds = [dict(element.attrib) for element in root.iter()
              if element.tag.rsplit("}", 1)[-1] == "releaseDetails"]
    return {"reportedProfile": report.get("profileName"), "validatorBuilds": builds}


def check_mustang(output, is_pdf):
    # Mustang may surround its XML result with CLI banners. Parse the result,
    # never infer success from the process exit code or an isolated 'valid'.
    matches = re.findall(r"<validation\b.*?</validation>", output, flags=re.DOTALL)
    require(len(matches) == 1, "Expected exactly one Mustang validation result")
    root = ET.fromstring(matches[0])
    require(root.find("summary") is not None and root.find("summary").get("status") == "valid",
            "Mustang overall validation failed")
    for section in (("pdf", "xml") if is_pdf else ("xml",)):
        summary = root.find(f"{section}/summary")
        require(summary is not None and summary.get("status") == "valid", f"Mustang {section} validation failed")
    validator = root.find("xml/info/validator")
    require(validator is not None and validator.get("version") == "2.26.0",
            "Mustang result did not report validator version 2.26.0")
    return {"pdfValidated": is_pdf, "xmlValidated": True,
            "validatorVersion": validator.get("version"), "ruleSet": "ZF_250"}


def check_signature(path, timestamped):
    from pyhanko.keys import load_certs_from_pemder
    from pyhanko.pdf_utils.reader import PdfFileReader
    from pyhanko.sign.validation import validate_pdf_signature
    from pyhanko.sign.validation.status import SignatureCoverageLevel
    from pyhanko_certvalidator import ValidationContext

    roots = list(load_certs_from_pemder([str(path.parent / "signer-root.pem"),
                                       str(path.parent / "tsa-root.pem")]))
    # Explicit synthetic roots only; no system trust or fetching of remote data.
    context = ValidationContext(trust_roots=roots, allow_fetching=False, revocation_mode="soft-fail")
    with path.open("rb") as stream:
        reader = PdfFileReader(stream, strict=True)
        require(len(reader.embedded_signatures) == 1, "Expected exactly one approval signature")
        signature = reader.embedded_signatures[0]
        require(str(signature.sig_object.get("/SubFilter")) == "/ETSI.CAdES.detached", "Not a PAdES CMS subfilter")
        require(str(signature.sig_object.get("/Type")) == "/Sig", "Expected an approval signature dictionary")
        attrs = {attr["type"].native for attr in signature.signer_info["signed_attrs"]}
        require({"content_type", "message_digest", "signing_certificate_v2"} <= attrs,
                "Missing baseline signed CMS attributes")
        require("signing_time" not in attrs, "PAdES must not include the CMS signing_time signed attribute")
        status = validate_pdf_signature(signature, signer_validation_context=context,
                                        ts_validation_context=context)
        require(status.intact and status.valid and status.trusted and status.bottom_line,
                "pyHanko signature validation failed: " + status.summary())
        require(status.coverage == SignatureCoverageLevel.ENTIRE_FILE, "Signature does not cover the entire file")
        timestamp = status.timestamp_validity
        if timestamped:
            require(timestamp is not None and timestamp.intact and timestamp.valid and timestamp.trusted,
                    "Missing or invalid signature timestamp/imprint")
        else:
            require(timestamp is None, "B-B fixture unexpectedly has a signature timestamp")
        return {"intact": status.intact, "valid": status.valid, "trusted": status.trusted,
                "coverage": status.coverage.name, "baselineAttributesPresent": True,
                "cmsSigningTimeAbsent": True,
                "signatureTimestampValid": timestamped, "summary": status.summary(),
                "trustScope": "Ephemeral fixture roots; no revocation or public trust qualification"}


def check_invoice_attachment(path):
    from pyhanko.pdf_utils.reader import PdfFileReader

    xml = path.with_suffix(".xml").read_bytes()
    with path.open("rb") as stream:
        reader = PdfFileReader(stream, strict=True)
        files = reader.root["/AF"]
        require(len(files) == 1, "Expected one associated invoice XML file")
        file_spec = files[0].get_object()
        require(str(file_spec["/F"]) == "factur-x.xml", "Wrong invoice attachment filename")
        require(file_spec["/EF"]["/F"].get_object().data == xml, "Embedded and standalone XML differ")
    return {"embeddedXmlMatches": True, "xml": fingerprint(path.with_suffix(".xml"))}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--proof", type=Path, required=True)
    parser.add_argument("--verapdf", required=True)
    parser.add_argument("--mustang", type=Path, required=True)
    args = parser.parse_args()
    directory = args.proof.resolve() / "standards"
    require(directory.is_dir(), f"Missing fixture directory {directory}")
    rows = []
    artifacts = {}

    def record(path, validator, profile, operation):
        diagnostic = directory / f"{validator}-{path.stem}-{profile}.txt"
        row = {"artifactFile": path.name, "validator": validator, "profile": profile,
               "validatedUtc": datetime.now(timezone.utc).isoformat(), "status": "Failed"}
        try:
            before = fingerprint(path)
            artifacts[path.name] = before
            row.update(artifactSha256=before["sha256"], artifactSizeBytes=before["sizeBytes"])
            details, text = operation()
            row["details"] = details
            row["status"] = "Passed"
            require(fingerprint(path) == before, "Artifact changed during validation")
        except Exception as exc:
            row["status"] = "Failed"
            text = exc.diagnostic if isinstance(exc, ValidatorFailure) else f"{type(exc).__name__}: {exc}\n"
            print(f"FAIL {path.name} {validator}/{profile}: {exc}", file=sys.stderr)
        diagnostic.write_text(text, encoding="utf-8")
        row["diagnostic"] = fingerprint(diagnostic)
        rows.append(row)

    def external(command, check):
        exit_code, output, error = execute(command)
        text = json.dumps({"command": command, "exitCode": exit_code}) + "\nSTDOUT:\n" + output + "\nSTDERR:\n" + error
        # Save failed validator output too; the record handler retains it in the error.
        try:
            require(exit_code == 0, f"Validator exited with code {exit_code}")
            details = check(output)
        except Exception as exc:
            raise ValidatorFailure(str(exc), text) from exc
        return details, text

    versions = {}
    for name, command in (("veraPDF", [args.verapdf, "--version"]), ("qpdf", ["qpdf", "--version"])):
        code, output, error = execute(command)
        require(code == 0 and (output + error).strip(), f"Cannot identify {name}")
        versions[name] = (output + error).strip()
    require(hashlib.sha256(args.mustang.read_bytes()).hexdigest() ==
            "42d7868cb68264874a7b8cab4c3587b03b23ccc7cd72373da917f66758bb9736", "Wrong Mustang 2.26.0 artifact")
    versions["Mustang"] = "2.26.0"
    # This hash-pinned CLI selects ZF_250 for all five supported CII profiles.
    versions["MustangCiiRuleSet"] = "ZF_250"
    versions["pyHanko"] = importlib.metadata.version("pyhanko")

    for name, profiles in PROFILES.items():
        pdf = directory / f"{name}.pdf"
        for flavor in profiles:
            record(pdf, "veraPDF", flavor, lambda p=pdf, f=flavor: external(
                [args.verapdf, "--format", "xml", "--flavour", f, str(p)], lambda out: check_verapdf(out, f)))
        if name.startswith("invoice-"):
            for artifact in (pdf, pdf.with_suffix(".xml")):
                record(artifact, "Mustang", artifact.suffix[1:], lambda p=artifact: external(
                    ["java", "-jar", str(args.mustang.resolve()), "--action", "validate", "--source", str(p)],
                    lambda out: check_mustang(out, p.suffix == ".pdf")))
            record(pdf, "pyHanko", "attachment", lambda p=pdf: (check_invoice_attachment(p), "Exact embedded XML matches standalone XML\n"))

    for name, timestamped in (("pades-bb", False), ("pades-bt", True)):
        pdf = directory / f"{name}.pdf"
        def signature_check(p=pdf, t=timestamped):
            details = check_signature(p, t)
            return details, json.dumps(details, indent=2) + "\n"
        record(pdf, "pyHanko", name, signature_check)

    for name in (*PROFILES, "pades-bb", "pades-bt"):
        pdf = directory / f"{name}.pdf"
        record(pdf, "qpdf", "syntax", lambda p=pdf: external(["qpdf", "--check", str(p)], lambda _: {}))

    for name in ("signer-root.pem", "tsa-root.pem"):
        artifacts[name] = fingerprint(directory / name)
    manifest = {"schemaVersion": 1, "toolVersions": versions,
                "artifacts": list(artifacts.values()), "validations": rows,
                "allPassed": all(row["status"] == "Passed" for row in rows)}
    (directory / "proof.json").write_text(json.dumps(manifest, indent=2) + "\n", encoding="utf-8")
    require(manifest["allPassed"], "Independent standards proof failed; see standards/proof.json and diagnostics")
    print(f"Independent standards proof: {len(rows)} checks passed")


if __name__ == "__main__":
    main()
