#!/usr/bin/env python3
"""Validate exact optimizer artifacts; --assert-only rechecks the retained proof bundle."""

import argparse
import importlib.util
import json
import hashlib
import sys
from datetime import datetime, timezone
from pathlib import Path

# Reuse the lane's semantic validators and hash/command handling, including its pins.
sys.dont_write_bytecode = True
_spec = importlib.util.spec_from_file_location("standards_proof", Path(__file__).with_name("Validate-PdfStandardsProof.py"))
standards = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(standards)
require = standards.require
fingerprint = standards.fingerprint

PROFILES = ("Custom", "Balanced", "MaximumCompression", "Web", "Archival")
SETTINGS = (*PROFILES, "CustomRewrite", "CustomXrefStream", "CustomObjectStreams")
SOURCES = {"pdfa2b": ("2b",), "invoice-pdfa3b-ua1": ("3b", "ua1"),
           "pdfa4": ("4",), "pdfua2": ("ua2",), "pdfa1b": ("1b",)}
A1_REFUSED = ("MaximumCompression", "CustomXrefStream", "CustomObjectStreams")
standards.PROFILE_NAMES["1b"] = "PDF/A-1b validation profile"


def matrix(directory):
    data = json.loads((directory / "measurements.json").read_text(encoding="utf-8"))
    require(data["schemaVersion"] == 1 and set(data["profiles"]) == set(PROFILES), "Wrong optimizer profile coverage")
    expected = {(source, setting) for source in SOURCES for setting in SETTINGS
                if not (source == "pdfa1b" and setting in A1_REFUSED)}
    rows = data["rows"]
    require(len(rows) == len(expected) and {(r["source"], r["setting"]) for r in rows} == expected,
            "Incomplete or duplicate optimizer source/setting matrix")
    for row in rows:
        source, setting = row["source"], row["setting"]
        require(row["input"] == f"{source}-before.pdf" and row["output"] == f"{source}-{setting}-after.pdf",
                "Unexpected optimizer artifact filename")
        options = row["options"]
        profile = setting if setting in PROFILES else "Custom"
        require(options["profile"] == profile, "Wrong requested optimizer profile")
        require(options["useObjectStreams"] == (setting in ("MaximumCompression", "CustomObjectStreams")),
                "Object stream setting mismatch")
        require(options["xrefFormat"] == ("XrefStream" if setting in ("MaximumCompression", "CustomXrefStream") else "ClassicTable"),
                "Requested xref setting mismatch")
        require(options["linearize"] == (setting == "Web"), "Linearization setting mismatch")
        require(options["keepOriginalWhenNotSmaller"] == (setting in ("Custom", "Balanced", "Archival")),
                "Original-retention setting mismatch")
        for name in ("compressUnfilteredStreams", "removeUnreferencedObjects", "deduplicateIdenticalStreams",
                     "deduplicateImages", "deduplicateFonts", "deduplicateResources"):
            require(options[name] is True, f"Missing enabled optimizer action {name}")
        before, after = directory / row["input"], directory / row["output"]
        require(before.stat().st_size == row["originalLengthBytes"] and after.stat().st_size == row["optimizedLengthBytes"],
                "Optimizer size measurements do not match artifacts")
        require(row["sizeDeltaBytes"] == after.stat().st_size - before.stat().st_size, "Wrong size delta")
        require(row["candidateLengthBytes"] > 0, "Missing candidate size")
        if row["returnedOriginal"]:
            require(options["keepOriginalWhenNotSmaller"] and before.read_bytes() == after.read_bytes(),
                    "Returned-original measurement does not match artifacts")
        else:
            require(row["candidateLengthBytes"] == after.stat().st_size, "Returned candidate length mismatch")
        require(row["candidateXrefFormat"] == ("XrefStream" if options["useObjectStreams"] else options["xrefFormat"]),
                "Wrong candidate xref format")
        require(row["candidateUsesObjectStreams"] == options["useObjectStreams"] and
                row["candidateLinearized"] == options["linearize"], "Wrong candidate layout")
    refusals = data["refusals"]
    expected_refusals = {("signed", s) for s in SETTINGS} | {("pdfa1-streams", s) for s in A1_REFUSED}
    require(len(refusals) == len(expected_refusals) and {(r["reason"], r["setting"]) for r in refusals} == expected_refusals,
            "Incomplete optimizer refusal matrix")
    for refusal in refusals:
        source = "signed" if refusal["reason"] == "signed" else "pdfa1b"
        name = f"{source}-before-{refusal['setting']}"
        require(refusal["input"] == f"{source}-before.pdf" and refusal["outputAbsent"] == f"{name}-must-not-exist.pdf"
                and refusal["existingOutput"] == f"{name}-sentinel.bin", "Unexpected refusal filenames")
        require(refusal["exceptionType"] in ("NotSupportedException", "PdfMutationBlockedException") and refusal["diagnostic"],
                "Refusal must identify a supported-operation rejection")
        setting_options = next(r["options"] for r in rows if r["source"] == "pdfa2b" and r["setting"] == refusal["setting"])
        require(refusal["options"] == setting_options, "Refusal used different options than the measured optimizer setting")
        require(refusal["sourceUnchanged"] and refusal["existingOutputUnchanged"], "Refusal changed existing bytes")
        require(not (directory / refusal["outputAbsent"]).exists(), "Refused optimizer wrote an output")
        require((directory / refusal["existingOutput"]).read_bytes() == b"Existing destination must survive optimizer refusal",
                "Refused optimizer overwrote an existing destination")
    return data


def preserved(before, after):
    """Inspect protected raw streams and Info using an independent PDF parser."""
    from pyhanko.pdf_utils.reader import PdfFileReader
    from pyhanko.pdf_utils import generic

    def inspect(path):
        protected, seen = set(), set()
        with path.open("rb") as stream:
            reader = PdfFileReader(stream, strict=True)

            def resolve(obj):
                return obj.get_object() if isinstance(obj, generic.IndirectObject) else obj

            def capture(kind, obj):
                obj = resolve(obj)
                require(isinstance(obj, generic.StreamObject), f"Invalid protected {kind} stream")
                protected.add((kind, hashlib.sha256(obj.encoded_data).hexdigest(),
                               hashlib.sha256(obj.data).hexdigest(), str(obj.get("/Filter", ""))))

            def walk(obj):
                if isinstance(obj, generic.IndirectObject):
                    key = (obj.idnum, obj.generation)
                    if key in seen:
                        return
                    seen.add(key)
                    obj = obj.get_object()
                if isinstance(obj, generic.DictionaryObject):
                    if str(obj.get("/Type", "")) == "/Metadata":
                        capture("Metadata", obj)
                    for key, value in obj.items():
                        if key in ("/Metadata", "/DestOutputProfile"):
                            capture(str(key), value)
                        if key == "/EF":
                            for embedded in resolve(value).values():
                                capture("EmbeddedFile", embedded)
                        walk(value)
                elif isinstance(obj, generic.ArrayObject):
                    for value in obj:
                        walk(value)

            walk(reader.root)
            try:
                info = reader.trailer["/Info"]
            except KeyError:
                info = None
            info = {str(k): str(resolve(v)) for k, v in resolve(info).items()} if info is not None else None
            return sorted(protected), info

    first, second = inspect(before), inspect(after)
    require(first == second, "Protected metadata, embedded-file, ICC stream bytes/filters or original Info changed")
    require(before.read_bytes().splitlines()[0] == after.read_bytes().splitlines()[0], "Declared PDF header version changed")
    return {"protectedStreamsUnchanged": True, "infoUnchanged": True, "headerUnchanged": True,
            "protectedStreams": first[0]}


def signed_source(path):
    from pyhanko.pdf_utils.reader import PdfFileReader
    from pyhanko.sign.validation import validate_pdf_signature
    from pyhanko.sign.validation.status import SignatureCoverageLevel
    from pyhanko_certvalidator import ValidationContext

    with path.open("rb") as stream:
        reader = PdfFileReader(stream, strict=True)
        require(len(reader.embedded_signatures) == 1, "Signed refusal source has no unique signature")
        signature = reader.embedded_signatures[0]
        context = ValidationContext(trust_roots=[signature.signer_cert], allow_fetching=False)
        status = validate_pdf_signature(signature, signer_validation_context=context, ts_validation_context=context)
        require(status.intact and status.valid and status.coverage == SignatureCoverageLevel.ENTIRE_FILE,
                "Refusal source is not an intact signed PDF")
        return {"intact": True, "valid": True, "coverage": status.coverage.name,
                "trustScope": "Ephemeral software fixture certificate; no public trust claim"}


def invoice_attachment(path, before):
    """Bind AF, the embedded-file name tree and both file streams to the captured XML."""
    from pyhanko.pdf_utils.reader import PdfFileReader
    from pyhanko.pdf_utils import generic

    def inspect(pdf):
        xml = pdf.with_suffix(".xml").read_bytes()
        with pdf.open("rb") as stream:
            reader = PdfFileReader(stream, strict=True)
            associated = reader.root["/AF"]
            require(len(associated) == 1, "Expected exactly one associated invoice file")
            reference = associated.raw_get(0)
            require(isinstance(reference, generic.IndirectObject), "Invoice filespec must be indirect")
            spec = reference.get_object()
            require(str(spec["/Type"]) == "/Filespec", "Wrong invoice filespec type")
            require(str(spec["/AFRelationship"]) == "/Alternative", "Wrong invoice AF relationship")
            require(str(spec["/F"]) == str(spec["/UF"]) == "factur-x.xml", "Wrong invoice F/UF filename")
            embedded = spec["/EF"]
            f, uf = embedded.raw_get("/F"), embedded.raw_get("/UF")
            require(isinstance(f, generic.IndirectObject) and isinstance(uf, generic.IndirectObject)
                    and (f.idnum, f.generation) == (uf.idnum, uf.generation), "F/UF do not identify the same embedded stream")
            payload = f.get_object()
            require(str(payload["/Type"]) == "/EmbeddedFile", "Wrong embedded invoice stream type")
            mime = str(payload["/Subtype"])
            require(mime in ("/text/xml", "/application/xml"), f"Wrong invoice MIME subtype: {mime}")
            require(payload.data == xml, "Embedded F/UF bytes differ from captured invoice XML")

            entries, visited = [], set()
            def name_tree(node):
                if isinstance(node, generic.IndirectObject):
                    identity = (node.idnum, node.generation)
                    require(identity not in visited, "Cyclic embedded-file name tree")
                    visited.add(identity)
                    node = node.get_object()
                names = node.get("/Names", [])
                require(len(names) % 2 == 0, "Invalid embedded-file name tree")
                entries.extend((str(names[i]), names.raw_get(i + 1)) for i in range(0, len(names), 2))
                for child in node.get("/Kids", []):
                    name_tree(child)
            name_tree(reader.root["/Names"].raw_get("/EmbeddedFiles"))
            require(len(entries) == 1 and entries[0][0] == "factur-x.xml", "Wrong embedded-file name-tree entry")
            named = entries[0][1]
            require(isinstance(named, generic.IndirectObject) and
                    (named.idnum, named.generation) == (reference.idnum, reference.generation),
                    "Name tree and AF do not identify the same filespec")

            def semantic(obj):
                if isinstance(obj, generic.IndirectObject):
                    obj = obj.get_object()
                if isinstance(obj, generic.DictionaryObject):
                    value = {str(k): semantic(v) for k, v in obj.items()}
                    if isinstance(obj, generic.StreamObject):
                        value["encodedSha256"] = hashlib.sha256(obj.encoded_data).hexdigest()
                        value["decodedSha256"] = hashlib.sha256(obj.data).hexdigest()
                    return value
                if isinstance(obj, generic.ArrayObject):
                    return [semantic(v) for v in obj]
                return str(obj)
            return {"filespec": semantic(spec), "mimeSubtype": mime,
                    "afAndNameTreeSameFilespec": True, "fAndUfSameStream": True}

    original, optimized = inspect(before), inspect(path)
    require(original == optimized, "Invoice filespec or embedded-file semantics changed during optimization")
    require(before.with_suffix(".xml").read_bytes() == path.with_suffix(".xml").read_bytes(),
            "Before/after invoice XML differ")
    return {**standards.check_invoice_attachment(path), **optimized, "filespecUnchanged": True}


def expected_checks(data):
    checks = set()
    for source, flavors in SOURCES.items():
        for name in (f"{source}-before.pdf", *(r["output"] for r in data["rows"] if r["source"] == source)):
            checks.add((name, "qpdf", "syntax"))
            for flavor in flavors:
                checks.add((name, "veraPDF", flavor))
            if source.startswith("invoice-"):
                checks.add((name, "pyHanko", "attachment"))
                checks.add((name, "Mustang", "pdf"))
                checks.add((name[:-4] + ".xml", "Mustang", "xml"))
    for row in data["rows"]:
        checks.add((row["output"], "pyHanko", "preservation"))
    checks.add(("signed-before.pdf", "pyHanko", "signature"))
    checks.add(("signed-before.pdf", "qpdf", "syntax"))
    for refusal in data["refusals"]:
        checks.add((refusal["input"], "refusal", refusal["setting"]))
    return checks


def verify_manifest(directory):
    proof = json.loads((directory / "proof.json").read_text(encoding="utf-8"))
    require(proof["schemaVersion"] == 1 and proof["allPassed"] is True, "Optimizer external proof failed")
    data = matrix(directory)
    expected_files = {"measurements.json", "signed-before.pdf"}
    expected_files.update(f"{s}-before.pdf" for s in SOURCES)
    expected_files.update(r["output"] for r in data["rows"])
    expected_files.update(r["existingOutput"] for r in data["refusals"])
    expected_files.update(p[:-4] + ".xml" for p in list(expected_files) if p.startswith("invoice-") and p.endswith(".pdf"))
    files = proof["artifacts"]
    require(len(files) == len(expected_files) and {e["file"] for e in files} == expected_files, "Incomplete optimizer artifact inventory")
    inventory = {e["file"]: e for e in files}
    def check_file(entry):
        require(Path(entry["file"]).name == entry["file"], "Unsafe optimizer proof filename")
        require(fingerprint(directory / entry["file"]) == entry, f"Stale optimizer artifact {entry['file']}")
    for entry in files:
        check_file(entry)
    validations = proof["validations"]
    checks = expected_checks(data)
    require(len(validations) == len(checks) and {(r["artifactFile"], r["validator"], r["profile"]) for r in validations} == checks,
            "Incomplete or duplicate optimizer validation matrix")
    for row in validations:
        require(row["status"] == "Passed" and row["validatedUtc"], "Unsuccessful optimizer validation")
        entry = inventory[row["artifactFile"]]
        require(row["artifactSha256"] == entry["sha256"] and row["artifactSizeBytes"] == entry["sizeBytes"],
                "Optimizer validator evidence is bound to different bytes")
        check_file(row["diagnostic"])
        if row["validator"] == "veraPDF":
            require(row["details"]["reportedProfile"] == standards.PROFILE_NAMES[row["profile"]]
                    and row["details"]["validatorBuilds"], "Missing exact veraPDF profile or release details")
        elif row["validator"] == "Mustang":
            require(row["details"]["validatorVersion"] == "2.26.0" and row["details"]["ruleSet"] == "ZF_250"
                    and row["details"]["xmlValidated"] is True
                    and row["details"]["pdfValidated"] == (row["profile"] == "pdf"), "Wrong Mustang validation scope")
        elif row["profile"] == "attachment":
            require(row["details"]["filespecUnchanged"] and row["details"]["afAndNameTreeSameFilespec"]
                    and row["details"]["fAndUfSameStream"] and row["details"]["embeddedXmlMatches"],
                    "Incomplete independent invoice filespec preservation proof")
            check_file(row["details"]["xml"])
    require(proof["toolVersions"]["veraPDF"].strip(), "Missing veraPDF version")
    require(proof["toolVersions"]["Mustang"] == "2.26.0" and proof["toolVersions"]["pyHanko"] == "0.37.0",
            "Wrong optimizer validator versions")
    require(proof["toolVersions"]["qpdf"], "Missing qpdf version")
    return proof


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--proof", type=Path, required=True, help="Compliance proof root; reads its optimizer subdirectory")
    parser.add_argument("--verapdf")
    parser.add_argument("--mustang", type=Path)
    parser.add_argument("--assert-only", action="store_true")
    args = parser.parse_args()
    directory = args.proof.resolve() / "optimizer"
    if args.assert_only:
        verify_manifest(directory)
        print("Optimizer proof inventory, matrix and exact-artifact bindings passed")
        return
    require(args.verapdf and args.mustang, "Both --verapdf and --mustang are required for validation")
    data = matrix(directory)
    versions = {"Mustang": "2.26.0", "pyHanko": standards.importlib.metadata.version("pyhanko")}
    for name, command in (("veraPDF", [args.verapdf, "--version"]), ("qpdf", ["qpdf", "--version"])):
        code, out, err = standards.execute(command)
        require(code == 0 and (out + err).strip(), f"Cannot identify {name}")
        versions[name] = (out + err).strip()
    require(hashlib.sha256(args.mustang.read_bytes()).hexdigest() ==
            "42d7868cb68264874a7b8cab4c3587b03b23ccc7cd72373da917f66758bb9736", "Wrong Mustang artifact")
    rows, artifacts = [], {}
    def external(command, check):
        code, out, err = standards.execute(command)
        diagnostic = json.dumps({"command": command, "exitCode": code}) + "\nSTDOUT:\n" + out + "\nSTDERR:\n" + err
        try:
            require(code == 0, f"Validator exited {code}")
            details = check(out)
        except Exception as exc:
            raise standards.ValidatorFailure(str(exc), diagnostic) from exc
        return details, diagnostic
    def record(name, validator, profile, operation):
        path = directory / name
        row = {"artifactFile": name, "validator": validator, "profile": profile,
               "validatedUtc": datetime.now(timezone.utc).isoformat(), "status": "Failed"}
        try:
            before = fingerprint(path)
            artifacts[name] = before
            row.update(artifactSha256=before["sha256"], artifactSizeBytes=before["sizeBytes"])
            details, text = operation()
            row["details"] = details
            require(fingerprint(path) == before, "Artifact changed during validation")
            row["status"] = "Passed"
        except Exception as exc:
            text = exc.diagnostic if isinstance(exc, standards.ValidatorFailure) else f"{type(exc).__name__}: {exc}\n"
            print(f"FAIL {name} {validator}/{profile}: {exc}")
        diagnostic = directory / f"{validator}-{path.stem}-{profile}.txt"
        diagnostic.write_text(text, encoding="utf-8")
        row["diagnostic"] = fingerprint(diagnostic)
        rows.append(row)
    for name, validator, profile in sorted(expected_checks(data)):
        path = directory / name
        if validator == "veraPDF":
            record(name, validator, profile, lambda p=path, f=profile: external(
                [args.verapdf, "--format", "xml", "--flavour", f, str(p)], lambda out: standards.check_verapdf(out, f)))
        elif validator == "qpdf":
            record(name, validator, profile, lambda p=path: external(["qpdf", "--check", str(p)], lambda _: {}))
        elif validator == "Mustang":
            record(name, validator, profile, lambda p=path: external(
                ["java", "-jar", str(args.mustang.resolve()), "--action", "validate", "--source", str(p)],
                lambda out: standards.check_mustang(out, p.suffix == ".pdf")))
        elif validator == "refusal":
            refusal = next(r for r in data["refusals"] if r["input"] == name and r["setting"] == profile)
            record(name, validator, profile, lambda r=refusal: (r, json.dumps(r, indent=2)))
        else:
            if profile == "attachment":
                operation = lambda p=path: invoice_attachment(p, directory / "invoice-pdfa3b-ua1-before.pdf")
            elif profile == "signature":
                operation = lambda p=path: signed_source(p)
            else:
                measurement = next(r for r in data["rows"] if r["output"] == name)
                operation = lambda p=path, r=measurement: preserved(directory / r["input"], p)
            def inspect(op=operation):
                details = op()
                return details, json.dumps(details, indent=2) + "\n"
            record(name, validator, profile, inspect)
    for path in (directory / "measurements.json", *(directory / r["existingOutput"] for r in data["refusals"])):
        artifacts[path.name] = fingerprint(path)
    manifest = {"schemaVersion": 1, "toolVersions": versions, "artifacts": list(artifacts.values()),
                "validations": rows, "allPassed": all(r["status"] == "Passed" for r in rows)}
    (directory / "proof.json").write_text(json.dumps(manifest, indent=2) + "\n", encoding="utf-8")
    verify_manifest(directory)
    print(f"Optimizer proof: {len(data['rows'])} measured rewrites, {len(data['refusals'])} refusals, {len(rows)} checks passed")


if __name__ == "__main__":
    main()
