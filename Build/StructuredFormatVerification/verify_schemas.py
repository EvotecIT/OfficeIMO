"""Opt-in validation of native output with pinned official schemas; no runtime dependency."""
import argparse
import hashlib
import json
from pathlib import Path
import subprocess
import urllib.request

import jsonschema
from lxml import etree

SCHEMAS = {
    "adf": ("https://go.atlassian.com/adf-json-schema", "5128562b75278c8a83e7e3619a570205bc80d59696985ec31a7a7883cff66fbe"),
    "docbook": ("https://cdn.docbook.org/schema/5.2/rng/docbook.rng", "e44da3bfa9db2a4b58b21c33f2ec441913f40b99dc04430ab30ddc745c729421"),
}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--output", required=True, type=Path, help="Task-owned output directory outside the checkout")
    args = parser.parse_args()
    output = args.output.resolve()
    output.mkdir(parents=True, exist_ok=True)
    validators = {}
    for name, (url, digest) in SCHEMAS.items():
        with urllib.request.urlopen(url, timeout=30) as response:
            data = response.read(4 * 1024 * 1024 + 1)
        if hashlib.sha256(data).hexdigest() != digest:
            raise RuntimeError(f"Official {name} schema hash changed; review it before updating the pin")
        (output / (name + ".schema")).write_bytes(data)
        validators[name] = jsonschema.Draft4Validator(json.loads(data)) if name == "adf" else etree.RelaxNG(etree.fromstring(data))

    project = Path(__file__).with_name("StructuredFormatVerification.csproj")
    subprocess.run(["dotnet", "run", "--configuration", "Release", "--project", str(project), "--", str(output)], cwd=project.parents[2], check=True)
    results = []
    for case in json.loads((output / "manifest.json").read_text()):
        path = output / case["file"]
        if case["format"] == "adf":
            errors = list(validators["adf"].iter_errors(json.loads(path.read_text())))
            valid = not errors
            detail = [error.message for error in errors[:3]]
        else:
            validator = validators["docbook"]
            valid = validator.validate(etree.parse(str(path), etree.XMLParser(resolve_entities=False, no_network=True)))
            detail = [str(error) for error in list(validator.error_log)[:3]]
        results.append({**case, "officialValid": valid, "errors": detail})
    (output / "results.json").write_text(json.dumps({"schemaHashes": {name: value[1] for name, value in SCHEMAS.items()}, "cases": results}, indent=2))
    failures = [case["file"] for case in results if case["officialValid"] != case["expectedValid"] or case["boundedValid"] != case["expectedValid"]]
    if failures:
        raise RuntimeError("Conformance mismatch: " + ", ".join(failures))
    print(f"PASS | {len(results)} cases match pinned official schemas and bounded validation")


if __name__ == "__main__":
    main()
