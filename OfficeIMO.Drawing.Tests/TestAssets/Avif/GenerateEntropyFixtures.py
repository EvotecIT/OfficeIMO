#!/usr/bin/env python3
"""Opt-in AOM reference generator; Python stdlib and a C11 compiler only.

Run with --work-dir pointing to task-owned scratch. This downloads hash-pinned
reference sources and their notices, builds an isolated native oracle, and writes
entropy-reference.json there. Normal OfficeIMO builds never execute this script.
Use --output explicitly to replace the checked-in fixture after inspecting it.
"""

import argparse
import base64
import hashlib
import json
import pathlib
import shutil
import subprocess
import urllib.request


def sha256(data):
    return hashlib.sha256(data).hexdigest()


def write_json(path, data):
    path.write_text(json.dumps(data, indent=2) + "\n", encoding="utf-8")


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--work-dir", required=True, type=pathlib.Path)
    parser.add_argument("--output", type=pathlib.Path)
    parser.add_argument("--cc", default="clang", help="C11 compiler executable")
    args = parser.parse_args()
    here = pathlib.Path(__file__).resolve().parent
    repo = here.parents[2]
    manifest = json.loads((here / "entropy-oracle.json").read_text(encoding="utf-8"))
    work = args.work_dir.resolve()
    work.mkdir(parents=True, exist_ok=True)
    native = work / "native-aom"
    for relative, expected in manifest["sourceSha256"].items():
        path = native / relative
        if path.is_file():
            data = path.read_bytes()
        else:
            url = manifest["sourceUrl"] + relative + "?format=TEXT"
            with urllib.request.urlopen(url, timeout=60) as response:
                data = base64.b64decode(response.read(), validate=True)
        if sha256(data) != expected:
            raise ValueError("Reference source hash mismatch: " + relative)
        if not path.is_file():
            path.parent.mkdir(parents=True, exist_ok=True)
            path.write_bytes(data)
    config = native / "config" / "aom_config.h"
    config.parent.mkdir(parents=True, exist_ok=True)
    config.write_text("#define INLINE inline\n#define CONFIG_DEBUG 1\n"
                      "#define ARCH_X86_64 0\n#define ARCH_X86 0\n"
                      "#define HAVE_NEON 0\n#define CONFIG_AV1_HIGHBITDEPTH 1\n", encoding="ascii")
    compiler = shutil.which(args.cc)
    if not compiler:
        raise FileNotFoundError("C11 compiler not found: " + args.cc)
    harness = here / "GenerateEntropyFixtures.c"
    executable = work / "generate-entropy"
    command = [compiler, "-std=c11", "-Wall", "-Wextra", "-Werror", "-O2", "-I", str(native), str(harness)]
    command += [str(native / "aom_dsp" / name) for name in ("entenc.c", "entdec.c", "entcode.c")]
    subprocess.run(command + ["-o", str(executable)], check=True)
    vector_output = work / "native-vectors.json"
    subprocess.run([str(executable), str(vector_output)], check=True)
    vectors = json.loads(vector_output.read_text(encoding="utf-8"))
    prefixes = []
    for item in manifest["framePrefixes"]:
        fixture = repo / item["path"]
        if sha256(fixture.read_bytes()) != item["sha256"]:
            raise ValueError("Frozen AVIF hash mismatch: " + item["name"])
        result = work / (item["name"] + "-prefix.json")
        subprocess.run([str(executable), str(fixture), str(item["offset"]), str(item["length"]), str(result)], check=True)
        prefix = {key: item[key] for key in ("name", "alpha", "offset", "length")}
        prefix.update(json.loads(result.read_text(encoding="utf-8")))
        prefixes.append(prefix)
    vectors["framePrefixes"] = prefixes
    output = args.output.resolve() if args.output else work / "entropy-reference.json"
    write_json(output, vectors)
    receipt = {
        "tag": manifest["tag"], "license": manifest["license"],
        "compiler": subprocess.run([compiler, "--version"], check=True, capture_output=True, text=True).stdout.splitlines()[0],
        "sourceSha256": manifest["sourceSha256"],
        "harnessSha256": sha256(harness.read_bytes()),
        "generatorSha256": sha256(pathlib.Path(__file__).read_bytes()),
        "manifestSha256": sha256((here / "entropy-oracle.json").read_bytes()),
        "fixtureSha256": sha256(output.read_bytes()),
        "nativeSelfCheck": vectors["nativeSelfCheck"],
        "vectorCases": len(vectors["cases"]), "booleanCases": len(vectors["booleanCases"]),
        "framePrefixes": manifest["framePrefixes"],
    }
    write_json(work / "oracle-receipt.json", receipt)
    print("Generated " + str(output) + " (SHA256 " + receipt["fixtureSha256"] + ")")


if __name__ == "__main__":
    main()
